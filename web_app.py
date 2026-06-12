"""
Literature Auto-Downloader — Flask web app
Local:  python web_app.py  →  http://127.0.0.1:5000
Online: deploy to Render / Railway (see README section in this file)

Environment variables (set on hosting platform):
  SCOPUS_API_KEY   — Elsevier/Scopus API key (required for search)
  SECRET_KEY       — Flask secret key (any random string; optional locally)
  PORT             — Port to bind (set automatically by Render)
"""

import os, sys, json, re, time, csv, threading, webbrowser, zipfile, tempfile
import shutil, sqlite3, struct, base64, ctypes, ctypes.wintypes
from pathlib import Path
from queue import Queue, Empty
from uuid import uuid4

import openpyxl
import requests
from flask import (Flask, jsonify, request, Response,
                   stream_with_context, render_template_string, send_file)

# ── CONSTANTS ─────────────────────────────────────────────────────────────────
BASE_DIR         = Path(__file__).parent
ABS_XLSX         = BASE_DIR / "ABS.xlsx"
CONFIG_FILE      = BASE_DIR / ".lit_web_config.json"
QUB_PROXY        = "qub.idm.oclc.org"
UNPAYWALL_EMAIL  = "zhangmin1006@gmail.com"
IS_ONLINE        = bool(os.environ.get("RENDER") or os.environ.get("RAILWAY_ENVIRONMENT")
                        or os.environ.get("FLY_APP_NAME") or os.environ.get("ONLINE_MODE"))

FIELD_LABELS = {
    "ACCOUNT":  "Accounting",
    "BUS HIST & ECON HIST": "Business History & Economic History",
    "ECON":     "Economics",
    "ENT-SBM":  "Entrepreneurship & Small Business Mgmt",
    "ETHICS-CSR-MAN": "Ethics, CSR & Management",
    "FINANCE":  "Finance",
    "HRM&EMP":  "HRM & Employment Studies",
    "IB&AREA":  "International Business & Area Studies",
    "INFO MAN": "Information Management",
    "INNOV":    "Innovation",
    "MDEV&EDU": "Management Development & Education",
    "MKT":      "Marketing",
    "OPS&TECH": "Operations & Technology Management",
    "OR&MANSCI":"Operations Research & Management Science",
    "ORG STUD": "Organisation Studies",
    "PSYCH (GENERAL)": "Psychology (General)",
    "PSYCH (WOP-OB)":  "Psychology (Work/Org/OB)",
    "PUB SEC":  "Public Sector Management",
    "REGIONAL STUDIES, PLANNING AND ENVIRONMENT": "Regional Studies, Planning & Environment",
    "SECTOR":   "Sector Studies",
    "SOC SCI":  "Social Sciences",
    "STRAT":    "Strategy",
}
RATING_ORDER = {"4*": 0, "4": 1, "3": 2, "2": 3, "1": 4}

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "dev-secret-change-me")

# ── GLOBAL DOWNLOAD STATE ─────────────────────────────────────────────────────
# One download job at a time (fine for a personal research tool)
_dl_queue:  Queue = Queue()
_dl_active: bool  = False
_dl_zip_path: str = ""        # path to completed ZIP (online mode)

# ── CDP AUTO-COOKIE STATE ──────────────────────────────────────────────────────
_cdp_port: int = 0             # remote-debug port used by auto-configure flow


# ── HELPERS ───────────────────────────────────────────────────────────────────
def load_config() -> dict:
    if CONFIG_FILE.exists():
        try:
            return json.loads(CONFIG_FILE.read_text("utf-8-sig"))
        except Exception:
            pass
    return {
        "api_key":           "",
        "qub_session_cookie": "",
        "dl_folder":         "" if IS_ONLINE else str(Path.home() / "Downloads"),
        "max_results":       200,
        "try_oa":            True,
        "try_proxy":         True,
    }

def save_config(cfg: dict) -> None:
    try:
        CONFIG_FILE.write_text(json.dumps(cfg, indent=2), encoding="utf-8", newline="\n")
    except Exception:
        pass  # read-only filesystem on some platforms — silently ignore

def get_api_key() -> str:
    """Env var takes priority over config file."""
    return (os.environ.get("SCOPUS_API_KEY", "").strip()
            or load_config().get("api_key", "").strip())

def load_abs_data() -> list:
    wb = openpyxl.load_workbook(ABS_XLSX, read_only=True, data_only=True)
    ws = wb.active
    data = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        field, title, rating = row[0], row[1], row[2]
        if field and title and rating is not None:
            data.append({
                "field":  str(field).strip(),
                "title":  str(title).strip(),
                "rating": str(rating).strip(),
            })
    wb.close()
    data.sort(key=lambda j: (
        FIELD_LABELS.get(j["field"], j["field"]),
        RATING_ORDER.get(j["rating"], 9),
        j["title"],
    ))
    return data

def build_scopus_query(journal_titles: list, keyword_groups: list, row_op: str = "AND") -> str:
    row_parts = []
    for group in keyword_groups:
        terms = []
        for i, item in enumerate(group):
            text = item.get("text", "").strip()
            if not text:
                continue
            kw = f'"{text}"' if " " in text else text
            if i > 0:
                terms.append(item.get("op", "AND"))
            terms.append(kw)
        if terms:
            row_parts.append("( " + " ".join(terms) + " )")
    q = ""
    if row_parts:
        q = "TITLE-ABS-KEY " + f"\n{row_op} ".join(row_parts)
    if journal_titles:
        src = "SRCTITLE (\n  " + " OR\n  ".join(f'"{t}"' for t in journal_titles) + "\n)"
        q = (q + "\nAND " + src) if q else src
    return q

def safe_filename(title: str, max_len: int = 110) -> str:
    name = re.sub(r'[\\/*?:"<>|]', "", title)
    name = re.sub(r'\s+', "_", name.strip())[:max_len]
    return name + ".pdf"

def get_unpaywall_urls(doi: str) -> list:
    """Return [(pdf_url, host_type), …] for the DOI, repositories first."""
    try:
        r = requests.get(
            f"https://api.unpaywall.org/v2/{doi}?email={UNPAYWALL_EMAIL}",
            timeout=10)
        if r.status_code != 200:
            return []
        data = r.json()
        if not data.get("is_oa"):
            return []
        locs = data.get("oa_locations", [])
        repos = [(l["url_for_pdf"], "repo") for l in locs
                 if l.get("host_type") == "repository" and l.get("url_for_pdf")]
        pubs  = [(l["url_for_pdf"], "pub") for l in locs
                 if l.get("host_type") == "publisher"  and l.get("url_for_pdf")]
        return repos + pubs
    except Exception:
        return []

def get_semantic_scholar_pdf(doi: str) -> str:
    """Return direct OA PDF url from Semantic Scholar, or ''."""
    try:
        r = requests.get(
            f"https://api.semanticscholar.org/graph/v1/paper/DOI:{doi}",
            params={"fields": "openAccessPdf"},
            timeout=10)
        if r.status_code != 200:
            return ""
        data = r.json()
        oa = data.get("openAccessPdf") or {}
        return oa.get("url", "")
    except Exception:
        return ""

def download_pdf_to(url: str, dest: str, session: requests.Session,
                    timeout: int = 30, extra_headers: dict = None) -> bool:
    """Fetch url → dest. Returns True only if dest contains a valid PDF."""
    tmp = dest + ".tmp"
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}
    if extra_headers:
        headers.update(extra_headers)
    try:
        r = session.get(url, timeout=timeout, allow_redirects=True,
                        headers=headers, stream=True)
        if r.status_code != 200:
            return False
        with open(tmp, "wb") as f:
            for chunk in r.iter_content(65536):
                f.write(chunk)
        with open(tmp, "rb") as f:
            magic = f.read(4)
        # close file BEFORE os.replace — Windows locks open files
        if magic == b"%PDF":
            os.replace(tmp, dest)
            return True
        os.remove(tmp)
    except Exception:
        try:
            if os.path.exists(tmp):
                os.remove(tmp)
        except Exception:
            pass
    return False

def try_elsevier_fulltext(doi: str, dest: str, api_key: str, session: requests.Session) -> bool:
    """Try Elsevier Full-Text API — returns PDF directly for entitled content."""
    if not doi or not api_key:
        return False
    url = f"https://api.elsevier.com/content/article/doi/{requests.utils.quote(doi, safe='')}"
    return download_pdf_to(url, dest, session, timeout=30,
                           extra_headers={"X-ELS-APIKey": api_key, "Accept": "application/pdf"})

# ── BROWSER COOKIE EXTRACTION ─────────────────────────────────────────────────

def _dpapi_decrypt(ciphertext: bytes) -> bytes:
    """Decrypt a DPAPI-protected blob using Windows CryptUnprotectData."""
    class _BLOB(ctypes.Structure):
        _fields_ = [("cbData", ctypes.wintypes.DWORD),
                    ("pbData", ctypes.POINTER(ctypes.c_char))]
    buf  = ctypes.create_string_buffer(ciphertext, len(ciphertext))
    inp  = _BLOB(len(ciphertext), buf)
    out  = _BLOB()
    ok   = ctypes.windll.crypt32.CryptUnprotectData(
               ctypes.byref(inp), None, None, None, None, 0, ctypes.byref(out))
    if not ok:
        return b""
    data = (ctypes.c_char * out.cbData)()
    ctypes.memmove(data, out.pbData, out.cbData)
    ctypes.windll.kernel32.LocalFree(out.pbData)
    return bytes(data)

def _chrome_aes_key(user_data_dir: Path) -> bytes:
    """Return the AES-256 key used by Chrome/Edge for cookie encryption."""
    try:
        local_state = user_data_dir / "Local State"
        state       = json.loads(local_state.read_text(encoding="utf-8"))
        enc_key     = base64.b64decode(state["os_crypt"]["encrypted_key"])
        return _dpapi_decrypt(enc_key[5:])   # strip "DPAPI" prefix
    except Exception:
        return b""

def _chrome_decrypt_cookie(value: bytes, key: bytes) -> str:
    """Decrypt a Chrome v80+ cookie value (AES-256-GCM)."""
    try:
        if value[:3] == b"v10" or value[:3] == b"v11":
            from Crypto.Cipher import AES as _AES
            iv      = value[3:15]
            payload = value[15:-16]
            tag     = value[-16:]
            cipher  = _AES.new(key, _AES.MODE_GCM, nonce=iv)
            return cipher.decrypt_and_verify(payload, tag).decode("utf-8")
    except Exception:
        pass
    # Older DPAPI-only format
    try:
        return _dpapi_decrypt(value).decode("utf-8")
    except Exception:
        return ""

def _copy_locked_file(src: Path, dst: str) -> bool:
    """Copy a file that may be locked by another process (Windows CreateFile with share flags)."""
    GENERIC_READ   = 0x80000000
    FILE_SHARE_ALL = 0x00000007   # READ | WRITE | DELETE
    OPEN_EXISTING  = 3
    FILE_ATTR_NORM = 0x80
    INVALID        = ctypes.c_void_p(-1).value

    k32    = ctypes.windll.kernel32
    handle = k32.CreateFileW(str(src), GENERIC_READ, FILE_SHARE_ALL,
                              None, OPEN_EXISTING, FILE_ATTR_NORM, None)
    if handle == INVALID:
        return False
    try:
        size_hi = ctypes.c_ulong(0)
        size_lo = k32.GetFileSize(handle, ctypes.byref(size_hi))
        total   = (size_hi.value << 32) | size_lo
        CHUNK   = 1 << 20   # 1 MB
        buf     = ctypes.create_string_buffer(CHUNK)
        nread   = ctypes.c_ulong(0)
        with open(dst, "wb") as f:
            remaining = total
            while remaining > 0:
                want = min(CHUNK, remaining)
                if not k32.ReadFile(handle, buf, want, ctypes.byref(nread), None):
                    break
                if nread.value == 0:
                    break
                f.write(buf.raw[: nread.value])
                remaining -= nread.value
        return True
    finally:
        k32.CloseHandle(handle)

def _extract_firefox_cookies(domain: str) -> dict:
    profiles_base = Path(os.environ.get("APPDATA", "")) / "Mozilla" / "Firefox" / "Profiles"
    if not profiles_base.exists():
        return {}
    for profile in sorted(profiles_base.iterdir(), key=lambda p: p.stat().st_mtime, reverse=True):
        db = profile / "cookies.sqlite"
        if not db.exists():
            continue
        tmp = tempfile.mktemp(suffix=".sqlite")
        try:
            if not _copy_locked_file(db, tmp):
                shutil.copy2(str(db), tmp)
            conn = sqlite3.connect(tmp)
            rows = conn.execute(
                "SELECT name, value FROM moz_cookies WHERE host LIKE ?",
                (f"%{domain}%",)).fetchall()
            conn.close()
            if rows:
                return {n: v for n, v in rows}
        except Exception:
            pass
        finally:
            try: os.remove(tmp)
            except Exception: pass
    return {}

def _chromium_open_cookie_db(cookie_db: Path):
    """Open a Chromium cookie SQLite database, even while the browser is running.

    Uses SQLite's immutable=1 URI flag which bypasses SQLite-level locking.
    Proper URI encoding (forward slashes, %20 for spaces, file:/// prefix) is
    required on Windows for the URI mode to resolve correctly.
    Chrome holds OS-level locks intermittently; retries get the idle window.
    """
    # Build a properly encoded file URI (Windows needs forward slashes + %20)
    fwd = str(cookie_db).replace("\\", "/").replace(" ", "%20")
    uri = f"file:///{fwd}?mode=ro&immutable=1"
    for _ in range(5):
        try:
            conn = sqlite3.connect(uri, uri=True, timeout=0)
            conn.execute("SELECT 1 FROM cookies LIMIT 1")
            return conn, None
        except Exception:
            time.sleep(0.05)
    # Fallback: copy then open (only works when browser is closed)
    tmp = tempfile.mktemp(suffix=".sqlite")
    try:
        shutil.copy2(str(cookie_db), tmp)
        tmp_fwd = tmp.replace("\\", "/").replace(" ", "%20")
        conn = sqlite3.connect(f"file:///{tmp_fwd}?mode=ro&immutable=1", uri=True)
        return conn, tmp
    except Exception:
        try: os.remove(tmp)
        except Exception: pass
        return None, None

def _extract_chromium_cookies(domain: str) -> dict:
    local = Path(os.environ.get("LOCALAPPDATA", ""))
    candidates = [
        local / "Google"       / "Chrome"         / "User Data",
        local / "Microsoft"    / "Edge"            / "User Data",
        local / "BraveSoftware"/ "Brave-Browser"   / "User Data",
    ]
    for user_data in candidates:
        if not user_data.exists():
            continue
        key       = _chrome_aes_key(user_data)
        cookie_db = user_data / "Default" / "Network" / "Cookies"
        if not cookie_db.exists():
            cookie_db = user_data / "Default" / "Cookies"
        if not cookie_db.exists():
            continue
        conn, tmp = _chromium_open_cookie_db(cookie_db)
        if conn is None:
            continue
        try:
            rows = conn.execute(
                "SELECT name, encrypted_value FROM cookies WHERE host_key LIKE ?",
                (f"%{domain}%",)).fetchall()
        except Exception:
            rows = []
        finally:
            try: conn.close()
            except Exception: pass
            if tmp:
                try: os.remove(tmp)
                except Exception: pass

        result = {}
        for name, enc_val in rows:
            # v20 = Chrome 127+ App-Bound Encryption: requires a signed Google
            # process to decrypt — cannot be done from a third-party app.
            if isinstance(enc_val, bytes) and enc_val[:3] == b"v20":
                continue
            val = _chrome_decrypt_cookie(enc_val, key) if key else ""
            if not val and isinstance(enc_val, bytes):
                try:
                    val = enc_val.decode("utf-8", errors="ignore")
                except Exception:
                    pass
            if val:
                result[name] = val
        if result:
            return result
    return {}

def _find_chrome_exe() -> str:
    """Return path to Chrome executable, or '' if not found."""
    candidates = [
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        os.path.join(os.environ.get("LOCALAPPDATA", ""),
                     "Google", "Chrome", "Application", "chrome.exe"),
    ]
    for p in candidates:
        if os.path.exists(p):
            return p
    try:
        import winreg
        with winreg.OpenKey(
                winreg.HKEY_LOCAL_MACHINE,
                r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe") as k:
            return winreg.QueryValueEx(k, "")[0]
    except Exception:
        return ""


def _ws_build_frame(data: bytes) -> bytes:
    """Build a masked WebSocket text frame (client→server)."""
    mask = os.urandom(4)
    n = len(data)
    if n < 126:
        hdr = bytes([0x81, 0x80 | n])
    elif n < 65536:
        hdr = struct.pack("!BBH", 0x81, 0xFE, n)
    else:
        hdr = struct.pack("!BBQ", 0x81, 0xFF, n)
    return hdr + mask + bytes(b ^ mask[i % 4] for i, b in enumerate(data))


def _ws_recv_message(sock) -> bytes:
    """Read one complete (possibly fragmented) WebSocket message from sock."""
    payload = b""
    while True:
        h = b""
        while len(h) < 2:
            h += sock.recv(2 - len(h))
        fin     = (h[0] & 0x80) != 0
        masked  = (h[1] & 0x80) != 0
        n       = h[1] & 0x7F
        if n == 126:
            ex = b""
            while len(ex) < 2: ex += sock.recv(2 - len(ex))
            n = struct.unpack("!H", ex)[0]
        elif n == 127:
            ex = b""
            while len(ex) < 8: ex += sock.recv(8 - len(ex))
            n = struct.unpack("!Q", ex)[0]
        mk = b""
        if masked:
            while len(mk) < 4: mk += sock.recv(4 - len(mk))
        chunk = b""
        while len(chunk) < n: chunk += sock.recv(n - len(chunk))
        payload += bytes(b ^ mk[i % 4] for i, b in enumerate(chunk)) if masked else chunk
        if fin:
            return payload


def cdp_get_qub_cookies(port: int) -> dict:
    """Connect to Chrome's DevTools Protocol and return QUB cookies as {name: value}.
    Chrome decrypts v20 cookies internally, so values are always plaintext here."""
    import http.client as _http
    import socket as _sock
    conn = _http.HTTPConnection("localhost", port, timeout=3)
    conn.request("GET", "/json")
    targets = json.loads(conn.getresponse().read())
    conn.close()
    page = next((t for t in targets if t.get("type") == "page"), None)
    if not page:
        return {}
    ws_path = page["webSocketDebuggerUrl"].split(f"localhost:{port}")[1]
    sock = _sock.create_connection(("localhost", port), timeout=10)
    sock.settimeout(10)
    ws_key = base64.b64encode(os.urandom(16)).decode()
    sock.sendall((
        f"GET {ws_path} HTTP/1.1\r\nHost: localhost:{port}\r\n"
        f"Upgrade: websocket\r\nConnection: Upgrade\r\n"
        f"Sec-WebSocket-Key: {ws_key}\r\nSec-WebSocket-Version: 13\r\n\r\n"
    ).encode())
    buf = b""
    while b"\r\n\r\n" not in buf:
        buf += sock.recv(4096)
    sock.sendall(_ws_build_frame(
        json.dumps({"id": 1, "method": "Network.getAllCookies"}).encode()
    ))
    while True:
        msg = json.loads(_ws_recv_message(sock))
        if msg.get("id") == 1:
            break
    sock.close()
    return {
        c["name"]: c["value"]
        for c in msg.get("result", {}).get("cookies", [])
        if QUB_PROXY in c.get("domain", "")
    }


def check_chromium_cookie_status(domain: str) -> dict:
    """Scan Chrome, Edge and Brave; return info on the browser with the most QUB cookies."""
    local = Path(os.environ.get("LOCALAPPDATA", ""))
    browsers = [
        ("Chrome", local / "Google"       / "Chrome"       / "User Data"),
        ("Edge",   local / "Microsoft"    / "Edge"          / "User Data"),
        ("Brave",  local / "BraveSoftware"/ "Brave-Browser" / "User Data"),
    ]
    best: dict = {}
    for bname, user_data in browsers:
        if not user_data.exists():
            continue
        cookie_db = user_data / "Default" / "Network" / "Cookies"
        if not cookie_db.exists():
            cookie_db = user_data / "Default" / "Cookies"
        if not cookie_db.exists():
            continue
        conn, tmp = _chromium_open_cookie_db(cookie_db)
        if conn is None:
            continue
        try:
            rows = conn.execute(
                "SELECT name, encrypted_value FROM cookies WHERE host_key LIKE ?",
                (f"%{domain}%",)).fetchall()
        except Exception:
            rows = []
        finally:
            try: conn.close()
            except Exception: pass
            if tmp:
                try: os.remove(tmp)
                except Exception: pass
        if not rows:
            continue
        v20  = sum(1 for _, v in rows if isinstance(v, bytes) and v[:3] == b"v20")
        info = {"browser": bname, "total": len(rows), "v20": v20,
                "decryptable": len(rows) - v20, "all_v20": v20 == len(rows)}
        # Prefer browser with more decryptable cookies, then more total
        if not best or info["decryptable"] > best.get("decryptable", -1) or (
                info["decryptable"] == best.get("decryptable") and
                info["total"] > best.get("total", 0)):
            best = info
    return best

def get_browser_cookies(domain: str) -> dict:
    """Try Firefox then Chrome/Edge; return cookie dict for domain."""
    cookies = _extract_firefox_cookies(domain)
    if not cookies:
        cookies = _extract_chromium_cookies(domain)
    return cookies

def get_qub_cookies() -> dict:
    """Return QUB proxy cookies: manual config value first, browser extraction as fallback."""
    cfg    = load_config()
    manual = cfg.get("qub_session_cookie", "").strip()
    if manual:
        # Support "name=value" format or bare value (assumed EZproxy cookie)
        if "=" in manual and not manual.startswith("="):
            name, _, val = manual.partition("=")
            return {name.strip(): val.strip()}
        return {"EZproxy": manual}
    return get_browser_cookies(QUB_PROXY)

# ── QUB PROXY DOWNLOAD (with HTML fallback) ───────────────────────────────────

_PDF_META_RE = re.compile(
    r'<meta[^>]+name=["\']citation_pdf_url["\'][^>]+content=["\']([^"\']+)["\']',
    re.IGNORECASE)
_PDF_HREF_RE = re.compile(
    r'href=["\']([^"\']*(?:/pdf[^"\']*|\.pdf(?:[?#][^"\']*)?|pdfft[^"\']*|/epdf/[^"\']*|download/pdf[^"\']*|full\.pdf[^"\']*|reader[^"\']*\.pdf[^"\']*)["\'])',
    re.IGNORECASE)

def _find_pdf_url_in_html(html: str, base_url: str) -> str:
    """Extract direct PDF URL from a publisher HTML page."""
    # 1. citation_pdf_url meta tag (most publishers support this)
    m = _PDF_META_RE.search(html)
    if m:
        return m.group(1).strip()
    # 2. href patterns for common publishers
    m = _PDF_HREF_RE.search(html)
    if m:
        href = m.group(1).strip()
        if href.startswith("http"):
            return href
        if href.startswith("//"):
            proto = base_url.split("://")[0]
            return proto + ":" + href
        if href.startswith("/"):
            from urllib.parse import urlparse
            p = urlparse(base_url)
            return f"{p.scheme}://{p.netloc}{href}"
    return ""

def try_proxy_download(doi: str, dest: str, session: requests.Session,
                       extra_cookies: dict = None) -> bool:
    """
    Download via QUB EZproxy:
      1. GET the proxy DOI URL (with any supplied auth cookies)
      2. If response is a PDF → save directly
      3. If response is HTML → find citation_pdf_url / PDF href → download that
    """
    if not doi:
        return False

    proxy_url = f"https://doi-org.{QUB_PROXY}/{doi}"
    headers   = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
                 "Accept": "text/html,application/pdf,*/*"}

    # Inject browser cookies into the session for this request
    saved_cookies = {}
    if extra_cookies:
        for k, v in extra_cookies.items():
            saved_cookies[k] = session.cookies.get(k)
            session.cookies.set(k, v, domain=QUB_PROXY)

    try:
        r = session.get(proxy_url, headers=headers, timeout=30,
                        allow_redirects=True, stream=True)
        ct = r.headers.get("content-type", "")

        # Direct PDF
        if "pdf" in ct.lower() and r.status_code == 200:
            tmp = dest + ".tmp"
            with open(tmp, "wb") as f:
                for chunk in r.iter_content(65536):
                    f.write(chunk)
            with open(tmp, "rb") as f:
                magic = f.read(4)
            if magic == b"%PDF":
                os.replace(tmp, dest)
                return True
            os.remove(tmp)
            return False

        # HTML page — parse for PDF link
        if "html" in ct.lower() and r.status_code == 200:
            html     = r.content.decode("utf-8", errors="replace")
            pdf_url  = _find_pdf_url_in_html(html, r.url)
            if pdf_url:
                return download_pdf_to(pdf_url, dest, session, timeout=30)

    except Exception:
        pass
    finally:
        # Restore session cookies
        for k, v in saved_cookies.items():
            if v is None:
                session.cookies.clear(domain=QUB_PROXY, name=k)
            else:
                session.cookies.set(k, v, domain=QUB_PROXY)

    return False

def scopus_search(query: str, api_key: str, max_results: int, push) -> list:
    push("log", f"Searching Scopus — up to {max_results} results…")
    results, start = [], 0
    headers = {"X-ELS-APIKey": api_key, "Accept": "application/json"}

    while start < max_results:
        params = {
            "query": query,
            "start": start,
            "count": min(25, max_results - start),
            "field": ("dc:title,prism:doi,dc:creator,prism:publicationName,"
                      "prism:coverDate,dc:identifier,prism:volume,prism:pageRange"),
        }
        try:
            r = requests.get("https://api.elsevier.com/content/search/scopus",
                             headers=headers, params=params, timeout=20)
            if r.status_code == 401:
                push("error", "Invalid Scopus API key.")
                return []
            if r.status_code == 429:
                push("error", "Scopus API rate limit — wait a moment and retry.")
                return results
            if r.status_code != 200:
                push("error", f"Scopus API error {r.status_code}: {r.text[:200]}")
                break
            sr = r.json().get("search-results", {})
            entries = sr.get("entry", [])
            if not entries:
                break
            for e in entries:
                doi = e.get("prism:doi", "")
                results.append({
                    "title":   e.get("dc:title", ""),
                    "doi":     doi,
                    "authors": e.get("dc:creator", ""),
                    "year":    e.get("prism:coverDate", "")[:4],
                    "source":  e.get("prism:publicationName", ""),
                    "volume":  e.get("prism:volume", ""),
                    "pages":   e.get("prism:pageRange", ""),
                    "url":     f"https://doi.org/{doi}" if doi else "",
                    "status":  "pending",
                })
            total = int(sr.get("opensearch:totalResults", 0))
            push("log", f"Retrieved {start + len(entries)} / {min(total, max_results)}")
            start += len(entries)
            if start >= total:
                break
            time.sleep(0.3)
        except Exception as ex:
            push("error", str(ex))
            break
    return results


# ── DOWNLOAD WORKER ───────────────────────────────────────────────────────────
def _download_worker(papers: list, folder: str, try_oa: bool, try_prx: bool) -> None:
    global _dl_active, _dl_zip_path
    api_key  = get_api_key()
    session  = requests.Session()
    total    = len(papers)
    ok = fail = skip = 0
    downloaded_paths = []

    # Defer QUB cookie extraction until we actually need it (lazy — avoids
    # slow browser DB access for empty lists or OA-only jobs)
    qub_cookies = None   # fetched lazily on first proxy attempt

    for idx, paper in enumerate(papers):
        if not _dl_active:
            break

        doi   = paper.get("doi", "").strip()
        title = paper.get("title", f"paper_{idx+1}")
        fname = safe_filename(title)
        dest  = str(Path(folder) / fname)

        _dl_queue.put({"type": "progress",
                       "pct": round(idx / total * 100),
                       "current": idx + 1, "total": total})

        if Path(dest).exists():
            skip += 1
            paper["status"] = "exists"
            downloaded_paths.append(dest)
            _dl_queue.put({"type": "item", "idx": idx, "status": "exists", "tag": "skip"})
            continue

        done  = False
        tried = []

        # 1) Elsevier Full-Text API (uses existing Scopus API key)
        if not done and doi and try_oa:
            tried.append("Elsevier")
            if try_elsevier_fulltext(doi, dest, api_key, session):
                ok += 1; done = True
                paper["status"] = "Elsevier ✓"
                downloaded_paths.append(dest)
                _dl_queue.put({"type": "item", "idx": idx,
                               "status": "Elsevier ✓", "tag": "ok"})

        # 2) Unpaywall — repository URLs first, then publisher
        if not done and doi and try_oa:
            tried.append("Unpaywall")
            for oa_url, host_type in get_unpaywall_urls(doi):
                if download_pdf_to(oa_url, dest, session):
                    ok += 1; done = True
                    label = "OA-repo ✓" if host_type == "repo" else "OA-pub ✓"
                    paper["status"] = label
                    downloaded_paths.append(dest)
                    _dl_queue.put({"type": "item", "idx": idx,
                                   "status": label, "tag": "ok"})
                    break

        # 3) Semantic Scholar open-access PDF
        if not done and doi and try_oa:
            tried.append("SemanticScholar")
            ss_url = get_semantic_scholar_pdf(doi)
            if ss_url and download_pdf_to(ss_url, dest, session):
                ok += 1; done = True
                paper["status"] = "OA-SS ✓"
                downloaded_paths.append(dest)
                _dl_queue.put({"type": "item", "idx": idx,
                               "status": "OA-SS ✓", "tag": "ok"})

        # 4) QUB proxy — with browser cookies + HTML citation_pdf_url parsing
        if not done and doi and try_prx:
            if qub_cookies is None:          # fetch once, on first proxy attempt
                qub_cookies = get_qub_cookies()
            tried.append("QUB-proxy")
            if try_proxy_download(doi, dest, session, extra_cookies=qub_cookies):
                ok += 1; done = True
                paper["status"] = "QUB ✓"
                downloaded_paths.append(dest)
                _dl_queue.put({"type": "item", "idx": idx,
                               "status": "QUB ✓", "tag": "ok"})

        if not done:
            fail += 1
            reason = "no DOI" if not doi else ("tried: " + ", ".join(tried))
            paper["status"] = "no PDF"
            proxy_link = f"https://doi-org.{QUB_PROXY}/{doi}" if doi else ""
            _dl_queue.put({"type": "item", "idx": idx,
                           "status": "no PDF", "tag": "fail",
                           "reason": reason,
                           "proxy_url": proxy_link})

    # ── Save summary CSV ──────────────────────────────────────────────────────
    csv_path = str(Path(folder) / "_download_summary.csv")
    try:
        with open(csv_path, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.DictWriter(f,
                fieldnames=["title","authors","year","source","doi","url","status"],
                extrasaction="ignore")
            w.writeheader()
            w.writerows(papers)
        downloaded_paths.append(csv_path)
    except Exception:
        pass

    # ── Build ZIP (for browser download in online mode) ───────────────────────
    zip_path = ""
    if downloaded_paths:
        try:
            zd = tempfile.mkdtemp()
            zip_path = str(Path(zd) / "papers.zip")
            with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
                for p in downloaded_paths:
                    if os.path.exists(p):
                        zf.write(p, os.path.basename(p))
            _dl_zip_path = zip_path
        except Exception:
            pass

    _dl_queue.put({
        "type": "done",
        "ok": ok, "fail": fail, "skip": skip,
        "csv": csv_path,
        "has_zip": bool(zip_path),
    })


# ── API ROUTES ────────────────────────────────────────────────────────────────
@app.route("/")
def index():
    return render_template_string(FRONTEND_HTML)

@app.route("/api/journals")
def api_journals():
    data   = load_abs_data()
    fields: dict = {}
    for j in data:
        label = FIELD_LABELS.get(j["field"], j["field"])
        fields.setdefault(j["field"], {"label": label, "journals": []})["journals"].append(
            {"title": j["title"], "rating": j["rating"]}
        )
    return jsonify({
        "fields": [
            {"code": k, "label": v["label"], "journals": v["journals"]}
            for k, v in fields.items()
        ],
        "is_online": IS_ONLINE,
    })

@app.route("/api/query", methods=["POST"])
def api_query():
    body = request.json or {}
    return jsonify({"query": build_scopus_query(
        body.get("journals", []),
        body.get("groups",   []),
        body.get("row_op",   "AND"),
    )})

@app.route("/api/search", methods=["POST"])
def api_search():
    body    = request.json or {}
    api_key = get_api_key()
    if not api_key:
        return jsonify({"error": "No Scopus API key configured."}), 400
    query   = body.get("query", "")
    max_r   = int(body.get("max_results", load_config().get("max_results", 200)))
    results = scopus_search(query, api_key, max_r, lambda _t, _m: None)
    return jsonify({"results": results, "count": len(results)})

@app.route("/api/download", methods=["POST"])
def api_download():
    global _dl_active, _dl_queue, _dl_zip_path
    body    = request.json or {}
    papers  = body.get("papers", [])
    try_oa  = body.get("try_oa",    True)
    try_prx = body.get("try_proxy", True)

    # Online mode: always use a temp folder; local mode: honour configured folder
    cfg = load_config()
    folder = body.get("folder", "").strip() or cfg.get("dl_folder", "").strip()
    if not folder or IS_ONLINE:
        folder = tempfile.mkdtemp()

    Path(folder).mkdir(parents=True, exist_ok=True)
    _dl_queue    = Queue()
    _dl_active   = True
    _dl_zip_path = ""

    threading.Thread(
        target=_download_worker,
        args=(papers, folder, try_oa, try_prx),
        daemon=True,
    ).start()
    return jsonify({"started": True, "folder": folder})

@app.route("/api/download/stream")
def api_download_stream():
    """SSE endpoint — open after POST /api/download."""
    def generate():
        # Send ~64 KB of SSE comment padding to immediately flush Werkzeug's
        # dev-server TCP send buffer (default ~65 KB on Windows).  Without this,
        # a small single-event response (e.g. an empty-list "done") sits in the
        # OS send buffer for up to 30 s until a ping fires.
        yield ": " + "x" * 65536 + "\n\n"
        while True:
            try:
                msg = _dl_queue.get(timeout=5)   # 5 s → frequent pings keep socket alive
                yield f"data: {json.dumps(msg)}\n\n"
                if msg.get("type") == "done":
                    break
            except Empty:
                yield 'data: {"type":"ping"}\n\n'
    return Response(
        stream_with_context(generate()),
        mimetype="text/event-stream",
        headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"},
    )

@app.route("/api/download/get-zip")
def api_download_get_zip():
    """Serve the completed ZIP for browser download."""
    if not _dl_zip_path or not Path(_dl_zip_path).exists():
        return jsonify({"error": "No ZIP ready — run a download first."}), 404
    return send_file(
        _dl_zip_path,
        as_attachment=True,
        download_name="papers.zip",
        mimetype="application/zip",
    )

@app.route("/api/download/stop", methods=["POST"])
def api_download_stop():
    global _dl_active
    _dl_active = False
    return jsonify({"stopped": True})

@app.route("/api/config", methods=["GET", "POST"])
def api_config():
    if request.method == "GET":
        cfg = load_config()
        cfg.pop("api_key", None)          # never expose key to frontend
        cfg["is_online"] = IS_ONLINE
        return jsonify(cfg)
    cfg = load_config()
    cfg.update(request.json or {})
    save_config(cfg)
    return jsonify({"saved": True})

@app.route("/api/config/key", methods=["POST"])
def api_config_key():
    cfg = load_config()
    cfg["api_key"] = (request.json or {}).get("api_key", "")
    save_config(cfg)
    return jsonify({"saved": True})

@app.route("/api/config/key/exists")
def api_config_key_exists():
    return jsonify({"exists": bool(get_api_key())})

@app.route("/api/proxy-cookies/check")
def api_proxy_cookies_check():
    cookies    = get_qub_cookies()
    cfg        = load_config()
    manual_set = bool(cfg.get("qub_session_cookie", "").strip())
    source     = "manual" if manual_set else "browser"
    # Detect v20 (App-Bound Encrypted) cookies even when they can't be decrypted
    v20_info   = {} if IS_ONLINE else check_chromium_cookie_status(QUB_PROXY)
    return jsonify({"found": bool(cookies), "count": len(cookies),
                    "names": list(cookies.keys())[:6], "source": source,
                    "v20_info": v20_info})

@app.route("/api/proxy-cookies/test")
def api_proxy_cookies_test():
    """Test whether the stored QUB proxy cookie actually grants access (not just login redirect)."""
    cookies = get_qub_cookies()
    if not cookies:
        return jsonify({"success": False, "reason": "no_cookie",
                        "message": "No QUB session cookie configured."})
    # Use a well-known Elsevier DOI to test proxy access
    TEST_DOI = "10.1016/j.omega.2020.102215"
    session  = requests.Session()
    for k, v in cookies.items():
        session.cookies.set(k, v, domain=QUB_PROXY)
    try:
        r = session.get(
            f"https://doi-org.{QUB_PROXY}/{TEST_DOI}",
            headers={"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
                     "Accept": "text/html,application/pdf,*/*"},
            timeout=20, allow_redirects=True,
        )
        final_url = r.url
        ct = r.headers.get("content-type", "").lower()
        if "login" in final_url or "shibboleth" in final_url.lower() or "idp" in final_url:
            return jsonify({"success": False, "reason": "redirected_to_login",
                            "message": "Cookie is expired or invalid — QUB redirected to login page.",
                            "url": final_url})
        if r.status_code == 200 and "pdf" in ct:
            return jsonify({"success": True, "reason": "got_pdf",
                            "message": "Connected successfully — received PDF directly."})
        if r.status_code == 200 and "html" in ct:
            preview = r.content[:3000].decode("utf-8", errors="replace").lower()
            if "sign in" in preview or ("login" in preview and "qub" not in preview):
                return jsonify({"success": False, "reason": "login_in_html",
                                "message": "Cookie expired — publisher page shows a login prompt."})
            return jsonify({"success": True, "reason": "got_publisher_page",
                            "message": "Connected successfully — reached publisher page via QUB proxy.",
                            "url": final_url})
        return jsonify({"success": False, "reason": f"http_{r.status_code}",
                        "message": f"Unexpected response: HTTP {r.status_code}"})
    except Exception as e:
        return jsonify({"success": False, "reason": "error", "message": str(e)})

@app.route("/api/proxy-cookies/launch-browser", methods=["POST"])
def api_launch_browser_for_cookie():
    """Launch a fresh Chrome window with remote debugging so CDP can read cookies."""
    global _cdp_port
    if IS_ONLINE:
        return jsonify({"error": "not available in online mode"}), 400
    chrome = _find_chrome_exe()
    if not chrome:
        return jsonify({"error": "Chrome not found — please install Google Chrome."}), 404
    import subprocess
    import socket as _sock
    # Find a free debug port (9222–9229)
    for port in range(9222, 9230):
        try:
            s = _sock.create_connection(("localhost", port), timeout=0.2)
            s.close()
        except OSError:
            _cdp_port = port
            break
    else:
        _cdp_port = 9222
    tmp_profile = tempfile.mkdtemp(prefix="qub_cdp_")
    subprocess.Popen(
        [chrome,
         f"--remote-debugging-port={_cdp_port}",
         f"--user-data-dir={tmp_profile}",
         "--no-first-run", "--no-default-browser-check",
         "--disable-extensions",
         "https://qub.idm.oclc.org/login"],
        stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
    )
    return jsonify({"launched": True, "port": _cdp_port})

@app.route("/api/proxy-cookies/grab-from-browser")
def api_grab_cookie_from_browser():
    """Poll Chrome CDP for the EZproxy cookie; save and return when found."""
    if IS_ONLINE:
        return jsonify({"error": "not available"}), 400
    if not _cdp_port:
        return jsonify({"ready": False, "reason": "not_launched"})
    try:
        cookies = cdp_get_qub_cookies(_cdp_port)
    except ConnectionRefusedError:
        return jsonify({"ready": False, "reason": "chrome_starting"})
    except Exception as e:
        return jsonify({"ready": False, "reason": str(e)})
    ez_val = cookies.get("EZproxy")
    if not ez_val:
        return jsonify({"ready": True, "found": False,
                        "available": list(cookies.keys())[:10]})
    cfg = load_config()
    cfg["qub_session_cookie"] = ez_val   # get_qub_cookies() wraps it as {"EZproxy": ez_val}
    save_config(cfg)
    return jsonify({"ready": True, "found": True, "cookie_name": "EZproxy"})

@app.route("/api/config/qub-cookie", methods=["POST"])
def api_config_qub_cookie():
    cfg = load_config()
    cfg["qub_session_cookie"] = (request.json or {}).get("cookie", "").strip()
    save_config(cfg)
    return jsonify({"saved": True})

@app.route("/api/config/qub-cookie/exists")
def api_config_qub_cookie_exists():
    exists = bool(load_config().get("qub_session_cookie", "").strip())
    return jsonify({"exists": exists})

@app.route("/api/open-folder", methods=["POST"])
def api_open_folder():
    if IS_ONLINE:
        return jsonify({"ok": False, "reason": "not available in online mode"})
    import subprocess
    folder = (request.json or {}).get("folder", load_config().get("dl_folder", ""))
    if folder and Path(folder).exists():
        if sys.platform == "win32":
            os.startfile(folder)
        else:
            subprocess.Popen(["xdg-open", folder])
    return jsonify({"ok": True})


# ── FRONTEND ──────────────────────────────────────────────────────────────────
FRONTEND_HTML = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Literature Downloader · QUB</title>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Segoe UI',Arial,sans-serif;background:#f0f2f5;color:#222;height:100vh;display:flex;flex-direction:column;overflow:hidden}

/* HEADER */
header{background:linear-gradient(135deg,#1a3a5c,#0d2340);color:#fff;padding:10px 18px;display:flex;align-items:center;gap:12px;flex-shrink:0;box-shadow:0 2px 8px rgba(0,0,0,.3)}
header h1{font-size:1.05rem;font-weight:700}
header .sub{font-size:.72rem;opacity:.65;margin-top:1px}
.hbadge{background:rgba(255,255,255,.15);border-radius:20px;padding:3px 12px;font-size:.7rem;font-weight:600;white-space:nowrap}
.hbadge.online{background:#2e7d32}

/* TABS */
.tabs{display:flex;background:#fff;border-bottom:2px solid #dde3ea;flex-shrink:0;padding:0 12px}
.tab{padding:9px 16px;font-size:.82rem;font-weight:600;color:#888;cursor:pointer;border-bottom:3px solid transparent;transition:color .15s,border-color .15s;user-select:none;white-space:nowrap}
.tab:hover{color:#1a3a5c}.tab.active{color:#1a3a5c;border-bottom-color:#1a3a5c}

/* PANELS */
.panel{display:none;flex:1;overflow:hidden}.panel.active{display:flex}

/* SEARCH PANEL */
#panel-search{flex-direction:row}

/* SIDEBAR */
.sidebar{width:290px;min-width:220px;background:#fff;border-right:1px solid #dde3ea;display:flex;flex-direction:column;flex-shrink:0}
.sidebar-hd{padding:9px 11px;border-bottom:1px solid #dde3ea;background:#f7f9fb;flex-shrink:0}
.sidebar-hd h2{font-size:.68rem;text-transform:uppercase;letter-spacing:.08em;color:#999;font-weight:700;margin-bottom:6px}
.sidebar-topbar{display:flex;gap:5px;flex-wrap:wrap;margin-bottom:6px}
.sidebar-topbar button{font-size:.7rem;padding:3px 8px;border:1px solid #c5cdd8;background:#fff;border-radius:4px;cursor:pointer;color:#555;transition:background .1s}
.sidebar-topbar button:hover{background:#e8edf3}
.sel-info{font-size:.7rem;color:#1a3a5c;font-weight:600}
.rating-bar{display:flex;gap:4px;flex-wrap:wrap;margin-bottom:2px}
.r-chip{display:inline-flex;align-items:center;gap:4px;padding:3px 8px;border-radius:12px;font-size:.72rem;font-weight:700;cursor:pointer;user-select:none;border:2px solid transparent;transition:opacity .15s}
.r-chip input{width:12px;height:12px;cursor:pointer;accent-color:#fff}
.r4s{background:#1a3a5c;color:#fff}.r4{background:#2e7d32;color:#fff}
.r3{background:#e65100;color:#fff}.r2{background:#7b1fa2;color:#fff}.r1{background:#757575;color:#fff}
.r-chip.off{opacity:.35}
.sidebar-body{flex:1;overflow-y:auto}

/* FIELD ACCORDION */
.field-group{border-bottom:1px solid #eef1f5}
.fgh{display:flex;align-items:center;gap:6px;padding:7px 10px;cursor:pointer;user-select:none;transition:background .1s;position:sticky;top:0;background:#fff;z-index:1}
.fgh:hover{background:#f7f9fb}.fgh.open{background:#eef3f9}
.fgh-arrow{font-size:.68rem;color:#bbb;transition:transform .18s;flex-shrink:0}
.fgh.open .fgh-arrow{transform:rotate(90deg)}
.fgh-cb{width:13px;height:13px;accent-color:#1a3a5c;cursor:pointer;flex-shrink:0}
.fgh-label{font-size:.77rem;font-weight:600;color:#333;flex:1;line-height:1.3}
.fgh-counts{font-size:.65rem;color:#bbb;white-space:nowrap}
.fgh-sel{font-size:.65rem;background:#1a3a5c;color:#fff;border-radius:9px;padding:1px 6px;white-space:nowrap;display:none}
.fgh-sel.on{display:inline}
.jlist{display:none;background:#fafbfc}.jlist.open{display:block}
.ji{display:flex;align-items:center;gap:6px;padding:4px 10px 4px 24px;cursor:pointer;transition:background .1s}
.ji:hover{background:#f0f4fa}
.ji input{width:12px;height:12px;accent-color:#1a3a5c;cursor:pointer;flex-shrink:0}
.ji label{font-size:.74rem;cursor:pointer;color:#444;flex:1;line-height:1.3}
.jbadge{font-size:.63rem;font-weight:700;padding:1px 5px;border-radius:3px;white-space:nowrap;flex-shrink:0}
.b4s{background:#1a3a5c;color:#fff}.b4{background:#2e7d32;color:#fff}
.b3{background:#e65100;color:#fff}.b2{background:#7b1fa2;color:#fff}.b1{background:#757575;color:#fff}

/* RIGHT PANEL */
.search-right{flex:1;display:flex;flex-direction:column;overflow:hidden;padding:12px 16px;gap:10px}
.kw-card{background:#fff;border-radius:8px;border:1px solid #dde3ea;padding:12px 14px;flex-shrink:0}
.kw-card h3{font-size:.7rem;text-transform:uppercase;letter-spacing:.08em;color:#999;font-weight:700;margin-bottom:10px}
.kw-group{display:flex;align-items:flex-start;gap:8px;margin-bottom:8px;flex-wrap:wrap}
.kw-group-label{font-size:.72rem;color:#888;min-width:72px;padding-top:7px;font-weight:600}
.tags-box{display:flex;flex-wrap:wrap;align-items:center;gap:4px;background:#f7f9fb;border:1.5px solid #c5cdd8;border-radius:6px;padding:4px 8px;flex:1;min-width:240px;cursor:text;transition:border-color .15s}
.tags-box:focus-within{border-color:#1a3a5c;background:#fff}
.tag{display:inline-flex;align-items:center;gap:3px;border-radius:4px;padding:3px 7px;font-size:.75rem;font-weight:600;white-space:nowrap}
.tag-and{background:#1a3a5c;color:#fff}.tag-or{background:#e67e00;color:#fff}
.tag-x{cursor:pointer;opacity:.7;font-size:.85rem;margin-left:1px}.tag-x:hover{opacity:1}
.tags-input{border:none;outline:none;font-size:.82rem;background:transparent;color:#222;min-width:140px;flex:1;padding:3px 0}
.within-op{display:flex;align-items:center;gap:6px;font-size:.71rem;color:#aaa}
.within-op label{cursor:pointer;color:#555}
.row-between{display:flex;align-items:center;gap:8px;font-size:.71rem;color:#888;padding-left:80px;margin-bottom:4px}
.add-grp{font-size:.73rem;color:#1a3a5c;background:none;border:none;cursor:pointer;font-weight:600;padding:0;opacity:.8}
.add-grp:hover{opacity:1}
.preview-card{background:#0d2340;border-radius:8px;padding:10px 13px;flex-shrink:0;position:relative}
.preview-card pre{font-family:Consolas,monospace;font-size:.76rem;color:#a8c4e0;white-space:pre-wrap;word-break:break-all;max-height:120px;overflow-y:auto;line-height:1.55}
.preview-card .pc-label{font-size:.65rem;color:#5a7a9c;text-transform:uppercase;letter-spacing:.07em;margin-bottom:5px}
.action-bar{display:flex;gap:7px;flex-wrap:wrap;flex-shrink:0}

/* BUTTONS */
.btn{display:inline-flex;align-items:center;gap:5px;font-size:.79rem;padding:7px 13px;border-radius:6px;border:none;cursor:pointer;font-weight:600;transition:background .15s,transform .1s;white-space:nowrap}
.btn:active{transform:scale(.97)}
.btn-blue{background:#1a3a5c;color:#fff}.btn-blue:hover{background:#15314f}
.btn-orange{background:#e67e00;color:#fff}.btn-orange:hover{background:#c96e00}
.btn-red{background:#c0392b;color:#fff}.btn-red:hover{background:#a93226}
.btn-ghost{background:#e8edf3;color:#333;border:1px solid #c5cdd8}.btn-ghost:hover{background:#dde3ea}
.btn-green{background:#2e7d32;color:#fff}.btn-green:hover{background:#256328}
.btn-sm{padding:4px 9px;font-size:.72rem}
.btn:disabled{opacity:.45;cursor:not-allowed;transform:none!important}

/* RESULTS PANEL */
#panel-results{flex-direction:column}
.results-toolbar{background:#fff;border-bottom:1px solid #dde3ea;padding:8px 14px;display:flex;align-items:center;gap:8px;flex-wrap:wrap;flex-shrink:0}
.res-count{font-size:.82rem;color:#555;font-weight:500}.res-count strong{color:#1a3a5c}
.prog-wrap{display:flex;align-items:center;gap:8px;margin-left:auto}
.prog-bar-outer{width:160px;height:8px;background:#e8edf3;border-radius:4px;overflow:hidden}
.prog-bar-inner{height:100%;background:#2e7d32;border-radius:4px;transition:width .3s}
.prog-label{font-size:.72rem;color:#888;white-space:nowrap}
.results-body{flex:1;overflow:auto;padding:10px 14px}
.rtable{width:100%;border-collapse:collapse;background:#fff;border-radius:8px;overflow:hidden;box-shadow:0 1px 4px rgba(0,0,0,.07)}
.rtable th{background:#eef1f5;font-size:.67rem;font-weight:700;color:#666;text-transform:uppercase;letter-spacing:.05em;padding:8px 12px;text-align:left;border-bottom:1px solid #dde3ea;white-space:nowrap;cursor:pointer;user-select:none}
.rtable th:hover{background:#e3e9f2}
.rtable td{padding:7px 12px;font-size:.78rem;border-bottom:1px solid #f0f2f5;vertical-align:middle}
.rtable tr:last-child td{border-bottom:none}
.rtable tr:hover td{background:#f7f9fb}
.rtable tr.ok td{background:#f1f8f1}.rtable tr.fail td{background:#fff5f5}.rtable tr.skip td{background:#fffde7}
.st-badge{display:inline-flex;align-items:center;padding:2px 7px;border-radius:10px;font-size:.68rem;font-weight:700}
.st-ok{background:#e8f5e9;color:#2e7d32}.st-fail{background:#ffebee;color:#c0392b}
.st-skip{background:#fff8e1;color:#e67e00}.st-pend{background:#eee;color:#888}
.doi-link,.qub-link{font-size:.7rem;text-decoration:none;padding:2px 6px;border-radius:3px;white-space:nowrap}
.doi-link{color:#1a3a5c}.doi-link:hover{text-decoration:underline}
.qub-link{background:#e67e00;color:#fff;margin-left:4px}.qub-link:hover{background:#c96e00}
.empty-state{text-align:center;padding:60px 20px;color:#ccc;font-size:.9rem}
.empty-icon{font-size:2.5rem;margin-bottom:10px}

/* SETTINGS PANEL */
#panel-settings{flex-direction:column;overflow-y:auto}
.settings-body{padding:20px 24px;max-width:720px}
.settings-body h2{font-size:1rem;font-weight:700;color:#1a3a5c;margin-bottom:4px}
.settings-body .sdesc{font-size:.78rem;color:#888;margin-bottom:20px;line-height:1.6}
.form-row{display:flex;flex-direction:column;gap:4px;margin-bottom:14px}
.form-row label{font-size:.75rem;font-weight:600;color:#555}
.form-row input[type=text],.form-row input[type=password],.form-row input[type=number]{padding:8px 11px;border:1.5px solid #c5cdd8;border-radius:6px;font-size:.84rem;width:100%;max-width:480px;transition:border-color .15s}
.form-row input:focus{outline:none;border-color:#1a3a5c}
.form-row .hint{font-size:.72rem;color:#aaa;margin-top:2px}
.form-row .folder-row{display:flex;gap:6px;align-items:center}
.form-row .folder-row input{flex:1}
.check-row{display:flex;align-items:center;gap:8px;margin-bottom:8px;font-size:.82rem;cursor:pointer}
.check-row input{width:15px;height:15px;cursor:pointer;accent-color:#1a3a5c}
.sep{border:none;border-top:1px solid #dde3ea;margin:20px 0}
.help-box{background:#f7f9fb;border:1px solid #dde3ea;border-radius:8px;padding:14px 16px;font-size:.78rem;color:#444;line-height:1.75}
.help-box h4{font-size:.82rem;font-weight:700;color:#1a3a5c;margin-bottom:8px}
.help-box .step{display:flex;gap:10px;margin-bottom:6px}
.help-box .sn{background:#1a3a5c;color:#fff;border-radius:50%;width:20px;height:20px;display:flex;align-items:center;justify-content:center;font-size:.65rem;font-weight:700;flex-shrink:0;margin-top:1px}
.info-banner{background:#e8f4fd;border:1px solid #b3d9f7;border-radius:7px;padding:10px 14px;font-size:.78rem;color:#1a3a5c;margin-bottom:16px;display:none}
.info-banner.show{display:block}

/* LOG */
.log-card{background:#fff;border-top:1px solid #dde3ea;padding:8px 14px;flex-shrink:0}
.log-card pre{font-family:Consolas,monospace;font-size:.75rem;color:#555;max-height:110px;overflow-y:auto;white-space:pre-wrap;word-break:break-all}

/* TOAST */
#toast{position:fixed;bottom:20px;right:20px;background:#1a3a5c;color:#fff;padding:9px 16px;border-radius:6px;font-size:.78rem;opacity:0;transition:opacity .3s;pointer-events:none;z-index:9999}
#toast.show{opacity:1}
</style>
</head>
<body>

<header>
  <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="white" stroke-width="2">
    <path d="M2 3h6a4 4 0 0 1 4 4v14a3 3 0 0 0-3-3H2z"/>
    <path d="M22 3h-6a4 4 0 0 0-4 4v14a3 3 0 0 1 3-3h7z"/>
  </svg>
  <div>
    <h1>Literature Auto-Downloader</h1>
    <div class="sub">AJG 2024 · QUB Library · Scopus</div>
  </div>
  <div class="hbadge" id="modeBadge">Local</div>
</header>

<div class="tabs">
  <div class="tab active" data-tab="search">Search &amp; Select</div>
  <div class="tab" data-tab="results">Results <span id="res-tab-badge"></span></div>
  <div class="tab" data-tab="settings">Settings</div>
</div>

<!-- ── SEARCH PANEL ── -->
<div class="panel active" id="panel-search">
  <div class="sidebar">
    <div class="sidebar-hd">
      <h2>ABS Journals <span class="sel-info" id="selCount"></span></h2>
      <div class="rating-bar" id="ratingBar"></div>
      <div class="sidebar-topbar">
        <button onclick="selAll()">All</button>
        <button onclick="selNone()">None</button>
        <button onclick="collapseAll()">Collapse</button>
        <input id="sbSearch" type="text" placeholder="Filter journals…"
               style="font-size:.72rem;padding:3px 6px;border:1px solid #c5cdd8;border-radius:4px;width:100%;margin-top:4px"
               oninput="filterSidebar()">
      </div>
    </div>
    <div class="sidebar-body" id="sidebarBody"></div>
  </div>

  <div class="search-right">
    <div class="kw-card">
      <h3>Keywords — searched in Title, Abstract &amp; Keywords (TITLE-ABS-KEY)</h3>
      <div id="kwGroups"></div>
      <div style="display:flex;align-items:center;gap:16px;margin-top:6px;flex-wrap:wrap">
        <button class="add-grp" onclick="addKwGroup()">＋ Add keyword group</button>
        <div class="row-between" style="padding-left:0">
          <span>Between groups:</span>
          <label style="cursor:pointer;font-size:.75rem"><input type="radio" name="rowOp" value="AND" checked> AND</label>
          <label style="cursor:pointer;font-size:.75rem"><input type="radio" name="rowOp" value="OR"> OR</label>
        </div>
      </div>
    </div>

    <div class="preview-card">
      <div class="pc-label">Scopus query preview</div>
      <pre id="queryPreview">(select journals and enter keywords to build query)</pre>
    </div>

    <div class="action-bar">
      <button class="btn btn-blue" onclick="refreshPreview()">↻ Refresh</button>
      <button class="btn btn-ghost" onclick="copyQuery()">📋 Copy query</button>
      <button class="btn btn-orange" onclick="openScopus()">🔍 Open Scopus (QUB)</button>
      <button class="btn btn-blue" id="btnSearch" onclick="doSearch()">🔎 Search Scopus API</button>
      <button class="btn btn-ghost" onclick="importCSV()">⬆ Import Scopus CSV</button>
    </div>
  </div>
</div>

<!-- ── RESULTS PANEL ── -->
<div class="panel" id="panel-results">
  <div class="results-toolbar">
    <span class="res-count" id="resCount">No results yet</span>
    <button class="btn btn-red btn-sm" id="btnDl" onclick="startDownload()" disabled>⬇ Download PDFs</button>
    <button class="btn btn-ghost btn-sm" id="btnStop" onclick="stopDownload()" disabled>■ Stop</button>
    <button class="btn btn-green btn-sm" id="btnZip" onclick="downloadZip()" style="display:none">📦 Save ZIP</button>
    <button class="btn btn-orange btn-sm" id="btnQUB" onclick="openFailedInQUB()" style="display:none">🔗 Open failed in QUB</button>
    <button class="btn btn-ghost btn-sm" onclick="exportCSV()">💾 Export CSV</button>
    <button class="btn btn-ghost btn-sm" id="btnFolder" onclick="openFolder()">📁 Open folder</button>
    <div class="prog-wrap" id="progWrap" style="display:none">
      <div class="prog-bar-outer"><div class="prog-bar-inner" id="progBar" style="width:0"></div></div>
      <span class="prog-label" id="progLabel"></span>
    </div>
  </div>
  <div class="results-body">
    <div class="empty-state" id="emptyState">
      <div class="empty-icon">📄</div>
      <div>Run a search or import a Scopus CSV to see results here</div>
    </div>
    <table class="rtable" id="resTable" style="display:none">
      <thead><tr>
        <th onclick="sortTable('title')">Title ↕</th>
        <th onclick="sortTable('authors')">Authors ↕</th>
        <th onclick="sortTable('year')">Year ↕</th>
        <th onclick="sortTable('source')">Journal ↕</th>
        <th onclick="sortTable('_rating')">AJG ↕</th>
        <th>DOI / Access</th>
        <th onclick="sortTable('status')">Status ↕</th>
      </tr></thead>
      <tbody id="resTbody"></tbody>
    </table>
  </div>
  <div class="log-card" id="logCard" style="display:none">
    <pre id="logPre"></pre>
  </div>
</div>

<!-- ── SETTINGS PANEL ── -->
<div class="panel" id="panel-settings">
  <div class="settings-body">
    <h2>Settings</h2>
    <p class="sdesc">Configure your Scopus API key and download preferences.</p>

    <div class="info-banner" id="onlineBanner">
      🌐 <strong>Online mode</strong> — the Scopus API key is pre-configured on the server.
      Downloaded papers are packaged as a <strong>ZIP file</strong> sent to your browser.
      The QUB proxy download is available for campus/VPN sessions.
    </div>

    <div class="form-row" id="apiKeyRow">
      <label>Scopus API Key</label>
      <input type="password" id="cfgApiKey" placeholder="Paste your Elsevier API key here">
      <span class="hint">Get a free key at <strong>dev.elsevier.com</strong> — sign in with your QUB email for institutional access.</span>
    </div>

    <div class="form-row" id="folderRow">
      <label>Download Folder <span style="font-weight:400;color:#aaa">(local mode only)</span></label>
      <div class="folder-row">
        <input type="text" id="cfgFolder" placeholder="e.g. C:\Users\you\Downloads\papers">
        <button class="btn btn-ghost btn-sm" onclick="browseFolder()">Browse…</button>
      </div>
      <span class="hint">Leave blank to receive papers as a ZIP download in your browser.</span>
    </div>

    <div class="form-row">
      <label>Max results per search</label>
      <input type="number" id="cfgMax" value="200" min="10" max="2000" style="max-width:120px">
    </div>

    <label class="check-row"><input type="checkbox" id="cfgOA" checked> Try Unpaywall for open-access PDFs (works in all modes)</label>
    <label class="check-row"><input type="checkbox" id="cfgProxy" checked> Try QUB proxy for paywalled papers (requires campus network or VPN)</label>

    <div id="cookieStatus" style="display:none;margin:8px 0 12px 24px;padding:9px 14px;border-radius:6px;font-size:.78rem;border:1px solid #dde3ea;background:#f7f9fb">
      <strong id="cookieStatusIcon"></strong> <span id="cookieStatusMsg"></span>
      <div id="cookieStatusNames" style="font-size:.72rem;color:#888;margin-top:3px"></div>
    </div>
    <div style="margin:-4px 0 14px 24px">
      <button class="btn btn-ghost btn-sm" onclick="checkProxyCookies()">🔍 Check browser cookies for QUB proxy</button>
    </div>

    <div class="form-row" id="qubCookieRow" style="margin-top:4px;padding-top:14px;border-top:1px solid #eef1f5">
      <label>QUB EZproxy Session Cookie <span style="font-weight:400;color:#aaa">(paste manually if auto-detect fails)</span></label>
      <input type="password" id="cfgQubCookie" placeholder="(paste EZproxy cookie value here)">
      <span class="hint">
        <strong>How to get it:</strong> In Edge, open any QUB library page while logged in → press <kbd>F12</kbd> →
        <strong>Application</strong> tab → <strong>Cookies</strong> → click <strong>qub.idm.oclc.org</strong> →
        find the cookie named <strong>EZproxy</strong> → copy the <strong>Value</strong> column → paste above.<br>
        The session expires when you log out of QUB Library — re-paste if downloads stop working.
      </span>
      <div style="margin-top:7px;display:flex;gap:8px;align-items:center;flex-wrap:wrap">
        <button class="btn btn-blue btn-sm" onclick="autoConfigQubCookie()">🔑 Auto-configure</button>
        <button class="btn btn-ghost btn-sm" onclick="testQubCookie()">🔌 Test QUB connection</button>
        <button class="btn btn-ghost btn-sm" onclick="clearQubCookie()" id="btnClearQubCookie" style="display:none">🗑 Clear saved cookie</button>
        <span id="qubCookieExistsMsg" style="font-size:.73rem;color:#2e7d32;display:none">✅ Cookie saved</span>
      </div>
      <div id="qubTestStatus" style="display:none;margin-top:8px;padding:8px 12px;border-radius:5px;font-size:.78rem;border:1px solid #dde3ea;background:#f7f9fb">
        <strong id="qubTestIcon"></strong> <span id="qubTestMsg"></span>
      </div>
    </div>

    <button class="btn btn-blue" onclick="saveSettings()">Save Settings</button>

    <hr class="sep">

    <div class="help-box">
      <h4>How to use</h4>
      <div class="step"><div class="sn">1</div><div>On <strong>Search &amp; Select</strong>, choose AJG rating levels and tick the journals you want to include.</div></div>
      <div class="step"><div class="sn">2</div><div>Enter keywords — press <kbd>Enter</kbd> or click <strong>Add</strong>. Terms in a group are joined AND/OR; multiple groups are also joined AND/OR.</div></div>
      <div class="step"><div class="sn">3</div><div>Click <strong>Search Scopus API</strong> (requires API key in Settings), or open Scopus manually and use <strong>Import Scopus CSV</strong>.</div></div>
      <div class="step"><div class="sn">4</div><div>On the <strong>Results</strong> tab, click <strong>⬇ Download PDFs</strong>. Papers are tried in order: Elsevier API → Unpaywall OA → Semantic Scholar → QUB proxy (auto-uses your browser cookies).</div></div>
      <div class="step"><div class="sn">5</div><div>For QUB proxy: log into the QUB Library portal in Edge, then in Settings paste the <strong>EZproxy</strong> cookie value (F12 → Application → Cookies → qub.idm.oclc.org) and click <strong>Test QUB connection</strong>.</div></div>
      <div class="step"><div class="sn">6</div><div>Papers that still can't be auto-downloaded show a <strong>QUB link</strong> in the table — click it to open in your browser as a last resort.</div></div>
    </div>
  </div>
</div>

<div id="toast"></div>
<input type="file" id="csvInput" accept=".csv" style="display:none" onchange="handleCSV(event)">

<script>
// ── STATE ──────────────────────────────────────────────────────
let allFields   = [];
let selected    = new Set();
let kwGroups    = [];
let kwInputRefs = [];   // [{group, tagsBox, inp}] — for flushing uncommitted text
let results     = [];
let sortCol     = null, sortAsc = true;
let dlEs        = null;
let ratingMap   = {};
let isOnline    = false;

const RATING_ORDER = {'4*':0,'4':1,'3':2,'2':3,'1':4};
const BADGE = {'4*':'b4s','4':'b4','3':'b3','2':'b2','1':'b1'};
const ACTIVE_RATINGS = new Set(['4*','4','3']);

// ── INIT ────────────────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', async () => {
  buildRatingBar();
  await loadJournals();
  buildSidebar();
  addKwGroup();
  await loadSettings();
  document.querySelectorAll('.tab').forEach(t =>
    t.addEventListener('click', () => switchTab(t.dataset.tab)));
});

// ── TABS ────────────────────────────────────────────────────────
function switchTab(name) {
  document.querySelectorAll('.tab').forEach(t =>
    t.classList.toggle('active', t.dataset.tab === name));
  document.querySelectorAll('.panel').forEach(p =>
    p.classList.toggle('active', p.id === 'panel-' + name));
  if (name === 'search') refreshPreview();
}

// ── RATING BAR ──────────────────────────────────────────────────
function buildRatingBar() {
  const bar = document.getElementById('ratingBar');
  [['4*','r4s','4★'],['4','r4','4'],['3','r3','3'],['2','r2','2'],['1','r1','1']].forEach(([r,cls,lbl]) => {
    const chip = document.createElement('label');
    chip.className = 'r-chip ' + cls + (ACTIVE_RATINGS.has(r) ? '' : ' off');
    chip.innerHTML = '<input type="checkbox" ' + (ACTIVE_RATINGS.has(r)?'checked':'') + '> ' + lbl;
    chip.querySelector('input').addEventListener('change', e => {
      if (e.target.checked) ACTIVE_RATINGS.add(r); else ACTIVE_RATINGS.delete(r);
      chip.classList.toggle('off', !e.target.checked);
      buildSidebar(); refreshPreview();
    });
    bar.appendChild(chip);
  });
}

// ── JOURNALS ────────────────────────────────────────────────────
async function loadJournals() {
  const res = await fetch('/api/journals');
  const data = await res.json();
  allFields  = data.fields;
  isOnline   = data.is_online || false;
  ratingMap  = {};
  allFields.forEach(f => f.journals.forEach(j => ratingMap[j.title] = j.rating));
}

function buildSidebar(filter) {
  filter = (filter || '').toLowerCase();
  const body = document.getElementById('sidebarBody');
  body.innerHTML = '';

  allFields.forEach(field => {
    const journals = field.journals.filter(j =>
      ACTIVE_RATINGS.has(j.rating) &&
      (!filter || j.title.toLowerCase().includes(filter) || field.label.toLowerCase().includes(filter))
    );
    if (!journals.length) return;

    const selCount = journals.filter(j => selected.has(j.title)).length;
    const fg = document.createElement('div');
    fg.className = 'field-group';

    const hd = document.createElement('div');
    hd.className = 'fgh';
    hd.innerHTML =
      '<span class="fgh-arrow">▶</span>' +
      '<input type="checkbox" class="fgh-cb">' +
      '<span class="fgh-label">' + escHtml(field.label) + '</span>' +
      '<span class="fgh-counts">' + journals.length + '</span>' +
      '<span class="fgh-sel' + (selCount ? ' on' : '') + '">' + selCount + '</span>';

    const cb = hd.querySelector('.fgh-cb');
    cb.checked = selCount === journals.length;
    if (selCount > 0 && selCount < journals.length) cb.indeterminate = true;

    const jl = document.createElement('div');
    jl.className = 'jlist';

    journals.forEach(j => {
      const ji = document.createElement('div');
      ji.className = 'ji';
      ji.innerHTML =
        '<input type="checkbox" ' + (selected.has(j.title) ? 'checked' : '') + '>' +
        '<label>' + escHtml(j.title) + '</label>' +
        '<span class="jbadge ' + (BADGE[j.rating]||'b1') + '">' + j.rating + '</span>';
      const jcb = ji.querySelector('input');
      jcb.addEventListener('change', () => {
        if (jcb.checked) selected.add(j.title); else selected.delete(j.title);
        updateSelCount(); updateFieldHeader(hd, journals); refreshPreview();
      });
      ji.querySelector('label').addEventListener('click', () => jcb.click());
      jl.appendChild(ji);
    });

    cb.addEventListener('change', () => {
      journals.forEach(j => { if (cb.checked) selected.add(j.title); else selected.delete(j.title); });
      jl.querySelectorAll('input[type=checkbox]').forEach(c => c.checked = cb.checked);
      updateSelCount(); updateFieldHeader(hd, journals); refreshPreview();
    });

    const arrow = hd.querySelector('.fgh-arrow');
    hd.addEventListener('click', e => {
      if (e.target === cb || e.target.type === 'checkbox') return;
      const open = jl.classList.toggle('open');
      hd.classList.toggle('open', open);
      arrow.style.transform = open ? 'rotate(90deg)' : '';
    });

    fg.appendChild(hd);
    fg.appendChild(jl);
    body.appendChild(fg);
  });
  updateSelCount();
}

function updateFieldHeader(hd, journals) {
  const n = journals.filter(j => selected.has(j.title)).length;
  const cb = hd.querySelector('.fgh-cb');
  cb.checked = n === journals.length;
  cb.indeterminate = n > 0 && n < journals.length;
  const badge = hd.querySelector('.fgh-sel');
  badge.textContent = n;
  badge.classList.toggle('on', n > 0);
}

function updateSelCount() {
  const n = [...selected].filter(t => ACTIVE_RATINGS.has(ratingMap[t])).length;
  document.getElementById('selCount').textContent = n ? n + ' selected' : '';
}

function selAll() {
  allFields.forEach(f => f.journals.filter(j => ACTIVE_RATINGS.has(j.rating))
    .forEach(j => selected.add(j.title)));
  buildSidebar(document.getElementById('sbSearch').value); refreshPreview();
}
function selNone() {
  allFields.forEach(f => f.journals.forEach(j => selected.delete(j.title)));
  buildSidebar(document.getElementById('sbSearch').value); refreshPreview();
}
function collapseAll() {
  document.querySelectorAll('.jlist.open').forEach(el => {
    el.classList.remove('open');
    el.previousSibling.classList.remove('open');
    el.previousSibling.querySelector('.fgh-arrow').style.transform = '';
  });
}
function filterSidebar() { buildSidebar(document.getElementById('sbSearch').value); }

// ── KEYWORD GROUPS ───────────────────────────────────────────────
function addKwGroup() {
  const idx   = kwGroups.length;
  const group = { terms: [], within: 'AND' };
  kwGroups.push(group);

  const container  = document.getElementById('kwGroups');
  const row        = document.createElement('div');
  row.className    = 'kw-group';

  const lbl        = document.createElement('span');
  lbl.className    = 'kw-group-label';
  lbl.textContent  = idx === 0 ? 'Keywords:' : 'Group ' + (idx + 1) + ':';

  const tagsBox    = document.createElement('div');
  tagsBox.className= 'tags-box';

  const inp        = document.createElement('input');
  inp.className    = 'tags-input';
  inp.placeholder  = 'Type keyword, press Enter';

  inp.addEventListener('keydown', e => {
    if (e.key === 'Enter') { e.preventDefault(); addTerm(group, tagsBox, inp); }
    if (e.key === 'Backspace' && !inp.value && group.terms.length)
      removeLast(group, tagsBox);
  });
  inp.addEventListener('blur', () => addTerm(group, tagsBox, inp));
  kwInputRefs.push({ group, tagsBox, inp });
  tagsBox.appendChild(inp);
  tagsBox.addEventListener('click', () => inp.focus());

  const addBtn     = document.createElement('button');
  addBtn.className = 'btn btn-ghost btn-sm';
  addBtn.textContent = 'Add';
  addBtn.onclick   = () => addTerm(group, tagsBox, inp);

  const opRow      = document.createElement('div');
  opRow.className  = 'within-op';
  opRow.innerHTML  =
    '<span style="color:#aaa;font-size:.7rem">Join:</span>' +
    '<label><input type="radio" name="within' + idx + '" value="AND" checked> AND</label>' +
    '<label><input type="radio" name="within' + idx + '" value="OR"> OR</label>';
  opRow.querySelectorAll('input').forEach(r =>
    r.addEventListener('change', e => { group.within = e.target.value; }));

  const right      = document.createElement('div');
  right.style.cssText = 'display:flex;align-items:center;gap:8px;flex-wrap:wrap;flex:1;min-width:240px';
  right.appendChild(tagsBox);
  right.appendChild(addBtn);
  right.appendChild(opRow);

  if (idx > 0) {
    const del     = document.createElement('button');
    del.className = 'btn btn-ghost btn-sm';
    del.textContent = '✕';
    del.onclick   = () => { kwGroups.splice(idx, 1, null); row.remove(); refreshPreview(); };
    right.appendChild(del);
  }

  row.appendChild(lbl);
  row.appendChild(right);
  container.appendChild(row);
}

function addTerm(group, tagsBox, inp) {
  const text = inp.value.trim();
  if (!text) return;
  const term = { text, op: group.within };
  group.terms.push(term);

  const tag = document.createElement('span');
  tag.className = 'tag ' + (group.within === 'AND' ? 'tag-and' : 'tag-or');
  tag.innerHTML = escHtml(text) + '<span class="tag-x">×</span>';
  tag.querySelector('.tag-x').onclick = () => {
    const i = group.terms.indexOf(term);
    if (i >= 0) group.terms.splice(i, 1);
    tag.remove();
    refreshPreview();
  };
  tagsBox.insertBefore(tag, inp);
  inp.value = '';
  refreshPreview();
}

function removeLast(group, tagsBox) {
  group.terms.pop();
  const tags = tagsBox.querySelectorAll('.tag');
  if (tags.length) tags[tags.length - 1].remove();
  refreshPreview();
}

// ── QUERY ────────────────────────────────────────────────────────
function flushKeywordInputs() {
  kwInputRefs.forEach(({ group, tagsBox, inp }) => addTerm(group, tagsBox, inp));
}

function buildQuery() {
  flushKeywordInputs();
  const activeJ = [...selected].filter(t => ACTIVE_RATINGS.has(ratingMap[t]));
  const groups  = kwGroups
    .filter(g => g && g.terms.length)
    .map(g => g.terms.map((t, i) => ({ text: t.text, op: i === 0 ? 'AND' : t.op })));
  const rowOp   = (document.querySelector('input[name="rowOp"]:checked') || {}).value || 'AND';

  const parts = groups.map(grp => {
    const ts = [];
    grp.forEach((t, i) => { if (i > 0) ts.push(t.op); ts.push(t.text.includes(' ') ? '"' + t.text + '"' : t.text); });
    return '( ' + ts.join(' ') + ' )';
  });

  let q = '';
  if (parts.length === 1)      q = 'TITLE-ABS-KEY ' + parts[0];
  else if (parts.length > 1)   q = 'TITLE-ABS-KEY ' + parts.join('\n' + rowOp + ' ');
  if (activeJ.length) {
    const src = 'SRCTITLE (\n  ' + activeJ.map(t => '"' + t + '"').join(' OR\n  ') + '\n)';
    q = q ? q + '\nAND ' + src : src;
  }
  return q;
}

function refreshPreview() {
  const q = buildQuery();
  document.getElementById('queryPreview').textContent = q || '(no keywords or journals selected)';
}
function copyQuery() {
  const q = buildQuery();
  if (!q) { toast('Build a query first'); return; }
  navigator.clipboard.writeText(q).then(() => toast('Query copied to clipboard'));
}
function openScopus() {
  const q = buildQuery();
  if (q) navigator.clipboard.writeText(q).then(() =>
    toast('Query copied — paste into Scopus Advanced Search'));
  setTimeout(() =>
    window.open('https://www-scopus-com.qub.idm.oclc.org/search/form.uri?display=advanced', '_blank'), 400);
}

// ── SEARCH ───────────────────────────────────────────────────────
async function doSearch() {
  const q = buildQuery();
  if (!q) { toast('Select journals or enter keywords first'); return; }

  const { exists } = await fetch('/api/config/key/exists').then(r => r.json());
  if (!exists) {
    toast('No API key — go to Settings to add one');
    switchTab('settings');
    return;
  }

  const btn = document.getElementById('btnSearch');
  btn.disabled = true; btn.textContent = '⏳ Searching…';
  logClear(); showLog(true);
  log('Sending query to Scopus API…\n' + q);

  try {
    const cfg  = await fetch('/api/config').then(r => r.json());
    const data = await fetch('/api/search', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ query: q, max_results: cfg.max_results })
    }).then(r => r.json());

    if (data.error) { log('Error: ' + data.error); toast('Search error — see log'); return; }
    results = data.results || [];
    log('Found ' + results.length + ' papers.');
    renderResults();
    switchTab('results');
    updateResBadge();
  } catch (err) {
    log('Network error: ' + err.message);
    toast('Search failed — is the server running?');
  } finally {
    btn.disabled = false; btn.textContent = '🔎 Search Scopus API';
  }
}

// ── CSV IMPORT ───────────────────────────────────────────────────
function importCSV() { document.getElementById('csvInput').click(); }

function handleCSV(event) {
  const file = event.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = e => {
    const text = e.target.result;
    const lines = text.split(/\r?\n/);
    if (!lines.length) return;

    function parseCSVLine(line) {
      const res = []; let cur = '', inq = false;
      for (let i = 0; i < line.length; i++) {
        const c = line[i];
        if (c === '"') { inq = !inq; }
        else if (c === ',' && !inq) { res.push(cur.trim()); cur = ''; }
        else { cur += c; }
      }
      res.push(cur.trim());
      return res;
    }

    const hdrs = parseCSVLine(lines[0]).map(h => h.toLowerCase().replace(/"/g,''));
    const ci   = h => hdrs.indexOf(h);
    const ti   = Math.max(ci('title'));
    const di   = ci('doi');
    const ai   = ci('authors');
    const yi   = ci('year');
    const si   = Math.max(ci('source title'), ci('source'));
    if (ti < 0) { toast('CSV missing "Title" column'); return; }

    results = [];
    for (let i = 1; i < lines.length; i++) {
      const row   = parseCSVLine(lines[i]);
      const title = (row[ti] || '').replace(/^"|"$/g,'').trim();
      if (!title) continue;
      const doi   = di >= 0 ? (row[di] || '').replace(/^"|"$/g,'').trim() : '';
      results.push({
        title,
        doi,
        authors: ai >= 0 ? (row[ai] || '').replace(/^"|"$/g,'').slice(0, 70) : '',
        year:    yi >= 0 ? (row[yi] || '').replace(/^"|"$/g,'').slice(0, 4)  : '',
        source:  si >= 0 ? (row[si] || '').replace(/^"|"$/g,'').trim()       : '',
        url:     doi ? 'https://doi.org/' + doi : '',
        status:  'pending',
      });
    }
    renderResults(); switchTab('results'); updateResBadge();
    toast('Imported ' + results.length + ' records from ' + file.name);
    event.target.value = '';
  };
  reader.readAsText(file, 'utf-8');
}

// ── RESULTS ──────────────────────────────────────────────────────
function renderResults() {
  const tbody  = document.getElementById('resTbody');
  const empty  = document.getElementById('emptyState');
  const table  = document.getElementById('resTable');

  if (!results.length) {
    empty.style.display = ''; table.style.display = 'none';
    document.getElementById('resCount').textContent = 'No results';
    document.getElementById('btnDl').disabled = true;
    return;
  }
  empty.style.display = 'none'; table.style.display = '';
  document.getElementById('resCount').innerHTML =
    '<strong>' + results.length + '</strong> paper' + (results.length !== 1 ? 's' : '');
  document.getElementById('btnDl').disabled = false;

  tbody.innerHTML = '';
  results.forEach((r, idx) => {
    r._rating = ratingMap[r.source] || '';
    const st  = r.status || 'pending';
    const sc  = st.includes('✓') || st === 'exists' ? 'ok'
              : st === 'no PDF' || st === 'fail'     ? 'fail'
              : st === 'pending'                     ? 'pend' : 'skip';
    const rc  = st.includes('✓') || st === 'exists' ? 'ok'
              : st === 'no PDF' || st === 'fail'     ? 'fail'
              : st.includes('skip') || st === 'exists' ? 'skip' : '';

    const doi = r.doi || '';
    const qubHref = doi ? 'https://doi-org.qub.idm.oclc.org/' + doi : '';
    const doiCell = doi
      ? '<a class="doi-link" href="https://doi.org/' + doi + '" target="_blank">' + doi + '</a>'
        + (qubHref ? '<a class="qub-link" href="' + qubHref + '" target="_blank">QUB</a>' : '')
      : '';

    const tr = document.createElement('tr');
    tr.dataset.idx = idx;
    if (rc) tr.className = rc;
    tr.innerHTML =
      '<td style="max-width:300px">' + escHtml(r.title || '') + '</td>' +
      '<td style="max-width:120px;color:#666">' + escHtml((r.authors||'').slice(0,55)) + '</td>' +
      '<td>' + escHtml(r.year || '') + '</td>' +
      '<td style="max-width:170px">' + escHtml(r.source || '') + '</td>' +
      '<td>' + (r._rating ? '<span class="jbadge ' + (BADGE[r._rating]||'b1') + '">' + r._rating + '</span>' : '') + '</td>' +
      '<td style="white-space:nowrap">' + doiCell + '</td>' +
      '<td><span class="st-badge st-' + sc + '">' + escHtml(st) + '</span></td>';
    tbody.appendChild(tr);
  });
}

function sortTable(col) {
  if (sortCol === col) sortAsc = !sortAsc; else { sortCol = col; sortAsc = true; }
  results.sort((a, b) => {
    const va = String(col === '_rating' ? (RATING_ORDER[a._rating] ?? 9) : (a[col] || '')).toLowerCase();
    const vb = String(col === '_rating' ? (RATING_ORDER[b._rating] ?? 9) : (b[col] || '')).toLowerCase();
    return sortAsc ? va.localeCompare(vb) : vb.localeCompare(va);
  });
  renderResults();
}

function updateResBadge() {
  const n = results.length;
  document.getElementById('res-tab-badge').textContent = n ? '(' + n + ')' : '';
}

// ── DOWNLOAD ─────────────────────────────────────────────────────
async function startDownload() {
  if (!results.length) return;
  const cfg = await fetch('/api/config').then(r => r.json());

  document.getElementById('btnDl').disabled    = true;
  document.getElementById('btnStop').disabled  = false;
  document.getElementById('btnZip').style.display = 'none';
  document.getElementById('progWrap').style.display = 'flex';
  document.getElementById('logCard').style.display  = '';
  logClear();
  log('Starting download of ' + results.length + ' papers…');
  if (cfg.dl_folder && !cfg.is_online)
    log('Saving to: ' + cfg.dl_folder);
  else
    log('Online mode: papers will be packaged as a ZIP for browser download.');

  try {
    const sr = await fetch('/api/download', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        papers:     results,
        folder:     cfg.is_online ? '' : (cfg.dl_folder || ''),
        try_oa:     cfg.try_oa    !== false,
        try_proxy:  cfg.try_proxy !== false,
      })
    });
    const sd = await sr.json();
    if (sd.error) {
      log('Error: ' + sd.error); toast('Download error: ' + sd.error);
      document.getElementById('btnDl').disabled   = false;
      document.getElementById('btnStop').disabled = true;
      return;
    }
  } catch (err) {
    log('Failed to start: ' + err.message);
    document.getElementById('btnDl').disabled   = false;
    document.getElementById('btnStop').disabled = true;
    return;
  }

  if (dlEs) dlEs.close();
  dlEs = new EventSource('/api/download/stream');

  dlEs.onmessage = e => {
    let msg;
    try { msg = JSON.parse(e.data); } catch(_) { return; }
    if (msg.type === 'ping') return;

    if (msg.type === 'progress') {
      document.getElementById('progBar').style.width = msg.pct + '%';
      document.getElementById('progLabel').textContent = msg.current + '/' + msg.total;

    } else if (msg.type === 'item') {
      const tr = document.querySelector('tr[data-idx="' + msg.idx + '"]');
      if (tr) {
        const badge = tr.querySelector('.st-badge');
        if (badge) {
          const sc = msg.tag === 'ok' ? 'ok' : msg.tag === 'fail' ? 'fail' : 'skip';
          badge.className = 'st-badge st-' + sc;
          badge.textContent = msg.status;
        }
        tr.className = msg.tag === 'ok' ? 'ok' : msg.tag === 'fail' ? 'fail' : msg.tag === 'skip' ? 'skip' : '';
      }
      if (results[msg.idx]) results[msg.idx].status = msg.status;
      const r   = results[msg.idx];
      const ttl = r ? r.title.slice(0, 65) : '';
      const suffix = msg.tag === 'fail' && msg.reason ? ' (' + msg.reason + ')' : '';
      log((msg.tag === 'ok' ? '[✓] ' : msg.tag === 'fail' ? '[✗] ' : '[=] ') + ttl + suffix);

    } else if (msg.type === 'done') {
      dlEs.close(); dlEs = null;
      document.getElementById('btnDl').disabled   = false;
      document.getElementById('btnStop').disabled = true;
      document.getElementById('progBar').style.width = '100%';
      document.getElementById('progLabel').textContent = 'Done';
      log('\n✓ ' + msg.ok + ' downloaded  ✗ ' + msg.fail + ' failed  = ' + msg.skip + ' skipped');
      toast('Done — ' + msg.ok + ' downloaded, ' + msg.fail + ' failed');
      if (msg.has_zip) {
        document.getElementById('btnZip').style.display = '';
        log('📦 Click "Save ZIP" to download all papers.');
      }
      if (msg.fail > 0) {
        const failedWithDoi = results.filter(r => r.status === 'no PDF' && r.doi).length;
        if (failedWithDoi > 0) {
          document.getElementById('btnQUB').style.display = '';
          log('🔗 ' + failedWithDoi + ' paper(s) could not be auto-downloaded. Click "Open failed in QUB" to open them in your browser via the QUB proxy (requires QUB login).');
        }
      }
    }
  };

  dlEs.onerror = () => {
    if (dlEs && dlEs.readyState === EventSource.CLOSED) {
      log('Stream closed.');
      document.getElementById('btnDl').disabled   = false;
      document.getElementById('btnStop').disabled = true;
    }
  };
}

async function stopDownload() {
  await fetch('/api/download/stop', { method: 'POST' });
  if (dlEs) { dlEs.close(); dlEs = null; }
  document.getElementById('btnDl').disabled   = false;
  document.getElementById('btnStop').disabled = true;
  log('Download stopped.');
}

function downloadZip() {
  window.location.href = '/api/download/get-zip';
}

function openFailedInQUB() {
  const QUB_PROXY = 'qub.idm.oclc.org';
  const failed = results.filter(r => r.status === 'no PDF' && r.doi);
  if (!failed.length) { toast('No failed papers with DOIs'); return; }
  const MAX = 10;
  const batch = failed.slice(0, MAX);
  if (failed.length > MAX)
    toast(`Opening first ${MAX} of ${failed.length} failed papers in QUB — repeat as needed`);
  else
    toast(`Opening ${batch.length} paper${batch.length > 1 ? 's' : ''} in QUB proxy tabs`);
  batch.forEach(r => {
    window.open('https://doi-org.' + QUB_PROXY + '/' + r.doi, '_blank');
  });
  // Mark opened papers so button cycles through remaining
  batch.forEach(r => { r._qub_opened = true; });
  results.filter(r => r._qub_opened).forEach(r => { delete r._qub_opened; });
}

function exportCSV() {
  if (!results.length) { toast('No results to export'); return; }
  const cols = ['title','authors','year','source','doi','url','status'];
  const rows = [cols.join(','),
    ...results.map(r => cols.map(c => '"' + String(r[c]||'').replace(/"/g,'""') + '"').join(','))];
  const a = document.createElement('a');
  a.href = URL.createObjectURL(new Blob([rows.join('\n')], {type:'text/csv'}));
  a.download = 'literature_results.csv';
  a.click();
}

async function openFolder() {
  if (isOnline) { toast('Folder access not available in online mode'); return; }
  await fetch('/api/open-folder', {
    method: 'POST', headers: {'Content-Type':'application/json'},
    body: JSON.stringify({})
  });
}

// ── SETTINGS ─────────────────────────────────────────────────────
async function loadSettings() {
  const cfg = await fetch('/api/config').then(r => r.json());
  isOnline   = cfg.is_online || false;

  document.getElementById('cfgFolder').value = cfg.dl_folder || '';
  document.getElementById('cfgMax').value    = cfg.max_results || 200;
  document.getElementById('cfgOA').checked   = cfg.try_oa    !== false;
  document.getElementById('cfgProxy').checked= cfg.try_proxy !== false;

  if (isOnline) {
    document.getElementById('modeBadge').textContent = 'Online';
    document.getElementById('modeBadge').classList.add('online');
    document.getElementById('onlineBanner').classList.add('show');
    document.getElementById('apiKeyRow').style.display  = 'none';
    document.getElementById('folderRow').style.display  = 'none';
    document.getElementById('btnFolder').style.display  = 'none';
  } else {
    const { exists } = await fetch('/api/config/key/exists').then(r => r.json());
    document.getElementById('cfgApiKey').placeholder =
      exists ? '(API key saved — enter to replace)' : 'Paste your Elsevier API key here';
  }

  // QUB cookie status
  const { exists: cookieExists } = await fetch('/api/config/qub-cookie/exists').then(r => r.json());
  _updateQubCookieUI(cookieExists);
}

function _updateQubCookieUI(exists) {
  document.getElementById('cfgQubCookie').placeholder =
    exists ? '(cookie saved — paste to replace)' : '(paste EZproxy cookie value here)';
  document.getElementById('qubCookieExistsMsg').style.display = exists ? '' : 'none';
  document.getElementById('btnClearQubCookie').style.display  = exists ? '' : 'none';
}

async function saveSettings() {
  const apiKey    = document.getElementById('cfgApiKey').value.trim();
  const qubCookie = document.getElementById('cfgQubCookie').value.trim();
  const folder    = document.getElementById('cfgFolder').value.trim();
  const max       = parseInt(document.getElementById('cfgMax').value) || 200;
  const tryOA     = document.getElementById('cfgOA').checked;
  const tryPrx    = document.getElementById('cfgProxy').checked;

  if (apiKey && !isOnline) {
    await fetch('/api/config/key', {
      method: 'POST', headers: {'Content-Type':'application/json'},
      body: JSON.stringify({api_key: apiKey})
    });
    document.getElementById('cfgApiKey').value       = '';
    document.getElementById('cfgApiKey').placeholder = '(API key saved — enter to replace)';
  }
  if (qubCookie) {
    await fetch('/api/config/qub-cookie', {
      method: 'POST', headers: {'Content-Type':'application/json'},
      body: JSON.stringify({cookie: qubCookie})
    });
    document.getElementById('cfgQubCookie').value = '';
    _updateQubCookieUI(true);
  }
  await fetch('/api/config', {
    method: 'POST', headers: {'Content-Type':'application/json'},
    body: JSON.stringify({dl_folder: folder, max_results: max, try_oa: tryOA, try_proxy: tryPrx})
  });
  toast('Settings saved');
}

function browseFolder() {
  const val = prompt('Enter the full folder path for downloads:',
    document.getElementById('cfgFolder').value);
  if (val !== null) document.getElementById('cfgFolder').value = val;
}

async function checkProxyCookies() {
  const box  = document.getElementById('cookieStatus');
  const icon = document.getElementById('cookieStatusIcon');
  const msg  = document.getElementById('cookieStatusMsg');
  const names= document.getElementById('cookieStatusNames');
  box.style.display = '';
  icon.textContent = '⏳'; msg.textContent = 'Checking…'; names.textContent = '';
  try {
    const d = await fetch('/api/proxy-cookies/check').then(r => r.json());
    const vi = d.v20_info || {};

    if (d.found) {
      box.style.background  = '#e8f5e9'; box.style.borderColor = '#a5d6a7';
      icon.textContent = '✅';
      const src = d.source === 'manual' ? 'manually saved cookie' : 'browser cookies';
      msg.textContent  = d.count + ' QUB proxy cookie(s) ready (' + src + ') — proxy downloads are enabled.';
      names.textContent = 'Cookies: ' + d.names.join(', ');
    } else if (vi.total && vi.all_v20) {
      // Chrome/Edge found with cookies, but all v20-encrypted (Chrome 127+)
      box.style.background  = '#fff3e0'; box.style.borderColor = '#ffb74d';
      icon.textContent = '🔒';
      msg.textContent  = vi.browser + ' has ' + vi.total + ' QUB cookies but they use Chrome 127+ App-Bound Encryption — cannot be read automatically.';
      names.textContent = 'Fix: paste the EZproxy cookie manually using the field below (F12 → Application → Cookies → qub.idm.oclc.org).';
    } else if (vi.total && vi.decryptable > 0) {
      // Some decryptable, some v20
      box.style.background  = '#fff8e1'; box.style.borderColor = '#ffe082';
      icon.textContent = '⚠️';
      msg.textContent  = vi.browser + ': ' + vi.decryptable + ' readable + ' + vi.v20 + ' v20-encrypted cookies.';
    } else {
      box.style.background  = '#fff8e1'; box.style.borderColor = '#ffe082';
      icon.textContent = '⚠️';
      msg.textContent  = 'No QUB proxy cookies found in browser. Log in to the QUB Library portal first, or paste the EZproxy cookie below.';
    }
  } catch(e) {
    icon.textContent = '❌'; msg.textContent = 'Error: ' + e.message;
  }
}

async function testQubCookie() {
  // Save any unsaved cookie value first
  const val = document.getElementById('cfgQubCookie').value.trim();
  if (val) {
    await fetch('/api/config/qub-cookie', {
      method: 'POST', headers: {'Content-Type':'application/json'},
      body: JSON.stringify({cookie: val})
    });
    document.getElementById('cfgQubCookie').value = '';
    _updateQubCookieUI(true);
  }

  const box  = document.getElementById('qubTestStatus');
  const icon = document.getElementById('qubTestIcon');
  const msg  = document.getElementById('qubTestMsg');
  box.style.display = '';
  box.style.background = '#f7f9fb'; box.style.borderColor = '#dde3ea';
  icon.textContent = '⏳'; msg.textContent = 'Testing QUB proxy connection…';

  try {
    const d = await fetch('/api/proxy-cookies/test').then(r => r.json());
    if (d.success) {
      box.style.background = '#e8f5e9'; box.style.borderColor = '#a5d6a7';
      icon.textContent = '✅';
      msg.textContent  = d.message + (d.url ? ' (' + d.url.slice(0,60) + '…)' : '');
    } else {
      box.style.background = '#fff8e1'; box.style.borderColor = '#ffe082';
      icon.textContent = '⚠️';
      msg.textContent  = d.message || ('Failed: ' + d.reason);
      if (d.reason === 'no_cookie') {
        msg.textContent += ' — paste the cookie value above and click Test again.';
      }
    }
  } catch(e) {
    icon.textContent = '❌'; msg.textContent = 'Error: ' + e.message;
  }
}

async function clearQubCookie() {
  await fetch('/api/config/qub-cookie', {
    method: 'POST', headers: {'Content-Type':'application/json'},
    body: JSON.stringify({cookie: ''})
  });
  _updateQubCookieUI(false);
  const box = document.getElementById('qubTestStatus');
  box.style.display = 'none';
  toast('QUB session cookie cleared');
}

// ── LOG ───────────────────────────────────────────────────────────
function log(msg) {
  const el = document.getElementById('logPre');
  el.textContent += msg + '\n';
  el.scrollTop = el.scrollHeight;
}
function logClear() { document.getElementById('logPre').textContent = ''; }
function showLog(v) { document.getElementById('logCard').style.display = v ? '' : 'none'; }

// ── UTILS ─────────────────────────────────────────────────────────
function escHtml(s) {
  return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
}
function toast(msg, dur=2800) {
  const t = document.getElementById('toast');
  t.textContent = '✓ ' + msg;
  t.classList.add('show');
  setTimeout(() => t.classList.remove('show'), dur);
}

// ── AUTO-CONFIGURE QUB COOKIE ──────────────────────────────────────
let _cgTimer = null;

async function autoConfigQubCookie() {
  const modal = document.getElementById('cookieGrabModal');
  modal.style.display = 'flex';
  document.getElementById('cgStatus').textContent = '⏳';
  document.getElementById('cgMsg').textContent = 'Launching Chrome…';

  let resp, data;
  try {
    resp = await fetch('/api/proxy-cookies/launch-browser', {method: 'POST'});
    data = await resp.json();
  } catch (e) {
    document.getElementById('cgMsg').textContent = 'Network error: ' + e;
    return;
  }
  if (data.error) {
    document.getElementById('cgStatus').textContent = '❌';
    document.getElementById('cgMsg').textContent = data.error;
    return;
  }

  document.getElementById('cgMsg').innerHTML =
    'A Chrome window has opened — log in to QUB.<br>' +
    '<span style="color:#888;font-size:.82rem">This dialog closes automatically once the cookie is detected.</span>';

  let tries = 0;
  _cgTimer = setInterval(async () => {
    if (++tries > 90) {
      clearInterval(_cgTimer);
      document.getElementById('cgStatus').textContent = '❌';
      document.getElementById('cgMsg').textContent = 'Timed out after 3 minutes. Please try again.';
      return;
    }
    try {
      const r = await fetch('/api/proxy-cookies/grab-from-browser');
      const d = await r.json();
      if (d.found) {
        clearInterval(_cgTimer);
        document.getElementById('cgStatus').textContent = '✅';
        document.getElementById('cgMsg').textContent = 'Cookie saved — QUB proxy downloads are enabled!';
        _updateQubCookieUI(true);
        document.getElementById('btnClearQubCookie').style.display = '';
        setTimeout(closeCookieGrabModal, 2000);
      }
    } catch (_) {}
  }, 2000);
}

function closeCookieGrabModal() {
  clearInterval(_cgTimer);
  document.getElementById('cookieGrabModal').style.display = 'none';
}
</script>

<!-- Auto-configure QUB cookie modal -->
<div id="cookieGrabModal" style="display:none;position:fixed;inset:0;background:rgba(0,0,0,.45);z-index:9999;align-items:center;justify-content:center">
  <div style="background:#fff;border-radius:10px;padding:28px 32px;max-width:460px;width:90%;box-shadow:0 8px 32px rgba(0,0,0,.2)">
    <h3 style="margin:0 0 12px;font-size:1.05rem">Auto-configure QUB Cookie</h3>
    <div id="cgStatus" style="font-size:2rem;text-align:center;padding:10px 0">⏳</div>
    <p id="cgMsg" style="color:#555;font-size:.88rem;margin:8px 0 18px;text-align:center;line-height:1.5">Launching Chrome…</p>
    <div style="text-align:center">
      <button class="btn btn-ghost btn-sm" onclick="closeCookieGrabModal()">Cancel</button>
    </div>
  </div>
</div>
</body>
</html>"""


# ── ENTRY POINT ───────────────────────────────────────────────────────────────
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    print("=" * 55)
    print("  Literature Auto-Downloader · QUB")
    print(f"  Open: http://127.0.0.1:{port}")
    print("=" * 55)
    if not IS_ONLINE:
        webbrowser.open(f"http://127.0.0.1:{port}")
    app.run(debug=False, host="0.0.0.0", port=port, threaded=True)
