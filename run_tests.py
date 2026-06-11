"""Comprehensive test suite for web_app.py"""
import urllib.request, urllib.error, json, sys, os, copy, tempfile
sys.stdout.reconfigure(encoding='utf-8')

BASE = 'http://127.0.0.1:5000'
PASS_N = 0; FAIL_N = 0; ISSUES = []

def ok(label):
    global PASS_N; PASS_N += 1
    print(f'  PASS  {label}')

def fail(label, detail=''):
    global FAIL_N; FAIL_N += 1
    msg = f'  FAIL  {label}' + (f': {detail}' if detail else '')
    print(msg)
    ISSUES.append(msg)

def get(path, timeout=8):
    with urllib.request.urlopen(BASE + path, timeout=timeout) as r:
        return r.status, json.loads(r.read())

def post(path, data, timeout=30):
    req = urllib.request.Request(
        BASE + path,
        data=json.dumps(data).encode(),
        headers={'Content-Type': 'application/json'})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return r.status, json.loads(r.read())

def stream_events(timeout=90):
    events = []
    req = urllib.request.Request(BASE + '/api/download/stream')
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        for raw in resp:
            line = raw.decode('utf-8').strip()
            if line.startswith('data: '):
                msg = json.loads(line[6:])
                if msg.get('type') != 'ping':
                    events.append(msg)
                if msg.get('type') == 'done':
                    break
    return events

CFG_PATH = os.path.join(os.path.dirname(__file__), '.lit_web_config.json')

# ── SECTION 1: Basic routes ──────────────────────────────────
print('=== SECTION 1: Basic routes ===')

try:
    with urllib.request.urlopen(BASE + '/', timeout=5) as r:
        html = r.read().decode('utf-8')
    assert 'Literature Auto-Downloader' in html and len(html) > 10000
    ok(f'GET /  returns full page ({len(html):,} chars)')
except Exception as e:
    fail('GET /', str(e))

try:
    s, d = get('/api/config/key/exists')
    assert s == 200 and d['exists'] is True
    ok('GET /api/config/key/exists  ->  exists=True')
except Exception as e:
    fail('/api/config/key/exists', str(e))

try:
    s, d = get('/api/config')
    assert all(k in d for k in ['dl_folder', 'max_results', 'try_oa', 'try_proxy'])
    assert 'api_key' not in d, 'api_key must NOT be sent to frontend'
    ok(f'GET /api/config  ->  keys={sorted(d.keys())}')
except Exception as e:
    fail('/api/config', str(e))

try:
    s, d = get('/api/journals')
    fields = d['fields']
    total = sum(len(f['journals']) for f in fields)
    ratings = set(j['rating'] for f in fields for j in f['journals'])
    assert len(fields) == 22, f'Expected 22 fields, got {len(fields)}'
    assert total == 1822, f'Expected 1822 journals, got {total}'
    assert ratings == {'1', '2', '3', '4', '4*'}, f'Ratings: {ratings}'
    ok(f'GET /api/journals  ->  22 fields, 1822 journals, all 5 ratings present')
except Exception as e:
    fail('/api/journals', str(e))

# ── SECTION 2: Query builder ─────────────────────────────────
print()
print('=== SECTION 2: Query builder ===')

qtests = [
    ('AND within group',
     {'journals': ['Management Science'],
      'groups': [[{'text': 'supply chain', 'op': 'AND'}, {'text': 'resilience', 'op': 'AND'}]],
      'row_op': 'AND'},
     lambda q: 'TITLE-ABS-KEY' in q and 'AND resilience' in q and '"Management Science"' in q),

    ('OR within group',
     {'journals': [], 'row_op': 'AND',
      'groups': [[{'text': 'machine learning', 'op': 'AND'}, {'text': 'deep learning', 'op': 'OR'}]]},
     lambda q: 'OR' in q and 'machine learning' in q),

    ('Two groups joined by AND',
     {'journals': [], 'row_op': 'AND',
      'groups': [[{'text': 'innovation', 'op': 'AND'}], [{'text': 'performance', 'op': 'AND'}]]},
     lambda q: q.count('(') >= 2),

    ('Two groups joined by OR',
     {'journals': [], 'row_op': 'OR',
      'groups': [[{'text': 'innovation', 'op': 'AND'}], [{'text': 'performance', 'op': 'AND'}]]},
     lambda q: 'OR' in q),

    ('Empty query',
     {'journals': [], 'groups': [], 'row_op': 'AND'},
     lambda q: q == ''),

    ('Keywords only - no SRCTITLE',
     {'journals': [], 'row_op': 'AND',
      'groups': [[{'text': 'blockchain', 'op': 'AND'}]]},
     lambda q: 'TITLE-ABS-KEY' in q and 'SRCTITLE' not in q),

    ('Multi-word term is quoted',
     {'journals': [], 'row_op': 'AND',
      'groups': [[{'text': 'digital transformation', 'op': 'AND'}]]},
     lambda q: '"digital transformation"' in q),

    ('Single word is NOT quoted',
     {'journals': [], 'row_op': 'AND',
      'groups': [[{'text': 'sustainability', 'op': 'AND'}]]},
     lambda q: 'sustainability' in q and '"sustainability"' not in q),

    ('Journals only - no TITLE-ABS-KEY',
     {'journals': ['Nature', 'Science'], 'groups': [], 'row_op': 'AND'},
     lambda q: 'SRCTITLE' in q and 'TITLE-ABS-KEY' not in q),

    ('Large journal list in SRCTITLE',
     {'journals': ['Journal A', 'Journal B', 'Journal C', 'Journal D', 'Journal E'],
      'groups': [[{'text': 'test', 'op': 'AND'}]], 'row_op': 'AND'},
     lambda q: q.count('"Journal') == 5),
]

for label, body, check in qtests:
    try:
        _, r = post('/api/query', body)
        q = r['query']
        assert check(q), f'Check failed on: {repr(q[:120])}'
        ok(label)
    except Exception as e:
        fail(label, str(e))

# ── SECTION 3: Scopus search ─────────────────────────────────
print()
print('=== SECTION 3: Scopus search ===')

try:
    _, r = post('/api/search', {
        'query': 'TITLE-ABS-KEY("supply chain resilience") AND SRCTITLE("Journal of Operations Management")',
        'max_results': 5
    }, timeout=40)
    assert 'error' not in r, r.get('error')
    assert r['count'] > 0
    p = r['results'][0]
    required = ['title', 'doi', 'authors', 'year', 'source', 'url', 'status']
    missing = [k for k in required if k not in p]
    assert not missing, f'Missing keys: {missing}'
    assert 'abstract' not in p, 'Stale abstract field present'
    ok(f'Search returns {r["count"]} results with correct schema')
    ok(f'First result: "{p["title"][:55]}" ({p["year"]})')
except Exception as e:
    fail('Scopus basic search', str(e))

try:
    _, r = post('/api/search', {
        'query': 'TITLE-ABS-KEY(innovation)', 'max_results': 3
    }, timeout=40)
    assert r['count'] <= 3
    ok(f'max_results=3 honoured (got {r["count"]})')
except Exception as e:
    fail('max_results', str(e))

try:
    orig = json.loads(open(CFG_PATH).read())
    tmp = copy.deepcopy(orig)
    tmp['api_key'] = ''
    open(CFG_PATH, 'w').write(json.dumps(tmp))
    try:
        post('/api/search', {'query': 'test', 'max_results': 1})
        fail('Empty API key should return 400')
    except urllib.error.HTTPError as e:
        assert e.code == 400
        ok('Empty API key  ->  HTTP 400')
    finally:
        open(CFG_PATH, 'w').write(json.dumps(orig))
except Exception as e:
    fail('Empty API key test', str(e))

# ── SECTION 4: Download pipeline ─────────────────────────────
print()
print('=== SECTION 4: Download pipeline ===')

folder = tempfile.mkdtemp()

# OA paper — PLOS Medicine (fully open-access journal, reliable PDF)
OA_DOI   = '10.1371/journal.pmed.0020124'
OA_TITLE = 'Why Most Published Research Findings Are False'
try:
    _, r = post('/api/download', {
        'papers': [{'title': OA_TITLE, 'doi': OA_DOI,
                    'authors': 'Vaswani et al.', 'year': '2017', 'source': 'arXiv',
                    'url': '', 'status': 'pending'}],
        'folder': folder, 'try_oa': True, 'try_proxy': False
    })
    assert r.get('started')
    events = stream_events(timeout=90)
    done = next(m for m in events if m['type'] == 'done')
    item = next(m for m in events if m['type'] == 'item')
    pdfs = [f for f in os.listdir(folder) if f.endswith('.pdf')]
    assert done['ok'] == 1, f'expected ok=1 got {done}'
    assert done['fail'] == 0
    assert len(pdfs) == 1
    size = os.path.getsize(os.path.join(folder, pdfs[0]))
    assert size > 10000, f'PDF too small: {size} bytes'
    assert item['status'].startswith('OA-repo'), f'Status: {item["status"]}'
    ok(f'OA-repo download: {pdfs[0]!r} ({size:,} bytes)')
except Exception as e:
    fail('OA-repo download', str(e))

# Skip already-existing file
try:
    _, r = post('/api/download', {
        'papers': [{'title': OA_TITLE, 'doi': OA_DOI,
                    'authors': 'Vaswani et al.', 'year': '2017', 'source': 'arXiv',
                    'url': '', 'status': 'pending'}],
        'folder': folder, 'try_oa': True, 'try_proxy': False
    })
    events = stream_events(timeout=15)
    done = next(m for m in events if m['type'] == 'done')
    assert done['skip'] == 1 and done['ok'] == 0
    ok('Already-existing file is skipped (not re-downloaded)')
except Exception as e:
    fail('Skip existing', str(e))

# No-DOI paper -> fail gracefully
try:
    folder2 = tempfile.mkdtemp()
    _, r = post('/api/download', {
        'papers': [{'title': 'No DOI paper', 'doi': '',
                    'authors': 'B', 'year': '2020', 'source': 'Test',
                    'url': '', 'status': 'pending'}],
        'folder': folder2, 'try_oa': True, 'try_proxy': False
    })
    events = stream_events(timeout=15)
    done = next(m for m in events if m['type'] == 'done')
    item = next(m for m in events if m['type'] == 'item')
    assert done['fail'] == 1 and item['status'] == 'no PDF'
    assert item['proxy_url'] == ''
    ok('No-DOI paper fails with status "no PDF" and empty proxy_url')
except Exception as e:
    fail('No-DOI paper', str(e))

# Invalid DOI (not in Unpaywall) -> fails gracefully
try:
    folder3 = tempfile.mkdtemp()
    _, r = post('/api/download', {
        'papers': [{'title': 'Bad DOI paper', 'doi': '10.9999/notreal.12345',
                    'authors': 'C', 'year': '2020', 'source': 'X',
                    'url': '', 'status': 'pending'}],
        'folder': folder3, 'try_oa': True, 'try_proxy': False
    })
    events = stream_events(timeout=20)
    done = next(m for m in events if m['type'] == 'done')
    assert done['fail'] == 1
    ok('Non-existent DOI fails gracefully')
except Exception as e:
    fail('Non-existent DOI', str(e))

# Empty paper list
try:
    folder4 = tempfile.mkdtemp()
    _, r = post('/api/download', {'papers': [], 'folder': folder4})
    events = stream_events(timeout=10)
    done = next(m for m in events if m['type'] == 'done')
    assert done['ok'] == 0 and done['fail'] == 0 and done['skip'] == 0
    ok('Empty paper list completes cleanly')
except Exception as e:
    fail('Empty paper list', str(e))

# Empty folder -> auto-uses temp dir, succeeds
try:
    _, r = post('/api/download', {
        'papers': [{'title': 'x', 'doi': '', 'authors': '',
                    'year': '', 'source': '', 'url': '', 'status': ''}],
        'folder': ''
    })
    assert r.get('started'), f'Expected started=True, got {r}'
    events = stream_events(timeout=15)
    done = next(m for m in events if m['type'] == 'done')
    assert done['fail'] == 1
    ok('Empty folder: uses temp dir, download completes cleanly')
except Exception as e:
    fail('Empty folder test', str(e))

# Summary CSV created and valid
try:
    csv_path = os.path.join(folder, '_download_summary.csv')
    assert os.path.exists(csv_path), 'CSV not created'
    with open(csv_path, encoding='utf-8-sig') as f:
        lines = f.readlines()
    assert len(lines) >= 2, 'CSV has no data rows'
    header = lines[0].lower()
    assert 'title' in header and 'doi' in header and 'status' in header
    ok(f'Summary CSV: {len(lines)-1} data row(s), correct columns')
except Exception as e:
    fail('Summary CSV', str(e))

# ── SECTION 5: Config persistence ────────────────────────────
print()
print('=== SECTION 5: Config persistence ===')

try:
    test_folder = tempfile.mkdtemp()
    _, r = post('/api/config', {'dl_folder': test_folder, 'max_results': 77, 'try_oa': False, 'try_proxy': True})
    assert r.get('saved')
    _, d = get('/api/config')
    assert d['dl_folder'] == test_folder
    assert d['max_results'] == 77
    assert d['try_oa'] == False
    assert d['try_proxy'] == True
    post('/api/config', {'dl_folder': str(os.path.expanduser('~/Downloads')), 'max_results': 200, 'try_oa': True, 'try_proxy': True})
    ok('Config round-trip: dl_folder, max_results, try_oa, try_proxy')
except Exception as e:
    fail('Config persistence', str(e))

try:
    _, r = post('/api/config/key', {'api_key': 'temp_test_key'})
    assert r.get('saved')
    _, d = get('/api/config/key/exists')
    assert d['exists']
    post('/api/config/key', {'api_key': '660d6aec95c672b9fd51fa04d4453bb1'})
    _, d2 = get('/api/config/key/exists')
    assert d2['exists']
    ok('API key save + restore works')
except Exception as e:
    fail('API key save/restore', str(e))

# ── SECTION 6: open-folder endpoint ──────────────────────────
print()
print('=== SECTION 6: Misc endpoints ===')

try:
    _, r = post('/api/open-folder', {'folder': folder})
    assert r.get('ok')
    ok('/api/open-folder returns ok for valid folder')
except Exception as e:
    fail('/api/open-folder', str(e))

# ── SUMMARY ───────────────────────────────────────────────────
print()
print(f'{"="*50}')
print(f'  PASSED: {PASS_N}   FAILED: {FAIL_N}')
if ISSUES:
    print()
    print('  Failures:')
    for i in ISSUES:
        print(i)
print(f'{"="*50}')
sys.exit(0 if FAIL_N == 0 else 1)
