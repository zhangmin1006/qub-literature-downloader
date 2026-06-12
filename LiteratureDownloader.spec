# -*- mode: python ; coding: utf-8 -*-
# PyInstaller spec for Literature Auto-Downloader
# Build:  pyinstaller LiteratureDownloader.spec
#    or:  powershell -File build.ps1

from PyInstaller.utils.hooks import collect_data_files, collect_submodules

# Bundle openpyxl's internal template/style data (required for reading xlsx)
openpyxl_datas = collect_data_files("openpyxl")

a = Analysis(
    ["web_app.py"],
    pathex=[],
    binaries=[],
    datas=[
        ("ABS.xlsx", "."),          # journal list — read at runtime
        *openpyxl_datas,            # openpyxl templates/styles
    ],
    hiddenimports=[
        # Flask stack
        "flask", "flask.json",
        "werkzeug", "werkzeug.serving", "werkzeug.exceptions",
        "werkzeug.middleware.shared_data",
        "jinja2", "jinja2.ext",
        "click", "click.exceptions",
        # HTTP
        "requests", "requests.adapters", "requests.auth",
        "urllib3", "urllib3.util", "urllib3.util.retry",
        "certifi",
        "charset_normalizer",
        "idna",
        # Windows-specific (for Chrome cookie extraction)
        "winreg",
        "ctypes", "ctypes.wintypes",
        # Stdlib that PyInstaller sometimes misses
        "sqlite3", "csv", "zipfile", "struct", "base64",
        "email.mime.text",
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        # Trim packages that are definitely not used
        "tkinter", "matplotlib", "numpy", "pandas", "scipy",
        "PIL", "cv2", "PyQt5", "PySide2",
        "IPython", "jupyter", "notebook",
        "test", "unittest",
        "gunicorn",                 # deployment-only; not needed in the exe
    ],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="LiteratureDownloader",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,              # UPX can trigger AV false positives; leave off
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,           # show server window so user can see status / stop it
    icon=None,              # add an .ico file path here to set a custom icon
    version=None,
)
