# -*- mode: python ; coding: utf-8 -*-
import os
import sys

from PyInstaller.utils.hooks import collect_all

# Real, previously-undetected bug (found 2026-08-17): with pathex=[], PyInstaller's
# Analysis can only see `xgen_seepage` if the *build-time* cwd happens to be the
# project root (Python prepends cwd to sys.path) - it silently is NOT visible when
# building from this packaging/ directory, which is how this spec is normally
# invoked. The build itself doesn't fail; the frozen exe just crashes at startup
# with "ModuleNotFoundError: No module named 'xgen_seepage.connector'". SPECPATH
# (a PyInstaller builtin, the absolute dir containing this .spec file) makes this
# independent of whatever directory the build is actually run from.
_PROJECT_ROOT = os.path.abspath(os.path.join(SPECPATH, '..'))

datas = []
binaries = []
hiddenimports = []
for pkg in ('keyring', 'uvicorn', 'starlette'):
    tmp_ret = collect_all(pkg)
    datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]

# Real, found-by-building bug (2026-08-17): on a conda Python, OpenSSL lives in
# <env>/Library/bin (libssl-3-x64.dll, libcrypto-3-x64.dll), a directory
# PyInstaller does NOT scan. The build succeeds but the frozen exe crashes on the
# very first HTTPS-touching import with "ImportError: DLL load failed while
# importing _ssl: The specified procedure could not be found" - _ssl.pyd loads
# some other (wrong-version) libcrypto from the system PATH instead. That means
# the connector can't reach the XGEN server at all. Bundle the correct DLLs
# explicitly. On a non-conda Python these globs simply find nothing (no-op).
_py_dir = os.path.dirname(sys.executable)
_ssl_dll_dirs = [
    os.path.join(_py_dir, 'Library', 'bin'),  # conda
    os.path.join(_py_dir, 'DLLs'),             # standard CPython
    _py_dir,
]
_ssl_dll_names = (
    'libssl-3-x64.dll', 'libcrypto-3-x64.dll',  # OpenSSL 3.x, 64-bit
    'libssl-3.dll', 'libcrypto-3.dll',
    'libssl-1_1-x64.dll', 'libcrypto-1_1-x64.dll',  # OpenSSL 1.1.x fallback
)
_seen_ssl = set()
for _d in _ssl_dll_dirs:
    for _n in _ssl_dll_names:
        _p = os.path.join(_d, _n)
        if _n not in _seen_ssl and os.path.exists(_p):
            binaries += [(_p, '.')]
            _seen_ssl.add(_n)

# uvicorn picks its event-loop/HTTP-protocol backends at runtime via
# importlib string dispatch (loop="auto"/http="auto"), which PyInstaller's
# static import analysis can't see - collect_all('uvicorn') above pulls in
# uvicorn's own package data but has historically still missed some of these
# dynamically-loaded submodules, so list the ones actually needed explicitly.
hiddenimports += [
    'uvicorn.loops.auto',
    'uvicorn.loops.asyncio',
    'uvicorn.protocols.http.auto',
    'uvicorn.protocols.http.h11_impl',
    'uvicorn.protocols.websockets.auto',
    'uvicorn.lifespan.on',
    'uvicorn.lifespan.off',
]

# libreoffice_adapter.py launches _uno_worker.py by file path
# (Path(__file__).with_name(...)), never by import, so PyInstaller's static
# analysis has no reason to bundle it - it must be listed explicitly or the
# frozen exe silently can't find it at runtime.
datas += [(os.path.join(_PROJECT_ROOT, 'xgen_seepage', '_uno_worker.py'), 'xgen_seepage')]

# taskpane_server.py serves this directory's contents by path
# (StaticFiles(directory=...)), never by import - same class of gap as
# _uno_worker.py above. Must be listed explicitly or it's silently missing
# from the frozen bundle and the taskpane has nothing to serve.
datas += [(os.path.join(_PROJECT_ROOT, 'xgen_seepage', 'taskpane'), 'xgen_seepage/taskpane')]


a = Analysis(
    ['run_connector.py'],
    pathex=[_PROJECT_ROOT],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='xgen-seepage-connector',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
