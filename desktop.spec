# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec file for Micro Automation desktop app.
Build with: pyinstaller desktop.spec
"""

import os

block_cipher = None

# Collect all hidden imports for uvicorn/fastapi
hidden_imports = [
    # FastAPI / Starlette / uvicorn
    "uvicorn",
    "uvicorn.logging",
    "uvicorn.lifespan",
    "uvicorn.lifespan.on",
    "uvicorn.protocols",
    "uvicorn.protocols.http",
    "uvicorn.protocols.http.auto",
    "uvicorn.protocols.http.h11_impl",
    "uvicorn.protocols.http.httptools_impl",
    "uvicorn.protocols.websockets",
    "uvicorn.protocols.websockets.auto",
    "uvicorn.loops",
    "uvicorn.loops.auto",
    "uvicorn.loops.asyncio",
    "fastapi",
    "fastapi.middleware",
    "fastapi.middleware.cors",
    "starlette",
    "starlette.staticfiles",
    "starlette.responses",
    "starlette.routing",
    "multipart",
    "multipart.multipart",
    # App dependencies
    "pyembroidery",
    "openpyxl",
    "sqlite3",
    "json",
    # System tray
    "pystray",
    "PIL",
    "PIL.Image",
    "PIL.ImageDraw",
    # API modules
    "api",
    "api.server",
    "api.database",
]

a = Analysis(
    ["desktop_main.py"],
    pathex=[],
    binaries=[],
    datas=[
        ("static_web", "static_web"),  # Pre-built Next.js static site
        ("api", "api"),                 # Python API module
    ],
    hiddenimports=hidden_imports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        "customtkinter",  # Legacy desktop UI — not needed
        "pywinauto",      # Windows automation — not needed for web app
        "tkinter",
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="MicroAutomation",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,  # Windowed mode — no console window
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon="web/public/favicon.ico" if os.path.exists("web/public/favicon.ico") else None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name="MicroAutomation",
)
