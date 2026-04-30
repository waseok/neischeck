# -*- mode: python ; coding: utf-8 -*-

from pathlib import Path

block_cipher = None
project_root = Path(__file__).resolve().parents[1]

datas = [
    (str(project_root / "config" / "settings.json"), "config"),
    (str(project_root / "config" / "forbidden_rules.json"), "config"),
    (str(project_root / "config" / "suggestion_rules.json"), "config"),
    (str(project_root / "config" / "allowlist.json"), "config"),
    (str(project_root / "config" / "category_rules.json"), "config"),
]

a = Analysis(
    [str(project_root / "main.py")],
    pathex=[str(project_root)],
    binaries=[],
    datas=datas,
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)
exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name="생기부특기사항자동검토기",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
)
