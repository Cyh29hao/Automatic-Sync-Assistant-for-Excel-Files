# -*- mode: python ; coding: utf-8 -*-
from pathlib import Path
import sys

project_dir = Path(SPECPATH)
python_base = Path(sys.base_prefix)
datas = [
    (str(project_dir / 'tasks.template.json'), '.'),
    (str(project_dir / 'README_RELEASE.md'), '.'),
]

stdlib_tkinter = python_base / 'Lib' / 'tkinter'
if stdlib_tkinter.exists():
    datas.append((str(stdlib_tkinter), 'tkinter'))
binaries = []

for folder_name, target in (
    ('tcl8.6', '_tcl_data'),
    ('tk8.6', '_tk_data'),
):
    folder = python_base / 'tcl' / folder_name
    if folder.exists():
        datas.append((str(folder), target))

for dll_name in ('_tkinter.pyd', 'tcl86t.dll', 'tk86t.dll'):
    dll_path = python_base / 'DLLs' / dll_name
    if dll_path.exists():
        binaries.append((str(dll_path), '.'))

a = Analysis(
    [str(project_dir / 'app.py')],
    pathex=[str(project_dir)],
    binaries=binaries,
    datas=datas,
    hiddenimports=['tkinter', '_tkinter'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['torch', 'matplotlib', 'IPython', 'jupyter', 'notebook'],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='Excel Sync Manager',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='Excel Sync Manager',
)
