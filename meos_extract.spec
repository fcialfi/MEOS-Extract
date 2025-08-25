# meos_extract.spec
# Generate the executables with:  pyinstaller meos_extract.spec

from pathlib import Path
import sys
block_cipher = None
project_root = Path(__file__).resolve().parent if "__file__" in globals() else Path.cwd()

# --- CLI analysis/executable ---
a_cli = Analysis(
    ['Extract_all_charts.py'],
    pathex=[str(project_root)],
    binaries=[],
    datas=[],
    hiddenimports=['bs4', 'pandas', 'numpy', 'openpyxl'],  # add others if needed
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    cipher=block_cipher,
)

pyz_cli = PYZ(a_cli.pure, a_cli.zipped_data, cipher=block_cipher)

exe_cli = EXE(
    pyz_cli,
    a_cli.scripts,
    [],
    exclude_binaries=True,
    name='extract_cli',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=True,        # console app
)

# --- GUI analysis/executable ---
a_gui = Analysis(
    ['gui_app.py'],
    pathex=[str(project_root)],
    binaries=[],
    datas=[],
    hiddenimports=['bs4', 'pandas', 'numpy', 'openpyxl'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    cipher=block_cipher,
)

pyz_gui = PYZ(a_gui.pure, a_gui.zipped_data, cipher=block_cipher)

exe_gui = EXE(
    pyz_gui,
    a_gui.scripts,
    [],
    exclude_binaries=True,
    name='extract_gui',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,       # windowed app
)

# --- Collect both executables into one dist folder ---
coll = COLLECT(
    exe_cli,
    exe_gui,
    a_cli.binaries + a_gui.binaries,
    a_cli.zipfiles + a_gui.zipfiles,
    a_cli.datas + a_gui.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='MEOS-Extract'
)
