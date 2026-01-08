# -*- mode: python ; coding: utf-8 -*-
import sys
from pathlib import Path
from PyInstaller.building.build_main import Analysis, PYZ, EXE, COLLECT


block_cipher = None

# Fallback so running `python polyboard_production_gui.spec` does not crash
spec_path = Path(__file__).resolve() if "__file__" in globals() else Path(sys.argv[0]).resolve()
base_path = spec_path.parent
project_root = base_path.parent

# Data files that the GUI expects at runtime
datas = []


def add_data(src: Path, dest: str):
    if src.exists():
        datas.append((str(src), dest))


def add_dir(dir_path: Path, dest_root: str, patterns=("*",)):
    if not dir_path.exists():
        return
    for pat in patterns:
        for p in dir_path.rglob(pat):
            if p.is_file():
                rel = p.relative_to(dir_path)
                datas.append((str(p), str(Path(dest_root) / rel)))


add_data(base_path / "Polyboard_convention.json", "Polyboard_Infra_PP")
add_data(base_path / "Polyboard_convention.xlsx", "Polyboard_Infra_PP")
add_data(base_path / "mpr_format_reference.json", "Polyboard_Infra_PP")
add_data(base_path / "Polyboard_Convention_Column_Summary.md", "Polyboard_Infra_PP")
add_data(base_path / "woodwop-mpr4x-format-pdf-free.pdf", "Polyboard_Infra_PP")

edge_dir = base_path / "Edge_Diagram_Ref"
add_dir(edge_dir, "Polyboard_Infra_PP/Edge_Diagram_Ref")

hiddenimports = ["PIL._tkinter_finder"]

a = Analysis(
    ["polyboard_production_gui.py"],
    pathex=[str(base_path)],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
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
    name="PolyboardProduction",
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
    icon=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name="PolyboardProduction",
)

