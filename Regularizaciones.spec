"""Distribución Windows onedir sin datos de ningún despacho."""

import os
import sys
from pathlib import Path

from PyInstaller.utils.hooks import collect_data_files, collect_dynamic_libs, collect_submodules

ROOT = Path(SPECPATH).resolve()
sys.path.insert(0, str(ROOT / "core"))
from app_paths import PUBLIC_RESOURCES

datas = []
for relative in PUBLIC_RESOURCES:
    source = ROOT / relative
    files = source.rglob("*") if source.is_dir() else (source,)
    for item in files:
        if item.is_file():
            datas.append((str(item), str(item.parent.relative_to(ROOT))))
datas += collect_data_files("customtkinter")
datas += collect_data_files("rapidocr")

binaries = collect_dynamic_libs("onnxruntime")
poppler = os.environ.get("POPPLER_BIN_DIR")
if poppler:
    folder = Path(poppler)
    if not (folder / "pdftoppm.exe").is_file():
        raise ValueError("POPPLER_BIN_DIR debe contener pdftoppm.exe")
    binaries += [(str(item), "poppler/bin") for item in folder.iterdir() if item.is_file()]

hiddenimports = [
    item.stem for item in (ROOT / "core").glob("*.py")
    if item.stem not in {"app", "private_658_validation"}
]
hiddenimports += collect_submodules("rapidocr")
hiddenimports += collect_submodules("keyring.backends")

a = Analysis(
    [str(ROOT / "core" / "app.py")],
    pathex=[str(ROOT / "core")],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=["private_658_validation", "pytest"],
    noarchive=False,
)
pyz = PYZ(a.pure)
exe = EXE(pyz, a.scripts, [], exclude_binaries=True,
          name="Regularizaciones", console=False, debug=False, strip=False, upx=False)
coll = COLLECT(exe, a.binaries, a.datas, strip=False, upx=False,
               name="Regularizaciones")
