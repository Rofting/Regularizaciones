"""Comprueba que este equipo tiene todo lo necesario para usar Regularizaciones.

Uso:
    python scripts/comprobar_instalacion.py

Lo ejecuta el instalador de Windows al terminar y sirve para diagnosticar un
equipo nuevo. Cada comprobación dice qué falta y cómo arreglarlo. Devuelve 0
si lo imprescindible está listo (Python, librerías, LibreOffice) y 1 si no.
Poppler sólo hace falta para leer PDF escaneados: su falta es un aviso.
"""

from __future__ import annotations

import importlib
import sys
from dataclasses import dataclass
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT / "core"))

# Nombre en requirements.txt -> módulo que se importa.
MODULES = (
    ("customtkinter", "customtkinter"),
    ("openpyxl", "openpyxl"),
    ("python-docx", "docx"),
    ("pdfplumber", "pdfplumber"),
    ("matplotlib", "matplotlib"),
    ("numpy", "numpy"),
    ("xlrd", "xlrd"),
    ("xlwt", "xlwt"),
    ("pdf2image", "pdf2image"),
    ("pytesseract", "pytesseract"),
    ("rapidocr", "rapidocr"),
    ("onnxruntime", "onnxruntime"),
    ("rapidfuzz", "rapidfuzz"),
    ("keyring", "keyring"),
)
FILES = (
    "core/app.py",
    "plantillas/Plantilla_Cartas.docx",
    "plantillas/modelo/modelo_acs_v1.xlsx",
    "config/proveedores.json",
    "config/palabras_clave.json",
)


@dataclass(frozen=True)
class Check:
    name: str
    ok: bool
    detail: str
    required: bool = True


def check_python() -> Check:
    version = ".".join(str(part) for part in sys.version_info[:3])
    ok = sys.version_info[:2] >= (3, 11)
    return Check("Python", ok, f"{version} en {sys.executable}" if ok else
                 f"{version}: se necesita Python 3.12 (ejecuta INSTALAR.bat)")


def check_tkinter() -> Check:
    try:
        import tkinter
        return Check("Ventanas (tkinter)", True, f"Tk {tkinter.TkVersion}")
    except Exception as error:  # Python instalado sin la opción tcl/tk
        return Check("Ventanas (tkinter)", False,
                     f"{error}. Reinstala Python 3.12 desde python.org con la opción «tcl/tk and IDLE».")


def check_modules() -> list[Check]:
    results = []
    for package, module in MODULES:
        try:
            imported = importlib.import_module(module)
            version = getattr(imported, "__version__", "")
            results.append(Check(f"Librería {package}", True, str(version) or "instalada"))
        except Exception as error:
            results.append(Check(f"Librería {package}", False,
                                 f"{error}. Vuelve a ejecutar INSTALAR.bat."))
    return results


def check_files() -> Check:
    missing = [item for item in FILES if not (PROJECT_ROOT / item).is_file()]
    if missing:
        return Check("Archivos del programa", False,
                     "Faltan " + ", ".join(missing) + ". Descarga de nuevo la carpeta completa.")
    return Check("Archivos del programa", True, str(PROJECT_ROOT))


def check_libreoffice() -> Check:
    from office_recalculation import (
        MINIMUM_LIBREOFFICE_VERSION, LibreOfficeRecalculator, RecalculationError, version_text,
    )
    recalculator = LibreOfficeRecalculator()
    try:
        executable = recalculator._find_executable()
        version = recalculator.check_minimum_version()
    except RecalculationError as error:
        return Check("LibreOffice (Excel oficial)", False, f"{error} Ejecuta INSTALAR.bat.")
    return Check("LibreOffice (Excel oficial)", True,
                 f"{version_text(version)} en {executable} (mínima {version_text(MINIMUM_LIBREOFFICE_VERSION)})")


def check_poppler() -> Check:
    import shutil
    try:
        from lector_pdf import _poppler_path
        folder = _poppler_path()
    except Exception as error:
        return Check("Poppler (PDF escaneados)", False, str(error), required=False)
    if folder or shutil.which("pdftoppm"):
        return Check("Poppler (PDF escaneados)", True, folder or "en el PATH", required=False)
    return Check("Poppler (PDF escaneados)", False,
                 "No encontrado: los PDF escaneados no se podrán leer por OCR. "
                 "Ejecuta INSTALAR.bat de nuevo.", required=False)


def check_writable() -> Check:
    target = PROJECT_ROOT / "data"
    try:
        target.mkdir(parents=True, exist_ok=True)
        probe = target / ".prueba_escritura"
        probe.write_text("ok", encoding="utf-8")
        probe.unlink()
    except OSError as error:
        return Check("Carpeta de datos", False,
                     f"No se puede escribir en {target}: {error}. "
                     "Coloca el programa en una carpeta propia, p. ej. C:\\Regularizaciones.")
    return Check("Carpeta de datos", True, str(target))


def run_checks() -> list[Check]:
    checks = [check_python(), check_tkinter(), *check_modules(), check_files(), check_writable()]
    checks.append(check_libreoffice())
    checks.append(check_poppler())
    return checks


def main() -> int:
    checks = run_checks()
    width = max(len(item.name) for item in checks)
    for item in checks:
        mark = "OK " if item.ok else ("FALTA" if item.required else "AVISO")
        print(f"[{mark:5}] {item.name:<{width}}  {item.detail}")
    failed = [item for item in checks if item.required and not item.ok]
    print()
    if failed:
        print(f"Faltan {len(failed)} componente(s) imprescindible(s). Revisa los mensajes de arriba.")
        return 1
    print("Todo listo. Abre el programa con el acceso directo «Regularizaciones».")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
