"""Genera una copia limpia de la aplicación para instalarla en otro despacho.

La carpeta de trabajo de un despacho mezcla el producto (código, plantillas
genéricas, catálogo global de proveedores) con sus propios datos (base de
datos, fuentes, salidas, perfiles Excel de sus comunidades, CUPS, identidad de
cartas). Este script copia sólo lo primero y comprueba que en la copia no
queda ningún dato de comunidades o personas.

Uso:
    python scripts/exportar_producto.py DESTINO [--bd data/gestion.db]

Con ``--bd`` se comprueba además que no aparezcan nombres ni CIF de las
comunidades y propietarios registrados en esa base.

El despacho nuevo arranca con una base vacía: la aplicación la crea y pide sus
datos del despacho la primera vez.
"""

from __future__ import annotations

import argparse
import json
import re
import shutil
import sqlite3
import sys
import zipfile
from dataclasses import dataclass
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT / "core"))

from provider_registry import find_tax_ids  # noqa: E402


# Lo que forma el producto. Todo lo demás se considera del despacho.
PRODUCT_PATHS = (
    "core",
    "scripts",
    "config/proveedores.json",
    "config/palabras_clave.json",
    "config/proveedores_despacho.ejemplo.json",
    "config/rutas.ejemplo.json",
    "config/modelo_excel",
    "plantillas/Plantilla_Cartas.docx",
    "plantillas/modelo",
    "requirements.txt",
    "GUIA_PROCESAR_TODO.txt",
)

# Dentro de las rutas del producto, piezas propias de un despacho concreto.
EXCLUDED = (
    "__pycache__",
    "core/private_658_validation.py",
)

# Ejemplos sintéticos usados en textos de ayuda; no son datos de nadie.
SYNTHETIC_TAX_IDS = frozenset({"H12345674", "B12345674", "B00000000", "B50000009"})
CUPS_PATTERN = re.compile(r"\bES\d{16}[A-Z0-9]{2}(?:\d[A-Z])?\b")
TEXT_SUFFIXES = {".py", ".json", ".txt", ".md", ".csv", ".ps1", ".spec"}


@dataclass(frozen=True)
class Leak:
    path: str
    kind: str
    value: str


def _excluded(relative: str) -> bool:
    return any(part == "__pycache__" for part in Path(relative).parts) or any(
        relative == item or relative.startswith(item + "/") for item in EXCLUDED
    )


def copy_product(source_root: Path, destination: Path) -> list[str]:
    copied: list[str] = []
    for item in PRODUCT_PATHS:
        source = source_root / item
        if not source.exists():
            continue
        files = [source] if source.is_file() else sorted(p for p in source.rglob("*") if p.is_file())
        for path in files:
            relative = path.relative_to(source_root).as_posix()
            if _excluded(relative):
                continue
            target = destination / relative
            target.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(path, target)
            copied.append(relative)
    return copied


def _document_text(path: Path) -> str:
    suffix = path.suffix.lower()
    if suffix in TEXT_SUFFIXES:
        return path.read_text(encoding="utf-8", errors="replace")
    if suffix in {".docx", ".xlsx"}:
        parts = []
        with zipfile.ZipFile(path) as archive:
            for name in archive.namelist():
                if name.endswith(".xml"):
                    xml = archive.read(name).decode("utf-8", errors="replace")
                    parts.append(re.sub(r"<[^>]+>", " ", xml))
        return " ".join(parts)
    return ""


def _known_private_values(database: Path | None) -> dict[str, set[str]]:
    values: dict[str, set[str]] = {"nombre de comunidad": set(), "propietario": set(), "CIF de comunidad": set()}
    if database is None or not database.is_file():
        return values
    connection = sqlite3.connect(f"file:{database}?mode=ro", uri=True)
    try:
        for name, cif in connection.execute("SELECT nombre, cif FROM comunidades"):
            if name and len(name.strip()) >= 6:
                values["nombre de comunidad"].add(name.strip().upper())
            if cif:
                values["CIF de comunidad"].add(re.sub(r"[^A-Z0-9]", "", cif.upper()))
        for (name,) in connection.execute("SELECT nombre_propietario FROM propietarios"):
            if name and len(name.strip()) >= 8 and " " in name.strip():
                values["propietario"].add(name.strip().upper())
    finally:
        connection.close()
    return values


def scan_for_leaks(root: Path, database: Path | None = None) -> list[Leak]:
    private = _known_private_values(database)
    leaks: list[Leak] = []
    for path in sorted(p for p in root.rglob("*") if p.is_file()):
        relative = path.relative_to(root).as_posix()
        text = _document_text(path)
        if not text:
            continue
        upper = text.upper()
        for tax_id in find_tax_ids(text):
            if tax_id in SYNTHETIC_TAX_IDS:
                continue
            if tax_id[0] == "H":
                leaks.append(Leak(relative, "CIF de comunidad", tax_id))
            elif tax_id[0].isdigit() or tax_id[0] in "XYZ":
                leaks.append(Leak(relative, "NIF/NIE de persona", tax_id))
        for cups in set(CUPS_PATTERN.findall(upper)):
            if not re.fullmatch(r"ES0{16}[A-Z0-9]*", cups):
                leaks.append(Leak(relative, "CUPS", cups))
        for kind, values in private.items():
            for value in values:
                if value in upper:
                    leaks.append(Leak(relative, kind, value))
    return leaks


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("destino", type=Path, help="Carpeta nueva donde crear la copia")
    parser.add_argument("--bd", type=Path, default=None,
                        help="Base del despacho actual para buscar sus nombres y CIF en la copia")
    args = parser.parse_args(argv)
    destination = args.destino.resolve()
    if destination.exists() and any(destination.iterdir()):
        print(f"La carpeta {destination} no está vacía; elige una nueva.", file=sys.stderr)
        return 2
    try:
        destination.relative_to(PROJECT_ROOT)
    except ValueError:
        pass
    else:
        print("El destino no puede estar dentro del proyecto.", file=sys.stderr)
        return 2
    copied = copy_product(PROJECT_ROOT, destination)
    leaks = scan_for_leaks(destination, args.bd)
    print(f"Copiados {len(copied)} archivos del producto en {destination}")
    if leaks:
        print("\nATENCIÓN: la copia contiene datos que parecen de un despacho:", file=sys.stderr)
        for leak in leaks:
            print(f"  - {leak.path}: {leak.kind} {leak.value}", file=sys.stderr)
        print("\nRevísalos antes de entregar la copia.", file=sys.stderr)
        return 1
    manifest = {"archivos": copied}
    (destination / "PRODUCTO.json").write_text(
        json.dumps(manifest, ensure_ascii=False, indent=2) + "\n", encoding="utf-8",
    )
    print("Sin datos de comunidades ni personas. Lista para instalar en otro despacho.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
