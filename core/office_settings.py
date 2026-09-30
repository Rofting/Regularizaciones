"""Datos del despacho que usa esta instalación.

Cada despacho tiene su nombre, CIF, ciudad, firma, logo y calendario de
ejercicio. Se guardan en la propia base de datos (una fila), de modo que la
aplicación se instala igual en cualquier despacho y la identidad no viaja en
el código ni en archivos que se distribuyen con el producto.
"""

from __future__ import annotations

import json
import shutil
import sqlite3
from collections import Counter
from dataclasses import asdict, dataclass, fields, replace
from pathlib import Path

from provider_registry import normalize_tax_id, valid_spanish_tax_id


LOGO_RELATIVE_DIR = Path("data") / "despacho"


@dataclass(frozen=True)
class OfficeSettings:
    name: str = ""
    tax_id: str = ""
    city: str = ""
    address: str = ""
    phone: str = ""
    email: str = ""
    signature: str = ""
    footer: str = ""
    logo_path: str = ""  # relativa al proyecto
    fiscal_start_month: int = 1

    @property
    def configured(self) -> bool:
        return bool(self.name.strip())

    @property
    def letter_signature(self) -> str:
        return self.signature.strip() or self.name.strip() or "La Administración"

    @property
    def letter_footer(self) -> str:
        if self.footer.strip():
            return self.footer.strip()
        parts = [self.address, self.phone, self.email]
        return " · ".join(part.strip() for part in parts if part and part.strip())


def _ensure_table(connection: sqlite3.Connection) -> None:
    connection.execute(
        """CREATE TABLE IF NOT EXISTS office_settings (
            id INTEGER PRIMARY KEY CHECK (id = 1),
            settings_json TEXT NOT NULL,
            updated_at TEXT NOT NULL DEFAULT (datetime('now'))
        )"""
    )


def load_office_settings(connection: sqlite3.Connection | None) -> OfficeSettings:
    """Datos guardados o, si aún no hay, valores neutros (sin marca de nadie)."""
    if connection is None:
        return OfficeSettings()
    try:
        row = connection.execute(
            "SELECT settings_json FROM office_settings WHERE id=1"
        ).fetchone()
    except sqlite3.OperationalError:
        return OfficeSettings()
    if row is None:
        return OfficeSettings()
    try:
        payload = json.loads(row[0])
    except (TypeError, ValueError):
        return OfficeSettings()
    known = {item.name for item in fields(OfficeSettings)}
    values = {key: value for key, value in payload.items() if key in known}
    try:
        values["fiscal_start_month"] = int(values.get("fiscal_start_month") or 1)
    except (TypeError, ValueError):
        values["fiscal_start_month"] = 1
    return OfficeSettings(**values)


def validate_office_settings(settings: OfficeSettings) -> OfficeSettings:
    """Normaliza y valida; lanza ``ValueError`` con un mensaje para el usuario."""
    cleaned = replace(
        settings,
        **{
            item.name: str(getattr(settings, item.name) or "").strip()
            for item in fields(OfficeSettings)
            if item.name != "fiscal_start_month"
        },
    )
    if not cleaned.name:
        raise ValueError("Indica el nombre del despacho.")
    if cleaned.tax_id:
        tax_id = normalize_tax_id(cleaned.tax_id)
        if not valid_spanish_tax_id(tax_id):
            raise ValueError(f"El CIF/NIF «{cleaned.tax_id}» no es válido.")
        cleaned = replace(cleaned, tax_id=tax_id)
    if not 1 <= int(cleaned.fiscal_start_month) <= 12:
        raise ValueError("El mes de inicio del ejercicio debe estar entre 1 y 12.")
    return replace(cleaned, fiscal_start_month=int(cleaned.fiscal_start_month))


def save_office_settings(connection: sqlite3.Connection, settings: OfficeSettings) -> OfficeSettings:
    cleaned = validate_office_settings(settings)
    _ensure_table(connection)
    connection.execute(
        """INSERT INTO office_settings(id, settings_json, updated_at)
           VALUES (1, ?, datetime('now'))
           ON CONFLICT(id) DO UPDATE SET settings_json=excluded.settings_json,
                                         updated_at=excluded.updated_at""",
        (json.dumps(asdict(cleaned), ensure_ascii=False),),
    )
    connection.commit()
    return cleaned


def install_logo(project_root: Path, source: Path) -> str:
    """Copia el logo dentro del proyecto y devuelve su ruta relativa."""
    source = Path(source)
    if source.suffix.lower() not in {".png", ".jpg", ".jpeg"}:
        raise ValueError("El logo debe ser una imagen PNG o JPG.")
    root = Path(project_root).resolve()
    destination = root / LOGO_RELATIVE_DIR / f"logo{source.suffix.lower()}"
    destination.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(source, destination)
    return destination.relative_to(root).as_posix()


def resolve_logo(project_root: Path, settings: OfficeSettings) -> Path | None:
    if not settings.logo_path:
        return None
    root = Path(project_root).resolve()
    candidate = (root / settings.logo_path).resolve()
    try:
        candidate.relative_to(root)
    except ValueError:
        return None
    return candidate if candidate.is_file() else None


def infer_fiscal_start_month(connection: sqlite3.Connection) -> int:
    """Mes de inicio más habitual de los expedientes existentes (1 si no hay)."""
    try:
        rows = connection.execute(
            "SELECT fecha_inicio FROM regularization_cases WHERE fecha_inicio IS NOT NULL"
        ).fetchall()
    except sqlite3.OperationalError:
        return 1
    months = Counter(int(str(row[0])[5:7]) for row in rows if len(str(row[0])) >= 7)
    return months.most_common(1)[0][0] if months else 1


def suggested_settings(connection: sqlite3.Connection) -> OfficeSettings:
    """Punto de partida del asistente: lo guardado o lo deducible de la base."""
    current = load_office_settings(connection)
    if current.configured:
        return current
    return replace(current, fiscal_start_month=infer_fiscal_start_month(connection))


__all__ = [
    "OfficeSettings",
    "infer_fiscal_start_month",
    "install_logo",
    "load_office_settings",
    "resolve_logo",
    "save_office_settings",
    "suggested_settings",
    "validate_office_settings",
]
