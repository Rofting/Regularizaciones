"""Importación transaccional de propietarios desde listados CSV."""

from __future__ import annotations

import csv
import re
import sqlite3
from dataclasses import dataclass
from pathlib import Path
from collections.abc import Callable


ProgressCallback = Callable[[str, dict], None]
EMAIL_RE = re.compile(r"[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}", re.IGNORECASE)


@dataclass(frozen=True)
class OwnerImportSummary:
    inserted: int
    updated: int
    skipped: int


def _clean(value) -> str:
    return " ".join(str(value or "").strip().split())


def _decimal(value) -> float:
    text = _clean(value).replace(".", "").replace(",", ".")
    return float(text or 0)


def _email(value) -> str | None:
    matches = EMAIL_RE.findall(str(value or ""))
    return matches[0].lower() if matches else None


def import_owner_csv(
    connection: sqlite3.Connection,
    id_comunidad: int,
    path: str | Path,
    *,
    expected_count: int | None = None,
    progress: ProgressCallback | None = None,
) -> OwnerImportSummary:
    path = Path(path)
    if progress:
        progress("reading_owners", {"file": path.name})
    with path.open("r", encoding="utf-8-sig", errors="replace", newline="") as handle:
        rows = list(csv.DictReader(handle, delimiter=";"))
    valid_rows = [row for row in rows if _clean(row.get("Fdenominacion"))]
    if expected_count is not None and len(valid_rows) != expected_count:
        raise ValueError(
            f"Se encontraron {len(valid_rows)} propiedades y se esperaban {expected_count}"
        )

    if connection.in_transaction:
        connection.commit()
    connection.execute("BEGIN IMMEDIATE")
    inserted = updated = skipped = 0
    try:
        for row in valid_rows:
            property_code = _clean(row.get("Fdenominacion"))
            owner_name = _clean(row.get("Nombre"))
            if not owner_name:
                skipped += 1
                continue
            values = (
                owner_name,
                _decimal(row.get("Coeficiente")),
                _email(row.get("Email")),
            )
            existing = connection.execute(
                """
                SELECT id_propietario FROM propietarios
                WHERE id_comunidad=? AND codigo_vivienda=?
                """,
                (id_comunidad, property_code),
            ).fetchone()
            if existing:
                connection.execute(
                    """
                    UPDATE propietarios
                    SET nombre_propietario=?, coeficiente=?, email=?, activo=1
                    WHERE id_propietario=?
                    """,
                    (*values, existing[0]),
                )
                updated += 1
            else:
                connection.execute(
                    """
                    INSERT INTO propietarios
                        (id_comunidad, codigo_vivienda, nombre_propietario,
                         coeficiente, email)
                    VALUES (?, ?, ?, ?, ?)
                    """,
                    (id_comunidad, property_code, *values),
                )
                inserted += 1
        connection.commit()
    except Exception:
        connection.rollback()
        raise
    if progress:
        progress(
            "owners_imported",
            {"inserted": inserted, "updated": updated, "skipped": skipped},
        )
    return OwnerImportSummary(inserted, updated, skipped)
