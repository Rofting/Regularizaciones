"""CUPS confirmados que identifican una comunidad y su suministro."""

from __future__ import annotations

import re
import sqlite3
from dataclasses import dataclass


_CUPS = re.compile(r"ES\d{16}[A-Z]{2}(?:\d[A-Z])?")
_CUPS_IN_TEXT = re.compile(
    r"\bES\s?\d{4}\s?\d{4}\s?\d{4}\s?\d{4}\s?[A-Z]{2}(?:\s?\d\s?[A-Z])?\b",
    re.IGNORECASE,
)


@dataclass(frozen=True)
class LearnedSupplyPoint:
    cups: str
    community_id: int
    community_code: str
    community_name: str
    supply_type: str
    source_name: str | None


def normalize_cups(value: object) -> str | None:
    """Acepta un CUPS español completo; descarta referencias genéricas."""
    normalized = re.sub(r"[\s-]", "", str(value or "").upper())
    return normalized if _CUPS.fullmatch(normalized) else None


def cups_in_text(text: str) -> tuple[str, ...]:
    return tuple(dict.fromkeys(
        normalized for match in _CUPS_IN_TEXT.finditer(text or "")
        if (normalized := normalize_cups(match.group(0)))
    ))


def lookup_supply_point(connection: sqlite3.Connection, cups: str) -> LearnedSupplyPoint | None:
    normalized = normalize_cups(cups)
    if normalized is None:
        return None
    row = connection.execute("""SELECT p.cups,p.id_comunidad,c.codigo,c.nombre,
                                      p.tipo_suministro,d.original_name
        FROM learned_supply_points p
        JOIN comunidades c ON c.id_comunidad=p.id_comunidad
        LEFT JOIN source_documents d ON d.id_document=p.id_document
        WHERE p.cups=?""", (normalized,)).fetchone()
    if row is None:
        return None
    return LearnedSupplyPoint(*row)


def assert_supply_point_compatible(
    connection: sqlite3.Connection, cups: str, community_id: int, supply_type: str,
    *, current_source: str,
) -> str | None:
    """Rechaza una identidad contradictoria e incluye ambas evidencias."""
    normalized = normalize_cups(cups)
    if normalized is None:
        return None
    existing = lookup_supply_point(connection, normalized)
    supply = supply_type.strip().upper()
    if existing and (existing.community_id != community_id or existing.supply_type != supply):
        current = connection.execute(
            "SELECT codigo,nombre FROM comunidades WHERE id_comunidad=?", (community_id,),
        ).fetchone()
        if current is None:
            raise LookupError("La comunidad no existe")
        raise ValueError(
            f"CUPS {normalized} ya confirmado para {existing.community_code} — "
            f"{existing.community_name} ({existing.supply_type}; "
            f"{existing.source_name or 'fuente anterior'}). "
            f"La fuente {current_source} propone {current[0]} — {current[1]} "
            f"({supply}). Revise ambas facturas antes de asignarlo."
        )
    return normalized


def learn_supply_point(
    connection: sqlite3.Connection, cups: str, community_id: int, supply_type: str,
    *, document_id: int, invoice_id: int, current_source: str,
) -> str | None:
    normalized = assert_supply_point_compatible(
        connection, cups, community_id, supply_type, current_source=current_source,
    )
    if normalized is None:
        return None
    connection.execute("""INSERT OR IGNORE INTO learned_supply_points
        (cups,id_comunidad,tipo_suministro,id_document,id_factura)
        VALUES (?,?,?,?,?)""", (
            normalized, community_id, supply_type.strip().upper(), document_id, invoice_id,
        ))
    return normalized
