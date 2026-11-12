"""Gastos de «OTROS GASTOS» que se deducen de las facturas del expediente.

El estudio manual de la comunidad no obliga a anotar a mano ni el mantenimiento
mensual ni el servicio de lectura de contadores: salen de las facturas. Las
reparaciones se reparten en cinco años, una quinta parte en cada ejercicio.

Reglas (decididas por el despacho):

* Cuenta la factura emitida dentro del ejercicio (``fecha_factura``), aunque su
  consumo sea del mes anterior.
* Lo anotado a mano en «Gastos fijos» manda sobre lo deducido.
* Una reparación entra por una quinta parte en el ejercicio de la factura y en
  los cuatro siguientes.
"""

from __future__ import annotations

import re
import sqlite3
import unicodedata
from dataclasses import dataclass
from datetime import date
from decimal import ROUND_HALF_UP, Decimal

REPAIR_YEARS = 5
_CENT = Decimal("0.01")

METER_READING_KEY = "fixed_cost_meter_reading_acs"
MAINTENANCE_KEY = "fixed_cost_boiler_maintenance"

_REPAIR = re.compile(
    r"\b(reparaci[oó]n(?:es)?|reparar|rep\b\.?|aver[ií]as?|fugas?|sustituci[oó]n|sustituir|cambio de)",
    re.IGNORECASE,
)
_METERING = re.compile(r"contador|lectura", re.IGNORECASE)


@dataclass(frozen=True)
class RepairShare:
    invoice_id: int
    invoice_date: str
    description: str
    total: Decimal
    share: Decimal
    years: int


def _plain(text: object) -> str:
    cleaned = unicodedata.normalize("NFKD", str(text or ""))
    return "".join(char for char in cleaned if not unicodedata.combining(char))


def _cents(value: object) -> Decimal:
    return Decimal(str(value)).quantize(_CENT, rounding=ROUND_HALF_UP)


def _document_text(connection: sqlite3.Connection, archived_path: str | None) -> str:
    if not archived_path:
        return ""
    row = connection.execute(
        """SELECT c.text_content FROM source_documents d
           JOIN document_text_cache c ON c.sha256=d.sha256
           WHERE d.archived_path=? ORDER BY c.created_at DESC LIMIT 1""",
        (archived_path,),
    ).fetchone()
    if row is None:
        # La carpeta de datos puede haberse movido: el nombre archivado empieza
        # por la huella del documento y sigue identificándolo.
        name = str(archived_path).replace("\\", "/").rsplit("/", 1)[-1]
        row = connection.execute(
            """SELECT c.text_content FROM source_documents d
               JOIN document_text_cache c ON c.sha256=d.sha256
               WHERE replace(d.archived_path,'\\','/') LIKE ? ESCAPE '!'
               ORDER BY c.created_at DESC LIMIT 1""",
            ("%/" + name.replace("!", "!!").replace("%", "!%").replace("_", "!_"),),
        ).fetchone()
    return str(row[0]) if row and row[0] else ""


def classify(connection: sqlite3.Connection, invoice: sqlite3.Row) -> str:
    """Devuelve ``meter_reading``, ``repair`` o ``maintenance``."""
    provider = _plain(invoice["proveedor"])
    text = _plain(_document_text(connection, invoice["archivo_origen"]))
    if _METERING.search(provider):
        return "meter_reading"
    file_name = _plain(str(invoice["archivo_origen"] or "").replace("\\", "/").rsplit("/", 1)[-1])
    file_name = file_name.replace("_", " ")
    if _REPAIR.search(text) or _REPAIR.search(_plain(invoice["notas"])) or _REPAIR.search(file_name):
        return "repair"
    return "maintenance"


def _maintenance_invoices(connection, community_id: int, start: str, end: str):
    rows = connection.execute(
        """SELECT id_factura,proveedor,fecha_factura,importe_total,archivo_origen,notas,num_factura
           FROM facturas
           WHERE id_comunidad=? AND tipo_suministro='MANTENIMIENTO'
             AND fecha_factura BETWEEN ? AND ?
           ORDER BY fecha_factura,id_factura""",
        (community_id, start, end),
    ).fetchall()
    return rows


def automatic_fixed_costs(
    connection: sqlite3.Connection, community_id: int, start: str, end: str,
) -> dict[str, Decimal]:
    """Importe mensual medio del ejercicio, tal como lo espera «Gastos fijos».

    El modelo multiplica por doce la celda mensual, de modo que el total anual
    es exactamente la suma de las facturas del ejercicio.
    """
    totals = {METER_READING_KEY: Decimal(0), MAINTENANCE_KEY: Decimal(0)}
    for invoice in _maintenance_invoices(connection, community_id, start, end):
        kind = classify(connection, invoice)
        if kind == "repair":
            continue
        key = METER_READING_KEY if kind == "meter_reading" else MAINTENANCE_KEY
        totals[key] += _cents(invoice["importe_total"])
    return {key: total / 12 for key, total in totals.items() if total}


def repair_shares(
    connection: sqlite3.Connection, community_id: int, start: str, end: str,
) -> list[RepairShare]:
    """Quinta parte de cada reparación facturada en este ejercicio o en los cuatro anteriores."""
    exercise_start = date.fromisoformat(start)
    window_start = date(exercise_start.year - (REPAIR_YEARS - 1), exercise_start.month, min(exercise_start.day, 28))
    shares = []
    for invoice in _maintenance_invoices(connection, community_id, window_start.isoformat(), end):
        if classify(connection, invoice) != "repair":
            continue
        total = _cents(invoice["importe_total"])
        shares.append(RepairShare(
            invoice_id=int(invoice["id_factura"]),
            invoice_date=str(invoice["fecha_factura"]),
            description=f"Reparación {invoice['num_factura'] or ''}".strip(),
            total=total,
            share=(total / REPAIR_YEARS).quantize(_CENT, rounding=ROUND_HALF_UP),
            years=REPAIR_YEARS,
        ))
    return shares


__all__ = [
    "METER_READING_KEY", "MAINTENANCE_KEY", "REPAIR_YEARS", "RepairShare",
    "automatic_fixed_costs", "classify", "repair_shares",
]
