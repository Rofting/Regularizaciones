"""Gastos fijos mensuales del estudio (hoja OTROS GASTOS del modelo común).

El modelo reparte cuatro gastos periódicos que no siempre llegan como factura
del expediente: el servicio de lecturas de contadores de ACS y de calefacción
y el mantenimiento de placas solares y de la sala de calderas. Son de cada
comunidad, así que el modelo común los trae vacíos y aquí se guardan por
período para escribirlos al generar el Excel.
"""

from __future__ import annotations

import sqlite3
import unicodedata
from dataclasses import dataclass
from decimal import Decimal

from cuotas_servicio import _importe


SHEET = "OTROS GASTOS"


@dataclass(frozen=True)
class FixedCost:
    key: str
    label: str
    cell: str
    sheet_label: str  # texto de la columna A que confirma que la fila es la esperada


FIXED_COSTS = (
    FixedCost("fixed_cost_meter_reading_acs", "Lecturas de contadores de ACS", "B6",
              "LECTURAS CONTADORES ACS"),
    FixedCost("fixed_cost_meter_reading_heating", "Lecturas de contadores de calefacción", "B7",
              "LECTURAS CONTADORES CALEF"),
    FixedCost("fixed_cost_solar_maintenance", "Mantenimiento de placas solares", "B8",
              "MANTENIMIENTO PLACAS SOLARES"),
    FixedCost("fixed_cost_boiler_maintenance", "Mantenimiento de la sala de calderas", "B9",
              "MANTENIMIENTO SALA CALDERAS"),
)
FIXED_COST_CELLS = {SHEET: tuple(item.cell for item in FIXED_COSTS)}


def _normalise(value: object) -> str:
    text = unicodedata.normalize("NFKD", str(value or ""))
    text = "".join(char for char in text if not unicodedata.combining(char))
    return " ".join(text.upper().split())


def load_fixed_costs(
    connection: sqlite3.Connection, community_id: int, period_id: int,
) -> dict[str, Decimal]:
    keys = tuple(item.key for item in FIXED_COSTS)
    rows = connection.execute(
        f"""SELECT parameter_key, numeric_value FROM period_parameters
            WHERE id_comunidad=? AND id_periodo=? AND parameter_key IN ({",".join("?" * len(keys))})""",
        (community_id, period_id, *keys),
    ).fetchall()
    return {
        str(row[0]): Decimal(str(row[1])) for row in rows if row[1] is not None
    }


def save_fixed_costs(
    connection: sqlite3.Connection,
    community_id: int,
    period_id: int,
    values: dict[str, object],
) -> dict[str, Decimal]:
    """Guarda importes mensuales; un campo vacío borra el valor anterior."""
    known = {item.key for item in FIXED_COSTS}
    saved: dict[str, Decimal] = {}
    for key, raw in values.items():
        if key not in known:
            raise ValueError(f"Gasto fijo desconocido: {key}")
        text = str(raw or "").strip()
        if not text:
            connection.execute(
                """DELETE FROM period_parameters
                   WHERE id_comunidad=? AND id_periodo=? AND parameter_key=?""",
                (community_id, period_id, key),
            )
            continue
        amount = _importe(text, "El importe mensual")
        if amount < 0:
            raise ValueError("Los gastos fijos no pueden ser negativos")
        connection.execute(
            """INSERT INTO period_parameters
                   (id_comunidad,id_periodo,parameter_key,numeric_value,unit)
               VALUES (?,?,?,?,'EUR/mes')
               ON CONFLICT(id_comunidad,id_periodo,parameter_key)
               DO UPDATE SET numeric_value=excluded.numeric_value, unit=excluded.unit""",
            (community_id, period_id, key, float(amount)),
        )
        saved[key] = amount
    connection.commit()
    return saved


def write_fixed_costs(workbook, values: dict[str, Decimal], set_cell) -> list[str]:
    """Escribe los importes en su celda si la fila del libro es la esperada.

    Devuelve las celdas escritas. Una plantilla propia con otra disposición no
    se toca: sólo se escribe donde la etiqueta de la columna A coincide.
    """
    if SHEET not in workbook.sheetnames or not values:
        return []
    sheet = workbook[SHEET]
    written = []
    for item in FIXED_COSTS:
        if item.key not in values:
            continue
        row = sheet[item.cell].row
        if not _normalise(sheet[f"A{row}"].value).startswith(item.sheet_label):
            continue
        set_cell(sheet[item.cell], float(values[item.key]))
        written.append(item.cell)
    return written


__all__ = [
    "FIXED_COSTS",
    "FIXED_COST_CELLS",
    "load_fixed_costs",
    "save_fixed_costs",
    "write_fixed_costs",
]
