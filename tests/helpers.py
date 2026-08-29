import csv
import sqlite3
import tempfile
from contextlib import contextmanager
from datetime import date
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
from typing import ContextManager, Iterator

from openpyxl import Workbook


def cents(value: str | int | float | Decimal) -> int:
    """Convierte un importe a céntimos con redondeo comercial explícito."""
    amount = Decimal(str(value)).quantize(Decimal("0.01"), ROUND_HALF_UP)
    return int(amount * 100)


@contextmanager
def temporary_database() -> Iterator[tuple[sqlite3.Connection, Path]]:
    """Entrega una conexión SQLite temporal y elimina el archivo al salir."""
    with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as directory:
        path = Path(directory) / "gestion-test.db"
        connection = sqlite3.connect(path)
        connection.row_factory = sqlite3.Row
        connection.execute("PRAGMA foreign_keys = ON")
        try:
            yield connection, path
        finally:
            connection.close()


def make_reference_workbook(
    path: Path,
    *,
    period: str,
    actual_cents: int,
    billed_cents: int,
) -> Path:
    """Crea una referencia económica sintética sin datos personales."""
    workbook = Workbook()
    data_sheet = workbook.active
    data_sheet.title = "DATOS"
    data_sheet["A1"] = "COMUNIDAD DE PRUEBA"
    first_year, second_year = (int(part) for part in period.split("-"))
    data_sheet["A3"] = (
        f"PERIODO: EJERCICIO 01/08/{first_year} - 31/07/{second_year}"
    )
    data_sheet["A4"] = "Nº de VIVIENDAS:"
    data_sheet["D4"] = 2

    readings_sheet = workbook.create_sheet("LECTURAS ACS M3")
    readings_sheet["B8"] = date(first_year, 8, 1)
    readings_sheet["C8"] = date(second_year, 7, 31)

    analysis_sheet = workbook.create_sheet("ANALISIS")
    fixed_actual = actual_cents * 30 // 100
    variable_actual = actual_cents - fixed_actual
    fixed_billed = billed_cents * 40 // 100
    variable_billed = billed_cents - fixed_billed
    analysis_sheet["G54"] = "RESULTADO DEL ANALISIS Y DEL REPARTO DE COSTES"
    analysis_sheet["G56"] = "GASTO"
    analysis_sheet["H56"] = "CUOTA FIJA"
    analysis_sheet["J56"] = "CUOTA VARIABLE"
    analysis_sheet["F64"] = "TOTAL €."
    analysis_sheet["G64"] = actual_cents / 100
    analysis_sheet["H64"] = fixed_actual / 100
    analysis_sheet["J64"] = variable_actual / 100
    analysis_sheet["F66"] = "IMPORTE COBRADO"
    analysis_sheet["G66"] = billed_cents / 100
    analysis_sheet["H66"] = fixed_billed / 100
    analysis_sheet["J66"] = variable_billed / 100
    analysis_sheet["F67"] = "DIFERENCIA"
    analysis_sheet["G67"] = (billed_cents - actual_cents) / 100
    analysis_sheet["H67"] = (fixed_billed - fixed_actual) / 100
    analysis_sheet["J67"] = (variable_billed - variable_actual) / 100
    analysis_sheet["H69"] = "DIFERENCIA COSTE REAL - FACTURADO"
    analysis_sheet["E70"] = "Cuota fija vivienda/mes (€/mes)"
    analysis_sheet["F70"] = 8
    analysis_sheet["G70"] = 7.5
    analysis_sheet["E71"] = "Cuota variable m3 consumido (€/m3)"
    analysis_sheet["F71"] = 12
    analysis_sheet["G71"] = 13.5
    analysis_sheet["F80"] = "Cuota fija vivienda/mes (€/mes)"
    analysis_sheet["F81"] = "Cuota variable m3 consumido (€/m3)"

    workbook.save(path)
    return path


def make_owner_csv(path: Path, rows: list[dict[str, str]]) -> Path:
    """Crea un listado sintético con las cabeceras del origen real."""
    fieldnames = ["Codigo", "Nombre", "Fdenominacion", "Coeficiente", "Email"]
    with path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames, delimiter=";")
        writer.writeheader()
        writer.writerows(rows)
    return path
