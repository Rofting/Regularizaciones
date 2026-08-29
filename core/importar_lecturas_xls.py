"""Importa lecturas individuales XLS y detecta reinicios anuales."""

from __future__ import annotations

import re
import sqlite3
import unicodedata
from collections.abc import Callable
from dataclasses import dataclass
from pathlib import Path

import xlrd


ProgressCallback = Callable[[str, dict], None]


@dataclass(frozen=True)
class ReadingImportSummary:
    inserted: int
    duplicates: int
    unresolved_properties: tuple[str, ...]
    reset_detected: bool
    participating_properties: int


def _key(value) -> str:
    text = unicodedata.normalize("NFKD", str(value or "").upper())
    text = "".join(character for character in text if not unicodedata.combining(character))
    return re.sub(r"[^A-Z0-9]", "", text.replace("�", ""))


def _is_period_header(value) -> bool:
    text = str(value or "").strip().lower()
    return bool(
        re.fullmatch(r"\d{1,2}/\d{2,4}", text)
        or re.fullmatch(r"[a-záéíóú]{3,10}-\d{2,4}", text)
    )


def import_readings_xls(
    connection: sqlite3.Connection,
    id_comunidad: int,
    id_periodo: int,
    path: str | Path,
    *,
    service: str = "ACS",
    progress: ProgressCallback | None = None,
) -> ReadingImportSummary:
    path = Path(path)
    if progress:
        progress("reading_individual_readings", {"file": path.name})
    sheet = xlrd.open_workbook(path).sheet_by_index(0)
    header_row = None
    for row_index in range(sheet.nrows):
        labels = [_key(sheet.cell_value(row_index, column)) for column in range(sheet.ncols)]
        if "COD" in labels and "PROPIEDAD" in labels:
            header_row = row_index
            break
    if header_row is None:
        raise ValueError("No se encuentran las cabeceras Cod. y Propiedad")
    headers = [sheet.cell_value(header_row, column) for column in range(sheet.ncols)]
    property_column = next(index for index, value in enumerate(headers) if _key(value) == "PROPIEDAD")
    period_columns = [index for index, value in enumerate(headers) if _is_period_header(value)]
    if len(period_columns) < 2:
        raise ValueError("No se encuentran dos columnas de lectura por periodo")
    initial_column, final_column = period_columns[0], period_columns[-1]

    owners = {
        _key(row["codigo_vivienda"]): int(row["id_propietario"])
        for row in connection.execute(
            "SELECT id_propietario, codigo_vivienda FROM propietarios WHERE id_comunidad=?",
            (id_comunidad,),
        )
    }
    period = connection.execute(
        "SELECT fecha_inicio, fecha_fin FROM periodos WHERE id_periodo=? AND id_comunidad=?",
        (id_periodo, id_comunidad),
    ).fetchone()
    if not period or not period[1]:
        raise ValueError("El periodo debe tener fecha inicial y final")

    candidates = []
    unresolved = []
    for row_index in range(header_row + 1, sheet.nrows):
        property_code = str(sheet.cell_value(row_index, property_column) or "").strip()
        if not property_code:
            continue
        try:
            initial = float(sheet.cell_value(row_index, initial_column) or 0)
            final = float(sheet.cell_value(row_index, final_column) or 0)
        except (TypeError, ValueError):
            continue
        owner_id = owners.get(_key(property_code))
        if owner_id is None:
            if initial or final:
                unresolved.append(property_code)
            continue
        participated_before = connection.execute(
            "SELECT 1 FROM lecturas_vecino WHERE id_propietario=? AND tipo=? LIMIT 1",
            (owner_id, service),
        ).fetchone()
        if initial == 0 and final == 0 and not participated_before:
            continue
        candidates.append((owner_id, property_code, initial, final))

    if unresolved:
        raise ValueError("Propiedades sin correspondencia: " + ", ".join(sorted(unresolved)))
    comparable = [(initial, final) for _, _, initial, final in candidates if initial > 0]
    reset_ratio = (
        sum(1 for initial, final in comparable if final < initial) / len(comparable)
        if comparable
        else 0
    )
    reset_detected = reset_ratio >= 0.8

    if connection.in_transaction:
        connection.commit()
    connection.execute("BEGIN IMMEDIATE")
    inserted = duplicates = 0
    try:
        for owner_id, _, initial, final in candidates:
            stored_initial = 0.0 if reset_detected else initial
            note = f"reinicio anual; lectura previa={initial}" if reset_detected else None
            for reading_date, value in ((period[0], stored_initial), (period[1], final)):
                cursor = connection.execute(
                    """
                    INSERT OR IGNORE INTO lecturas_vecino
                        (id_propietario, id_periodo, tipo, fecha_lectura,
                         valor_acumulado, estado, fuente, notas)
                    VALUES (?, ?, ?, ?, ?, 'real', 'excel_lecturas', ?)
                    """,
                    (owner_id, id_periodo, service, reading_date, value, note),
                )
                if cursor.rowcount:
                    inserted += 1
                else:
                    duplicates += 1
        connection.commit()
    except Exception:
        connection.rollback()
        raise
    if progress:
        progress(
            "individual_readings_imported",
            {
                "inserted": inserted,
                "duplicates": duplicates,
                "participants": len(candidates),
                "reset_detected": reset_detected,
            },
        )
    return ReadingImportSummary(
        inserted,
        duplicates,
        tuple(),
        reset_detected,
        len(candidates),
    )
