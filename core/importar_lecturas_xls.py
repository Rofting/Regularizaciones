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
    carried_forward_properties: tuple[str, ...] = ()


def _key(value) -> str:
    text = unicodedata.normalize("NFKD", str(value or "").upper())
    text = "".join(character for character in text if not unicodedata.combining(character))
    return re.sub(r"[^A-Z0-9]", "", text.replace("�", ""))


def _is_period_header(value) -> bool:
    text = str(value or "").strip().lower()
    return bool(
        # Algunos listados de Meditrade imprimen el año abreviado a un dígito
        # (por ejemplo, ``01/6`` para enero de 2026). El intervalo se confirma
        # después en el expediente, por lo que aquí basta reconocer la columna.
        re.fullmatch(r"\d{1,2}/\d{1,4}", text)
        or re.fullmatch(r"[a-záéíóú]{3,10}-\d{2,4}", text)
    )


def reading_columns(headers) -> tuple[int, tuple[int, ...]] | None:
    """Identify the legacy table without assuming service or exact dates."""
    labels = [_key(value) for value in headers]
    if "COD" not in labels or "PROPIEDAD" not in labels:
        return None
    return labels.index("PROPIEDAD"), tuple(
        index for index, value in enumerate(headers) if _is_period_header(value)
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
        columns = reading_columns(sheet.row_values(row_index))
        if columns is not None:
            header_row = row_index
            break
    if header_row is None:
        raise ValueError("No se encuentran las cabeceras Cod. y Propiedad")
    property_column, period_columns = columns
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
    carried_forward = tuple(
        property_code
        for _, property_code, initial, final in candidates
        if not reset_detected and final < initial
    )

    if connection.in_transaction:
        connection.commit()
    connection.execute("BEGIN IMMEDIATE")
    inserted = duplicates = 0
    try:
        for owner_id, property_code, initial, final in candidates:
            stored_initial = 0.0 if reset_detected else initial
            note = f"reinicio anual; lectura previa={initial}" if reset_detected else None
            final_is_carried = property_code in carried_forward
            final_value = initial if final_is_carried else final
            final_state = "estimado" if final_is_carried else "real"
            final_method = "arrastre_lectura_anterior" if final_is_carried else None
            final_note = (
                f"lectura informada={final}; se mantiene lectura anterior={initial}; "
                "pendiente de lectura posterior"
                if final_is_carried
                else note
            )
            for reading_date, value, state, method, reading_note in (
                (period[0], stored_initial, "real", None, note),
                (period[1], final_value, final_state, final_method, final_note),
            ):
                cursor = connection.execute(
                    """
                    INSERT OR IGNORE INTO lecturas_vecino
                        (id_propietario, id_periodo, tipo, fecha_lectura,
                        valor_acumulado, estado, metodo_estimacion, fuente, notas)
                    VALUES (?, ?, ?, ?, ?, ?, ?, 'excel_lecturas', ?)
                    """,
                    (owner_id, id_periodo, service, reading_date, value, state, method, reading_note),
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
                "carried_forward": len(carried_forward),
            },
        )
    return ReadingImportSummary(
        inserted,
        duplicates,
        tuple(),
        reset_detected,
        len(candidates),
        carried_forward,
    )
