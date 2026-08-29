"""Orquestación del flujo guiado de regularización desde archivos fuente."""

from __future__ import annotations

import sqlite3
from pathlib import Path
from collections.abc import Callable

from importar_excel_referencia import import_reference_workbook
from importar_lecturas_xls import import_readings_xls
from importar_propietarios_csv import import_owner_csv
from reconciliation import reconcile_period
from regularization_service import calculate_period


ProgressCallback = Callable[[str, dict], None]


def run_regularization(
    connection: sqlite3.Connection,
    id_comunidad: int,
    reference_path: str | Path,
    owners_path: str | Path,
    readings_path: str | Path,
    *,
    progress: ProgressCallback | None = None,
) -> dict:
    """Importa fuentes, calcula resultados y exige conciliación exacta."""
    owner_summary = import_owner_csv(connection, id_comunidad, owners_path, progress=progress)
    batch_id, period_id = import_reference_workbook(
        connection, id_comunidad, reference_path, progress=progress
    )
    reading_summary = import_readings_xls(
        connection, id_comunidad, period_id, readings_path, progress=progress
    )
    result_count = calculate_period(connection, period_id, batch_id, progress=progress)
    reconciliation = reconcile_period(connection, period_id)
    if reconciliation.status != "cuadrado":
        raise ValueError(
            f"La conciliación no cuadra ({reconciliation.difference_cents} céntimos)"
        )
    if progress:
        progress(
            "regularization_completed",
            {
                "period_id": period_id,
                "results": result_count,
                "participants": reading_summary.participating_properties,
                "carried_forward": len(reading_summary.carried_forward_properties),
                "reconciliation": reconciliation.status,
            },
        )
    return {
        "period_id": period_id,
        "batch_id": batch_id,
        "owner_summary": owner_summary,
        "reading_summary": reading_summary,
        "results": result_count,
        "reconciliation": reconciliation,
    }
