"""Cálculo genérico de resultados por concepto y propiedad."""

from __future__ import annotations

import sqlite3
from collections.abc import Callable, Mapping
from decimal import Decimal, ROUND_FLOOR, ROUND_HALF_UP


ProgressCallback = Callable[[str, dict], None]


def allocate_cents(total_cents: int, weights: Mapping[int, Decimal]) -> dict[int, int]:
    """Distribuye todos los céntimos por mayor resto con desempate estable."""
    if not weights or sum(weights.values(), Decimal(0)) <= 0:
        raise ValueError("Se necesitan pesos positivos para distribuir el importe")
    sign = -1 if total_cents < 0 else 1
    absolute_total = abs(total_cents)
    weight_total = sum(weights.values(), Decimal(0))
    exact = {
        key: Decimal(absolute_total) * Decimal(weight) / weight_total
        for key, weight in weights.items()
    }
    allocated = {
        key: int(value.to_integral_value(rounding=ROUND_FLOOR))
        for key, value in exact.items()
    }
    remaining = absolute_total - sum(allocated.values())
    order = sorted(weights, key=lambda key: (-(exact[key] - allocated[key]), key))
    for key in order[:remaining]:
        allocated[key] += 1
    return {key: sign * value for key, value in allocated.items()}


def _euros_to_cents(value: str) -> int:
    amount = Decimal(str(value)).quantize(Decimal("0.01"), ROUND_HALF_UP)
    return int(amount * 100)


def calculate_period(
    connection: sqlite3.Connection,
    id_periodo: int,
    id_batch: int,
    *,
    progress: ProgressCallback | None = None,
) -> int:
    period = connection.execute(
        "SELECT fecha_inicio,fecha_fin FROM periodos WHERE id_periodo=?", (id_periodo,)
    ).fetchone()
    if not period or not period[1]:
        raise ValueError("El periodo debe estar cerrado antes de calcular")
    rows = connection.execute(
        """
        SELECT p.id_propietario, ini.valor_acumulado, fin.valor_acumulado
        FROM propietarios p
        JOIN lecturas_vecino ini ON ini.id_propietario=p.id_propietario
            AND ini.id_periodo=? AND ini.tipo='ACS' AND ini.fecha_lectura=?
        JOIN lecturas_vecino fin ON fin.id_propietario=p.id_propietario
            AND fin.id_periodo=? AND fin.tipo='ACS' AND fin.fecha_lectura=?
        ORDER BY p.id_propietario
        """,
        (id_periodo, period[0], id_periodo, period[1]),
    ).fetchall()
    if not rows:
        raise ValueError("No hay lecturas individuales para calcular el periodo")
    consumptions = {
        int(row[0]): Decimal(str(row[2])) - Decimal(str(row[1])) for row in rows
    }
    if any(value < 0 for value in consumptions.values()):
        raise ValueError("Hay consumos negativos pendientes de revisión")

    source_rows = connection.execute(
        """
        SELECT entity_key,field_name,normalized_value FROM source_values
        WHERE id_batch=? AND entity_type='concept'
          AND field_name IN ('billed_cents','actual_cents')
        """,
        (id_batch,),
    ).fetchall()
    totals: dict[str, dict[str, int]] = {}
    for concept, field, value in source_rows:
        totals.setdefault(concept, {})[field] = _euros_to_cents(value)
    required = {"acs_fixed", "acs_variable"}
    if not required.issubset(totals):
        raise ValueError("La referencia no contiene los conceptos ACS requeridos")

    if connection.in_transaction:
        connection.commit()
    connection.execute("BEGIN IMMEDIATE")
    written = 0
    try:
        for concept in ("acs_fixed", "acs_variable"):
            weights = (
                {owner_id: Decimal(1) for owner_id in consumptions}
                if concept == "acs_fixed"
                else consumptions
            )
            billed = allocate_cents(totals[concept]["billed_cents"], weights)
            actual = allocate_cents(totals[concept]["actual_cents"], weights)
            for owner_id in sorted(weights):
                consumption = float(consumptions[owner_id]) if concept == "acs_variable" else None
                connection.execute(
                    """
                    INSERT INTO owner_concept_results
                        (id_propietario,id_periodo,concept_key,consumption,
                         consumption_unit,billed_cents,actual_cents,difference_cents,
                         id_batch,status)
                    VALUES (?,?,?,?,?,?,?,?,?,'calculated')
                    ON CONFLICT(id_propietario,id_periodo,concept_key) DO UPDATE SET
                        consumption=excluded.consumption,
                        consumption_unit=excluded.consumption_unit,
                        billed_cents=excluded.billed_cents,
                        actual_cents=excluded.actual_cents,
                        difference_cents=excluded.difference_cents,
                        id_batch=excluded.id_batch,
                        status='calculated'
                    """,
                    (
                        owner_id,
                        id_periodo,
                        concept,
                        consumption,
                        "m³" if consumption is not None else None,
                        billed[owner_id],
                        actual[owner_id],
                        actual[owner_id] - billed[owner_id],
                        id_batch,
                    ),
                )
                written += 1
        connection.commit()
    except Exception:
        connection.rollback()
        raise
    if progress:
        progress("period_calculated", {"period_id": id_periodo, "results": written})
    return written
