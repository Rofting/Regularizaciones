"""Conciliación al céntimo entre referencia y resultados calculados."""

from __future__ import annotations

import sqlite3
from dataclasses import dataclass
from decimal import Decimal, ROUND_HALF_UP
from typing import Literal


class ReconciliationError(ValueError):
    pass


@dataclass(frozen=True)
class ConceptReconciliation:
    concept_key: str
    reference_billed_cents: int
    calculated_billed_cents: int
    reference_actual_cents: int
    calculated_actual_cents: int
    difference_cents: int
    status: Literal["cuadrado", "descuadrado"]


@dataclass(frozen=True)
class ReconciliationResult:
    period_id: int
    status: Literal["cuadrado", "descuadrado"]
    difference_cents: int
    concepts: tuple[ConceptReconciliation, ...]


def _cents(value: str) -> int:
    return int(Decimal(str(value)).quantize(Decimal("0.01"), ROUND_HALF_UP) * 100)


def reconcile_period(connection: sqlite3.Connection, id_periodo: int) -> ReconciliationResult:
    reference_rows = connection.execute(
        """
        SELECT sv.entity_key,sv.field_name,sv.normalized_value
        FROM source_values sv JOIN import_batches b ON b.id_batch=sv.id_batch
        WHERE b.id_periodo=? AND b.status='validated' AND sv.entity_type='concept'
          AND sv.field_name IN ('billed_cents','actual_cents')
        """,
        (id_periodo,),
    ).fetchall()
    references: dict[str, dict[str, int]] = {}
    for concept, field, value in reference_rows:
        references.setdefault(concept, {})[field] = _cents(value)
    calculated = {
        row[0]: (int(row[1] or 0), int(row[2] or 0))
        for row in connection.execute(
            """SELECT concept_key,SUM(billed_cents),SUM(actual_cents)
               FROM owner_concept_results WHERE id_periodo=? GROUP BY concept_key""",
            (id_periodo,),
        )
    }
    details = []
    for concept in sorted(references):
        ref_billed = references[concept]["billed_cents"]
        ref_actual = references[concept]["actual_cents"]
        calc_billed, calc_actual = calculated.get(concept, (0, 0))
        difference = max(abs(calc_billed - ref_billed), abs(calc_actual - ref_actual))
        status = "cuadrado" if difference <= 1 else "descuadrado"
        connection.execute(
            """INSERT INTO reconciliations
               (id_periodo,concept_key,reference_billed_cents,calculated_billed_cents,
                reference_actual_cents,calculated_actual_cents,difference_cents,status)
               VALUES (?,?,?,?,?,?,?,?)
               ON CONFLICT(id_periodo,concept_key) DO UPDATE SET
                reference_billed_cents=excluded.reference_billed_cents,
                calculated_billed_cents=excluded.calculated_billed_cents,
                reference_actual_cents=excluded.reference_actual_cents,
                calculated_actual_cents=excluded.calculated_actual_cents,
                difference_cents=excluded.difference_cents,status=excluded.status,
                checked_at=datetime('now')""",
            (id_periodo, concept, ref_billed, calc_billed, ref_actual, calc_actual, difference, status),
        )
        details.append(ConceptReconciliation(concept, ref_billed, calc_billed, ref_actual, calc_actual, difference, status))
    connection.commit()
    overall = "cuadrado" if details and all(item.status == "cuadrado" for item in details) else "descuadrado"
    return ReconciliationResult(id_periodo, overall, max((item.difference_cents for item in details), default=0), tuple(details))


def assert_period_reconciled(connection: sqlite3.Connection, id_periodo: int) -> None:
    result = reconcile_period(connection, id_periodo)
    if result.status != "cuadrado":
        problems = ", ".join(
            f"{item.concept_key}: {item.difference_cents} céntimos"
            for item in result.concepts if item.status != "cuadrado"
        )
        raise ReconciliationError(f"El periodo no cuadra: {problems}")
