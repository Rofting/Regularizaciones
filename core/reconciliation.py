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


def reconcile_declared_totals(
    connection: sqlite3.Connection,
    *,
    id_periodo: int,
    references: dict[str, tuple[int, int]],
    commit: bool = True,
    tolerance_cents: int = 0,
) -> ReconciliationResult:
    """Conciliación reutilizable para fuentes normalizadas declaradas.

    ``references`` contiene ``(facturado, real)`` ya expresados en céntimos.
    El reparto por expediente usa tolerancia cero; la API anterior conserva su
    tolerancia histórica de un céntimo al llamar a esta función.
    """
    if tolerance_cents < 0:
        raise ValueError("La tolerancia de conciliación no puede ser negativa")
    calculated = {
        row[0]: (int(row[1] or 0), int(row[2] or 0))
        for row in connection.execute(
            """SELECT concept_key,SUM(billed_cents),SUM(actual_cents)
               FROM owner_concept_results WHERE id_periodo=? GROUP BY concept_key""",
            (id_periodo,),
        )
    }
    if "total_general" in references:
        calculated["total_general"] = (
            sum(calculated.get(key, (0, 0))[0] for key in references if key != "total_general"),
            sum(calculated.get(key, (0, 0))[1] for key in references if key != "total_general"),
        )
    details = []
    for concept in sorted(references):
        ref_billed, ref_actual = references[concept]
        calc_billed, calc_actual = calculated.get(concept, (0, 0))
        difference = max(abs(calc_billed - ref_billed), abs(calc_actual - ref_actual))
        status = "cuadrado" if difference <= tolerance_cents else "descuadrado"
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
    if commit:
        connection.commit()
    overall = "cuadrado" if details and all(item.status == "cuadrado" for item in details) else "descuadrado"
    return ReconciliationResult(id_periodo, overall, max((item.difference_cents for item in details), default=0), tuple(details))


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
    grouped: dict[str, dict[str, int]] = {}
    for concept, field, value in reference_rows:
        grouped.setdefault(concept, {})[field] = _cents(value)
    references = {
        concept: (values["billed_cents"], values["actual_cents"])
        for concept, values in grouped.items()
    }
    return reconcile_declared_totals(
        connection, id_periodo=id_periodo, references=references,
        tolerance_cents=1,
    )


def assert_period_reconciled(connection: sqlite3.Connection, id_periodo: int) -> None:
    result = reconcile_period(connection, id_periodo)
    if result.status != "cuadrado":
        problems = ", ".join(
            f"{item.concept_key}: {item.difference_cents} céntimos"
            for item in result.concepts if item.status != "cuadrado"
        )
        raise ReconciliationError(f"El periodo no cuadra: {problems}")
