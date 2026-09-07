"""Reparto de un expediente validado, exacto al céntimo y reproducible."""

from __future__ import annotations

import sqlite3
from contextlib import contextmanager
from dataclasses import dataclass
from decimal import Decimal, ROUND_FLOOR, ROUND_HALF_UP
from pathlib import Path
from typing import Any, Callable, Mapping

import document_review
from excel_profiles import ConceptRule, ExcelProfile, calculate_profile_sha256, load_profile
from excel_export_service import calculate_case_input_hash
from reconciliation import reconcile_declared_totals


class DistributionBlockedError(ValueError):
    """El expediente no tiene datos aprobados suficientes para repartir."""


@dataclass(frozen=True)
class DistributionResult:
    id_case: int
    id_periodo: int
    concept_totals_cents: dict[str, int]
    owner_result_count: int


ProgressCallback = Callable[[str, dict[str, Any]], None]


def _emit(progress: ProgressCallback | None, stage: str, **payload: Any) -> None:
    if progress is not None:
        progress(stage, payload)


def _cents(value: object) -> int:
    try:
        amount = Decimal(str(value)).quantize(Decimal("0.01"), ROUND_HALF_UP)
    except Exception as error:  # Decimal reports several implementation errors.
        raise DistributionBlockedError(f"Importe no válido para repartir: {value!r}") from error
    return int(amount * 100)


def allocate_concept_cents(
    total_cents: int,
    owner_weights: Mapping[int, Decimal],
) -> dict[int, int]:
    """Distribuye un importe por mayor resto, incluso si es un abono negativo.

    El identificador de propietario es el desempate estable: dos fracciones
    idénticas reciben el céntimo sobrante en orden ascendente de identificador.
    """
    if not owner_weights:
        raise DistributionBlockedError("No hay propietarios con peso para repartir")
    normalized = {int(owner_id): Decimal(str(weight)) for owner_id, weight in owner_weights.items()}
    if any(weight < 0 for weight in normalized.values()):
        raise DistributionBlockedError("Hay un peso negativo o inválido para repartir")
    weight_total = sum(normalized.values(), Decimal(0))
    if weight_total <= 0:
        raise DistributionBlockedError("La suma de pesos debe ser positiva")

    sign = -1 if total_cents < 0 else 1
    absolute_total = abs(int(total_cents))
    exact = {
        owner_id: Decimal(absolute_total) * weight / weight_total
        for owner_id, weight in normalized.items()
    }
    result = {
        owner_id: int(amount.to_integral_value(rounding=ROUND_FLOOR))
        for owner_id, amount in exact.items()
    }
    remaining = absolute_total - sum(result.values())
    order = sorted(
        normalized,
        key=lambda owner_id: (-(exact[owner_id] - result[owner_id]), owner_id),
    )
    for owner_id in order[:remaining]:
        result[owner_id] += 1
    return {owner_id: sign * cents for owner_id, cents in result.items()}


@contextmanager
def _transaction(connection: sqlite3.Connection):
    """Escritura atómica sin confirmar transacciones que pertenecen al caller."""
    if connection.in_transaction:
        connection.execute("SAVEPOINT case_distribution")
        try:
            yield
        except Exception:
            connection.execute("ROLLBACK TO SAVEPOINT case_distribution")
            connection.execute("RELEASE SAVEPOINT case_distribution")
            raise
        else:
            connection.execute("RELEASE SAVEPOINT case_distribution")
        return
    connection.execute("BEGIN IMMEDIATE")
    try:
        yield
    except Exception:
        connection.rollback()
        raise
    else:
        connection.commit()


def _case_and_profile(
    connection: sqlite3.Connection,
    id_case: int,
    project_root: Path,
) -> tuple[sqlite3.Row, ExcelProfile]:
    case = connection.execute(
        """SELECT c.id_case,c.id_comunidad,c.id_periodo,c.estado,c.fecha_inicio,c.fecha_fin,
                  co.codigo
           FROM regularization_cases c JOIN comunidades co ON co.id_comunidad=c.id_comunidad
           WHERE c.id_case=?""",
        (id_case,),
    ).fetchone()
    if case is None:
        raise DistributionBlockedError("El expediente no existe")
    if case["id_periodo"] is None:
        raise DistributionBlockedError("El expediente no tiene un período ligado")
    if case["estado"] not in ("ready_for_calculation", "calculated", "reconciled"):
        raise DistributionBlockedError("El expediente no está listo para repartir")

    period = connection.execute(
        """SELECT id_comunidad,fecha_inicio,fecha_fin FROM periodos WHERE id_periodo=?""",
        (case["id_periodo"],),
    ).fetchone()
    if (
        period is None
        or int(period["id_comunidad"]) != int(case["id_comunidad"])
        or period["fecha_inicio"] != case["fecha_inicio"]
        or period["fecha_fin"] != case["fecha_fin"]
    ):
        raise DistributionBlockedError("El período ligado no coincide con el expediente")

    open_issues = connection.execute(
        "SELECT COUNT(*) FROM review_issues WHERE id_case=? AND status='open'", (id_case,)
    ).fetchone()[0]
    if open_issues:
        raise DistributionBlockedError("Hay incidencias abiertas antes del reparto")
    try:
        document_review.assert_case_final_readings_approved(connection, id_case)
    except ValueError as error:
        raise DistributionBlockedError(str(error)) from error

    latest_export = connection.execute(
        """SELECT e.status,e.id_periodo,e.input_sha256,t.profile_key,t.profile_version,
                  t.profile_sha256
           FROM excel_export_runs e
           JOIN excel_template_profiles t ON t.id_template_profile=e.id_template_profile
           WHERE e.id_case=?
           ORDER BY e.created_at DESC,e.id_export_run DESC LIMIT 1""",
        (id_case,),
    ).fetchone()
    if latest_export is None or latest_export["status"] != "validated":
        raise DistributionBlockedError("Falta un Excel oficial validado para este expediente")
    if int(latest_export["id_periodo"]) != int(case["id_periodo"]):
        raise DistributionBlockedError("El Excel validado pertenece a otro período")
    try:
        profile = load_profile(latest_export["profile_key"], project_root)
    except (LookupError, ValueError) as error:
        raise DistributionBlockedError(f"No se puede cargar el perfil del Excel validado: {error}") from error
    if profile.version != latest_export["profile_version"]:
        raise DistributionBlockedError("La versión del perfil no coincide con el Excel validado")
    if calculate_profile_sha256(profile, project_root) != latest_export["profile_sha256"]:
        raise DistributionBlockedError(
            "La huella del perfil no coincide con el Excel validado; debe reimportar y revalidar"
        )
    if profile.community_code != str(case["codigo"]):
        raise DistributionBlockedError("El perfil del Excel no corresponde a la comunidad")
    try:
        current_input_hash = calculate_case_input_hash(
            connection,
            id_case=id_case,
            project_root=project_root,
            profile=profile,
        )
    except ValueError as error:
        raise DistributionBlockedError(str(error)) from error
    if latest_export["input_sha256"] != current_input_hash:
        raise DistributionBlockedError(
            "Las entradas cambiaron desde el Excel validado; debe regenerar Excel antes de repartir"
        )
    return case, profile


def _owners(connection: sqlite3.Connection, community_id: int) -> list[sqlite3.Row]:
    owners = connection.execute(
        """SELECT id_propietario,coeficiente FROM propietarios
           WHERE id_comunidad=? AND activo=1 ORDER BY id_propietario""",
        (community_id,),
    ).fetchall()
    if not owners:
        raise DistributionBlockedError("No hay propietarios activos para repartir")
    return owners


def _parameter_cents(
    connection: sqlite3.Connection,
    case: sqlite3.Row,
    source: str | None,
) -> int | None:
    if source is None:
        return 0
    prefix = "period_parameters."
    if not source.startswith(prefix):
        raise DistributionBlockedError(f"Fuente de concepto no soportada: {source}")
    key = source[len(prefix):]
    row = connection.execute(
        """SELECT numeric_value FROM period_parameters
           WHERE id_comunidad=? AND id_periodo=? AND parameter_key=?""",
        (case["id_comunidad"], case["id_periodo"], key),
    ).fetchone()
    if row is None or row["numeric_value"] is None:
        return None
    return _cents(row["numeric_value"])


def _source_key(source: str) -> str:
    prefix = "period_parameters."
    if not source.startswith(prefix):
        raise DistributionBlockedError(f"Fuente directa no soportada: {source}")
    return source[len(prefix):]


def _direct_amounts(
    connection: sqlite3.Connection,
    case: sqlite3.Row,
    source: str | None,
    owner_ids: list[int],
) -> dict[int, int]:
    """Lee importes directos por propietario sin inventar una regla de reparto.

    Un perfil que use ``direct`` declara el parámetro base y el importador debe
    guardar cada importe como ``<base>.owner.<id_propietario>``. Es deliberado:
    un total global no permite deducir un cargo individual directo.
    """
    if source is None:
        return {owner_id: 0 for owner_id in owner_ids}
    base = _source_key(source)
    result: dict[int, int] = {}
    for owner_id in owner_ids:
        row = connection.execute(
            """SELECT numeric_value FROM period_parameters
               WHERE id_comunidad=? AND id_periodo=? AND parameter_key=?""",
            (case["id_comunidad"], case["id_periodo"], f"{base}.owner.{owner_id}"),
        ).fetchone()
        if row is None or row["numeric_value"] is None:
            raise DistributionBlockedError(
                f"Falta el importe directo aprobado para el propietario {owner_id}"
            )
        result[owner_id] = _cents(row["numeric_value"])
    return result


def _has_direct_amounts(
    connection: sqlite3.Connection,
    case: sqlite3.Row,
    source: str | None,
    owner_ids: list[int],
) -> bool:
    if source is None:
        return False
    base = _source_key(source)
    placeholders = ",".join("?" for _ in owner_ids)
    rows = connection.execute(
        f"""SELECT parameter_key FROM period_parameters
            WHERE id_comunidad=? AND id_periodo=?
              AND parameter_key IN ({placeholders})""",
        (
            case["id_comunidad"], case["id_periodo"],
            *(f"{base}.owner.{owner_id}" for owner_id in owner_ids),
        ),
    ).fetchall()
    return bool(rows)


def _reading_type(concept: ConceptRule) -> str:
    if concept.key.startswith("acs_"):
        return "ACS"
    if concept.key.startswith("heating_"):
        return "CALEFACCION"
    raise DistributionBlockedError(
        f"El concepto de consumo {concept.key} no indica ACS ni calefacción"
    )


def _consumption_weights(
    connection: sqlite3.Connection,
    case: sqlite3.Row,
    concept: ConceptRule,
    owner_ids: list[int],
) -> dict[int, Decimal]:
    reading_type = _reading_type(concept)
    result: dict[int, Decimal] = {}
    for owner_id in owner_ids:
        rows = connection.execute(
            """SELECT fecha_lectura,valor_acumulado,estado,approved_by,approved_at
               FROM lecturas_vecino
               WHERE id_propietario=? AND id_periodo=? AND tipo=?
                 AND fecha_lectura IN (?,?)
               ORDER BY fecha_lectura""",
            (owner_id, case["id_periodo"], reading_type, case["fecha_inicio"], case["fecha_fin"]),
        ).fetchall()
        by_date = {row["fecha_lectura"]: row for row in rows}
        if case["fecha_inicio"] not in by_date or case["fecha_fin"] not in by_date:
            raise DistributionBlockedError("Falta una lectura de contador necesaria")
        initial = by_date[case["fecha_inicio"]]
        final = by_date[case["fecha_fin"]]
        for reading in (initial, final):
            status = reading["estado"] or "real"
            if status == "contador_averiado":
                raise DistributionBlockedError("Hay un contador averiado pendiente de estimación aprobada")
            if status == "sin_lectura":
                raise DistributionBlockedError("Hay una lectura de contador incompleta")
            if status == "estimado" and (not reading["approved_by"] or not reading["approved_at"]):
                raise DistributionBlockedError("Hay una estimación de contador sin aprobación explícita")
            if status not in ("real", "estimado"):
                raise DistributionBlockedError("Hay un estado de lectura no revisable")
        consumption = Decimal(str(final["valor_acumulado"])) - Decimal(str(initial["valor_acumulado"]))
        if consumption < 0:
            raise DistributionBlockedError("El contador se reinició y no tiene una estimación aprobada")
        result[owner_id] = consumption
    return result


def _weights_for(
    connection: sqlite3.Connection,
    case: sqlite3.Row,
    concept: ConceptRule,
    owners: list[sqlite3.Row],
) -> tuple[dict[int, Decimal], dict[int, float | None]]:
    owner_ids = [int(owner["id_propietario"]) for owner in owners]
    if concept.allocation_method == "equal":
        return {owner_id: Decimal(1) for owner_id in owner_ids}, {owner_id: None for owner_id in owner_ids}
    if concept.allocation_method == "coefficient":
        weights = {
            int(owner["id_propietario"]): Decimal(str(owner["coeficiente"]))
            for owner in owners
        }
        if any(weight <= 0 for weight in weights.values()):
            raise DistributionBlockedError("Hay un peso de coeficiente nulo o ausente")
        return weights, {owner_id: None for owner_id in owner_ids}
    if concept.allocation_method == "consumption":
        weights = _consumption_weights(connection, case, concept, owner_ids)
        return weights, {owner_id: float(weight) for owner_id, weight in weights.items()}
    if concept.allocation_method == "direct":
        return {}, {owner_id: None for owner_id in owner_ids}
    raise DistributionBlockedError(f"Método de reparto no soportado: {concept.allocation_method}")


def _source_batch_id(connection: sqlite3.Connection, case: sqlite3.Row) -> int | None:
    row = connection.execute(
        """SELECT id_batch FROM case_import_batches
           WHERE id_case=? AND id_periodo=? ORDER BY created_at DESC,id_batch DESC LIMIT 1""",
        (case["id_case"], case["id_periodo"]),
    ).fetchone()
    return None if row is None else int(row["id_batch"])


def _write_results(
    connection: sqlite3.Connection,
    *,
    period_id: int,
    concept: ConceptRule,
    allocations_billed: Mapping[int, int],
    allocations_actual: Mapping[int, int],
    consumption: Mapping[int, float | None],
    source_batch_id: int | None,
) -> int:
    connection.execute(
        """INSERT INTO regularization_concepts
           (concept_key,label,display_order,active) VALUES (?,?,?,1)
           ON CONFLICT(concept_key) DO NOTHING""",
        (concept.key, concept.key, 999),
    )
    for owner_id in sorted(allocations_actual):
        actual = int(allocations_actual[owner_id])
        billed = int(allocations_billed[owner_id])
        consumption_value = consumption.get(owner_id)
        connection.execute(
            """INSERT INTO owner_concept_results
               (id_propietario,id_periodo,concept_key,consumption,consumption_unit,
                billed_cents,actual_cents,difference_cents,id_batch,status)
               VALUES (?,?,?,?,?,?,?,?,?,'calculated')
               ON CONFLICT(id_propietario,id_periodo,concept_key) DO UPDATE SET
                consumption=excluded.consumption,
                consumption_unit=excluded.consumption_unit,
                billed_cents=excluded.billed_cents,
                actual_cents=excluded.actual_cents,
                difference_cents=excluded.difference_cents,
                id_batch=excluded.id_batch,status='calculated'""",
            (
                owner_id, period_id, concept.key, consumption_value,
                (
                    "kWh" if concept.key.startswith("heating_") else "m³"
                ) if consumption_value is not None else None,
                billed, actual, actual - billed, source_batch_id,
            ),
        )
    return len(allocations_actual)


def _write_legacy_projection(
    connection: sqlite3.Connection,
    case: sqlite3.Row,
    concepts: tuple[ConceptRule, ...],
) -> None:
    """Mantiene las APIs heredadas con una sola fila ACS/CALEFACCION por dueño."""
    services = {
        "ACS": tuple(concept.key for concept in concepts if concept.key.startswith("acs_")),
        "CALEFACCION": tuple(concept.key for concept in concepts if concept.key.startswith("heating_")),
    }
    for service, keys in services.items():
        connection.execute(
            """DELETE FROM repartos WHERE id_periodo=? AND tipo_suministro=?
               AND id_propietario IN (SELECT id_propietario FROM propietarios WHERE id_comunidad=?)""",
            (case["id_periodo"], service, case["id_comunidad"]),
        )
        if not keys:
            continue
        placeholders = ",".join("?" for _ in keys)
        rows = connection.execute(
            f"""SELECT r.id_propietario,SUM(r.billed_cents) AS billed,SUM(r.actual_cents) AS actual,
                       MAX(r.consumption) AS consumption
                FROM owner_concept_results r JOIN propietarios p ON p.id_propietario=r.id_propietario
                WHERE r.id_periodo=? AND p.id_comunidad=? AND p.activo=1
                  AND r.concept_key IN ({placeholders})
                GROUP BY r.id_propietario ORDER BY r.id_propietario""",
            (case["id_periodo"], case["id_comunidad"], *keys),
        ).fetchall()
        for row in rows:
            billed = int(row["billed"] or 0)
            actual = int(row["actual"] or 0)
            connection.execute(
                """INSERT INTO repartos
                   (id_propietario,id_periodo,tipo_suministro,consumo_real,
                    importe_cobrado,importe_real,diferencia,estado,notas)
                   VALUES (?,?,?,?,?,?,?,?,?)""",
                (
                    row["id_propietario"], case["id_periodo"], service,
                    float(row["consumption"] or 0), billed / 100, actual / 100,
                    (actual - billed) / 100, "calculado",
                    "Proyección del reparto por conceptos del expediente",
                ),
            )


def calculate_case_distribution(
    connection: sqlite3.Connection,
    *,
    id_case: int,
    project_root: Path | None = None,
    progress: ProgressCallback | None = None,
) -> DistributionResult:
    """Calcula todos los conceptos activos de un expediente sin reparto parcial."""
    _emit(progress, "validate_export", id_case=id_case)
    root = Path(project_root or Path(__file__).resolve().parents[1]).resolve()
    with _transaction(connection):
        case, profile = _case_and_profile(connection, id_case, root)
        owners = _owners(connection, int(case["id_comunidad"]))
        owner_ids = [int(owner["id_propietario"]) for owner in owners]
        source_batch_id = _source_batch_id(connection, case)

        # Toda fila de un concepto que el perfil conoce se sustituye; así un
        # concepto opcional que deja de estar presente no sobrevive del cálculo
        # anterior. Los demás conceptos heredados no se tocan.
        configured_keys = [concept.key for concept in profile.concepts]
        if configured_keys:
            placeholders = ",".join("?" for _ in configured_keys)
            connection.execute(
                f"""DELETE FROM owner_concept_results WHERE id_periodo=?
                    AND concept_key IN ({placeholders})
                    AND id_propietario IN (SELECT id_propietario FROM propietarios WHERE id_comunidad=?)""",
                (case["id_periodo"], *configured_keys, case["id_comunidad"]),
            )
            connection.execute(
                f"""DELETE FROM reconciliations WHERE id_periodo=?
                    AND concept_key IN ({placeholders})""",
                (case["id_periodo"], *configured_keys),
            )

        references: dict[str, tuple[int, int]] = {}
        total_actuals: dict[str, int] = {}
        written = 0
        active_concepts: list[ConceptRule] = []
        for concept in profile.concepts:
            if concept.allocation_method == "direct":
                actual_present = _has_direct_amounts(
                    connection, case, concept.actual_source, owner_ids
                )
                billed_present = _has_direct_amounts(
                    connection, case, concept.billed_source, owner_ids
                )
                if not actual_present and not billed_present and not concept.required:
                    continue
                if not actual_present:
                    raise DistributionBlockedError(f"Falta el total real del concepto {concept.key}")
                if concept.billed_source is not None and not billed_present:
                    raise DistributionBlockedError(f"Falta el total facturado del concepto {concept.key}")
                _emit(progress, f"calculate_{concept.key}", id_case=id_case)
                actual_allocations = _direct_amounts(
                    connection, case, concept.actual_source, owner_ids
                )
                billed_allocations = _direct_amounts(
                    connection, case, concept.billed_source, owner_ids
                )
                actual_total = sum(actual_allocations.values())
                billed_total = sum(billed_allocations.values())
                consumption = {owner_id: None for owner_id in owner_ids}
            else:
                actual_total = _parameter_cents(connection, case, concept.actual_source)
                billed_total = _parameter_cents(connection, case, concept.billed_source)
                if actual_total is None:
                    if not concept.required and (
                        concept.billed_source is None or billed_total is None
                    ):
                        continue
                    raise DistributionBlockedError(f"Falta el total real del concepto {concept.key}")
                if billed_total is None:
                    raise DistributionBlockedError(f"Falta el total facturado del concepto {concept.key}")
                _emit(progress, f"calculate_{concept.key}", id_case=id_case)
                weights, consumption = _weights_for(connection, case, concept, owners)
                actual_allocations = allocate_concept_cents(actual_total, weights)
                billed_allocations = allocate_concept_cents(billed_total, weights)
            written += _write_results(
                connection, period_id=int(case["id_periodo"]), concept=concept,
                allocations_billed=billed_allocations, allocations_actual=actual_allocations,
                consumption=consumption, source_batch_id=source_batch_id,
            )
            references[concept.key] = (billed_total, actual_total)
            total_actuals[concept.key] = actual_total
            active_concepts.append(concept)

        references["total_general"] = (
            sum(pair[0] for pair in references.values()),
            sum(pair[1] for pair in references.values()),
        )
        connection.execute(
            """INSERT INTO regularization_concepts
               (concept_key,label,display_order,active) VALUES
               ('total_general','Total general',1000,1)
               ON CONFLICT(concept_key) DO NOTHING"""
        )
        _emit(progress, "reconcile", id_case=id_case)
        reconciliation = reconcile_declared_totals(
            connection, id_periodo=int(case["id_periodo"]), references=references,
            commit=False, tolerance_cents=0,
        )
        if reconciliation.status != "cuadrado":
            raise DistributionBlockedError("El reparto no cuadra al céntimo")
        _write_legacy_projection(connection, case, tuple(active_concepts))
    _emit(progress, "complete", id_case=id_case, results=written)
    return DistributionResult(
        id_case=int(case["id_case"]), id_periodo=int(case["id_periodo"]),
        concept_totals_cents=total_actuals, owner_result_count=written,
    )
