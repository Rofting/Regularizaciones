import sqlite3
from contextlib import contextmanager
from decimal import Decimal, InvalidOperation
from itertools import count
from typing import Collection, Iterator, Mapping

from expedient_models import RegularizationCase, ReviewIssue, review_issue_from_row
from expedient_service import get_case, set_case_status


_savepoint_counter = count()
_MISSING_FIELD_CODE = "MISSING_REQUIRED_FIELD"
_COUNTER_RESET_CODE = "COUNTER_RESET"
_INVOICE_OUTSIDE_PERIOD_CODE = "INVOICE_OUTSIDE_PERIOD"


@contextmanager
def _transaction(connection: sqlite3.Connection) -> Iterator[None]:
    """Aísla cada operación sin confirmar la transacción de quien llama."""
    if connection.in_transaction:
        savepoint = f"document_review_{next(_savepoint_counter)}"
        connection.execute(f"SAVEPOINT {savepoint}")
        try:
            yield
        except Exception:
            connection.execute(f"ROLLBACK TO SAVEPOINT {savepoint}")
            connection.execute(f"RELEASE SAVEPOINT {savepoint}")
            raise
        else:
            connection.execute(f"RELEASE SAVEPOINT {savepoint}")
        return

    connection.execute("BEGIN IMMEDIATE")
    try:
        yield
    except Exception:
        connection.rollback()
        raise
    else:
        connection.commit()


def _normalise_value(value: str | None) -> str | None:
    if value is None:
        return None
    normalized = value.strip()
    return normalized or None


def _positive_consumption(value: object) -> Decimal:
    """Valida el consumo confirmado sin convertirlo en una lectura acumulada."""
    if isinstance(value, bool):
        raise ValueError("El consumo estimado debe ser un número positivo")
    try:
        parsed = Decimal(str(value).strip().replace(",", "."))
    except (InvalidOperation, ValueError):
        raise ValueError("El consumo estimado debe ser un número positivo") from None
    if not parsed.is_finite() or parsed <= 0:
        raise ValueError("El consumo estimado debe ser un número positivo")
    return parsed


def _decimal_text(value: Decimal | float | int) -> str:
    rendered = format(Decimal(str(value)), "f")
    return rendered.rstrip("0").rstrip(".") if "." in rendered else rendered


def _counter_reset_target(
    connection: sqlite3.Connection, issue: sqlite3.Row
) -> tuple[sqlite3.Row, sqlite3.Row, sqlite3.Row, sqlite3.Row]:
    """Obtiene la pareja canónica de lecturas que una incidencia puede corregir."""
    if issue["code"] != _COUNTER_RESET_CODE:
        raise LookupError("La incidencia no corresponde a un reinicio de contador")
    prefix, separator, remainder = str(issue["field_name"]).partition("reading.")
    if prefix or not separator:
        raise LookupError("La incidencia no apunta a una lectura de propietario")
    property_code, separator, service = remainder.rpartition(".")
    if not separator or not property_code or service not in {"ACS", "CALEFACCION"}:
        raise LookupError("La incidencia no apunta a una lectura de propietario")
    case = connection.execute(
        """SELECT id_case,id_comunidad,id_periodo,fecha_inicio,fecha_fin
           FROM regularization_cases WHERE id_case=?""",
        (issue["id_case"],),
    ).fetchone()
    if case is None or case["id_periodo"] is None:
        raise LookupError("El expediente no tiene un período de lecturas válido")
    owner = connection.execute(
        """SELECT id_propietario FROM propietarios
           WHERE id_comunidad=? AND codigo_vivienda=?""",
        (case["id_comunidad"], property_code),
    ).fetchone()
    if owner is None:
        raise LookupError("La lectura no pertenece a un propietario del expediente")
    readings = connection.execute(
        """SELECT id_lectura,fecha_lectura,valor_acumulado,estado,metodo_estimacion,
                  fuente,notas,approved_by,approved_at
           FROM lecturas_vecino
           WHERE id_propietario=? AND id_periodo=? AND tipo=?
             AND fecha_lectura IN (?,?)
           ORDER BY fecha_lectura""",
        (
            owner["id_propietario"], case["id_periodo"], service,
            case["fecha_inicio"], case["fecha_fin"],
        ),
    ).fetchall()
    by_date = {row["fecha_lectura"]: row for row in readings}
    initial = by_date.get(case["fecha_inicio"])
    final = by_date.get(case["fecha_fin"])
    if initial is None or final is None:
        raise LookupError("No se encuentran las lecturas inicial y final del período")
    if final["estado"] != "contador_averiado" or final["valor_acumulado"] >= initial["valor_acumulado"]:
        raise LookupError("La lectura ya no es un reinicio pendiente de aprobar")
    return case, owner, initial, final


def _issue_row(connection: sqlite3.Connection, issue_id: int):
    return connection.execute(
        """SELECT issues.id_issue, issues.id_case, issues.id_document,
                  documents.archived_path, issues.code, issues.field_name,
                  issues.detected_value, issues.message, issues.status
           FROM review_issues AS issues
           JOIN source_documents AS documents ON documents.id_document = issues.id_document
           WHERE issues.id_issue = ?""",
        (issue_id,),
    ).fetchone()


def record_candidates(connection: sqlite3.Connection, document_id: int,
                      candidates: Mapping[str, str | None], *, source: str) -> None:
    normalized_source = source.strip()
    if not normalized_source:
        raise ValueError("La fuente del candidato es obligatoria")

    with _transaction(connection):
        for field_name, value in candidates.items():
            normalized_value = _normalise_value(value)
            connection.execute(
                """INSERT INTO extraction_candidates
                   (id_document, field_name, value, source, validation_status)
                   VALUES (?, ?, ?, ?, ?)
                   ON CONFLICT(id_document, field_name) DO UPDATE SET
                       value = excluded.value,
                       source = excluded.source,
                       validation_status = excluded.validation_status""",
                (
                    document_id, field_name, normalized_value, normalized_source,
                    "validated" if normalized_value is not None else "candidate",
                ),
            )


def record_candidates_if_missing(
    connection: sqlite3.Connection,
    document_id: int,
    candidates: Mapping[str, str | None],
    *,
    source: str,
) -> None:
    """Completa una ingestión parcial sin alterar valores ya revisados."""
    normalized_source = source.strip()
    if not normalized_source:
        raise ValueError("La fuente del candidato es obligatoria")

    with _transaction(connection):
        for field_name, value in candidates.items():
            normalized_value = _normalise_value(value)
            connection.execute(
                """INSERT INTO extraction_candidates
                   (id_document, field_name, value, source, validation_status)
                   VALUES (?, ?, ?, ?, ?)
                   ON CONFLICT(id_document, field_name) DO NOTHING""",
                (
                    document_id, field_name, normalized_value, normalized_source,
                    "validated" if normalized_value is not None else "candidate",
                ),
            )


def create_missing_field_issues(connection: sqlite3.Connection, case_id: int,
                                document_id: int, required_fields: Collection[str]) -> tuple[ReviewIssue, ...]:
    with _transaction(connection):
        document = connection.execute(
            "SELECT id_case FROM source_documents WHERE id_document = ?", (document_id,)
        ).fetchone()
        if document is None or document["id_case"] != case_id:
            raise LookupError("El documento no pertenece al expediente")

        for field_name in required_fields:
            candidate = connection.execute(
                """SELECT value FROM extraction_candidates
                   WHERE id_document = ? AND field_name = ?
                     AND value IS NOT NULL AND trim(value) <> ''""",
                (document_id, field_name),
            ).fetchone()
            if candidate is None:
                connection.execute(
                    """INSERT INTO review_issues
                       (id_case, id_document, code, field_name, detected_value, message, status)
                       VALUES (?, ?, ?, ?, NULL, ?, 'open')
                       ON CONFLICT(id_document, code, field_name, status) DO NOTHING""",
                    (
                        case_id, document_id, _MISSING_FIELD_CODE, field_name,
                        f"Falta el campo requerido: {field_name}",
                    ),
                )

        open_count = connection.execute(
            "SELECT COUNT(*) FROM review_issues WHERE id_document = ? AND status = 'open'",
            (document_id,),
        ).fetchone()[0]
        if open_count:
            connection.execute(
                "UPDATE source_documents SET status = 'under_review' WHERE id_document = ?",
                (document_id,),
            )
            case = get_case(connection, case_id)
            if case.status == "gathering_sources":
                set_case_status(connection, case_id, "under_review")

        rows = connection.execute(
            """SELECT issues.id_issue, issues.id_case, issues.id_document,
                      documents.archived_path, issues.code, issues.field_name,
                      issues.detected_value, issues.message, issues.status
               FROM review_issues AS issues
               JOIN source_documents AS documents ON documents.id_document = issues.id_document
               WHERE issues.id_document = ? AND issues.status = 'open'
               ORDER BY issues.field_name""",
            (document_id,),
        ).fetchall()
    return tuple(review_issue_from_row(row) for row in rows)


def create_review_issue(
    connection: sqlite3.Connection,
    case_id: int,
    document_id: int,
    *,
    code: str,
    field_name: str,
    message: str,
    detected_value: str | None = None,
) -> ReviewIssue:
    """Registra de forma idempotente una incidencia genérica que sigue abierta."""
    normalized_code = code.strip()
    normalized_field = field_name.strip()
    normalized_message = message.strip()
    if not normalized_code or not normalized_field or not normalized_message:
        raise ValueError("Código, campo y mensaje de la incidencia son obligatorios")

    with _transaction(connection):
        document = connection.execute(
            "SELECT id_case FROM source_documents WHERE id_document = ?", (document_id,)
        ).fetchone()
        if document is None or document["id_case"] != case_id:
            raise LookupError("El documento no pertenece al expediente")
        connection.execute(
            """INSERT INTO review_issues
               (id_case,id_document,code,field_name,detected_value,message,status)
               VALUES (?,?,?,?,?,?,'open')
               ON CONFLICT(id_document,code,field_name,status) DO NOTHING""",
            (
                case_id,
                document_id,
                normalized_code,
                normalized_field,
                _normalise_value(detected_value),
                normalized_message,
            ),
        )
        connection.execute(
            "UPDATE source_documents SET status='under_review' WHERE id_document=?",
            (document_id,),
        )
        case = get_case(connection, case_id)
        if case.status == "draft":
            case = set_case_status(connection, case_id, "gathering_sources")
        if case.status == "gathering_sources":
            set_case_status(connection, case_id, "under_review")
        row = connection.execute(
            """SELECT issues.id_issue,issues.id_case,issues.id_document,
                      documents.archived_path,issues.code,issues.field_name,
                      issues.detected_value,issues.message,issues.status
               FROM review_issues AS issues
               JOIN source_documents AS documents
                 ON documents.id_document=issues.id_document
               WHERE issues.id_document=? AND issues.code=?
                 AND issues.field_name=? AND issues.status='open'""",
            (document_id, normalized_code, normalized_field),
        ).fetchone()
    return review_issue_from_row(row)


def list_open_issues(connection: sqlite3.Connection, case_id: int) -> tuple[ReviewIssue, ...]:
    rows = connection.execute(
        """SELECT issues.id_issue, issues.id_case, issues.id_document,
                  documents.archived_path, issues.code, issues.field_name,
                  issues.detected_value, issues.message, issues.status
           FROM review_issues AS issues
           JOIN source_documents AS documents ON documents.id_document = issues.id_document
           WHERE issues.id_case = ? AND issues.status = 'open'
           ORDER BY issues.id_issue""",
        (case_id,),
    ).fetchall()
    return tuple(review_issue_from_row(row) for row in rows)


def resolve_issue(connection: sqlite3.Connection, issue_id: int, *, value: str,
                  reason: str, resolved_by: str = "usuario_local") -> ReviewIssue:
    normalized_value = _normalise_value(value)
    normalized_reason = reason.strip()
    normalized_resolved_by = resolved_by.strip()
    if normalized_value is None:
        raise ValueError("El valor confirmado es obligatorio")
    if not normalized_reason:
        raise ValueError("El motivo de la corrección es obligatorio")
    if not normalized_resolved_by:
        raise ValueError("El responsable de la corrección es obligatorio")

    with _transaction(connection):
        issue = _issue_row(connection, issue_id)
        if issue is None or issue["status"] != "open":
            raise LookupError("Incidencia no encontrada")
        if issue["code"] == _COUNTER_RESET_CODE:
            raise ValueError(
                "Un reinicio de contador sólo se puede cerrar con una estimación aprobada"
            )
        candidate = connection.execute(
            """SELECT value FROM extraction_candidates
               WHERE id_document = ? AND field_name = ?""",
            (issue["id_document"], issue["field_name"]),
        ).fetchone()
        original_value = candidate["value"] if candidate is not None else None
        connection.execute(
            """INSERT INTO manual_corrections
               (id_issue, original_value, corrected_value, reason, resolved_by)
               VALUES (?, ?, ?, ?, ?)""",
            (
                issue_id, original_value, normalized_value, normalized_reason,
                normalized_resolved_by,
            ),
        )
        connection.execute(
            """INSERT INTO extraction_candidates
               (id_document, field_name, value, source, validation_status)
               VALUES (?, ?, ?, 'manual', 'validated')
               ON CONFLICT(id_document, field_name) DO UPDATE SET
                   value = excluded.value,
                   source = excluded.source,
                   validation_status = excluded.validation_status""",
            (issue["id_document"], issue["field_name"], normalized_value),
        )
        connection.execute(
            """UPDATE review_issues SET status = 'resolved', resolved_at = datetime('now')
               WHERE id_issue = ?""",
            (issue_id,),
        )
        resolved = _issue_row(connection, issue_id)
    return review_issue_from_row(resolved)


def approve_counter_reset_estimate(
    connection: sqlite3.Connection,
    issue_id: int,
    *,
    consumption: object,
    reason: str,
    approved_by: str,
) -> ReviewIssue:
    """Aprueba una estimación humana para un contador reiniciado.

    ``consumption`` es la variación del período confirmada por el gestor, no
    una lectura nueva. La lectura final guardada se transforma en un valor
    virtual continuo para que el motor de reparto conserve su contrato de
    lectura inicial/final sin utilizar el valor disminuido de origen.
    """
    confirmed_consumption = _positive_consumption(consumption)
    normalized_reason = reason.strip()
    normalized_approver = approved_by.strip()
    if not normalized_reason:
        raise ValueError("El motivo de la estimación es obligatorio")
    if not normalized_approver:
        raise ValueError("La persona que aprueba la estimación es obligatoria")

    with _transaction(connection):
        issue = _issue_row(connection, issue_id)
        if issue is None or issue["status"] != "open":
            raise LookupError("Incidencia no encontrada")
        case, owner, initial, final = _counter_reset_target(connection, issue)
        original_final = Decimal(str(final["valor_acumulado"]))
        corrected_final = Decimal(str(initial["valor_acumulado"])) + confirmed_consumption
        corrected_text = _decimal_text(corrected_final)
        original_text = _decimal_text(original_final)
        notes = (
            f"Estimación manual aprobada tras reinicio; valor original={original_text}; "
            f"consumo confirmado={_decimal_text(confirmed_consumption)}; "
            f"lectura virtual={corrected_text}; motivo={normalized_reason}; "
            f"aprobada por={normalized_approver}"
        )
        if final["notas"]:
            notes = f"{final['notas']} | {notes}"
        connection.execute(
            """INSERT INTO manual_corrections
               (id_issue,original_value,corrected_value,reason,resolved_by)
               VALUES (?,?,?,?,?)""",
            (issue_id, original_text, corrected_text, normalized_reason, normalized_approver),
        )
        connection.execute(
            """INSERT INTO extraction_candidates
               (id_document,field_name,value,source,validation_status)
               VALUES (?, ?, ?, 'manual_counter_reset', 'validated')
               ON CONFLICT(id_document,field_name) DO UPDATE SET
                   value=excluded.value,source=excluded.source,
                   validation_status=excluded.validation_status""",
            (issue["id_document"], issue["field_name"], corrected_text),
        )
        connection.execute(
            """UPDATE lecturas_vecino
               SET valor_acumulado=?,estado='estimado',metodo_estimacion='counter_reset_manual',
                   notas=?,approved_by=?,approved_at=datetime('now')
               WHERE id_lectura=?""",
            (corrected_text, notes, normalized_approver, final["id_lectura"]),
        )
        owner_state = "estado_contador_acs" if issue["field_name"].endswith(".ACS") else "estado_contador_cal"
        connection.execute(
            f"UPDATE propietarios SET {owner_state}='ok' WHERE id_propietario=?",
            (owner["id_propietario"],),
        )
        connection.execute(
            """UPDATE review_issues SET status='resolved',resolved_at=datetime('now')
               WHERE id_issue=?""",
            (issue_id,),
        )
        resolved = _issue_row(connection, issue_id)
    return review_issue_from_row(resolved)


def dismiss_invoice_outside_period(
    connection: sqlite3.Connection,
    issue_id: int,
    *,
    reason: str,
    dismissed_by: str,
) -> ReviewIssue:
    """Cierra con trazabilidad una factura que no debe entrar en el período.

    No corrige su fecha ni crea una factura: sólo documenta que la fuente
    identificada no corresponde a este expediente.
    """
    normalized_reason = reason.strip()
    normalized_actor = dismissed_by.strip()
    if not normalized_reason:
        raise ValueError("El motivo del cierre es obligatorio")
    if not normalized_actor:
        raise ValueError("La persona responsable del cierre es obligatoria")
    with _transaction(connection):
        issue = _issue_row(connection, issue_id)
        if issue is None or issue["status"] != "open":
            raise LookupError("Incidencia no encontrada")
        if issue["code"] != _INVOICE_OUTSIDE_PERIOD_CODE:
            raise LookupError("La incidencia no corresponde a una factura fuera del período")
        connection.execute(
            """INSERT INTO manual_corrections
               (id_issue,original_value,corrected_value,reason,resolved_by)
               VALUES (?,?,?,?,?)""",
            (
                issue_id, issue["detected_value"], "no_corresponde_al_periodo",
                normalized_reason, normalized_actor,
            ),
        )
        connection.execute(
            """UPDATE review_issues SET status='dismissed',resolved_at=datetime('now')
               WHERE id_issue=?""",
            (issue_id,),
        )
        dismissed = _issue_row(connection, issue_id)
    return review_issue_from_row(dismissed)


def assert_case_final_readings_approved(
    connection: sqlite3.Connection, case_id: int
) -> None:
    """Impide avanzar si una lectura final irregular sigue sin aprobación.

    Esta guarda se aplica aunque el perfil no reparta un concepto por consumo:
    así una incidencia cerrada indebidamente no puede convertir el expediente
    en calculable dejando una lectura canónica pendiente.
    """
    case = connection.execute(
        """SELECT id_comunidad,id_periodo,fecha_fin FROM regularization_cases
           WHERE id_case=?""",
        (case_id,),
    ).fetchone()
    if case is None:
        raise LookupError("Expediente no encontrado")
    if case["id_periodo"] is None:
        return
    unresolved = connection.execute(
        """SELECT COUNT(*) FROM lecturas_vecino AS reading
           JOIN propietarios AS owner ON owner.id_propietario=reading.id_propietario
           WHERE owner.id_comunidad=? AND reading.id_periodo=?
             AND reading.fecha_lectura=?
             AND (
                 reading.estado='contador_averiado'
                 OR (
                     reading.estado='estimado'
                     AND (
                         trim(COALESCE(reading.approved_by,''))=''
                         OR trim(COALESCE(reading.approved_at,''))=''
                     )
                 )
             )""",
        (case["id_comunidad"], case["id_periodo"], case["fecha_fin"]),
    ).fetchone()[0]
    if unresolved:
        raise ValueError(
            "Hay una lectura final de contador pendiente de estimación aprobada"
        )


def validate_case_ready(connection: sqlite3.Connection, case_id: int) -> RegularizationCase:
    with _transaction(connection):
        open_count = connection.execute(
            "SELECT COUNT(*) FROM review_issues WHERE id_case = ? AND status = 'open'",
            (case_id,),
        ).fetchone()[0]
        if open_count:
            raise ValueError(f"{open_count} incidencia abierta(s) por resolver")

        assert_case_final_readings_approved(connection, case_id)
        case = get_case(connection, case_id)
        if case.status == "draft":
            case = set_case_status(connection, case_id, "gathering_sources")
        if case.status == "gathering_sources":
            case = set_case_status(connection, case_id, "under_review")
        if case.status == "under_review":
            case = set_case_status(connection, case_id, "ready_for_calculation")
        elif case.status != "ready_for_calculation":
            case = set_case_status(connection, case_id, "ready_for_calculation")

        connection.execute(
            "UPDATE source_documents SET status = 'validated' WHERE id_case = ?",
            (case_id,),
        )
    return case
