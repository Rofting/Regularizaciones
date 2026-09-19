import re
import json
import sqlite3
from contextlib import contextmanager
from datetime import datetime
from decimal import Decimal, InvalidOperation
from itertools import count
from typing import Collection, Iterator, Mapping

from expedient_models import RegularizationCase, ReviewIssue, review_issue_from_row
from expedient_service import get_case, link_case_to_period, set_case_status


_savepoint_counter = count()
_MISSING_FIELD_CODE = "MISSING_REQUIRED_FIELD"
_COUNTER_RESET_CODE = "COUNTER_RESET"
_INVOICE_OUTSIDE_PERIOD_CODE = "INVOICE_OUTSIDE_PERIOD"
_CLASSIFICATION_REQUIRED_CODE = "DOCUMENT_CLASSIFICATION_REQUIRED"
_ARCHIVED_SOURCE_DUPLICATE_CODE = "ARCHIVED_SOURCE_DUPLICATE"
_ARCHIVED_SOURCE_MISSING_CODE = "ARCHIVED_SOURCE_MISSING"
_AUTOMATIC_REVIEW_CODES = (
    _MISSING_FIELD_CODE,
    _CLASSIFICATION_REQUIRED_CODE,
    _ARCHIVED_SOURCE_DUPLICATE_CODE,
    _ARCHIVED_SOURCE_MISSING_CODE,
)
_VALIDATION_STATUSES = {"candidate", "validated", "rejected"}
_ISSUE_ORIGINS = {"automatic", "manual"}


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


def _validate_issue_value(field_name: str, value: str) -> str:
    """Reject placeholders before they can become a manual canonical value."""
    if field_name in {"fecha_inicio", "fecha_fin", "fecha_factura"}:
        for pattern in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
            try:
                return datetime.strptime(value, pattern).date().isoformat()
            except ValueError:
                continue
        raise ValueError("Introduce una fecha válida (dd/mm/aaaa).")
    if field_name in {"importe_total", "consumo_total", "consumo_kwh", "consumo_m3"}:
        try:
            amount = Decimal(value.replace(".", "").replace(",", "."))
        except InvalidOperation:
            raise ValueError("Introduce un importe o consumo numérico válido.") from None
        if not amount.is_finite():
            raise ValueError("Introduce un importe o consumo numérico válido.")
    if field_name == "tipo_suministro" and (not re.search(r"[A-Za-zÁÉÍÓÚáéíóú]", value)):
        raise ValueError("Indica un tipo de suministro válido, por ejemplo GAS, ACS o MANTENIMIENTO.")
    if field_name == "tipo":
        reading_type = value.strip().upper()
        if reading_type not in {"ACS", "CALEFACCION"}:
            raise ValueError("El tipo de lectura debe ser ACS o CALEFACCION.")
        return reading_type
    if field_name == "document_kind" and value.casefold() not in {"invoice", "reading", "owners", "other"}:
        raise ValueError("El tipo de documento debe ser factura, lectura, propietarios u otro.")
    return value


def validate_candidate_value(field_name: str, value: str) -> str:
    """Aplica a una extracción las mismas reglas usadas al resolver incidencias."""
    return _validate_issue_value(field_name, value)


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
    field_name = str(issue["field_name"]).partition("|")[0]
    prefix, separator, remainder = field_name.partition("reading.")
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
    target = connection.execute("""SELECT a.fecha_lectura AS initial_date,b.fecha_lectura AS final_date
        FROM counter_reset_targets t
        JOIN lecturas_vecino a ON a.id_lectura=t.initial_reading_id
        JOIN lecturas_vecino b ON b.id_lectura=t.final_reading_id
        WHERE t.id_issue=? AND a.id_propietario=? AND b.id_propietario=?
          AND a.tipo=? AND b.tipo=?""",
        (issue['id_issue'], owner['id_propietario'], owner['id_propietario'], service, service)).fetchone()
    if target is None and connection.execute(
        "SELECT 1 FROM counter_reset_targets WHERE id_issue=?", (issue['id_issue'],)
    ).fetchone():
        raise LookupError("La pareja de lecturas de la incidencia no pertenece al propietario y servicio")
    initial_date = target['initial_date'] if target else case['fecha_inicio']
    final_date = target['final_date'] if target else case['fecha_fin']
    readings = connection.execute(
        """SELECT id_lectura,fecha_lectura,valor_acumulado,estado,metodo_estimacion,
                  fuente,notas,approved_by,approved_at
           FROM period_readings
           WHERE id_propietario=? AND id_periodo=? AND tipo=?
             AND fecha_lectura IN (?,?)
           ORDER BY fecha_lectura""",
        (
            owner["id_propietario"], case["id_periodo"], service,
            initial_date, final_date,
        ),
    ).fetchall()
    by_date = {row["fecha_lectura"]: row for row in readings}
    initial = by_date.get(initial_date)
    final = by_date.get(final_date)
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


def record_candidates(
    connection: sqlite3.Connection,
    document_id: int,
    candidates: Mapping[str, str | None],
    *,
    source: str,
    source_context: str | None = None,
    validation_status: str | None = None,
    preserve_validated: bool = False,
) -> None:
    """Guarda extracciones, sin reemplazar correcciones cuando se solicita."""
    normalized_source = source.strip()
    if not normalized_source:
        raise ValueError("La fuente del candidato es obligatoria")
    if validation_status is not None and validation_status not in _VALIDATION_STATUSES:
        raise ValueError("El estado de validación del candidato no es válido")

    with _transaction(connection):
        for field_name, value in candidates.items():
            normalized_value = _normalise_value(value)
            candidate_status = validation_status or (
                "validated" if normalized_value is not None else "candidate"
            )
            connection.execute(
                """INSERT INTO extraction_candidates
                   (id_document, field_name, value, source, validation_status, source_context)
                   VALUES (?, ?, ?, ?, ?, ?)
                   ON CONFLICT(id_document, field_name) DO UPDATE SET
                       value = excluded.value,
                       source = excluded.source,
                       validation_status = excluded.validation_status,
                       source_context = excluded.source_context
                   WHERE ? = 0
                      OR extraction_candidates.validation_status NOT IN ('validated', 'rejected')""",
                (
                    document_id, field_name, normalized_value, normalized_source, candidate_status,
                    source_context, int(preserve_validated),
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
                       (id_case, id_document, code, field_name, detected_value, message, status, origin)
                       VALUES (?, ?, ?, ?, NULL, ?, 'open', 'automatic')
                       ON CONFLICT(id_document, code, field_name) WHERE status='open' DO NOTHING""",
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


def _create_review_issue(
    connection: sqlite3.Connection,
    case_id: int,
    document_id: int,
    *,
    code: str,
    field_name: str,
    message: str,
    detected_value: str | None = None,
    origin: str,
) -> ReviewIssue:
    """Registra de forma idempotente una incidencia con procedencia explícita."""
    normalized_code = code.strip()
    normalized_field = field_name.strip()
    normalized_message = message.strip()
    if not normalized_code or not normalized_field or not normalized_message:
        raise ValueError("Código, campo y mensaje de la incidencia son obligatorios")
    if origin not in _ISSUE_ORIGINS:
        raise ValueError("El origen de la incidencia no es válido")

    with _transaction(connection):
        document = connection.execute(
            "SELECT id_case FROM source_documents WHERE id_document = ?", (document_id,)
        ).fetchone()
        if document is None or document["id_case"] != case_id:
            raise LookupError("El documento no pertenece al expediente")
        connection.execute(
            """INSERT INTO review_issues
               (id_case,id_document,code,field_name,detected_value,message,status,origin)
               VALUES (?,?,?,?,?,?,'open',?)
               ON CONFLICT(id_document,code,field_name) WHERE status='open' DO NOTHING""",
            (
                case_id,
                document_id,
                normalized_code,
                normalized_field,
                _normalise_value(detected_value),
                normalized_message,
                origin,
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
    """Registra una incidencia manual; sólo el análisis crea incidencias automáticas."""
    return _create_review_issue(
        connection,
        case_id,
        document_id,
        code=code,
        field_name=field_name,
        message=message,
        detected_value=detected_value,
        origin="manual",
    )


def create_invalid_field_issue(
    connection: sqlite3.Connection,
    case_id: int,
    document_id: int,
    *,
    field_name: str,
    message: str,
    detected_value: str | None = None,
) -> ReviewIssue:
    """Materializa como automática una extracción presente pero no utilizable."""
    return _create_review_issue(
        connection,
        case_id,
        document_id,
        code=_MISSING_FIELD_CODE,
        field_name=field_name,
        message=message,
        detected_value=detected_value,
        origin="automatic",
    )


def create_classification_required_issue(
    connection: sqlite3.Connection,
    case_id: int,
    document_id: int,
    *,
    message: str,
) -> ReviewIssue:
    """Crea la única incidencia abierta necesaria para una fuente desconocida."""
    return _create_review_issue(
        connection,
        case_id,
        document_id,
        code=_CLASSIFICATION_REQUIRED_CODE,
        field_name="document_kind",
        message=message,
        origin="automatic",
    )


def create_archived_source_issue(
    connection: sqlite3.Connection,
    case_id: int,
    document_id: int,
    *,
    duplicate_count: int,
) -> ReviewIssue:
    """Registra una ruta archivada no disponible sin inventar una copia válida."""
    if duplicate_count > 1:
        code = _ARCHIVED_SOURCE_DUPLICATE_CODE
        message = (
            f"Hay {duplicate_count} copias verificadas con la misma huella. "
            "Elige la copia correcta antes de reanalizar esta fuente."
        )
    else:
        code = _ARCHIVED_SOURCE_MISSING_CODE
        message = (
            "No se localiza la copia archivada con la misma huella. "
            "Vuelve a añadir el archivo original para poder analizarlo."
        )
    return _create_review_issue(
        connection,
        case_id,
        document_id,
        code=code,
        field_name="archived_path",
        message=message,
        origin="automatic",
    )


def resolved_classification_kind(
    connection: sqlite3.Connection, document_id: int
) -> str | None:
    """Devuelve una clasificación confirmada manualmente, si existe."""
    row = connection.execute(
        """SELECT candidates.value FROM review_issues AS issues
           JOIN extraction_candidates AS candidates
             ON candidates.id_document = issues.id_document
            AND candidates.field_name = issues.field_name
           WHERE issues.id_document = ?
             AND issues.code = ? AND issues.field_name = 'document_kind'
             AND issues.status = 'resolved'
             AND candidates.source = 'manual'
             AND candidates.validation_status = 'validated'
             AND candidates.value IS NOT NULL AND trim(candidates.value) <> ''""",
        (document_id, _CLASSIFICATION_REQUIRED_CODE),
    ).fetchone()
    return row["value"] if row is not None else None


def has_closed_classification_outcome(
    connection: sqlite3.Connection, document_id: int
) -> bool:
    """Una decisión cerrada no debe convertirse de nuevo en un bloqueo automático."""
    return connection.execute(
        """SELECT 1 FROM review_issues
           WHERE id_document = ? AND code = ? AND field_name = 'document_kind'
             AND status IN ('resolved', 'dismissed')""",
        (document_id, _CLASSIFICATION_REQUIRED_CODE),
    ).fetchone() is not None


def clear_open_automatic_issues(
    connection: sqlite3.Connection,
    case_id: int,
    document_id: int,
) -> None:
    """Elimina sólo incidencias abiertas que el análisis puede regenerar."""
    with _transaction(connection):
        document = connection.execute(
            "SELECT id_case FROM source_documents WHERE id_document = ?", (document_id,)
        ).fetchone()
        if document is None or document["id_case"] != case_id:
            raise LookupError("El documento no pertenece al expediente")
        placeholders = ", ".join("?" for _ in _AUTOMATIC_REVIEW_CODES)
        connection.execute(
            f"""DELETE FROM review_issues
                WHERE id_case = ? AND id_document = ? AND status = 'open'
                  AND origin = 'automatic' AND code IN ({placeholders})""",
            (case_id, document_id, *_AUTOMATIC_REVIEW_CODES),
        )


def resolve_archived_source_issue(
    connection: sqlite3.Connection,
    case_id: int,
    document_id: int,
) -> int:
    """Cierra las incidencias de archivo al recuperar una ruta verificada."""
    with _transaction(connection):
        document = connection.execute(
            "SELECT id_case FROM source_documents WHERE id_document = ?", (document_id,)
        ).fetchone()
        if document is None or document["id_case"] != case_id:
            raise LookupError("El documento no pertenece al expediente")
        result = connection.execute(
            """UPDATE review_issues SET status='resolved',resolved_at=datetime('now')
               WHERE id_case=? AND id_document=? AND status='open' AND origin='automatic'
                 AND code IN (?, ?)""",
            (
                case_id,
                document_id,
                _ARCHIVED_SOURCE_DUPLICATE_CODE,
                _ARCHIVED_SOURCE_MISSING_CODE,
            ),
        )
    return result.rowcount


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


def _normalise_reading_property(value: str) -> str:
    return re.sub(r"[\s\.º°]", "", value.upper())


def _resolve_reading_conflict(
    connection: sqlite3.Connection,
    issue: sqlite3.Row,
    value: str,
    *,
    reason: str,
    resolved_by: str,
) -> str:
    """Makes a reviewed reading authoritative before reapplying its source."""
    match = re.fullmatch(r"reading\.(.+)\.(ACS|CALEFACCION)", str(issue["field_name"]))
    if match is None:
        raise LookupError("La incidencia de lectura no identifica una vivienda válida")
    try:
        confirmed_value = Decimal(value.replace(",", "."))
    except InvalidOperation:
        raise ValueError("Introduce una lectura acumulada numérica válida.") from None
    if not confirmed_value.is_finite() or confirmed_value < 0:
        raise ValueError("Introduce una lectura acumulada numérica válida.")

    property_code, service = match.groups()
    case = get_case(connection, issue["id_case"])
    owners = connection.execute(
        """SELECT DISTINCT owner.id_propietario,owner.codigo_vivienda
           FROM reading_observations AS observation
           JOIN propietarios AS owner ON owner.id_propietario=observation.id_propietario
           WHERE observation.id_document=? AND observation.tipo=?
             AND owner.id_comunidad=?""",
        (issue["id_document"], service, case.community_id),
    ).fetchall()
    owner = next(
        (item for item in owners if _normalise_reading_property(item["codigo_vivienda"])
         == _normalise_reading_property(property_code)),
        None,
    )
    if owner is None:
        raise LookupError("La vivienda de la incidencia no pertenece al expediente")
    observation = connection.execute(
        """SELECT id_observation,fecha_lectura FROM reading_observations
           WHERE id_document=? AND id_propietario=? AND tipo=? AND status='conflict'
           ORDER BY id_observation DESC LIMIT 1""",
        (issue["id_document"], owner["id_propietario"], service),
    ).fetchone()
    if observation is None:
        raise LookupError("No se localiza la lectura en conflicto para corregirla")
    reading = connection.execute(
        """SELECT id_lectura FROM lecturas_vecino
           WHERE id_propietario=? AND tipo=? AND fecha_lectura=?""",
        (owner["id_propietario"], service, observation["fecha_lectura"]),
    ).fetchone()
    if reading is None:
        raise LookupError("No se localiza la lectura canónica en conflicto")

    rows_candidate = connection.execute(
        "SELECT value FROM extraction_candidates WHERE id_document=? AND field_name='vecinos'",
        (issue["id_document"],),
    ).fetchone()
    if rows_candidate is None:
        raise LookupError("La fuente no conserva las filas de lecturas para corregirla")
    try:
        source_rows = json.loads(rows_candidate["value"])
    except (TypeError, json.JSONDecodeError):
        raise LookupError("Las filas de lecturas de la fuente no tienen un formato válido") from None
    if not isinstance(source_rows, list):
        raise LookupError("Las filas de lecturas de la fuente no tienen un formato válido")

    defaults = {
        row["field_name"]: row["value"]
        for row in connection.execute(
            """SELECT field_name,value FROM extraction_candidates
               WHERE id_document=? AND field_name IN ('tipo','fecha_inicio','fecha_fin')""",
            (issue["id_document"],),
        )
    }
    updated = False
    for source_row in source_rows:
        if not isinstance(source_row, dict):
            continue
        if _normalise_reading_property(str(source_row.get("vivienda") or "")) != _normalise_reading_property(property_code):
            continue
        source_service = str(source_row.get("tipo") or defaults.get("tipo") or "").upper()
        if source_service != service:
            continue
        initial_date = str(source_row.get("fecha_ant") or defaults.get("fecha_inicio") or "")
        final_date = str(source_row.get("fecha_act") or defaults.get("fecha_fin") or "")
        if observation["fecha_lectura"] == initial_date:
            source_row["val_ant"] = float(confirmed_value)
        elif observation["fecha_lectura"] == final_date:
            source_row["val_act"] = float(confirmed_value)
        else:
            continue
        updated = True
        break
    if not updated:
        raise LookupError("No se localiza la fila de la lectura en la fuente")

    corrected_text = _decimal_text(confirmed_value)
    connection.execute(
        """UPDATE extraction_candidates
           SET value=?,source='manual',validation_status='validated'
           WHERE id_document=? AND field_name='vecinos'""",
        (json.dumps(source_rows, ensure_ascii=False), issue["id_document"]),
    )
    connection.execute(
        """UPDATE lecturas_vecino
           SET valor_acumulado=?,estado='real',metodo_estimacion='manual_conflict_resolution',
               notas=?,approved_by=?,approved_at=datetime('now')
           WHERE id_lectura=?""",
        (
            corrected_text,
            f"Lectura confirmada manualmente; motivo={reason}",
            resolved_by,
            reading["id_lectura"],
        ),
    )
    connection.execute(
        """UPDATE reading_observations
           SET status='observed',effective_reading_id=?,previous_reading_id=NULL
           WHERE id_observation=?""",
        (reading["id_lectura"], observation["id_observation"]),
    )
    return corrected_text


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
        normalized_value = _validate_issue_value(str(issue["field_name"]), normalized_value)
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
        if issue["code"] == "READING_CONFLICT":
            normalized_value = _resolve_reading_conflict(
                connection, issue, normalized_value,
                reason=normalized_reason, resolved_by=normalized_resolved_by,
            )
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
        if issue["code"] == _CLASSIFICATION_REQUIRED_CODE:
            next_status = "not_applicable" if normalized_value == "other" else "under_review"
            connection.execute(
                """UPDATE source_documents
                   SET document_kind=?, status=?, confirmed_by=NULL, confirmed_at=NULL
                   WHERE id_document=?""",
                (normalized_value, next_status, issue["id_document"]),
            )
        connection.execute(
            """UPDATE review_issues SET status = 'resolved', resolved_at = datetime('now')
               WHERE id_issue = ?""",
            (issue_id,),
        )
        resolved = _issue_row(connection, issue_id)
        _reapply_reviewed_document(connection, issue['id_case'], issue['id_document'])
    return review_issue_from_row(resolved)


def _reapply_reviewed_document(connection, case_id, document_id):
    document = connection.execute("SELECT confirmed_by FROM source_documents WHERE id_document=?",
                                  (document_id,)).fetchone()
    marker = connection.execute("SELECT 1 FROM archivos_procesados WHERE nombre_archivo=?",
                                (f'source_document:{document_id}',)).fetchone()
    if document['confirmed_by'] or marker:
        from case_ingestion import apply_confirmed_source
        apply_confirmed_source(connection, case_id, document_id)


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
        owner_state = "estado_contador_acs" if issue["field_name"].partition("|")[0].endswith(".ACS") else "estado_contador_cal"
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
        _reapply_reviewed_document(connection, issue['id_case'], issue['id_document'])
    return review_issue_from_row(resolved)


def carry_forward_counter_resets_for_source(
    connection: sqlite3.Connection,
    *,
    case_id: int,
    document_id: int,
    reason: str,
    approved_by: str,
) -> int:
    """Aprueba en bloque el criterio temporal para reinicios de una fuente.

    Conserva la última lectura acumulada conocida de cada contador afectado,
    por lo que el consumo provisional del intervalo es cero. Cada corrección
    queda auditada, pero el gestor toma una sola decisión para el documento.
    """
    normalized_reason = reason.strip()
    normalized_approver = approved_by.strip()
    if not normalized_reason:
        raise ValueError("El motivo del criterio temporal es obligatorio")
    if not normalized_approver:
        raise ValueError("La persona que aprueba el criterio temporal es obligatoria")

    with _transaction(connection):
        rows = connection.execute(
            """SELECT id_issue FROM review_issues
               WHERE id_case=? AND id_document=? AND code=? AND status='open'
               ORDER BY id_issue""",
            (case_id, document_id, _COUNTER_RESET_CODE),
        ).fetchall()
        if not rows:
            return 0
        for row in rows:
            issue = _issue_row(connection, row["id_issue"])
            if issue is None:
                raise LookupError("Incidencia de reinicio no encontrada")
            _case, owner, initial, final = _counter_reset_target(connection, issue)
            original_text = _decimal_text(Decimal(str(final["valor_acumulado"])))
            corrected_text = _decimal_text(Decimal(str(initial["valor_acumulado"])))
            notes = (
                f"Lectura previa conservada tras reinicio masivo; valor original={original_text}; "
                f"consumo provisional=0; lectura virtual={corrected_text}; "
                f"motivo={normalized_reason}; aprobada por={normalized_approver}"
            )
            if final["notas"]:
                notes = f"{final['notas']} | {notes}"
            connection.execute(
                """INSERT INTO manual_corrections
                   (id_issue,original_value,corrected_value,reason,resolved_by)
                   VALUES (?,?,?,?,?)""",
                (issue["id_issue"], original_text, corrected_text, normalized_reason, normalized_approver),
            )
            connection.execute(
                """INSERT INTO extraction_candidates
                   (id_document,field_name,value,source,validation_status)
                   VALUES (?, ?, ?, 'manual_counter_reset_carry_forward', 'validated')
                   ON CONFLICT(id_document,field_name) DO UPDATE SET
                       value=excluded.value,source=excluded.source,
                       validation_status=excluded.validation_status""",
                (issue["id_document"], issue["field_name"], corrected_text),
            )
            connection.execute(
                """UPDATE lecturas_vecino
                   SET valor_acumulado=?,estado='estimado',metodo_estimacion='counter_reset_carry_forward',
                       notas=?,approved_by=?,approved_at=datetime('now')
                   WHERE id_lectura=?""",
                (corrected_text, notes, normalized_approver, final["id_lectura"]),
            )
            owner_state = (
                "estado_contador_acs"
                if issue["field_name"].partition("|")[0].endswith(".ACS")
                else "estado_contador_cal"
            )
            connection.execute(
                f"UPDATE propietarios SET {owner_state}='ok' WHERE id_propietario=?",
                (owner["id_propietario"],),
            )
            connection.execute(
                "UPDATE review_issues SET status='resolved',resolved_at=datetime('now') WHERE id_issue=?",
                (issue["id_issue"],),
            )
        _reapply_reviewed_document(connection, case_id, document_id)
    return len(rows)


def confirm_initial_zero_readings_for_source(
    connection: sqlite3.Connection,
    *,
    case_id: int,
    document_id: int,
    reason: str,
    approved_by: str,
) -> int:
    """Confirma en bloque ceros iniciales de una fuente y reanuda su aplicación."""
    normalized_reason = reason.strip()
    normalized_approver = approved_by.strip()
    if not normalized_reason:
        raise ValueError("El motivo para confirmar los ceros iniciales es obligatorio")
    if not normalized_approver:
        raise ValueError("La persona que confirma los ceros iniciales es obligatoria")

    with _transaction(connection):
        case = get_case(connection, case_id)
        period_id = link_case_to_period(connection, case_id)
        document = connection.execute(
            "SELECT id_case,archived_path FROM source_documents WHERE id_document=?",
            (document_id,),
        ).fetchone()
        if document is None or document["id_case"] != case_id:
            raise LookupError("El documento no pertenece al expediente")
        rows = connection.execute(
            """SELECT id_issue,field_name FROM review_issues
               WHERE id_case=? AND id_document=? AND code='READING_ZERO_REVIEW'
                 AND status='open' ORDER BY id_issue""",
            (case_id, document_id),
        ).fetchall()
        if not rows:
            return 0

        for issue in rows:
            match = re.fullmatch(r"reading\.(.+)\.(ACS|CALEFACCION)", str(issue["field_name"]))
            if match is None:
                raise LookupError("La incidencia no apunta a una lectura inicial válida")
            property_code, service = match.groups()
            owner = connection.execute(
                "SELECT id_propietario FROM propietarios WHERE id_comunidad=? AND codigo_vivienda=?",
                (case.community_id, property_code),
            ).fetchone()
            if owner is None:
                raise LookupError("La vivienda de la lectura no pertenece al expediente")
            observation = connection.execute(
                """SELECT id_observation,fecha_lectura FROM reading_observations
                   WHERE id_document=? AND id_propietario=? AND tipo=?
                     AND observed_value=0
                   ORDER BY fecha_lectura,id_observation LIMIT 1""",
                (document_id, owner["id_propietario"], service),
            ).fetchone()
            if observation is None:
                raise LookupError("No existe una observación inicial a cero pendiente")
            current = connection.execute(
                """SELECT id_lectura,valor_acumulado,estado,metodo_estimacion,approved_by,approved_at
                   FROM lecturas_vecino
                   WHERE id_propietario=? AND tipo=? AND fecha_lectura=?""",
                (owner["id_propietario"], service, observation["fecha_lectura"]),
            ).fetchone()
            approved_counter_reset = current is not None and (
                current["estado"] == "estimado"
                and current["metodo_estimacion"] == "counter_reset_carry_forward"
                and str(current["approved_by"] or "").strip()
                and str(current["approved_at"] or "").strip()
            )
            if (
                current is not None
                and float(current["valor_acumulado"]) != 0
                and current["estado"] != "real"
                and not approved_counter_reset
            ):
                raise ValueError("Ya existe una lectura canónica distinta de cero para esta fecha")
            if current is None:
                cursor = connection.execute(
                    """INSERT INTO lecturas_vecino
                       (id_propietario,id_periodo,tipo,fecha_lectura,valor_acumulado,estado,
                        metodo_estimacion,fuente,notas,approved_by,approved_at)
                       VALUES (?,?,?,?,0,'real','confirmed_initial_zero',?,?,?,datetime('now'))""",
                    (
                        owner["id_propietario"], period_id, service, observation["fecha_lectura"],
                        str(document["archived_path"]),
                        f"Lectura inicial a 0 confirmada en bloque; motivo={normalized_reason}",
                        normalized_approver,
                    ),
                )
                reading_id = int(cursor.lastrowid)
                observation_status = "observed"
                corrected_value = "0"
            else:
                reading_id = int(current["id_lectura"])
                observation_status = "observed" if float(current["valor_acumulado"]) == 0 else "carried_forward"
                corrected_value = _decimal_text(Decimal(str(current["valor_acumulado"])))
            connection.execute(
                "INSERT OR IGNORE INTO reading_periods(id_lectura,id_periodo) VALUES (?,?)",
                (reading_id, period_id),
            )
            connection.execute(
                """UPDATE reading_observations
                   SET status=?,effective_reading_id=?,previous_reading_id=NULL
                   WHERE id_observation=?""",
                (observation_status, reading_id, observation["id_observation"]),
            )
            connection.execute(
                """INSERT INTO manual_corrections
                   (id_issue,original_value,corrected_value,reason,resolved_by)
                   VALUES (?,?,?,?,?)""",
                (issue["id_issue"], "0", corrected_value, normalized_reason, normalized_approver),
            )
            connection.execute(
                """UPDATE review_issues SET status='resolved',resolved_at=datetime('now')
                   WHERE id_issue=?""",
                (issue["id_issue"],),
            )

        from case_ingestion import apply_confirmed_source
        apply_confirmed_source(connection, case_id, document_id)
    return len(rows)


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
        """SELECT COUNT(*) FROM period_readings AS reading
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

        if case_has_unapplied_sources(connection, case_id):
            raise ValueError("Hay fuentes pendientes de confirmar y aplicar a los datos canónicos")

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
            """UPDATE source_documents SET status = 'validated' WHERE id_case = ?
               AND (classification_confidence IS NULL OR document_kind='other')""",
            (case_id,),
        )
    return case


def case_has_unapplied_sources(connection: sqlite3.Connection, case_id: int) -> bool:
    """Analysed sources may only pass readiness after canonical application."""
    return connection.execute("""SELECT 1 FROM source_documents d
        WHERE d.id_case=? AND d.classification_confidence IS NOT NULL
          AND d.document_kind<>'other' AND (
            d.status<>'validated' OR d.document_kind='unknown'
            OR NOT EXISTS (SELECT 1 FROM archivos_procesados a
                           WHERE a.nombre_archivo='source_document:' || d.id_document)
            OR EXISTS (SELECT 1 FROM extraction_candidates c WHERE c.id_document=d.id_document
                       AND c.validation_status='candidate' AND c.value IS NOT NULL AND trim(c.value)<>'')
          ) LIMIT 1""", (case_id,)).fetchone() is not None
