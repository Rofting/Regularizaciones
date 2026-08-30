import sqlite3
from contextlib import contextmanager
from itertools import count
from typing import Collection, Iterator, Mapping

from expedient_models import RegularizationCase, ReviewIssue, review_issue_from_row
from expedient_service import get_case, set_case_status


_savepoint_counter = count()
_MISSING_FIELD_CODE = "MISSING_REQUIRED_FIELD"


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


def validate_case_ready(connection: sqlite3.Connection, case_id: int) -> RegularizationCase:
    with _transaction(connection):
        open_count = connection.execute(
            "SELECT COUNT(*) FROM review_issues WHERE id_case = ? AND status = 'open'",
            (case_id,),
        ).fetchone()[0]
        if open_count:
            raise ValueError(f"{open_count} incidencia abierta(s) por resolver")

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
