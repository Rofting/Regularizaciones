import sqlite3
from dataclasses import dataclass
from pathlib import Path
from typing import Collection, Mapping

import document_review
import expedient_service
from expedient_models import RegularizationCase, SourceDocument


@dataclass(frozen=True)
class IngestionResult:
    document: SourceDocument
    created: bool
    open_issue_count: int


def _open_issue_count(connection: sqlite3.Connection, case_id: int) -> int:
    return connection.execute(
        """SELECT COUNT(*) FROM review_issues
           WHERE id_case = ? AND status = 'open'""",
        (case_id,),
    ).fetchone()[0]


def count_case_documents(connection: sqlite3.Connection, case_id: int) -> int:
    return connection.execute(
        "SELECT COUNT(*) FROM source_documents WHERE id_case = ?",
        (case_id,),
    ).fetchone()[0]


def assert_case_belongs_to_community(
    connection: sqlite3.Connection,
    case_id: int,
    community_id: int,
) -> RegularizationCase:
    case = expedient_service.get_case(connection, case_id)
    if case.community_id != community_id:
        raise LookupError("El expediente seleccionado pertenece a otra comunidad")
    return case


def add_document_to_case(
    connection: sqlite3.Connection,
    case_id: int,
    *,
    source_path: str | Path,
    archive_root: str | Path,
    document_kind: str,
    candidates: Mapping[str, str | None],
    required_fields: Collection[str],
) -> IngestionResult:
    normalized_kind = document_kind.strip()
    if not normalized_kind:
        raise ValueError("El tipo de documento es obligatorio")

    normalized_candidates = {key.strip(): value for key, value in candidates.items()}
    normalized_required = tuple(field.strip() for field in required_fields)
    if any(not field for field in normalized_required):
        raise ValueError("Los nombres de campos requeridos no pueden estar vacíos")

    document, created = expedient_service.register_source_document(
        connection,
        case_id,
        source_path=source_path,
        archive_root=archive_root,
        document_kind=normalized_kind,
    )
    if not created:
        return IngestionResult(document, False, _open_issue_count(connection, case_id))

    document_review.record_candidates(
        connection,
        document.id_document,
        normalized_candidates,
        source="ingestion",
    )
    document_review.create_missing_field_issues(
        connection,
        case_id,
        document.id_document,
        normalized_required,
    )
    return IngestionResult(document, True, _open_issue_count(connection, case_id))
