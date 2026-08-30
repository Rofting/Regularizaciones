from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Literal, cast


CaseStatus = Literal[
    "draft", "gathering_sources", "under_review", "ready_for_calculation",
    "calculated", "reconciled", "deliveries_generated", "closed",
]
DocumentStatus = Literal["registered", "under_review", "validated", "not_applicable"]
IssueStatus = Literal["open", "resolved", "dismissed"]


@dataclass(frozen=True)
class RegularizationCase:
    id_case: int
    community_id: int
    name: str
    start_date: date
    end_date: date
    status: CaseStatus


@dataclass(frozen=True)
class SourceDocument:
    id_document: int
    id_case: int
    original_name: str
    archived_path: Path
    sha256: str
    document_kind: str
    status: DocumentStatus


@dataclass(frozen=True)
class ReviewIssue:
    id_issue: int
    id_case: int
    id_document: int
    archived_path: Path
    code: str
    field_name: str
    detected_value: str | None
    message: str
    status: IssueStatus


def case_from_row(row) -> RegularizationCase:
    """Convierte una fila de ``regularization_cases`` al modelo público."""
    return RegularizationCase(
        id_case=row["id_case"],
        community_id=row["id_comunidad"],
        name=row["nombre"],
        start_date=date.fromisoformat(row["fecha_inicio"]),
        end_date=date.fromisoformat(row["fecha_fin"]),
        status=cast(CaseStatus, row["estado"]),
    )


def document_from_row(row) -> SourceDocument:
    """Convierte una fila de ``source_documents`` al modelo público."""
    return SourceDocument(
        id_document=row["id_document"],
        id_case=row["id_case"],
        original_name=row["original_name"],
        archived_path=Path(row["archived_path"]),
        sha256=row["sha256"],
        document_kind=row["document_kind"],
        status=cast(DocumentStatus, row["status"]),
    )


def review_issue_from_row(row) -> ReviewIssue:
    """Convierte una fila de incidencia, incluida su fuente archivada."""
    return ReviewIssue(
        id_issue=row["id_issue"],
        id_case=row["id_case"],
        id_document=row["id_document"],
        archived_path=Path(row["archived_path"]),
        code=row["code"],
        field_name=row["field_name"],
        detected_value=row["detected_value"],
        message=row["message"],
        status=cast(IssueStatus, row["status"]),
    )
