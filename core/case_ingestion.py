import sqlite3
import json
from contextlib import contextmanager
from dataclasses import dataclass
from itertools import count
from pathlib import Path
from typing import Callable, Collection, Iterator, Mapping

import document_review
import expedient_service
from expedient_models import RegularizationCase, SourceDocument, document_from_row
from source_analysis import SourceAnalysis, analyse_source


_savepoint_counter = count()


@contextmanager
def _transaction(connection: sqlite3.Connection) -> Iterator[None]:
    """Agrupa persistencia, candidatos e incidencias en una sola operación."""
    if connection.in_transaction:
        savepoint = f"case_ingestion_{next(_savepoint_counter)}"
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


@dataclass(frozen=True)
class IngestionResult:
    document: SourceDocument
    created: bool
    open_issue_count: int


def _compact_context(analysis: SourceAnalysis) -> str | None:
    locator = analysis.locator
    if locator is None:
        return None
    values = {
        key: value for key, value in (
            ("page", locator.page),
            ("fragment", locator.fragment),
            ("sheet", locator.sheet),
            ("cell", locator.cell),
        ) if value is not None
    }
    return json.dumps(values, ensure_ascii=False, separators=(",", ":")) or None


def _normalised_analysis(analysis: SourceAnalysis) -> SourceAnalysis:
    if not isinstance(analysis, SourceAnalysis):
        raise TypeError("El análisis de la fuente debe ser SourceAnalysis")
    if not analysis.kind.strip() or not analysis.confidence.strip():
        raise ValueError("El análisis debe incluir tipo y confianza")
    if any(not field.strip() for field in analysis.candidates):
        raise ValueError("Los nombres de campos extraídos no pueden estar vacíos")
    if any(not field.strip() for field in analysis.required_fields):
        raise ValueError("Los campos requeridos no pueden estar vacíos")
    return analysis


def _replace_unvalidated_candidates_not_in_analysis(
    connection: sqlite3.Connection,
    document_id: int,
    candidates: Mapping[str, str | None],
) -> None:
    """Evita que valores automáticos obsoletos bloqueen una nueva revisión."""
    candidate_names = {field.strip() for field in candidates}
    rows = connection.execute(
        """SELECT field_name FROM extraction_candidates
           WHERE id_document = ?
             AND validation_status NOT IN ('validated', 'rejected')""",
        (document_id,),
    ).fetchall()
    for row in rows:
        if row["field_name"] not in candidate_names:
            connection.execute(
                """DELETE FROM extraction_candidates
                   WHERE id_document = ? AND field_name = ?
                     AND validation_status NOT IN ('validated', 'rejected')""",
                (document_id, row["field_name"]),
            )


def _persist_analysis(
    connection: sqlite3.Connection,
    case_id: int,
    document: SourceDocument,
    analysis: SourceAnalysis,
    *,
    reanalysis: bool,
) -> None:
    analysis = _normalised_analysis(analysis)
    candidates = {field.strip(): value for field, value in analysis.candidates.items()}
    required_fields = tuple(field.strip() for field in analysis.required_fields)
    with _transaction(connection):
        confirmed_kind = document_review.resolved_classification_kind(
            connection, document.id_document,
        )
        effective_kind = confirmed_kind or analysis.kind.strip()
        connection.execute(
            """UPDATE source_documents
               SET document_kind = ?, classification_confidence = ?
               WHERE id_document = ? AND id_case = ?""",
            (effective_kind, analysis.confidence.strip(), document.id_document, case_id),
        )
        if reanalysis:
            _replace_unvalidated_candidates_not_in_analysis(
                connection, document.id_document, candidates,
            )
        document_review.record_candidates(
            connection,
            document.id_document,
            candidates,
            source="analysis",
            source_context=_compact_context(analysis),
            validation_status="candidate",
            preserve_validated=True,
        )
        if reanalysis:
            document_review.clear_open_automatic_issues(
                connection, case_id, document.id_document,
            )
        if (
            analysis.kind == "unknown"
            and not document_review.has_closed_classification_outcome(
                connection, document.id_document,
            )
        ):
            document_review.create_classification_required_issue(
                connection,
                case_id,
                document.id_document,
                message=analysis.review_message
                or "No se ha podido identificar el tipo de documento; revise la clasificación.",
            )
        else:
            document_review.create_missing_field_issues(
                connection, case_id, document.id_document, required_fields,
            )


def _source_document(connection: sqlite3.Connection, document_id: int) -> SourceDocument:
    row = connection.execute(
        """SELECT id_document, id_case, original_name, archived_path, sha256,
                  document_kind, status
           FROM source_documents WHERE id_document = ?""",
        (document_id,),
    ).fetchone()
    if row is None:
        raise LookupError("El documento no existe")
    return document_from_row(row)


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
    if created:
        document_review.record_candidates(
            connection,
            document.id_document,
            normalized_candidates,
            source="ingestion",
        )
    else:
        document_review.record_candidates_if_missing(
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
    return IngestionResult(document, created, _open_issue_count(connection, case_id))


def add_analysed_document_to_case(
    connection: sqlite3.Connection,
    case_id: int,
    *,
    source_path: str | Path,
    archive_root: str | Path,
    analysis: SourceAnalysis,
) -> IngestionResult:
    """Registra una fuente y conserva el resultado de su clasificación."""
    analysis = _normalised_analysis(analysis)
    document, created = expedient_service.register_source_document(
        connection,
        case_id,
        source_path=source_path,
        archive_root=archive_root,
        document_kind=analysis.kind.strip(),
    )
    _persist_analysis(
        connection, case_id, document, analysis, reanalysis=not created,
    )
    return IngestionResult(
        _source_document(connection, document.id_document),
        created,
        _open_issue_count(connection, case_id),
    )


def _case_analyser(
    connection: sqlite3.Connection,
    case_id: int,
) -> Callable[[Path], SourceAnalysis]:
    row = connection.execute(
        """SELECT comunidades.codigo FROM regularization_cases
           JOIN comunidades ON comunidades.id_comunidad = regularization_cases.id_comunidad
           WHERE regularization_cases.id_case = ?""",
        (case_id,),
    ).fetchone()
    if row is None:
        raise LookupError("El expediente no existe")
    return lambda path: analyse_source(path, community_code=row["codigo"])


def reanalyze_case_documents(
    connection: sqlite3.Connection,
    case_id: int,
    *,
    analyser: Callable[[Path], SourceAnalysis] | None = None,
) -> tuple[IngestionResult, ...]:
    """Actualiza sólo los datos automáticos de las fuentes de un expediente."""
    active_analyser = analyser or _case_analyser(connection, case_id)
    documents = connection.execute(
        """SELECT id_document, id_case, original_name, archived_path, sha256,
                  document_kind, status
           FROM source_documents WHERE id_case = ? ORDER BY id_document""",
        (case_id,),
    ).fetchall()
    results = []
    for row in documents:
        document = document_from_row(row)
        _persist_analysis(
            connection,
            case_id,
            document,
            active_analyser(document.archived_path),
            reanalysis=True,
        )
        results.append(IngestionResult(
            _source_document(connection, document.id_document),
            False,
            _open_issue_count(connection, case_id),
        ))
    return tuple(results)
