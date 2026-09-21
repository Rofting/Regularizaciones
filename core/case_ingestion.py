import sqlite3
import json
from decimal import Decimal, InvalidOperation
from contextlib import contextmanager
from dataclasses import dataclass
from itertools import count
from pathlib import Path
from typing import Callable, Collection, Iterator, Mapping

import document_review
import expedient_service
import gestor_bd
from importar_lecturas_metrigest import apply_confirmed_readings
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


def _required_fields_for_kind(analysis: SourceAnalysis, document_kind: str) -> tuple[str, ...]:
    """Uses the effective classification, not a conflicting automatic guess."""
    if document_kind == analysis.kind:
        return tuple(field.strip() for field in analysis.required_fields)
    required_by_kind = {
        "invoice": SourceAnalysis.invoice().required_fields,
        "reading": ("tipo", "fecha_inicio", "fecha_fin", "vecinos"),
        "owners": ("propietarios",),
    }
    return tuple(required_by_kind.get(document_kind, ()))


def _is_unmistakable_invoice(analysis: SourceAnalysis, candidates: Mapping[str, str | None]) -> bool:
    """Exige evidencia completa antes de corregir una clasificación manual."""
    if analysis.kind != "invoice" or analysis.confidence.strip().lower() != "high":
        return False
    if not candidates.get("proveedor") or candidates.get("vecinos") or candidates.get("propietarios"):
        return False
    for field_name in ("tipo_suministro", "fecha_inicio", "fecha_fin", "importe_total"):
        value = candidates.get(field_name)
        if value is None or not str(value).strip():
            return False
        try:
            document_review.validate_candidate_value(field_name, str(value))
        except ValueError:
            return False
    return True


def _recover_invalid_confirmed_candidates(
    connection: sqlite3.Connection,
    case_id: int,
    document_id: int,
    analysis: SourceAnalysis,
    candidates: Mapping[str, str | None],
) -> None:
    """Recupera sólo valores manuales inválidos con una extracción fiable."""
    if analysis.confidence.strip().lower() != "high":
        return
    rows = connection.execute(
        """SELECT field_name,value FROM extraction_candidates
           WHERE id_document=? AND source='manual' AND validation_status='validated'""",
        (document_id,),
    ).fetchall()
    for row in rows:
        field_name = row["field_name"]
        replacement = candidates.get(field_name)
        if replacement is None or not str(replacement).strip():
            continue
        try:
            document_review.validate_candidate_value(field_name, str(row["value"] or ""))
        except ValueError:
            try:
                document_review.validate_candidate_value(field_name, str(replacement))
            except ValueError:
                continue
            document_review.record_automatic_candidate_recovery(
                connection,
                case_id,
                document_id,
                field_name=field_name,
                original_value=row["value"],
                corrected_value=str(replacement),
                reason=(
                    "El valor confirmado no supera la validación semántica y "
                    "se recupera desde un análisis de confianza alta."
                ),
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
    candidates = {
        field.strip(): value
        for field, value in analysis.candidates.items()
        if field not in analysis.field_evidence
        or analysis.field_evidence[field].confidence.lower() != "low"
    }
    if analysis.disposition == "non_operational":
        eligibility_status, eligibility_reason = "not_applicable", "non_operational"
    elif analysis.kind == "unknown":
        eligibility_status, eligibility_reason = "review_required", "document_unknown"
    elif analysis.kind == "invoice" and analysis.field_evidence and not analysis.provider_key:
        eligibility_status, eligibility_reason = "review_required", "provider_unknown"
    else:
        # Legacy analyses already validated by their provider-specific parser
        # remain compatible while the new orchestrator is rolled out.
        eligibility_status, eligibility_reason = "eligible", None
    with _transaction(connection):
        stored = connection.execute(
            """SELECT document_kind,classification_confidence,source_context
               FROM source_documents WHERE id_document=? AND id_case=?""",
            (document.id_document, case_id),
        ).fetchone()
        if stored is None:
            raise LookupError("El documento no pertenece al expediente")
        confirmed_kind = document_review.resolved_classification_kind(
            connection, document.id_document,
        )
        if (
            confirmed_kind
            and analysis.kind == "other"
            and analysis.confidence.strip().lower() == "high"
            and analysis.disposition == "non_operational"
        ):
            document_review.record_automatic_candidate_recovery(
                connection,
                case_id,
                document.id_document,
                field_name="document_kind",
                original_value=confirmed_kind,
                corrected_value="other",
                reason=(
                    analysis.review_message
                    or "Documento informativo reconocido con confianza alta; no interviene en el cálculo."
                ),
            )
            confirmed_kind = None
        if (
            confirmed_kind in {"reading", "owners"}
            and _is_unmistakable_invoice(analysis, candidates)
        ):
            document_review.record_automatic_candidate_recovery(
                connection,
                case_id,
                document.id_document,
                field_name="document_kind",
                original_value=confirmed_kind,
                corrected_value="invoice",
                reason=(
                    "Proveedor y campos completos de factura detectados con confianza alta; "
                    "la clasificación anterior no contiene lecturas ni propietarios."
                ),
            )
            confirmed_kind = None
        uninformative = analysis.kind == "unknown" and not any(
            value is not None and str(value).strip() for value in candidates.values()
        )
        effective_kind = confirmed_kind or (
            stored["document_kind"] if uninformative else analysis.kind.strip()
        )
        effective_confidence = (
            stored["classification_confidence"]
            if uninformative and stored["classification_confidence"]
            else analysis.confidence.strip()
        )
        effective_context = (
            stored["source_context"]
            if uninformative and stored["source_context"]
            else _compact_context(analysis)
        )
        required_fields = _required_fields_for_kind(analysis, effective_kind)
        connection.execute(
            """UPDATE source_documents
               SET document_kind = ?, classification_confidence = ?, source_context = ?,
                   provider_key = ?, analysis_version = ?, eligibility_status = ?,
                   eligibility_reason = ?
               WHERE id_document = ? AND id_case = ?""",
            (
                effective_kind, effective_confidence, effective_context,
                analysis.provider_key, analysis.analysis_version,
                eligibility_status, eligibility_reason,
                document.id_document, case_id,
            ),
        )
        for field_name, evidence in analysis.field_evidence.items():
            connection.execute(
                """INSERT INTO source_field_evidence (
                       id_document,field_name,value,confidence,source,locator_json,
                       rule_id,extractor_version
                   ) VALUES (?,?,?,?,?,?,?,?)
                   ON CONFLICT(id_document,field_name,rule_id,extractor_version)
                   DO UPDATE SET value=excluded.value,confidence=excluded.confidence,
                                 source=excluded.source,locator_json=excluded.locator_json,
                                 created_at=datetime('now')""",
                (
                    document.id_document, field_name, evidence.value,
                    evidence.confidence, evidence.source,
                    json.dumps(dict(evidence.locator), ensure_ascii=False),
                    evidence.rule_id, evidence.extractor_version,
                ),
            )
        _recover_invalid_confirmed_candidates(
            connection, case_id, document.id_document, analysis, candidates,
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
        if eligibility_reason == "provider_unknown":
            document_review.create_review_issue(
                connection,
                case_id,
                document.id_document,
                code="PROVIDER_UNKNOWN",
                field_name="document.provider",
                detected_value=None,
                message=(
                    "No se ha identificado con seguridad el emisor. Confirma el proveedor "
                    "una sola vez; después se reanalizarán sus fechas e importes."
                ),
            )
            return
        classification_conflict = bool(
            confirmed_kind
            and effective_kind != analysis.kind
            and analysis.kind == "invoice"
            and _looks_like_invoice(candidates)
        )
        if classification_conflict:
            document_review.create_classification_required_issue(
                connection,
                case_id,
                document.id_document,
                message=(
                    "La clasificación guardada no coincide con el nuevo análisis, que parece una factura. "
                    "Confirma el tipo correcto antes de aplicar la fuente."
                ),
            )
        elif (
            effective_kind == "unknown"
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


def _community_has_owners(connection: sqlite3.Connection, case_id: int) -> bool:
    """¿Hay ya propietarios contra los que emparejar las lecturas?"""
    return connection.execute(
        """SELECT 1 FROM propietarios AS owners
           JOIN regularization_cases AS cases ON cases.id_comunidad = owners.id_comunidad
           WHERE cases.id_case = ? AND owners.activo = 1 LIMIT 1""",
        (case_id,),
    ).fetchone() is not None



def _auto_apply_clean_analysis(
    connection: sqlite3.Connection,
    case_id: int,
    document: SourceDocument,
    analysis: SourceAnalysis,
) -> SourceDocument:
    """Aplica fuentes inequívocas; las dudas siguen llegando a incidencias.

    La confirmación por pantalla era útil como salvaguarda inicial, pero no
    debe convertir un lote limpio de facturas en cientos de clics. Sólo se
    publica automáticamente un análisis de confianza alta, sin incidencias y
    de un tipo operativo conocido. Las extracciones incompletas permanecen
    pendientes para la revisión humana normal.
    """
    if analysis.confidence.strip().lower() != "high":
        return document
    eligibility = connection.execute(
        "SELECT eligibility_status FROM source_documents WHERE id_document=?",
        (document.id_document,),
    ).fetchone()
    if eligibility is not None and eligibility[0] != "eligible":
        return document
    if document.document_kind not in {"invoice", "owners", "reading"}:
        return document
    if document.document_kind == "reading" and not _community_has_owners(connection, case_id):
        # Sin el listado de propietarios, cada vivienda de cada fichero de
        # lecturas abriría su propia incidencia de "vivienda desconocida".
        # La fuente espera: auto_apply_case_sources la aplicará después, ya
        # en su orden (propietarios, facturas y luego lecturas).
        return document
    if _has_open_document_issues(connection, document.id_document):
        return document
    required_by_kind = {
        "invoice": ("tipo_suministro", "fecha_inicio", "fecha_fin", "importe_total"),
        "owners": ("propietarios",),
        "reading": ("tipo", "fecha_inicio", "fecha_fin", "vecinos"),
    }
    values = {
        row["field_name"]: row["value"]
        for row in connection.execute(
            """SELECT field_name,value FROM extraction_candidates
               WHERE id_document=? AND value IS NOT NULL AND trim(value)<>''""",
            (document.id_document,),
        ).fetchall()
    }
    if any(not values.get(field) for field in required_by_kind[document.document_kind]):
        return document
    try:
        return confirm_source_candidates(
            connection, case_id, document.id_document,
            confirmed_by="deteccion_automatica",
        )
    except (ValueError, LookupError):
        # ``confirm_source_candidates`` contiene las comprobaciones canónicas
        # adicionales (por ejemplo, tipo de suministro). Materializamos sólo
        # los campos que falten para que la interfaz los muestre como una
        # incidencia concreta, no como una confirmación masiva.
        document_review.create_missing_field_issues(
            connection, case_id, document.id_document,
            required_by_kind[document.document_kind],
        )
        return _source_document(connection, document.id_document)


def auto_apply_case_sources(connection: sqlite3.Connection, case_id: int) -> tuple[int, int]:
    """Aplica de una vez las fuentes seguras ya presentes en un expediente.

    También sanea expedientes creados antes de la validación automática. El
    resultado es ``(aplicadas, pendientes)`` y nunca cierra incidencias ni
    inventa valores para una fuente incompleta.
    """
    rows = connection.execute(
        """SELECT id_document FROM source_documents
           WHERE id_case=? AND classification_confidence='high'
             AND document_kind IN ('invoice','reading','owners')""",
        (case_id,),
    ).fetchall()
    order = {"owners": 0, "invoice": 1, "reading": 2}
    rows = sorted(
        rows,
        key=lambda row: order[_source_document(connection, int(row["id_document"])).document_kind],
    )
    applied = 0
    for row in rows:
        document = _source_document(connection, int(row["id_document"]))
        before = document.status
        result = _auto_apply_clean_analysis(
            connection, case_id, document,
            SourceAnalysis(document.document_kind, "high"),
        )
        if before != "validated" and result.status == "validated":
            applied += 1
    pending = _open_issue_count(connection, case_id)
    return applied, pending


def list_pending_sources(
    connection: sqlite3.Connection, case_id: int,
) -> tuple[SourceDocument, ...]:
    """Devuelve fuentes operativas que aún necesitan revisión o aplicación."""
    rows = connection.execute(
        """SELECT d.id_document
           FROM source_documents AS d
           WHERE d.id_case=? AND d.classification_confidence IS NOT NULL
             AND d.document_kind<>'other' AND (
               d.status<>'validated' OR d.document_kind='unknown'
               OR NOT EXISTS (
                   SELECT 1 FROM archivos_procesados AS a
                   WHERE a.nombre_archivo='source_document:' || d.id_document
               )
               OR EXISTS (
                   SELECT 1 FROM extraction_candidates AS c
                   WHERE c.id_document=d.id_document
                     AND c.validation_status='candidate'
                     AND c.value IS NOT NULL AND trim(c.value)<>''
               )
             )
           ORDER BY d.id_document""",
        (case_id,),
    ).fetchall()
    return tuple(_source_document(connection, int(row["id_document"])) for row in rows)


def _candidate_values(connection: sqlite3.Connection, document_id: int) -> dict[str, str]:
    rows = connection.execute(
        """SELECT field_name,value FROM extraction_candidates
           WHERE id_document=? AND value IS NOT NULL AND trim(value)<>''
             AND validation_status<>'rejected'""",
        (document_id,),
    ).fetchall()
    return {row["field_name"]: row["value"] for row in rows}


def _reading_required_fields(values: Mapping[str, str]) -> tuple[str, ...]:
    """Obtiene los datos de cabecera que faltan en filas de lectura heredadas."""
    raw_rows = values.get("vecinos")
    if not raw_rows:
        return ("tipo", "fecha_inicio", "fecha_fin", "vecinos")
    try:
        rows = json.loads(raw_rows)
    except (json.JSONDecodeError, TypeError):
        return ("vecinos",)
    if not isinstance(rows, list) or not rows or any(not isinstance(row, dict) for row in rows):
        return ("vecinos",)
    required = []
    for row_field, candidate_field in (
        ("tipo", "tipo"), ("fecha_ant", "fecha_inicio"), ("fecha_act", "fecha_fin"),
    ):
        if any(not row.get(row_field) for row in rows):
            required.append(candidate_field)
    return tuple(required)


def _structured_candidate_error(field_name: str, value: str) -> str | None:
    if field_name not in {"vecinos", "propietarios"}:
        return None
    try:
        rows = json.loads(value)
    except (json.JSONDecodeError, TypeError):
        return f"El campo {field_name} no tiene un formato de lista válido."
    if not isinstance(rows, list) or not rows or any(not isinstance(row, dict) for row in rows):
        return f"El campo {field_name} debe contener una lista no vacía de registros."
    if field_name == "propietarios" and any(
        not row.get("codigo_vivienda") or not row.get("nombre_propietario") for row in rows
    ):
        return "Cada propietario necesita una vivienda y un nombre."
    return None


def _invalid_candidate_message(field_name: str, value: str) -> str | None:
    structured_error = _structured_candidate_error(field_name, value)
    if structured_error:
        return structured_error
    try:
        document_review.validate_candidate_value(field_name, value)
    except ValueError as exc:
        return str(exc)
    return None


def _looks_like_invoice(values: Mapping[str, str]) -> bool:
    return bool(values.get("tipo_suministro") and values.get("importe_total"))


def ensure_pending_source_issues(connection: sqlite3.Connection, case_id: int) -> int:
    """Convierte toda fuente bloqueante en una decisión visible y resoluble."""
    created_before = _open_issue_count(connection, case_id)
    for document in list_pending_sources(connection, case_id):
        if _has_open_document_issues(connection, document.id_document):
            continue
        values = _candidate_values(connection, document.id_document)
        if (
            document.document_kind in {"reading", "owners"}
            and _looks_like_invoice(values)
            and not values.get("vecinos")
            and not values.get("propietarios")
        ):
            document_review.create_classification_required_issue(
                connection, case_id, document.id_document,
                message=(
                    "La clasificación guardada no coincide con el contenido extraído, que parece una factura. "
                    "Confirma el tipo correcto antes de aplicar la fuente."
                ),
            )
            continue
        if document.document_kind == "invoice":
            required = ("tipo_suministro", "fecha_inicio", "fecha_fin", "importe_total")
        elif document.document_kind == "owners":
            required = ("propietarios",)
        elif document.document_kind == "reading":
            required = _reading_required_fields(values)
        elif document.document_kind == "unknown":
            if not document_review.has_closed_classification_outcome(
                connection, document.id_document,
            ):
                document_review.create_classification_required_issue(
                    connection, case_id, document.id_document,
                    message="Confirma si la fuente es una factura, lectura, propietarios u otro documento.",
                )
            continue
        else:
            continue

        document_review.create_missing_field_issues(
            connection, case_id, document.id_document, required,
        )
        fields_to_validate = set(required)
        if document.document_kind == "owners":
            fields_to_validate.add("propietarios")
        elif document.document_kind == "reading":
            fields_to_validate.add("vecinos")
        for field_name in fields_to_validate:
            value = values.get(field_name)
            if not value:
                continue
            message = _invalid_candidate_message(field_name, value)
            if message:
                document_review.create_invalid_field_issue(
                    connection, case_id, document.id_document,
                    field_name=field_name, detected_value=value,
                    message=message,
                )
    return _open_issue_count(connection, case_id) - created_before


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


def apply_confirmed_source(
    connection: sqlite3.Connection,
    case_id: int,
    document_id: int,
) -> SourceDocument:
    """Copia exclusivamente las extracciones confirmadas a las tablas canónicas.

    La fuente sólo pasa a ``validated`` después de insertar sus datos canónicos
    y de registrar su identidad para que los reintentos sean idempotentes.
    """
    with _transaction(connection):
        document = _source_document(connection, document_id)
        if document.id_case != case_id:
            raise LookupError("El documento no pertenece al expediente")

        case = expedient_service.get_case(connection, case_id)
        period_id = expedient_service.link_case_to_period(connection, case_id)
        values = _confirmed_candidate_values(connection, document_id)
        marker_name = f"source_document:{document_id}"
        already_applied = connection.execute(
            "SELECT id_factura FROM archivos_procesados WHERE nombre_archivo=?",
            (marker_name,),
        ).fetchone()

        canonical_invoice_id = None
        if document.document_kind == "invoice":
            canonical_invoice_id = _apply_confirmed_invoice(
                connection, case, period_id, document, values, already_applied,
            )
        elif document.document_kind == "reading":
            _apply_confirmed_readings(
                connection, case, period_id, document, values, already_applied,
            )
        elif document.document_kind == "owners":
            _apply_confirmed_owners(connection, case, document, values)
        else:
            raise ValueError("Sólo se pueden aplicar facturas, lecturas o propietarios confirmados")

        has_open_issues = _has_open_document_issues(connection, document_id)
        if already_applied is None and (
            (document.document_kind == "invoice" and canonical_invoice_id is not None)
            or (document.document_kind in ("reading", "owners") and not has_open_issues)
        ):
            gestor_bd.marcar_archivo_procesado(
                connection,
                marker_name,
                hash_md5=document.sha256,
                id_factura=canonical_invoice_id,
                notas=f"Fuente confirmada del expediente {case_id}",
                commit=False,
            )
        if has_open_issues:
            return _source_document(connection, document_id)

        connection.execute(
            "UPDATE source_documents SET status='validated' WHERE id_document=?",
            (document_id,),
        )
        return _source_document(connection, document_id)


def confirm_source_candidates(connection: sqlite3.Connection, case_id: int,
                              document_id: int, *, confirmed_by: str) -> SourceDocument:
    """Explicit user confirmation of the displayed candidates and canonical application."""
    if not confirmed_by.strip():
        raise ValueError("La persona que confirma la fuente es obligatoria")
    with _transaction(connection):
        document = _source_document(connection, document_id)
        if document.id_case != case_id:
            raise LookupError("El documento no pertenece al expediente")
        if _has_open_document_issues(connection, document_id):
            raise ValueError("Resuelve las incidencias de la fuente antes de confirmarla")
        connection.execute("""UPDATE extraction_candidates SET validation_status='validated'
            WHERE id_document=? AND validation_status='candidate'
              AND value IS NOT NULL AND trim(value)<>''""", (document_id,))
        if document.document_kind == "invoice":
            _required_confirmed(_confirmed_candidate_values(connection, document_id),
                                "tipo_suministro", "fecha_inicio", "fecha_fin", "importe_total")
        result = apply_confirmed_source(connection, case_id, document_id)
        connection.execute("""UPDATE source_documents SET confirmed_by=?,confirmed_at=datetime('now')
            WHERE id_document=?""", (confirmed_by.strip(), document_id))
        return result


def _apply_confirmed_owners(
    connection: sqlite3.Connection,
    case: RegularizationCase,
    document: SourceDocument,
    values: Mapping[str, str],
) -> None:
    _required_confirmed(values, "propietarios")
    rows = json.loads(values["propietarios"])
    if not isinstance(rows, list) or not rows:
        raise ValueError("La fuente debe contener propietarios confirmados")
    coefficient_conflicts: list[dict[str, object]] = []
    for row in rows:
        if not isinstance(row, dict) or not row.get("codigo_vivienda") or not row.get("nombre_propietario"):
            raise ValueError("Cada propietario necesita vivienda y nombre")
        fields = {key: row[key] for key in ("nombre_propietario", "coeficiente", "email") if key in row}
        if "coeficiente" in fields:
            coefficient = _confirmed_number({"coeficiente": str(fields["coeficiente"])}, "coeficiente")
            if coefficient is None or coefficient < 0:
                raise ValueError("El coeficiente debe ser un número no negativo")
            fields["coeficiente"] = coefficient
        # El mismo piso llega de fuentes distintas ('MP-1ºA' y 'MP-1ºA    ');
        # sin normalizar, la búsqueda falla y se inserta una segunda fila que
        # duplica el coeficiente y parte las lecturas en dos.
        codigo = gestor_bd.normalizar_codigo_vivienda(row["codigo_vivienda"])
        owner = connection.execute(
            """SELECT id_propietario,coeficiente FROM propietarios
               WHERE id_comunidad=? AND codigo_vivienda=?""",
            (case.community_id, codigo),
        ).fetchone()
        if owner:
            incoming = fields.get("coeficiente")
            current = owner["coeficiente"]
            if (
                incoming is not None and current is not None
                and abs(float(incoming) - float(current)) > 0.0001
            ):
                coefficient_conflicts.append({
                    "vivienda": codigo,
                    "guardado": float(current),
                    "detectado": float(incoming),
                })
                # La participación registral sólo cambia mediante una
                # corrección explícita. El resto de datos de la fuente sí se
                # puede actualizar sin perder la discrepancia.
                fields.pop("coeficiente", None)
            assignments = ','.join(f'{key}=?' for key in fields)
            connection.execute(f"UPDATE propietarios SET {assignments} WHERE id_propietario=?",
                               (*fields.values(), owner[0]))
        else:
            fields["tipo_unidad"] = gestor_bd.clasificar_tipo_unidad(codigo)
            columns = ','.join(fields)
            placeholders = ','.join('?' for _ in fields)
            connection.execute(f"INSERT INTO propietarios(id_comunidad,codigo_vivienda,{columns}) VALUES (?,?,{placeholders})",
                               (case.community_id, codigo, *fields.values()))
    if coefficient_conflicts:
        document_review.create_review_issue(
            connection, case.id_case, document.id_document,
            code="OWNER_COEFFICIENT_CONFLICT", field_name="coeficientes",
            detected_value=json.dumps(coefficient_conflicts, ensure_ascii=False),
            message=(
                f"{len(coefficient_conflicts)} participación(es) difieren de los valores "
                "registrales guardados. Se han conservado los valores existentes; "
                "revise el listado y confirme sólo si realmente han cambiado."
            ),
        )


def _confirmed_candidate_values(
    connection: sqlite3.Connection, document_id: int,
) -> dict[str, str]:
    rows = connection.execute(
        """SELECT field_name,value FROM extraction_candidates
           WHERE id_document=? AND validation_status='validated'
             AND value IS NOT NULL AND trim(value)<>''""",
        (document_id,),
    ).fetchall()
    return {row["field_name"]: row["value"] for row in rows}


def _required_confirmed(values: Mapping[str, str], *fields: str) -> None:
    missing = [field for field in fields if field not in values]
    if missing:
        raise ValueError(
            "Faltan valores confirmados requeridos: " + ", ".join(missing)
        )


def _confirmed_number(values: Mapping[str, str], field: str) -> float | None:
    value = values.get(field)
    if value is None:
        return None
    try:
        number = Decimal(value.replace(",", "."))
    except (AttributeError, InvalidOperation):
        raise ValueError(f"El valor confirmado de {field} debe ser numérico") from None
    if not number.is_finite():
        raise ValueError(f"El valor confirmado de {field} debe ser finito")
    return float(number)


def _apply_confirmed_invoice(
    connection: sqlite3.Connection,
    case: RegularizationCase,
    period_id: int,
    document: SourceDocument,
    values: Mapping[str, str],
    already_applied: sqlite3.Row | None,
) -> int | None:
    _required_confirmed(
        values, "tipo_suministro", "fecha_inicio", "fecha_fin", "importe_total",
    )
    numeric_fields = (
        "consumo_total", "termino_fijo", "termino_variable", "impuestos", "iva",
    )
    invoice = {
        "id_comunidad": case.community_id,
        "id_periodo": period_id,
        "tipo_suministro": values["tipo_suministro"].strip().upper(),
        "importe_total": _confirmed_number(values, "importe_total"),
        "archivo_origen": str(document.archived_path),
    }
    for field in (
        "proveedor", "cups_o_referencia", "num_factura", "fecha_factura",
        "fecha_inicio", "fecha_fin", "unidad_consumo",
    ):
        if field in values:
            invoice[field] = values[field]
    for field in numeric_fields:
        if field in values:
            invoice[field] = _confirmed_number(values, field)

    invoice_id = already_applied["id_factura"] if already_applied is not None else None
    if invoice_id is not None:
        _update_confirmed_invoice(connection, invoice_id, invoice)
    else:
        invoice_id = gestor_bd.insertar_factura(
            connection, invoice, commit=False, preserve_missing_as_null=True,
        )
    if invoice_id is None:
        existing = connection.execute(
            """SELECT id_factura,id_periodo,tipo_suministro,importe_total FROM facturas
               WHERE id_comunidad=? AND num_factura IS ? AND cups_o_referencia IS ?""",
            (case.community_id, invoice.get("num_factura"), invoice.get("cups_o_referencia")),
        ).fetchone()
        if existing is None:
            raise ValueError("No se ha podido identificar la factura canónica")
        if existing["id_periodo"] not in (None, period_id):
            document_review.create_review_issue(
                connection, case.id_case, document.id_document,
                code="INVOICE_PERIOD_CONFLICT", field_name="id_periodo",
                message="La factura canónica ya pertenece a otro período",
                detected_value=str(existing["id_periodo"]),
            )
            return None
        if (
            existing["tipo_suministro"] != invoice["tipo_suministro"]
            or abs(existing["importe_total"] - invoice["importe_total"]) > 0.01
        ):
            document_review.create_review_issue(
                connection, case.id_case, document.id_document,
                code="INVOICE_CONFLICT", field_name="importe_total",
                message="La factura canónica contradice los importes o suministro confirmados",
                detected_value=str(existing["importe_total"]),
            )
            return None
        if existing["id_periodo"] is None:
            connection.execute(
                "UPDATE facturas SET id_periodo=? WHERE id_factura=?",
                (period_id, existing["id_factura"]),
            )
        invoice_id = existing["id_factura"]

    for component_key, candidate_name in (
        ("fixed", "termino_fijo"),
        ("variable", "termino_variable"),
        ("total", "importe_total"),
    ):
        amount = _confirmed_number(values, candidate_name)
        if amount is None:
            continue
        connection.execute(
            """INSERT INTO invoice_components
               (id_factura,component_key,amount,unit)
               VALUES (?, ?, ?, 'EUR')
               ON CONFLICT(id_factura,component_key) DO UPDATE SET
                   amount=excluded.amount,unit=excluded.unit""",
            (invoice_id, component_key, amount),
        )
    return invoice_id


def _update_confirmed_invoice(
    connection: sqlite3.Connection, invoice_id: int, values: Mapping[str, object],
) -> None:
    """Actualiza sólo los campos confirmados de la misma factura canónica."""
    updates = {
        "id_periodo": values["id_periodo"],
        "tipo_suministro": values["tipo_suministro"],
        "importe_total": values["importe_total"],
        "archivo_origen": values["archivo_origen"],
    }
    for field in (
        "proveedor", "cups_o_referencia", "num_factura", "fecha_factura",
        "fecha_inicio", "fecha_fin", "unidad_consumo", "consumo_total",
        "termino_fijo", "termino_variable", "impuestos", "iva",
    ):
        if field in values:
            updates[field] = values[field]
    assignments = ", ".join(f"{field}=?" for field in updates)
    connection.execute(
        f"UPDATE facturas SET {assignments} WHERE id_factura=?",
        (*updates.values(), invoice_id),
    )


def _apply_confirmed_readings(
    connection: sqlite3.Connection,
    case: RegularizationCase,
    period_id: int,
    document: SourceDocument,
    values: Mapping[str, str],
    already_applied: sqlite3.Row | None,
) -> None:
    _required_confirmed(values, "vecinos")
    try:
        readings = json.loads(values["vecinos"])
    except json.JSONDecodeError:
        raise ValueError("Las lecturas confirmadas no tienen un formato válido") from None
    if not isinstance(readings, list) or not readings:
        raise ValueError("Las lecturas confirmadas deben ser una lista no vacía")
    # Older spreadsheets carry month labels but no exact dates or service.
    # Only explicitly reviewed source fields may fill these missing row values.
    for reading in readings:
        if isinstance(reading, dict):
            for key, field in (("tipo", "tipo"), ("fecha_ant", "fecha_inicio"), ("fecha_act", "fecha_fin")):
                if not reading.get(key) and field in values:
                    reading[key] = values[field]
    apply_confirmed_readings(
        connection,
        case_id=case.id_case,
        document_id=document.id_document,
        community_id=case.community_id,
        period_id=period_id,
        source_path=str(document.archived_path),
        readings=readings,
    )


def _has_open_document_issues(connection: sqlite3.Connection, document_id: int) -> bool:
    return connection.execute(
        "SELECT 1 FROM review_issues WHERE id_document=? AND status='open'",
        (document_id,),
    ).fetchone() is not None


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
    document = _auto_apply_clean_analysis(connection, case_id, document, analysis)
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
    archive_root: str | Path | None = None,
) -> tuple[IngestionResult, ...]:
    """Actualiza los datos automáticos y recupera rutas archivadas verificables."""
    source_archive_root = (
        Path(archive_root)
        if archive_root is not None
        else Path(__file__).resolve().parents[1] / "data" / "expedientes"
    )
    expedient_service.repair_archived_source_paths(
        connection, archive_root=source_archive_root, case_id=case_id,
    )
    ambiguous_sources = {
        item.document_id: item
        for item in expedient_service.list_archived_path_candidates(
            connection, archive_root=source_archive_root, case_id=case_id,
        )
    }
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
        if not document.archived_path.is_file():
            candidate_set = ambiguous_sources.get(document.id_document)
            document_review.create_archived_source_issue(
                connection,
                case_id,
                document.id_document,
                duplicate_count=(len(candidate_set.candidate_paths) if candidate_set else 0),
            )
            results.append(IngestionResult(
                _source_document(connection, document.id_document),
                False,
                _open_issue_count(connection, case_id),
            ))
            continue
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
