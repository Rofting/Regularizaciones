"""Classify source documents before they enter the review workflow."""

from __future__ import annotations

import csv
import json
import math
import re
import sqlite3
import unicodedata
import calendar
from datetime import date, datetime
from dataclasses import dataclass, field
from pathlib import Path
from types import MappingProxyType
from typing import Callable, Mapping

from document_classifier import DocumentClassification, classify_document
from app_paths import ApplicationPaths
from document_text_service import TextExtraction, get_document_text
from invoice_extractors import FieldEvidence
from invoice_extractors import extract_invoice_fields
from provider_registry import (
    ProviderMatch,
    ProviderProfile,
    find_tax_ids,
    load_provider_registry,
    profile_for_name,
    provider_registry_from_payload,
    resolve_provider,
)


INVOICE_FIELDS = ("fecha_inicio", "fecha_fin", "importe_total")
_TABULAR_SUFFIXES = {".csv", ".xls", ".xlsx"}
_LOCATION_KEYS = {"page", "pagina", "fragment", "excerpt", "context", "sheet", "cell"}


@dataclass(frozen=True)
class SourceLocator:
    """A human-readable position supplied by a source extractor."""

    page: int | None = None
    fragment: str | None = None
    sheet: str | None = None
    cell: str | None = None


@dataclass(frozen=True)
class SourceAnalysis:
    """Immutable classification and extracted candidate values for one source."""

    kind: str
    confidence: str
    candidates: Mapping[str, str | None] = field(default_factory=dict)
    required_fields: tuple[str, ...] = ()
    locator: SourceLocator | None = None
    review_message: str | None = None
    disposition: str = "operational"
    provider_key: str | None = None
    field_evidence: Mapping[str, FieldEvidence] = field(default_factory=dict)
    analysis_version: str = "source-analysis-v2"

    def __post_init__(self) -> None:
        evidence = dict(self.field_evidence)
        candidates = dict(self.candidates)
        for name, item in evidence.items():
            candidates.setdefault(name, item.value)
        object.__setattr__(self, "candidates", MappingProxyType(candidates))
        object.__setattr__(self, "required_fields", tuple(self.required_fields))
        object.__setattr__(self, "field_evidence", MappingProxyType(evidence))

    @classmethod
    def invoice(cls, candidates: Mapping[str, str | None] | None = None, *, locator=None,
                confidence: str = "high", provider_key: str | None = None,
                field_evidence: Mapping[str, FieldEvidence] | None = None) -> "SourceAnalysis":
        return cls(
            "invoice", confidence, candidates or {}, INVOICE_FIELDS, locator,
            provider_key=provider_key, field_evidence=field_evidence or {},
        )

    @classmethod
    def reading(cls, candidates: Mapping[str, str | None] | None = None, *, locator=None) -> "SourceAnalysis":
        return cls("reading", "high", candidates or {}, (), locator)

    @classmethod
    def owners(cls, candidates: Mapping[str, str | None] | None = None, *, locator=None) -> "SourceAnalysis":
        return cls("owners", "high", candidates or {}, (), locator)

    @classmethod
    def reference(cls, *, locator=None) -> "SourceAnalysis":
        """Archivo histórico reconocido que no debe entrar como fuente operativa.

        Los modelos Excel completos se importan explícitamente desde el paso de
        arranque. Cuando llegan dentro de una carpeta de facturas no deben
        bloquear el expediente ni duplicar importes.
        """
        return cls(
            "other", "high", {}, (), locator,
            "Modelo Excel de referencia detectado; impórtalo desde «Importar modelo inicial» si quieres usarlo como histórico.",
            "non_operational",
        )

    @classmethod
    def informational(cls, message: str, *, locator=None) -> "SourceAnalysis":
        """Recognised document that must be archived without entering calculations."""
        return cls("other", "high", {}, (), locator, message, "non_operational")

    @classmethod
    def unknown(cls, message: str | None = None, *, locator=None) -> "SourceAnalysis":
        return cls(
            "unknown", "low", {}, (), locator,
            message or "No se ha podido identificar el tipo de documento; revise la clasificación.",
        )


_GENERIC_INVOICE_MARKERS = (
    re.compile(r"\bfactura\b", re.IGNORECASE),
    re.compile(r"n[.º°o]*\s*(?:de\s*)?factura", re.IGNORECASE),
)
_GENERIC_DATE_PATTERN = re.compile(r"\b\d{1,2}[/-]\d{1,2}[/-]\d{2,4}\b")
_GENERIC_TOTAL_PATTERN = re.compile(
    r"(?:total\s+(?:a\s+pagar|factura|importe)|importe\s+total)\D{0,32}"
    r"(?P<valor>\d{1,3}(?:\.\d{3})*,\d{2}|\d+\.\d{2})",
    re.IGNORECASE,
)


def classify_generic_invoice_text(text: str, *, locator: SourceLocator | None = None) -> SourceAnalysis | None:
    """Recognise an invoice without guessing a provider-specific supply type."""
    normalised = " ".join(str(text or "").split())
    if not normalised or not any(marker.search(normalised) for marker in _GENERIC_INVOICE_MARKERS):
        return None
    if not _GENERIC_DATE_PATTERN.search(normalised):
        return None
    total = _GENERIC_TOTAL_PATTERN.search(normalised)
    if total is None:
        return None
    return SourceAnalysis.invoice(
        {"importe_total": total.group("valor")}, locator=locator, confidence="medium",
    )


def _normalise_header(value: object) -> str:
    text = unicodedata.normalize("NFKD", str(value or "").lower())
    text = "".join(char for char in text if not unicodedata.combining(char))
    return re.sub(r"\s+", " ", text).strip()


def _string_value(value: object) -> str | None:
    if value is None:
        return None
    if isinstance(value, (dict, list, tuple)):
        return json.dumps(value, ensure_ascii=False, allow_nan=False)
    return str(value)


def _locator_from_mapping(result: Mapping[str, object], data: Mapping[str, object]) -> SourceLocator | None:
    values = {**result, **data}
    page = values.get("page", values.get("pagina"))
    try:
        page = int(page) if page is not None else None
    except (TypeError, ValueError):
        page = None
    fragment = values.get("fragment", values.get("excerpt", values.get("context")))
    sheet = values.get("sheet")
    cell = values.get("cell")
    if page is None and not any((fragment, sheet, cell)):
        return None
    return SourceLocator(
        page=page,
        fragment=_string_value(fragment),
        sheet=_string_value(sheet),
        cell=_string_value(cell),
    )


def analyse_pdf(path: Path, *, pdf_processor=None, community_code: str | None = None,
                providers: Mapping[str, object] | None = None) -> SourceAnalysis:
    """Classify a PDF using the existing provider/reading parser."""
    if pdf_processor is None:
        from lector_pdf import procesar_archivo
        provider_config = ApplicationPaths.resolve().home / "config" / "proveedores.json"
        processor_args = {"ruta_proveedores": str(provider_config)}
        if providers is not None:
            processor_args["proveedores"] = providers
        result = procesar_archivo(str(path), community_code, **processor_args)
    else:
        result = pdf_processor(path, community_code)
    if not isinstance(result, Mapping):
        return SourceAnalysis.unknown()
    if not result.get("ok"):
        locator = _locator_from_mapping(result, {})
        reason = _string_value(result.get("motivo"))
        detail = _string_value(result.get("detalle"))
        if reason == "COMUNIDAD_NO_COINCIDE":
            detected = _string_value(result.get("codigo_comunidad_detectado"))
            return SourceAnalysis.informational(
                detail or f"Documento perteneciente a la comunidad {detected or 'indicada en la fuente'}.",
                locator=locator,
            )
        if reason == "PROVEEDOR_NO_IDENTIFICADO":
            generic = classify_generic_invoice_text(
                _string_value(result.get("fragment")) or "", locator=locator,
            )
            if generic is not None:
                return generic
        message = ": ".join(part for part in (reason, detail) if part)
        return SourceAnalysis.unknown(message or None, locator=locator)
    raw_data = result.get("datos")
    data = raw_data if isinstance(raw_data, Mapping) else {}
    candidates = {
        str(key): _string_value(value)
        for key, value in data.items()
        if key not in _LOCATION_KEYS
    }
    locator = _locator_from_mapping(result, data)
    if result.get("tipo") == "FACTURA":
        return SourceAnalysis.invoice(
            candidates,
            locator=locator,
            provider_key=_string_value(
                result.get("provider_key", result.get("proveedor_clave"))
            ),
            field_evidence=(
                result.get("field_evidence")
                if isinstance(result.get("field_evidence"), Mapping)
                and all(
                    isinstance(item, FieldEvidence)
                    for item in result.get("field_evidence", {}).values()
                )
                else {}
            ),
        )
    if result.get("tipo") == "LECTURA_METRIGEST":
        return SourceAnalysis.reading(candidates, locator=locator)
    if result.get("tipo") == "JUSTIFICANTE_PAGO":
        return SourceAnalysis.informational(
            "Justificante bancario reconocido; se conserva como soporte y no se incorpora como factura.",
            locator=locator,
        )
    return SourceAnalysis.unknown(locator=locator)


def _classification_locator(
    extraction: TextExtraction,
) -> SourceLocator:
    return SourceLocator(
        page=extraction.pages[0] if extraction.pages else None,
        fragment=" ".join(extraction.text.split())[:500] or None,
    )


def _informational_classification(
    classification: DocumentClassification,
    locator: SourceLocator,
) -> SourceAnalysis | None:
    messages = {
        "quote": "Presupuesto reconocido; se archiva como soporte y no entra en el reparto.",
        "delivery_note": "Albarán reconocido; se archiva como soporte y no entra en el reparto.",
        "bank_receipt": "Justificante bancario reconocido; se conserva como soporte y no se incorpora como factura.",
        "report": "Informe reconocido; se conserva como soporte y no se incorpora como factura.",
        "other": "Documento auxiliar reconocido; se conserva como soporte y no entra en el reparto.",
    }
    message = messages.get(classification.kind)
    return SourceAnalysis.informational(message, locator=locator) if message else None


def _generic_invoice_profile() -> ProviderProfile:
    return ProviderProfile(
        key="UNKNOWN",
        display_name="Proveedor pendiente",
        tax_ids=(),
        aliases=(),
        document_types=("invoice", "credit_note"),
        service_family="",
        extractor_family="standard_spanish_invoice",
        required_signatures=(),
        excluded_signatures=(),
        legacy={},
    )


def _legacy_invoice_candidates(
    profile: ProviderProfile,
    text: str,
    locator: SourceLocator,
) -> tuple[dict[str, str | None], dict[str, FieldEvidence]]:
    """Reuse mature provider regexes without reopening or OCRing the PDF."""
    if not isinstance(profile.legacy.get("regex"), Mapping):
        return {}, {}
    from lector_pdf import (
        extraer_datos_agua_zaragoza,
        extraer_datos_factura,
        extraer_datos_naturgy,
    )

    config = dict(profile.legacy)
    if profile.key == "AGUA_ZARAGOZA":
        extracted = extraer_datos_agua_zaragoza(text, config)
    elif profile.key == "NATURGY_CLIENTES_GAS":
        extracted = extraer_datos_naturgy(text, config)
    else:
        extracted = extraer_datos_factura(text, config)
    candidates: dict[str, str | None] = {}
    evidence: dict[str, FieldEvidence] = {}
    for name, value in extracted.items():
        if value is None or isinstance(value, (dict, list, tuple)):
            continue
        string_value = _string_value(value)
        candidates[str(name)] = string_value
        evidence[str(name)] = FieldEvidence(
            value=string_value or "",
            confidence="medium",
            source="provider_regex",
            locator={"page": locator.page, "fragment": locator.fragment},
            rule_id=f"legacy-provider:{profile.key}:{name}:v1",
        )
    return candidates, evidence


def load_learned_tax_ids(connection: sqlite3.Connection | None) -> dict[str, str]:
    if connection is None:
        return {}
    try:
        rows = connection.execute(
            "SELECT tax_id, provider_key FROM provider_learned_tax_ids"
        ).fetchall()
    except sqlite3.OperationalError:
        return {}
    return {str(row[0]): str(row[1]) for row in rows}


def _community_tax_ids(connection: sqlite3.Connection | None) -> set[str]:
    if connection is None:
        return set()
    try:
        rows = connection.execute("SELECT cif FROM comunidades WHERE cif IS NOT NULL").fetchall()
    except sqlite3.OperationalError:
        return set()
    result = {re.sub(r"[^A-Z0-9]", "", str(row[0]).upper()) for row in rows}
    # El CIF del propio despacho aparece en facturas dirigidas «a la atención
    # del administrador»: nunca es el del emisor.
    from office_settings import load_office_settings
    office_tax_id = load_office_settings(connection).tax_id
    if office_tax_id:
        result.add(office_tax_id)
    return result


def learn_issuer_tax_id(
    connection: sqlite3.Connection | None,
    text: str,
    provider_key: str,
    registry: Mapping[str, ProviderProfile],
    *,
    source: str,
) -> str | None:
    """Guarda el CIF del emisor de una factura cuyo proveedor es seguro.

    Se descartan los CIF de comunidades de propietarios (letra H y los
    registrados) porque son el cliente, y los que ya pertenecen a otro
    proveedor. Sólo se aprende si queda exactamente un candidato: una factura
    que menciona también a la distribuidora no debe enseñar un CIF ajeno.
    """
    if connection is None:
        return None
    communities = _community_tax_ids(connection)
    learned = load_learned_tax_ids(connection)
    owners = {tax_id: profile.key for profile in registry.values() for tax_id in profile.tax_ids}
    owners.update(learned)
    candidates = [
        tax_id for tax_id in find_tax_ids(text)
        if not tax_id.startswith("H") and tax_id not in communities
        and owners.get(tax_id, provider_key) == provider_key
    ]
    if len(candidates) != 1 or candidates[0] in learned:
        return None
    try:
        connection.execute(
            """INSERT OR IGNORE INTO provider_learned_tax_ids(tax_id,provider_key,source)
               VALUES (?,?,?)""",
            (candidates[0], provider_key, source),
        )
    except sqlite3.OperationalError:
        return None
    return candidates[0]


def reading_table_analysis(table, locator: SourceLocator | None = None) -> SourceAnalysis:
    """Convierte una tabla de lecturas genérica en un análisis revisable."""
    complete = [
        {key: value for key, value in row.items() if key not in {"contador", "nombre"}}
        for row in table.rows
        if row.get("val_ant") is not None and row.get("val_act") is not None
    ]
    values: dict[str, str | None] = {"vecinos": _string_value(complete)}
    if table.service:
        values["tipo"] = table.service
    if table.start:
        values["fecha_inicio"] = table.start
    if table.end:
        values["fecha_fin"] = table.end
    if table.company:
        values["empresa_lecturas"] = table.company
    missing = tuple(
        name for name, row_key in (("tipo", "tipo"), ("fecha_inicio", "fecha_ant"), ("fecha_fin", "fecha_act"))
        if not values.get(name) and any(not row.get(row_key) for row in complete)
    )
    confidence = table.confidence
    notes = list(table.diagnostics)
    if len(complete) < len(table.rows):
        confidence = "medium"
        notes.append(f"filas_incompletas:{len(table.rows) - len(complete)}")
    message = "Lecturas detectadas automáticamente"
    if table.company:
        message += f" (informe de {table.company.replace('_', ' ').title()})"
    message += ". Revisa servicio, fechas y viviendas antes de confirmar."
    if notes:
        message += " Avisos: " + ", ".join(notes) + "."
    return SourceAnalysis(
        "reading", confidence if not missing else "medium", values, missing, locator, message,
    )


def _generic_reading_from_pdf(path: Path, text: str, locator: SourceLocator) -> SourceAnalysis | None:
    from reading_tables import parse_reading_document, tables_from_pdf
    table = parse_reading_document(text=text, table_rows=tables_from_pdf(path))
    if table is None or not any(
        row.get("val_ant") is not None and row.get("val_act") is not None for row in table.rows
    ):
        return None
    return reading_table_analysis(table, locator)


def analyse_pdf_pipeline(
    path: Path,
    *,
    connection: sqlite3.Connection | None,
    community_code: str | None,
    providers: Mapping[str, object] | None = None,
    provider_registry: Mapping[str, ProviderProfile] | None = None,
    text_extractor: Callable[[Path, int], TextExtraction] | None = None,
    provider_detail_extractor: Callable[[Path, int, int], str] | None = None,
    forced_provider: str | None = None,
) -> SourceAnalysis:
    """Run the cached global pipeline used by folder and bulk ingestion.

    ``forced_provider`` es el proveedor que el usuario confirmó para este
    documento: se usa cuando la detección automática no encuentra ninguno.
    """
    extraction = get_document_text(connection, path, extractor=text_extractor)
    classification = classify_document(extraction.text, path.name)
    locator = _classification_locator(extraction)

    if not extraction.text.strip():
        # Keep the mature reader as a recovery adapter for formats that the
        # cached text service cannot decode yet.  This path is exceptional;
        # normal documents continue through the single-pass global pipeline.
        fallback = analyse_pdf(
            path,
            community_code=community_code,
            providers=providers,
        )
        if fallback.kind != "unknown":
            return fallback
        detail = str(extraction.diagnostics.get("detail") or "No se obtuvo texto legible.")
        return SourceAnalysis.unknown(detail, locator=locator)

    informational = _informational_classification(classification, locator)
    if informational is not None:
        return informational

    if classification.kind == "owners":
        return SourceAnalysis.owners(locator=locator)

    if classification.kind == "reading":
        def process_preloaded(_path, selected_community):
            from lector_pdf import procesar_archivo
            provider_config = ApplicationPaths.resolve().home / "config" / "proveedores.json"
            arguments = {
                "ruta_proveedores": str(provider_config),
                "extracted_text": extraction.text,
            }
            if providers is not None:
                arguments["proveedores"] = providers
            return procesar_archivo(str(path), selected_community, **arguments)

        legacy = analyse_pdf(
            path,
            pdf_processor=process_preloaded,
            community_code=community_code,
            providers=providers,
        )
        if legacy.kind != "unknown":
            return legacy
        # Informe de una empresa de lecturas sin parser propio: se busca su
        # tabla por el significado de las cabeceras.
        generic = _generic_reading_from_pdf(path, extraction.text, locator)
        return generic or legacy

    if classification.kind not in {"invoice", "credit_note"}:
        if classification.kind == "unknown":
            generic = _generic_reading_from_pdf(path, extraction.text, locator)
            if generic is not None and generic.confidence == "high":
                return generic
        return SourceAnalysis.unknown(locator=locator)

    if provider_registry is None:
        if providers is not None:
            provider_registry = provider_registry_from_payload(providers)
        else:
            provider_registry = load_provider_registry(
                ApplicationPaths.resolve().home / "config" / "proveedores.json"
            )
    match = resolve_provider(
        provider_registry,
        extraction.text,
        path.name,
        classification.kind,
        learned_tax_ids=load_learned_tax_ids(connection),
    )
    if match is None and forced_provider:
        confirmed = profile_for_name(provider_registry, forced_provider)
        if confirmed is not None:
            match = ProviderMatch(confirmed.key, "high", "provider:manual:v1", (forced_provider,))
    if match is None:
        bundle = extract_invoice_fields(_generic_invoice_profile(), extraction.text)
        return SourceAnalysis.invoice(
            {"proveedor": forced_provider} if forced_provider else None,
            confidence="medium",
            locator=locator,
            field_evidence=bundle.fields,
        )
    if match.confidence == "high" and match.rule_id != "provider:tax-id:v1":
        learn_issuer_tax_id(
            connection, extraction.text, match.provider_key, provider_registry,
            source=match.rule_id,
        )

    profile = provider_registry[match.provider_key]
    provider_text = extraction.text
    detail_last_page = int(profile.legacy.get("ocr_detail_last_page") or 1)
    if detail_last_page > 1:
        detail_first_page = int(profile.legacy.get("ocr_detail_first_page") or 2)

        def extract_detail(document_path: Path, _max_pages: int) -> TextExtraction:
            import time
            from lector_pdf import extraer_texto_paginas_ocr

            started = time.monotonic()
            detail_text = (
                provider_detail_extractor(
                    document_path, detail_first_page, detail_last_page,
                )
                if provider_detail_extractor is not None
                else extraer_texto_paginas_ocr(
                    str(document_path), detail_first_page, detail_last_page,
                )
            )
            return TextExtraction(
                detail_text,
                "provider_detail_ocr" if detail_text.strip() else "no_text",
                tuple(range(detail_first_page, detail_last_page + 1)),
                {},
                int((time.monotonic() - started) * 1000),
                False,
            )

        detail = get_document_text(
            connection,
            path,
            extractor_version=(
                f"provider-detail:{profile.key}:"
                f"{detail_first_page}-{detail_last_page}:v1"
            ),
            max_pages=detail_last_page,
            extractor=extract_detail,
        )
        if detail.text.strip():
            provider_text = f"{provider_text}\n{detail.text}"

    bundle = extract_invoice_fields(profile, provider_text)
    candidates = {name: item.value for name, item in bundle.fields.items()}
    evidence = dict(bundle.fields)
    legacy_candidates, legacy_evidence = _legacy_invoice_candidates(
        profile, provider_text, locator,
    )
    specialised = profile.key in {"AGUA_ZARAGOZA", "NATURGY_CLIENTES_GAS"}
    for name, value in legacy_candidates.items():
        if specialised and value not in (None, ""):
            candidates[name] = value
        else:
            candidates.setdefault(name, value)
    for name, item in legacy_evidence.items():
        if specialised and item.value not in (None, ""):
            evidence[name] = item
        else:
            evidence.setdefault(name, item)
    return SourceAnalysis.invoice(
        candidates,
        locator=locator,
        # Provider resolution only returns a unique winner.  Combined with a
        # structurally strong invoice classification this is safe to apply
        # automatically even when the legacy profile has no extra signature.
        confidence="high" if classification.confidence == "high" else "medium",
        provider_key=match.provider_key,
        field_evidence=evidence,
    )


def classify_headers(
    headers: tuple[object, ...] | list[object], *, suffix: str, locator: SourceLocator | None = None
) -> SourceAnalysis:
    """Classify a tabular source from its headers without guessing missing data."""
    normalised = tuple(_normalise_header(header) for header in headers)
    has_reading_signal = any("lectura" in header or "contador" in header for header in normalised)
    has_owner_signal = any(
        "propietario" in header or "propiedad" in header or "vivienda" in header
        for header in normalised
    )
    if has_reading_signal:
        return SourceAnalysis.reading(locator=locator)
    if has_owner_signal:
        return SourceAnalysis.owners(locator=locator)
    return SourceAnalysis.unknown(locator=locator)


def analyse_tabular(path: Path) -> SourceAnalysis:
    """Extract explicit rows; incomplete or ambiguous tables remain reviewable.

    Every sheet is tried (active one first): reports often keep the data on a
    second sheet behind a cover page. The first recognised table wins; when none
    is, the first sheet's diagnostic is reported.
    """
    try:
        reference = _reference_workbook_analysis(path)
        if reference is not None:
            return reference
        first_failure: SourceAnalysis | None = None
        sheets = []
        for rows, sheet in _tabular_sheets(path):
            result = _analyse_rows(rows, sheet)
            if result.kind != "unknown":
                return result
            first_failure = first_failure or result
            sheets.append((rows, sheet))
        # Ningún formato conocido: informe de lecturas con cabeceras propias.
        from reading_tables import parse_reading_rows
        for rows, sheet in sheets:
            title = " ".join(str(cell) for row in rows[:6] for cell in row if cell)
            table = parse_reading_rows(rows, f"{sheet or ''} {title}")
            if table is not None:
                return reading_table_analysis(table, SourceLocator(sheet=sheet, cell="A1"))
        return first_failure or SourceAnalysis.unknown("El archivo no contiene filas.")
    except (OSError, ValueError, csv.Error, ImportError, TypeError) as error:
        return SourceAnalysis.unknown(str(error))


def _analyse_rows(rows, sheet) -> SourceAnalysis:
    locator = None
    try:
        for index, headers in enumerate(rows[:30]):
            keys = [_header_key(value) for value in headers]
            locator = SourceLocator(sheet=sheet, cell=f"A{index + 1}",
                                    fragment=" | ".join(str(v or "") for v in headers)[:1000])
            from importar_lecturas_xls import reading_columns
            legacy_columns = reading_columns(headers)
            if legacy_columns is not None:
                property_column, periods = legacy_columns
                if len(periods) < 2:
                    raise ValueError("Faltan dos columnas de lectura por periodo")
                candidates = _legacy_reading_rows(
                    rows[index + 1:], property_column, periods[0], periods[-1],
                )
                if not candidates:
                    raise ValueError("La tabla no contiene lecturas")
                for row in candidates:
                    for key in ("val_ant", "val_act"):
                        # Los informes Meditrade dejan la celda vacía cuando
                        # el contador informa 0. El importador histórico ya
                        # aplica esa convención; conservarla aquí permite que
                        # el flujo auditado decida si es un cero inicial o si
                        # debe arrastrar la última lectura fiable.
                        row[key] = (
                            0.0 if row[key] in (None, "")
                            else _tabular_number(row[key])
                        )
                metadata = _legacy_reading_metadata(headers, periods)
                values = {"vecinos": _string_value(candidates), **metadata}
                missing = tuple(
                    field for field in ("tipo", "fecha_inicio", "fecha_fin")
                    if not values.get(field)
                )
                return SourceAnalysis(
                    "reading", "high" if not missing else "medium", values, missing, locator,
                    "Revisa y confirma el servicio y el intervalo deducidos de las columnas de lectura.",
                )
            legacy_owners = _meditrade_owner_rows(rows[index + 1:], headers)
            if legacy_owners is not None:
                if not legacy_owners:
                    raise ValueError("El listado no contiene propietarios completos")
                return SourceAnalysis.owners(
                    {"propietarios": _string_value(legacy_owners)}, locator=locator,
                )
            reading_keys = {
                "vivienda": ("vivienda", "propiedad", "codigo vivienda", "inmueble", "piso"),
                "tipo": ("tipo", "servicio", "suministro", "tipo servicio", "tipo suministro"),
                "fecha_ant": (
                    "fecha ant", "fecha anterior", "fecha inicio", "fecha lectura anterior",
                    "fecha lectura ant", "fecha inicial",
                ),
                "fecha_act": (
                    "fecha act", "fecha actual", "fecha fin", "fecha lectura actual",
                    "fecha lectura act", "fecha final",
                ),
                "val_ant": (
                    "val ant", "lectura anterior", "lectura inicial", "lect ant",
                    "lectura ant", "valor anterior",
                ),
                "val_act": (
                    "val act", "lectura actual", "lectura final", "lect act",
                    "lectura act", "valor actual",
                ),
            }
            owner_keys = {
                "codigo_vivienda": ("fdenominacion", "vivienda", "propiedad", "codigo vivienda"),
                "nombre_propietario": (
                    "nombre", "propietario", "nombre propietario", "titular",
                    "nombre y apellidos", "apellidos y nombre",
                ),
                "coeficiente": (
                    "coeficiente", "participacion", "entero participacion",
                    "enteros participacion", "enteros de participacion", "coef",
                    "coeficiente participacion", "coeficiente de participacion",
                ),
                "email": ("email", "e mail", "correo", "correo electronico", "mail"),
            }
            reading_columns = _matching_columns(keys, reading_keys)
            owner_columns = _matching_columns(keys, owner_keys)
            if len(reading_columns) == len(reading_keys):
                candidates = _extract_rows(rows[index + 1:], reading_columns)
                if not candidates:
                    return SourceAnalysis.unknown("La tabla no contiene lecturas; revisa las filas.", locator=locator)
                for row in candidates:
                    if any(row.get(key) in (None, "") for key in reading_keys):
                        raise ValueError("Hay lecturas sin vivienda, servicio, fechas o valores")
                    row["tipo"] = _normalise_header(row["tipo"]).upper()
                    if row["tipo"] not in ("ACS", "CALEFACCION"):
                        raise ValueError("El servicio de las lecturas necesita confirmación")
                    for key in ("fecha_ant", "fecha_act"):
                        row[key] = _tabular_date(row[key])
                    for key in ("val_ant", "val_act"):
                        row[key] = _tabular_number(row[key])
                    if row["fecha_act"] <= row["fecha_ant"]:
                        raise ValueError("Las fechas de lectura no forman un intervalo válido")
                return SourceAnalysis.reading({"vecinos": _string_value(candidates)}, locator=locator)
            if "codigo_vivienda" in owner_columns and "nombre_propietario" in owner_columns:
                # A table with meter signals must not turn into an owners-only import.
                if any("lectura" in key or "contador" in key for key in keys):
                    return SourceAnalysis.unknown("Faltan fechas, servicio o valores de lectura inequívocos.", locator=locator)
                candidates = _extract_rows(rows[index + 1:], owner_columns)
                if not candidates:
                    return SourceAnalysis.unknown("El listado no contiene propietarios.", locator=locator)
                from importar_propietarios_csv import _email
                for row in candidates:
                    if not row.get("codigo_vivienda") or not row.get("nombre_propietario"):
                        raise ValueError("Hay propietarios sin vivienda o nombre")
                    if "coeficiente" in row:
                        row["coeficiente"] = _tabular_number(row["coeficiente"])
                    if "email" in row:
                        row["email"] = _email(row["email"])
                return SourceAnalysis.owners({"propietarios": _string_value(candidates)}, locator=locator)
        headers = rows[0] if rows else ()
        return SourceAnalysis.unknown(
            "No se ha podido extraer una tabla completa; revisa cabeceras, fechas, servicio y valores.",
            locator=SourceLocator(sheet=sheet, cell="A1", fragment=" | ".join(str(v or "") for v in headers)[:1000]),
        )
    except (ValueError, TypeError) as error:
        return SourceAnalysis.unknown(str(error), locator=locator)


def _header_key(value: object) -> str:
    """Comparable header: no accents, units in brackets or punctuation."""
    text = _normalise_header(value)
    text = re.sub(r"[\(\[][^)\]]*[)\]]", " ", text)
    text = re.sub(r"[^a-z0-9%]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def _matching_columns(keys, aliases):
    columns = {}
    for field, names in aliases.items():
        matches = [index for index, key in enumerate(keys) if key in names]
        if not matches:
            # "Lectura anterior m3" or "Fecha fin periodo" still name the
            # field; exact headers always take precedence over these.
            matches = [
                index for index, key in enumerate(keys)
                if any(key.startswith(f"{name} ") for name in names)
            ]
        if len(matches) > 1:
            raise ValueError(f"Cabeceras ambiguas para {field}; confirma qué columna corresponde")
        if matches:
            columns[field] = matches[0]
    return columns


def _extract_rows(rows, columns):
    result = []
    for row in rows:
        if not any(value not in (None, "") for value in row):
            continue
        result.append({key: row[index] if index < len(row) else None
                       for key, index in columns.items()})
    return result


def _legacy_reading_rows(rows, property_column: int, initial_column: int, final_column: int):
    """Return only dwelling rows from Meditrade-style printed reports.

    These reports end with a community control line (code 999) and ``Totales``.
    They are report footers, not missing homes. Rows with a real dwelling but a
    missing reading deliberately remain candidates so the normal numeric check
    can request a review rather than inventing a value.
    """
    result = []
    for row in rows:
        raw_vivienda = row[property_column] if property_column < len(row) else None
        vivienda = str(raw_vivienda or "").strip()
        if not vivienda:
            continue
        result.append({
            "vivienda": vivienda,
            "val_ant": row[initial_column] if initial_column < len(row) else None,
            "val_act": row[final_column] if final_column < len(row) else None,
        })
    return result


_READING_MONTHS = {
    "ene": 1, "enero": 1, "feb": 2, "febrero": 2, "mar": 3, "marzo": 3,
    "abr": 4, "abril": 4, "may": 5, "mayo": 5, "jun": 6, "junio": 6,
    "jul": 7, "julio": 7, "ago": 8, "agosto": 8, "sep": 9, "sept": 9,
    "septiembre": 9, "oct": 10, "octubre": 10, "nov": 11, "noviembre": 11,
    "dic": 12, "diciembre": 12,
}


def _legacy_reading_metadata(headers, periods: tuple[int, ...]) -> dict[str, str]:
    """Derive confirmable service and month boundaries from report headers."""
    label_text = " ".join(_normalise_header(value) for value in headers)
    service = None
    if "acs" in label_text or "agua caliente" in label_text:
        service = "ACS"
    elif "calefaccion" in label_text:
        service = "CALEFACCION"

    values = [headers[index] for index in periods]
    known_years = [
        year for value in values
        if (parts := re.search(r"(?:/|-)(\d{4})$", str(value or "").strip().lower()))
        for year in (int(parts.group(1)),)
    ]
    start = _legacy_header_month(values[0], known_years)
    end = _legacy_header_month(values[-1], known_years)
    metadata: dict[str, str] = {}
    if service:
        metadata["tipo"] = service
    if start:
        metadata["fecha_inicio"] = f"{start[0]:04d}-{start[1]:02d}-01"
    if end:
        metadata["fecha_fin"] = f"{end[0]:04d}-{end[1]:02d}-{calendar.monthrange(*end)[1]:02d}"
    return metadata


def _legacy_header_month(value, known_years: list[int]) -> tuple[int, int] | None:
    """Parse ``7/2025``, ``01/6`` or ``sep-24`` without guessing a day."""
    text = str(value or "").strip().lower()
    numeric = re.fullmatch(r"(\d{1,2})/(\d{1,4})", text)
    named = re.fullmatch(r"([a-záéíóú]+)-(\d{2,4})", text)
    if numeric:
        month, raw_year = int(numeric.group(1)), numeric.group(2)
    elif named:
        month = _READING_MONTHS.get(named.group(1))
        raw_year = named.group(2)
    else:
        return None
    if not month or not 1 <= month <= 12:
        return None
    if len(raw_year) == 4:
        year = int(raw_year)
    elif len(raw_year) == 2:
        year = 2000 + int(raw_year)
    else:
        suffix = int(raw_year)
        matching = [candidate for candidate in known_years if candidate % 10 == suffix]
        year = matching[0] if len(matching) == 1 else 2020 + suffix
    return year, month


def _meditrade_owner_rows(rows, headers):
    """Read Meditrade owner reports, whose emails are printed on the next row."""
    labels = [_normalise_header(value).replace("_", " ") for value in headers]
    try:
        code_column = next(index for index, value in enumerate(labels) if value in ("codigo", "cod"))
        property_column = labels.index("propiedad")
        name_column = labels.index("nombre")
    except StopIteration:
        return None
    except ValueError:
        return None
    coefficient_aliases = {
        "coeficiente", "participacion", "entero participacion",
        "enteros participacion", "enteros de participacion",
    }
    coefficient_column = next(
        (index for index, value in enumerate(labels) if value in coefficient_aliases), None,
    )
    result = []
    current = None
    from importar_propietarios_csv import _email
    for row in rows:
        values = [row[index] if index < len(row) else None for index in range(len(headers))]
        email = _email(" ".join(str(value or "") for value in values))
        code = str(values[code_column] or "").strip()
        vivienda = str(values[property_column] or "").strip()
        nombre = str(values[name_column] or "").strip()
        if vivienda and nombre and re.fullmatch(r"\d+(?:\.0)?", code):
            current = {
                "codigo_vivienda": vivienda,
                "nombre_propietario": nombre,
            }
            if coefficient_column is not None:
                coefficient = values[coefficient_column]
                if coefficient not in (None, ""):
                    current["coeficiente"] = _tabular_number(coefficient)
            if email:
                current["email"] = email
            result.append(current)
        elif email and current is not None:
            current["email"] = email
    return result


def _reference_workbook_analysis(path: Path) -> SourceAnalysis | None:
    """Recognise the validated multi-sheet Excel model before row analysis."""
    if path.suffix.lower() != ".xlsx":
        return None
    from openpyxl import load_workbook
    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        sheet_names = set(workbook.sheetnames)
        required = {"DATOS", "GAS", "ELECTRICIDAD", "AGUA", "LECTURAS ACS M3", "ANALISIS"}
        if not required.issubset(sheet_names):
            return None
        return SourceAnalysis.reference(locator=SourceLocator(
            sheet="DATOS", cell="A1",
            fragment="Modelo Excel de referencia detectado; no se añadirá como factura ni lectura operativa.",
        ))
    finally:
        workbook.close()


def _tabular_number(value):
    if value is None or isinstance(value, bool) or str(value).strip() == "":
        raise ValueError("Falta un valor numérico en la tabla")
    if isinstance(value, (int, float)):
        number = float(value)
    else:
        text = re.sub(r"(?i)\s*(?:m3|m³|kwh|€|eur|l)\s*$", "", str(value)).strip().replace(" ", "")
        if "," in text and "." in text:
            if text.rfind(",") > text.rfind("."):
                text = text.replace(".", "").replace(",", ".")
            else:
                text = text.replace(",", "")
        elif "," in text:
            text = text.replace(".", "").replace(",", ".")
        elif re.fullmatch(r"\d{1,3}(?:\.\d{3}){2,}", text):
            text = text.replace(".", "")
        number = float(text)
    if not math.isfinite(number) or number < 0:
        raise ValueError("Hay un valor numérico inválido en la tabla")
    return number


def _tabular_date(value):
    if isinstance(value, (date, datetime)):
        return value.strftime("%Y-%m-%d")
    text = str(value).strip()
    for format in ("%Y-%m-%d", "%Y-%m-%d %H:%M:%S", "%d/%m/%Y", "%d-%m-%Y", "%d.%m.%Y",
                   "%d/%m/%y", "%d-%m-%y", "%d.%m.%y"):
        try:
            return datetime.strptime(text, format).date().isoformat()
        except ValueError:
            pass
    raise ValueError("La fecha de lectura no es inequívoca")


def _sniff_csv(handle):
    sample = handle.read(4096)
    handle.seek(0)
    try:
        return csv.Sniffer().sniff(sample, delimiters=",;\t|")
    except csv.Error:
        return csv.excel


def _tabular_sheets(path):
    """Yield ``(rows, sheet_name)`` for each sheet, the active one first."""
    suffix = path.suffix.lower()
    if suffix == ".csv":
        for encoding in ("utf-8-sig", "cp1252"):
            try:
                with path.open("r", encoding=encoding, newline="") as handle:
                    yield list(csv.reader(handle, _sniff_csv(handle))), None
                return
            except UnicodeDecodeError:
                continue
        raise ValueError("No se puede leer la codificación del CSV")
    elif suffix == ".xlsx":
        from openpyxl import load_workbook
        workbook = load_workbook(path, read_only=True, data_only=True)
        try:
            active = workbook.active
            sheets = [active] + [item for item in workbook.worksheets if item is not active]
            for sheet in sheets:
                yield list(sheet.iter_rows(values_only=True)), sheet.title
        finally:
            workbook.close()
    elif suffix == ".xls":
        import xlrd
        workbook = xlrd.open_workbook(path)
        try:
            order = [workbook.sheet_by_index(0)]
            order += [workbook.sheet_by_index(i) for i in range(1, workbook.nsheets)]
            for sheet in order:
                rows = [[xlrd.xldate.xldate_as_datetime(cell.value, workbook.datemode)
                         if cell.ctype == xlrd.XL_CELL_DATE else cell.value
                         for cell in sheet.row(index)] for index in range(sheet.nrows)]
                yield rows, sheet.name
        finally:
            workbook.release_resources()
    else:
        raise ValueError("Formato tabular no compatible")


def analyse_source(path: Path, *, community_code: str, pdf_processor=None,
                   providers: Mapping[str, object] | None = None,
                   connection: sqlite3.Connection | None = None,
                   provider_registry: Mapping[str, ProviderProfile] | None = None,
                   text_extractor: Callable[[Path, int], TextExtraction] | None = None,
                   forced_provider: str | None = None) -> SourceAnalysis:
    """Dispatch a supported source file to the appropriate analyser."""
    path = Path(path)
    suffix = path.suffix.lower()
    if suffix == ".pdf":
        if connection is not None or provider_registry is not None or text_extractor is not None:
            return analyse_pdf_pipeline(
                path,
                connection=connection,
                community_code=community_code,
                providers=providers,
                provider_registry=provider_registry,
                text_extractor=text_extractor,
                forced_provider=forced_provider,
            )
        return analyse_pdf(
            path, pdf_processor=pdf_processor, community_code=community_code,
            providers=providers,
        )
    if suffix in _TABULAR_SUFFIXES:
        return analyse_tabular(path)
    return SourceAnalysis.unknown()
