"""Classify source documents before they enter the review workflow."""

from __future__ import annotations

import csv
import re
import unicodedata
from dataclasses import dataclass, field
from pathlib import Path
from types import MappingProxyType
from typing import Mapping


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

    def __post_init__(self) -> None:
        object.__setattr__(self, "candidates", MappingProxyType(dict(self.candidates)))
        object.__setattr__(self, "required_fields", tuple(self.required_fields))

    @classmethod
    def invoice(cls, candidates: Mapping[str, str | None] | None = None, *, locator=None) -> "SourceAnalysis":
        return cls("invoice", "high", candidates or {}, INVOICE_FIELDS, locator)

    @classmethod
    def reading(cls, candidates: Mapping[str, str | None] | None = None, *, locator=None) -> "SourceAnalysis":
        return cls("reading", "high", candidates or {}, (), locator)

    @classmethod
    def owners(cls, candidates: Mapping[str, str | None] | None = None, *, locator=None) -> "SourceAnalysis":
        return cls("owners", "high", candidates or {}, (), locator)

    @classmethod
    def unknown(cls, message: str | None = None, *, locator=None) -> "SourceAnalysis":
        return cls(
            "unknown", "low", {}, (), locator,
            message or "No se ha podido identificar el tipo de documento; revise la clasificación.",
        )


def _normalise_header(value: object) -> str:
    text = unicodedata.normalize("NFKD", str(value or "").lower())
    text = "".join(char for char in text if not unicodedata.combining(char))
    return re.sub(r"\s+", " ", text).strip()


def _string_value(value: object) -> str | None:
    if value is None:
        return None
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


def analyse_pdf(path: Path, *, pdf_processor=None, community_code: str | None = None) -> SourceAnalysis:
    """Classify a PDF using the existing provider/reading parser."""
    if pdf_processor is None:
        from lector_pdf import procesar_archivo
        pdf_processor = procesar_archivo
    result = pdf_processor(path, community_code)
    if not isinstance(result, Mapping) or not result.get("ok"):
        return SourceAnalysis.unknown()
    raw_data = result.get("datos")
    data = raw_data if isinstance(raw_data, Mapping) else {}
    candidates = {
        str(key): _string_value(value)
        for key, value in data.items()
        if key not in _LOCATION_KEYS
    }
    locator = _locator_from_mapping(result, data)
    if result.get("tipo") == "FACTURA":
        return SourceAnalysis.invoice(candidates, locator=locator)
    if result.get("tipo") == "LECTURA_METRIGEST":
        return SourceAnalysis.reading(candidates, locator=locator)
    return SourceAnalysis.unknown(locator=locator)


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


def _tabular_headers(path: Path) -> tuple[tuple[object, ...], SourceLocator | None]:
    suffix = path.suffix.lower()
    if suffix == ".csv":
        with path.open("r", encoding="utf-8-sig", errors="replace", newline="") as handle:
            sample = handle.read(4096)
            handle.seek(0)
            try:
                dialect = csv.Sniffer().sniff(sample, delimiters=",;\t")
            except csv.Error:
                dialect = csv.excel
            headers = next(csv.reader(handle, dialect), [])
        return tuple(headers), None
    if suffix == ".xlsx":
        from openpyxl import load_workbook
        workbook = load_workbook(path, read_only=True, data_only=True)
        try:
            sheet = workbook.active
            headers = tuple(cell.value for cell in next(sheet.iter_rows(max_row=1), ()))
            return headers, SourceLocator(sheet=sheet.title, cell="A1") if headers else None
        finally:
            workbook.close()
    if suffix == ".xls":
        import xlrd
        workbook = xlrd.open_workbook(path)
        sheet = workbook.sheet_by_index(0)
        headers = tuple(sheet.row_values(0)) if sheet.nrows else ()
        return headers, SourceLocator(sheet=sheet.name, cell="A1") if headers else None
    raise ValueError(f"Formato tabular no compatible: {suffix}")


def analyse_tabular(path: Path) -> SourceAnalysis:
    """Read tabular headers and retain their source position when available."""
    try:
        headers, locator = _tabular_headers(path)
    except (OSError, ValueError, csv.Error, ImportError):
        return SourceAnalysis.unknown()
    return classify_headers(headers, suffix=path.suffix.lower(), locator=locator)


def analyse_source(path: Path, *, community_code: str, pdf_processor=None) -> SourceAnalysis:
    """Dispatch a supported source file to the appropriate analyser."""
    path = Path(path)
    suffix = path.suffix.lower()
    if suffix == ".pdf":
        return analyse_pdf(path, pdf_processor=pdf_processor, community_code=community_code)
    if suffix in _TABULAR_SUFFIXES:
        return analyse_tabular(path)
    return SourceAnalysis.unknown()
