"""Classify source documents before they enter the review workflow."""

from __future__ import annotations

import csv
import json
import math
import re
import unicodedata
from datetime import date, datetime
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
    """Extract explicit rows; incomplete or ambiguous tables remain reviewable."""
    try:
        rows, sheet = _tabular_rows(path)
        for index, headers in enumerate(rows[:30]):
            keys = [_normalise_header(value).replace("_", " ") for value in headers]
            locator = SourceLocator(sheet=sheet, cell=f"A{index + 1}",
                                    fragment=" | ".join(str(v or "") for v in headers)[:1000])
            from importar_lecturas_xls import reading_columns
            legacy_columns = reading_columns(headers)
            if legacy_columns is not None:
                property_column, periods = legacy_columns
                if len(periods) < 2:
                    raise ValueError("Faltan dos columnas de lectura por periodo")
                candidates = _extract_rows(rows[index + 1:], {
                    "vivienda": property_column, "val_ant": periods[0], "val_act": periods[-1],
                })
                if not candidates:
                    raise ValueError("La tabla no contiene lecturas")
                for row in candidates:
                    if not str(row["vivienda"] or "").strip():
                        raise ValueError("Hay lecturas sin vivienda")
                    row["vivienda"] = str(row["vivienda"]).strip()
                    for key in ("val_ant", "val_act"):
                        row[key] = _tabular_number(row[key])
                return SourceAnalysis("reading", "high", {"vecinos": _string_value(candidates)},
                    ("tipo", "fecha_inicio", "fecha_fin"), locator,
                    "Confirma el servicio y las fechas exactas de las columnas inicial y final.")
            reading_keys = {
                "vivienda": ("vivienda", "propiedad", "codigo vivienda"),
                "tipo": ("tipo", "servicio", "suministro"),
                "fecha_ant": ("fecha ant", "fecha anterior", "fecha inicio"),
                "fecha_act": ("fecha act", "fecha actual", "fecha fin"),
                "val_ant": ("val ant", "lectura anterior", "lectura inicial"),
                "val_act": ("val act", "lectura actual", "lectura final"),
            }
            owner_keys = {
                "codigo_vivienda": ("fdenominacion", "vivienda", "propiedad", "codigo vivienda"),
                "nombre_propietario": ("nombre", "propietario", "nombre propietario"),
                "coeficiente": ("coeficiente",), "email": ("email", "correo"),
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
    except (OSError, ValueError, csv.Error, ImportError, TypeError) as error:
        return SourceAnalysis.unknown(str(error), locator=locals().get("locator"))


def _matching_columns(keys, aliases):
    columns = {}
    for field, names in aliases.items():
        matches = [index for index, key in enumerate(keys) if key in names]
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


def _tabular_number(value):
    if value is None or isinstance(value, bool) or str(value).strip() == "":
        raise ValueError("Falta un valor numérico en la tabla")
    text = str(value).strip()
    if "," in text:
        text = text.replace(".", "").replace(",", ".")
    number = float(text)
    if not math.isfinite(number) or number < 0:
        raise ValueError("Hay un valor numérico inválido en la tabla")
    return number


def _tabular_date(value):
    if isinstance(value, (date, datetime)):
        return value.strftime("%Y-%m-%d")
    for format in ("%Y-%m-%d", "%d/%m/%Y"):
        try:
            return datetime.strptime(str(value).strip(), format).date().isoformat()
        except ValueError:
            pass
    raise ValueError("La fecha de lectura no es inequívoca")


def _tabular_rows(path):
    if path.suffix.lower() == ".csv":
        with path.open("r", encoding="utf-8-sig", errors="replace", newline="") as handle:
            sample = handle.read(4096)
            handle.seek(0)
            try:
                dialect = csv.Sniffer().sniff(sample, delimiters=",;\t")
            except csv.Error:
                dialect = csv.excel
            return list(csv.reader(handle, dialect)), None
    if path.suffix.lower() == ".xlsx":
        from openpyxl import load_workbook
        workbook = load_workbook(path, read_only=True, data_only=True)
        try:
            sheet = workbook.active
            return list(sheet.iter_rows(values_only=True)), sheet.title
        finally:
            workbook.close()
    if path.suffix.lower() == ".xls":
        import xlrd
        workbook = xlrd.open_workbook(path)
        try:
            sheet = workbook.sheet_by_index(0)
            rows = [[xlrd.xldate.xldate_as_datetime(cell.value, workbook.datemode)
                     if cell.ctype == xlrd.XL_CELL_DATE else cell.value
                     for cell in sheet.row(index)] for index in range(sheet.nrows)]
            return rows, sheet.name
        finally:
            workbook.release_resources()
    raise ValueError("Formato tabular no compatible")


def analyse_source(path: Path, *, community_code: str, pdf_processor=None) -> SourceAnalysis:
    """Dispatch a supported source file to the appropriate analyser."""
    path = Path(path)
    suffix = path.suffix.lower()
    if suffix == ".pdf":
        return analyse_pdf(path, pdf_processor=pdf_processor, community_code=community_code)
    if suffix in _TABULAR_SUFFIXES:
        return analyse_tabular(path)
    return SourceAnalysis.unknown()
