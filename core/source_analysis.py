"""Classify source documents before they enter the review workflow."""

from __future__ import annotations

import csv
import json
import math
import re
import unicodedata
import calendar
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
    def invoice(cls, candidates: Mapping[str, str | None] | None = None, *, locator=None,
                confidence: str = "high") -> "SourceAnalysis":
        return cls("invoice", confidence, candidates or {}, INVOICE_FIELDS, locator)

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
        )

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
        provider_config = Path(__file__).resolve().parents[1] / "config" / "proveedores.json"
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
        reference = _reference_workbook_analysis(path)
        if reference is not None:
            return reference
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
                "coeficiente": (
                    "coeficiente", "participacion", "entero participacion",
                    "enteros participacion", "enteros de participacion",
                ),
                "email": ("email", "correo"),
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


def analyse_source(path: Path, *, community_code: str, pdf_processor=None,
                   providers: Mapping[str, object] | None = None) -> SourceAnalysis:
    """Dispatch a supported source file to the appropriate analyser."""
    path = Path(path)
    suffix = path.suffix.lower()
    if suffix == ".pdf":
        return analyse_pdf(
            path, pdf_processor=pdf_processor, community_code=community_code,
            providers=providers,
        )
    if suffix in _TABULAR_SUFFIXES:
        return analyse_tabular(path)
    return SourceAnalysis.unknown()
