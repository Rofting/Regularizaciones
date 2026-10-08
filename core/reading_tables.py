"""Lectura genérica de tablas de contadores (cualquier empresa de lecturas).

Cada empresa (ista, Techem, Metrigest, Zenner…) maqueta su informe a su manera,
pero todos traen lo mismo: una fila por vivienda con la lectura anterior y la
actual (a veces el consumo, el contador y el titular). Este módulo localiza
esa tabla por el significado de sus cabeceras, no por su posición:

* sinónimos configurables en ``config/palabras_clave.json`` y comparación
  aproximada para cabeceras mal leídas («LECTURA ANTERI0R», «Lect.Ant.»);
* columnas cuyo título es una fecha («Lectura 01/09/2025», «sep-25»), que dan
  también el período;
* informes combinados con columnas de ACS y de calefacción;
* cabeceras partidas en dos filas y tablas que continúan en varias páginas.

Sirve igual para filas de Excel/CSV, tablas extraídas de un PDF o líneas de
texto (PDF sin estructura u OCR). El consumo, cuando viene, se usa para
comprobar cada fila: es lo que permite fiarse del resultado.
"""

from __future__ import annotations

import calendar
import re
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path
from typing import Iterable, Sequence

import keywords
import reading_formats


ROLES_REQUIRED = ("vivienda", "val_ant", "val_act")
_MONTHS = {
    "ene": 1, "feb": 2, "mar": 3, "abr": 4, "may": 5, "jun": 6, "jul": 7,
    "ago": 8, "sep": 9, "set": 9, "oct": 10, "nov": 11, "dic": 12,
}


@dataclass(frozen=True)
class ReadingTable:
    rows: tuple[dict, ...]
    services: tuple[str, ...]
    start: str | None
    end: str | None
    company: str | None
    confidence: str
    diagnostics: tuple[str, ...] = field(default_factory=tuple)
    format_metadata: dict | None = None

    @property
    def service(self) -> str | None:
        return self.services[0] if len(self.services) == 1 else None


# ── Utilidades ───────────────────────────────────────────────────────────

def _number(value: object) -> float | None:
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = re.sub(r"(?i)\s*(m3|m³|kwh|uds?|l)\s*$", "", str(value)).strip().replace(" ", "")
    if not text or not re.fullmatch(r"-?[\d.,]+", text):
        return None
    if "," in text and "." in text:
        text = text.replace(".", "").replace(",", ".") if text.rfind(",") > text.rfind(".") \
            else text.replace(",", "")
    elif "," in text:
        text = text.replace(".", "").replace(",", ".")
    elif re.fullmatch(r"-?\d{1,3}(?:\.\d{3}){2,}", text):
        text = text.replace(".", "")
    try:
        return float(text)
    except ValueError:
        return None


def _date_in(value: object) -> tuple[date, bool] | None:
    """Fecha contenida en una cabecera. El booleano indica si sólo es mes/año."""
    if isinstance(value, datetime):
        return value.date(), False
    if isinstance(value, date):
        return value, False
    text = keywords.fold(value)
    full = re.search(r"\b(\d{1,2}) (\d{1,2}) (\d{4}|\d{2})\b", text)
    if full:
        day, month, year = int(full[1]), int(full[2]), int(full[3])
        year += 2000 if year < 100 else 0
        try:
            return date(year, month, day), False
        except ValueError:
            return None
    named = re.search(r"\b([a-z]{3})[a-z]* (\d{4}|\d{2})\b", text)
    if named and named[1] in _MONTHS:
        year = int(named[2]) + (2000 if len(named[2]) == 2 else 0)
        return date(year, _MONTHS[named[1]], 1), True
    numeric = re.fullmatch(r"(?:lectura |lect )?(\d{1,2}) (\d{4})", text)
    if numeric and 1 <= int(numeric[1]) <= 12:
        return date(int(numeric[2]), int(numeric[1]), 1), True
    return None


def _service_of(text: str) -> str | None:
    folded = f" {keywords.fold(text)} "
    found = [
        service for service, words in keywords.section("servicio_lecturas").items()
        if any(f" {word} " in folded for word in words)
    ]
    return found[0] if len(found) == 1 else None


def detect_company(text: str) -> str | None:
    """Empresa de lecturas que firma el informe (ista, Techem…), si es única."""
    folded = f" {keywords.fold(text)} "
    found = [
        company for company, words in keywords.section("empresas_lecturas").items()
        if any(f" {word} " in folded for word in words)
    ]
    return found[0] if len(found) == 1 else None


def _is_summary(value: str) -> bool:
    folded = keywords.fold(value)
    return any(folded == word or folded.startswith(word + " ")
               for word in keywords.word_list("filas_resumen"))


# ── Cabeceras ────────────────────────────────────────────────────────────

@dataclass
class _Column:
    index: int
    role: str
    service: str | None = None
    when: date | None = None
    month_only: bool = False


def _classify_header(cells: Sequence[object]) -> list[_Column]:
    roles = keywords.section("cabeceras_lecturas")
    columns: list[_Column] = []
    for index, cell in enumerate(cells):
        text = str(cell or "").strip()
        if not text:
            continue
        found_date = _date_in(cell)
        match = keywords.best_role(text, roles)
        service = _service_of(text)
        if found_date and (match is None or match[0] in {"val_ant", "val_act"}):
            columns.append(_Column(index, "lectura_fecha", service, found_date[0], found_date[1]))
        elif match is not None:
            columns.append(_Column(index, match[0], service))
        elif service and re.search(r"\blect", keywords.fold(text)):
            columns.append(_Column(index, "lectura_servicio", service))
    return columns


def _resolve_reading_columns(columns: list[_Column]) -> dict[str | None, tuple[_Column, _Column]]:
    """Pares (anterior, actual) por servicio a partir de las columnas detectadas."""
    pairs: dict[str | None, tuple[_Column, _Column]] = {}
    by_service: dict[str | None, list[_Column]] = {}
    for column in columns:
        if column.role in {"val_ant", "val_act", "lectura_fecha", "lectura_servicio"}:
            by_service.setdefault(column.service, []).append(column)
    for service, items in by_service.items():
        previous = [item for item in items if item.role == "val_ant"]
        current = [item for item in items if item.role == "val_act"]
        if previous and current:
            pairs[service] = (previous[0], current[0])
            continue
        dated = sorted((item for item in items if item.role == "lectura_fecha"),
                       key=lambda item: item.when)
        if len(dated) >= 2:
            pairs[service] = (dated[0], dated[-1])
            continue
        generic = [item for item in items if item.role == "lectura_servicio"]
        if len(generic) == 2:
            pairs[service] = (generic[0], generic[1])
    # Columnas sin servicio en un informe que sí distingue servicios: ambiguo.
    if None in pairs and len(pairs) > 1:
        pairs.pop(None)
    return pairs


def _find_header(rows: Sequence[Sequence[object]], learned=None):
    limit = min(len(rows), 40)
    for index in range(limit):
        candidates = [(index, list(rows[index]))]
        if index + 1 < len(rows):
            merged = [
                f"{rows[index][col] if col < len(rows[index]) and rows[index][col] is not None else ''} "
                f"{rows[index + 1][col] if col < len(rows[index + 1]) and rows[index + 1][col] is not None else ''}".strip()
                for col in range(max(len(rows[index]), len(rows[index + 1])))
            ]
            candidates.append((index + 1, merged))
        for last_header_row, cells in candidates:
            mapping = (learned or {}).get(reading_formats.header_signature(cells))
            if mapping and any(item["index"] >= len(cells) for item in mapping):
                mapping = None
            if mapping and any(item["role"] == "lectura_fecha" and not _date_in(cells[item["index"]])
                               for item in mapping):
                mapping = None
            if mapping:
                columns = []
                for item in mapping:
                    when = _date_in(cells[item["index"]])
                    columns.append(_Column(item["index"], item["role"], item["service"],
                                           when[0] if when else None, when[1] if when else False))
            else:
                columns = _classify_header(cells)
            if not any(column.role == "vivienda" for column in columns):
                continue
            pairs = _resolve_reading_columns(columns)
            if pairs:
                return last_header_row, cells, columns, pairs
    return None


# ── Tabla ────────────────────────────────────────────────────────────────

def parse_reading_rows(rows: Sequence[Sequence[object]], text: str = "", *, connection=None) -> ReadingTable | None:
    """Interpreta una tabla de lecturas; ``None`` si no hay una reconocible."""
    rows = [list(row) for row in rows if row is not None]
    company = detect_company(text + " " + " ".join(
        str(cell or "") for row in rows[:6] for cell in row
    ))
    learned = reading_formats.formats_for(connection, company)
    found = _find_header(rows, learned)
    if found is None:
        return None
    header_index, header_cells, columns, pairs = found
    role_index = {column.role: column.index for column in columns if column.role not in
                  {"val_ant", "val_act", "lectura_fecha", "lectura_servicio"}}
    vivienda_col = role_index["vivienda"]
    header_keys = {keywords.fold(cell) for cell in header_cells if cell}

    document_service = _service_of(" ".join(str(cell or "") for cell in header_cells)) or _service_of(text)
    services = tuple(
        service or document_service for service in pairs
    )
    extracted: list[dict] = []
    checked = matched = negatives = 0
    consumption_values: list[float] = []
    summary_consumptions: list[float] = []
    consumption_col = role_index.get("consumo")
    for raw in rows[header_index + 1:]:
        cells = list(raw)
        if {keywords.fold(cell) for cell in cells if cell} & header_keys == header_keys:
            continue  # cabecera repetida en otra página
        vivienda = str(cells[vivienda_col] if vivienda_col < len(cells) and cells[vivienda_col] is not None else "").strip()
        if _is_summary(vivienda):
            if consumption_col is not None and consumption_col < len(cells) and len(pairs) == 1:
                summary = _number(cells[consumption_col])
                if summary is not None:
                    summary_consumptions.append(summary)
            continue
        if not vivienda or len(vivienda) > 60:
            continue
        for (service_key, (previous, current)), service in zip(pairs.items(), services):
            before = _number(cells[previous.index]) if previous.index < len(cells) else None
            after = _number(cells[current.index]) if current.index < len(cells) else None
            if before is None and after is None:
                continue
            row = {"vivienda": vivienda, "val_ant": before, "val_act": after}
            if service:
                row["tipo"] = service
            for optional in ("contador", "nombre"):
                if optional in role_index and role_index[optional] < len(cells):
                    value = str(cells[role_index[optional]] or "").strip()
                    if value:
                        row[optional] = value
            if previous.when and not previous.month_only:
                row["fecha_ant"] = previous.when.isoformat()
            if current.when and not current.month_only:
                row["fecha_act"] = current.when.isoformat()
            if consumption_col is not None and consumption_col < len(cells) and len(pairs) == 1:
                consumption = _number(cells[consumption_col])
                if consumption is not None:
                    consumption_values.append(consumption)
                if consumption is not None and before is not None and after is not None:
                    checked += 1
                    matched += abs((after - before) - consumption) <= 0.6
            if before is not None and after is not None and after < before:
                negatives += 1
            extracted.append(row)
    if not extracted:
        return None

    diagnostics = []
    signature = reading_formats.header_signature(header_cells)
    start, end = _period(pairs, text)
    if checked:
        ratio = matched / checked
        confidence = "high" if ratio == 1 else "medium"
        if ratio < 1:
            diagnostics.append(f"consumo_no_cuadra:{checked - matched}/{checked}")
    else:
        complete = sum(row["val_ant"] is not None and row["val_act"] is not None for row in extracted)
        confidence = "high" if complete == len(extracted) and len(extracted) >= 2 else "medium"
    if negatives:
        confidence = "medium"
        diagnostics.append(f"lecturas_decrecientes:{negatives}")
    if summary_consumptions and len(consumption_values) == len(extracted):
        if abs(sum(consumption_values) - sum(summary_consumptions)) > 0.6:
            confidence = "medium"
            diagnostics.append("total_consumo_no_cuadra")
    if any(service is None for service in services):
        confidence = "medium"
        diagnostics.append("servicio_sin_identificar")
    if learned and signature not in learned:
        confidence = "medium"
        diagnostics.append("formato_lecturas_modificado")
    elif connection is not None and company and signature not in learned:
        confidence = "medium"  # el primer informe enseña el formato al confirmarlo
    metadata = {
        "company": company, "header_signature": signature,
        "extractor_version": reading_formats.EXTRACTOR_VERSION,
        "mapping": [{"index": c.index, "role": c.role, "service": c.service} for c in columns],
    } if company else None
    return ReadingTable(
        tuple(extracted), tuple(s for s in services if s), start, end,
        company, confidence, tuple(diagnostics), metadata,
    )


def _period(pairs, text: str) -> tuple[str | None, str | None]:
    dated = [column for pair in pairs.values() for column in pair if column.when]
    if len(dated) >= 2:
        first = min(dated, key=lambda column: column.when)
        last = max(dated, key=lambda column: column.when)
        end = last.when
        if last.month_only:
            end = end.replace(day=calendar.monthrange(end.year, end.month)[1])
        return first.when.isoformat(), end.isoformat()
    from invoice_extractors import _PERIOD_PATTERNS, _iso_date
    for pattern, _confidence in _PERIOD_PATTERNS:
        match = re.search(pattern, text or "", re.IGNORECASE | re.MULTILINE)
        if match:
            start, end = _iso_date(match.group("start")), _iso_date(match.group("end"))
            if start and end and start <= end:
                return start, end
    return None, None


# ── Fuentes: texto y PDF ─────────────────────────────────────────────────

_NUMBER_TOKEN = re.compile(r"-?\d{1,3}(?:[.\s]\d{3})*(?:,\d+)?|-?\d+(?:[.,]\d+)?")


def rows_from_text(text: str) -> list[list[str]]:
    """Convierte líneas de texto en filas: separa por 2+ espacios o tabuladores.

    Si una línea no tiene separadores claros, se separan los números finales
    del texto inicial (vivienda o titular).
    """
    rows: list[list[str]] = []
    for line in (text or "").splitlines():
        line = line.strip()
        if not line:
            continue
        cells = [cell for cell in re.split(r"\t+|\s{2,}|\s\|\s", line) if cell.strip()]
        if len(cells) < 3:
            tokens = line.split()
            numbers: list[str] = []
            while tokens and _number(tokens[-1]) is not None:
                numbers.insert(0, tokens.pop())
            if tokens and len(numbers) >= 2:
                cells = [" ".join(tokens), *numbers]
        rows.append(cells)
    return rows


def tables_from_pdf(path: str | Path, *, max_pages: int = 15, text: str = "") -> list[list[object]]:
    """Elige una estrategia de tablas válida, sin sumar filas duplicadas."""
    try:
        import pdfplumber
    except ImportError:
        return []
    settings = (
        None,
        {"vertical_strategy": "lines", "horizontal_strategy": "lines"},
        {"vertical_strategy": "text", "horizontal_strategy": "text"},
    )
    candidates: list[list[list[object]]] = [[] for _ in settings]
    try:
        with pdfplumber.open(str(path)) as pdf:
            for page in pdf.pages[:max_pages]:
                for index, setting in enumerate(settings):
                    try:
                        tables = page.extract_tables(table_settings=setting) if setting else page.extract_tables()
                    except Exception:
                        continue
                    for table in tables or ():
                        candidates[index].extend(
                            [cell if cell is None else str(cell).replace("\n", " ") for cell in row]
                            for row in table if row is not None
                        )
    except Exception:
        return []
    scored = []
    for index, rows in enumerate(candidates):
        parsed = parse_reading_rows(rows, text) if rows else None
        scored.append((len(parsed.rows) if parsed else 0,
                       parsed.confidence == "high" if parsed else False, -index, rows))
    return max(scored, key=lambda item: item[:3])[3]


def parse_reading_document(
    *, text: str = "", table_rows: Iterable[Sequence[object]] = (), connection=None,
) -> ReadingTable | None:
    """Prueba primero las tablas estructuradas y después las líneas de texto."""
    rows = list(table_rows)
    best = parse_reading_rows(rows, text, connection=connection) if rows else None
    if best is None or best.confidence != "high":
        from_text = parse_reading_rows(rows_from_text(text), text, connection=connection) if text else None
        if from_text is not None and (best is None or len(from_text.rows) > len(best.rows)
                                      or (len(from_text.rows) == len(best.rows)
                                          and from_text.confidence == "high"
                                          and not any(note.startswith(("consumo_no_cuadra", "total_consumo_no_cuadra"))
                                                      for note in best.diagnostics))):
            best = from_text
    return best


__all__ = [
    "ReadingTable",
    "detect_company",
    "parse_reading_document",
    "parse_reading_rows",
    "rows_from_text",
    "tables_from_pdf",
]
