"""Vista previa de la evidencia de una incidencia en su fuente original.

Las extracciones guardan un fragmento de texto y, a veces, la página o la
celda. Aquí se busca ese texto en la capa de texto del PDF para resaltar las
palabras exactas. Si no aparece (PDF escaneado, texto de OCR que no coincide,
fragmento resumido), no se resalta nada: se muestra la página y el fragmento
tal cual, sin inventar una posición.
"""

from __future__ import annotations

import json
import re
import sqlite3
import unicodedata
from dataclasses import dataclass, field
from pathlib import Path
from typing import Sequence

_MIN_FRAGMENT_RUN = 4


@dataclass(frozen=True)
class Evidence:
    """Lo que se sabe del dato: valor, fragmento y posición declarada."""

    value: str | None = None
    fragment: str | None = None
    page: int | None = None
    sheet: str | None = None
    cell: str | None = None
    confidence: str | None = None

    @property
    def empty(self) -> bool:
        return not any((self.value, self.fragment, self.page, self.sheet, self.cell))


@dataclass(frozen=True)
class PdfMatch:
    page: int                     # 1-based
    boxes: tuple[tuple[float, float, float, float], ...]
    matched: str                  # "value" | "fragment"
    text: str


@dataclass(frozen=True)
class SheetWindow:
    sheet: str
    first_row: int
    first_column: int
    rows: tuple[tuple[str, ...], ...]
    target: tuple[int, int] | None  # (fila, columna) absolutas resaltadas


@dataclass
class Preview:
    kind: str                     # "pdf" | "sheet" | "text"
    evidence: Evidence
    page_count: int = 0
    page: int = 1
    match: PdfMatch | None = None
    sheet: SheetWindow | None = None
    notes: list[str] = field(default_factory=list)

    @property
    def headline(self) -> str:
        if self.match is not None:
            what = "el valor detectado" if self.match.matched == "value" else "el fragmento extraído"
            return f"Resaltado {what} en la página {self.match.page} de {self.page_count}."
        if self.sheet is not None and self.sheet.target is not None:
            row, column = self.sheet.target
            return f"Celda {_column_letter(column)}{row} de la hoja «{self.sheet.sheet}»."
        if self.kind == "pdf":
            return ("No se ha localizado la posición exacta en el PDF. "
                    "Se muestra el fragmento guardado tal como se extrajo.")
        return "Sólo se conserva el fragmento de texto extraído."


# ── Evidencia guardada ──────────────────────────────────────────────────

def _context(raw: object) -> dict:
    if not raw:
        return {}
    try:
        value = json.loads(str(raw))
    except (TypeError, ValueError):
        return {"fragment": str(raw)}
    return value if isinstance(value, dict) else {}


def _int(value: object) -> int | None:
    try:
        number = int(value)  # type: ignore[arg-type]
    except (TypeError, ValueError):
        return None
    return number if number > 0 else None


_CONFIDENCE_RANK = {"high": 0, "medium": 1, "low": 2}


def evidence_for_issue(
    connection: sqlite3.Connection, id_document: int, field_name: str, detected_value: str | None = None,
) -> Evidence:
    """Reúne la mejor evidencia guardada para el campo de una incidencia."""
    value, fragment, page, sheet, cell, confidence = detected_value, None, None, None, None, None
    rows = connection.execute(
        """SELECT value, confidence, locator_json FROM source_field_evidence
           WHERE id_document=? AND field_name=?""",
        (id_document, field_name),
    ).fetchall()
    if rows:
        best = min(rows, key=lambda row: _CONFIDENCE_RANK.get(row[1], 3))
        locator = _context(best[2])
        value = value or best[0]
        confidence = best[1]
        fragment = locator.get("fragment")
        page, sheet, cell = _int(locator.get("page")), locator.get("sheet"), locator.get("cell")
    candidate = connection.execute(
        """SELECT value, source_context FROM extraction_candidates
           WHERE id_document=? ORDER BY (field_name=?) DESC, field_name LIMIT 1""",
        (id_document, field_name),
    ).fetchone()
    contexts = []
    if candidate is not None:
        own = connection.execute(
            "SELECT value FROM extraction_candidates WHERE id_document=? AND field_name=?",
            (id_document, field_name),
        ).fetchone()
        if own is not None and own[0] and not value:
            value = own[0]
        contexts.append(_context(candidate[1]))
    document = connection.execute(
        "SELECT source_context FROM source_documents WHERE id_document=?", (id_document,),
    ).fetchone()
    if document is not None:
        contexts.append(_context(document[0]))
    for context in contexts:
        fragment = fragment or context.get("fragment") or context.get("excerpt")
        page = page or _int(context.get("page"))
        sheet = sheet or context.get("sheet")
        cell = cell or context.get("cell")
    row_in_field = re.search(r"\.row_(\d+)\b", field_name)
    if row_in_field and not cell:
        cell = f"A{row_in_field.group(1)}"
    return Evidence(
        value=str(value).strip() if value not in (None, "") else None,
        fragment=" ".join(str(fragment).split()) if fragment else None,
        page=page, sheet=str(sheet) if sheet else None, cell=str(cell) if cell else None,
        confidence=confidence,
    )


# ── Búsqueda en el PDF ──────────────────────────────────────────────────

def _token(text: str) -> str:
    text = unicodedata.normalize("NFKD", str(text))
    text = "".join(char for char in text if not unicodedata.combining(char)).casefold()
    return text.strip(".,;:()[]{}«»\"'¡!¿?")


def _tokens(text: str | None) -> list[str]:
    return [token for token in (_token(part) for part in str(text or "").split()) if token]


def _find_run(words: Sequence[str], wanted: Sequence[str]) -> int | None:
    """Posición donde aparece ``wanted`` seguido, o None."""
    size = len(wanted)
    if not size:
        return None
    for start in range(len(words) - size + 1):
        if list(words[start:start + size]) == list(wanted):
            return start
    return None


def _longest_run(words: Sequence[str], wanted: Sequence[str]) -> tuple[int, int, int] | None:
    """Tramo contiguo más largo de ``wanted`` presente en ``words``.

    Devuelve (inicio en words, inicio en wanted, longitud).
    """
    best = None
    positions: dict[str, list[int]] = {}
    for index, word in enumerate(words):
        positions.setdefault(word, []).append(index)
    for offset, token in enumerate(wanted):
        for start in positions.get(token, ()):
            length = 0
            while (offset + length < len(wanted) and start + length < len(words)
                   and words[start + length] == wanted[offset + length]):
                length += 1
            if best is None or length > best[2]:
                best = (start, offset, length)
    return best


def _box(words: Sequence[dict]) -> tuple[float, float, float, float]:
    return (min(w["x0"] for w in words), min(w["top"] for w in words),
            max(w["x1"] for w in words), max(w["bottom"] for w in words))


def _line_boxes(words: Sequence[dict]) -> tuple[tuple[float, float, float, float], ...]:
    """Un rectángulo por línea para que un fragmento de varias líneas se lea bien."""
    lines: list[list[dict]] = []
    for word in words:
        if lines and abs(lines[-1][-1]["top"] - word["top"]) < 3:
            lines[-1].append(word)
        else:
            lines.append([word])
    return tuple(_box(line) for line in lines)


def value_variants(value: str | None) -> tuple[str, ...]:
    """Formas en que un valor normalizado puede aparecer impreso."""
    text = str(value or "").strip()
    if not text:
        return ()
    variants = [text]
    iso = re.fullmatch(r"(\d{4})-(\d{2})-(\d{2})", text)
    if iso:
        year, month, day = iso.groups()
        variants += [f"{day}/{month}/{year}", f"{day}-{month}-{year}", f"{day}.{month}.{year}",
                     f"{day}/{month}/{year[2:]}"]
    number = re.fullmatch(r"(-?)(\d+)\.(\d{1,4})", text)
    if number:
        sign, whole, decimals = number.groups()
        decimals = decimals.ljust(2, "0")
        grouped = f"{int(whole):,}".replace(",", ".")
        variants += [f"{sign}{whole},{decimals}", f"{sign}{grouped},{decimals}"]
    return tuple(dict.fromkeys(variants))


def locate_in_pdf(pages_words: Sequence[Sequence[dict]], evidence: Evidence) -> PdfMatch | None:
    """Prueba el valor tal cual y en sus formas impresas (fecha, importe)."""
    for variant in value_variants(evidence.value) or (None,):
        match = _locate(pages_words, evidence, variant)
        if match is not None and match.matched == "value":
            return match
    return _locate(pages_words, evidence, None)


def _locate(pages_words: Sequence[Sequence[dict]], evidence: Evidence, value: str | None) -> PdfMatch | None:
    """Busca el valor (dentro del fragmento si lo hay) y si no, el fragmento.

    ``pages_words`` son las palabras de cada página tal como las da
    pdfplumber (``text``, ``x0``, ``top``, ``x1``, ``bottom``).
    """
    order = list(range(len(pages_words)))
    if evidence.page and evidence.page <= len(order):
        order.remove(evidence.page - 1)
        order.insert(0, evidence.page - 1)
    value_tokens = _tokens(value)
    fragment_tokens = _tokens(evidence.fragment)
    needed = min(len(fragment_tokens), _MIN_FRAGMENT_RUN)
    fragment_hit = None
    for index in order:
        words = pages_words[index]
        tokens = [_token(word["text"]) for word in words]
        run = _longest_run(tokens, fragment_tokens) if fragment_tokens else None
        if run is not None and run[2] < needed:
            run = None
        if value_tokens:
            if run is not None:
                window = range(run[0], run[0] + run[2])
                start = _find_run(tokens[run[0]:run[0] + run[2]], value_tokens)
                if start is not None:
                    selected = words[window.start + start:window.start + start + len(value_tokens)]
                    return PdfMatch(index + 1, _line_boxes(selected), "value", value or "")
        if run is not None and fragment_hit is None:
            selected = words[run[0]:run[0] + run[2]]
            fragment_hit = PdfMatch(index + 1, _line_boxes(selected), "fragment",
                                    " ".join(word["text"] for word in selected))
    if fragment_hit is not None:
        return fragment_hit
    return _unique_value(pages_words, value_tokens, value or "")


def _unique_value(pages_words, value_tokens: list[str], value: str) -> PdfMatch | None:
    """Sin fragmento localizable, el valor sólo vale si es distintivo y único.

    «130» o «2026» aparecen en cualquier factura; un CUPS o «1.234,56» no.
    """
    if len("".join(value_tokens)) < 6:
        return None
    hits = []
    for index, words in enumerate(pages_words):
        tokens = [_token(word["text"]) for word in words]
        offset = 0
        while (start := _find_run(tokens[offset:], value_tokens)) is not None:
            hits.append((index, words[offset + start:offset + start + len(value_tokens)]))
            offset += start + 1
    if len(hits) != 1:
        return None
    index, selected = hits[0]
    return PdfMatch(index + 1, _line_boxes(selected), "value", value)


def pdf_words(path: Path) -> list[list[dict]]:
    import pdfplumber
    with pdfplumber.open(str(path)) as pdf:
        return [list(page.extract_words(keep_blank_chars=False, use_text_flow=True)) for page in pdf.pages]


def render_pdf_page(path: Path, page: int, boxes=(), *, resolution: int = 110):
    """Imagen PIL de la página con los rectángulos resaltados."""
    import pdfplumber
    with pdfplumber.open(str(path)) as pdf:
        image = pdf.pages[page - 1].to_image(resolution=resolution)
        if boxes:
            image.draw_rects(list(boxes), stroke=(220, 38, 38), fill=(250, 204, 21, 70), stroke_width=2)
        return image.annotated.convert("RGB")


# ── Hojas de cálculo ────────────────────────────────────────────────────

def _column_letter(index: int) -> str:
    letters = ""
    while index:
        index, remainder = divmod(index - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


def _cell_position(cell: str) -> tuple[int, int] | None:
    match = re.fullmatch(r"\$?([A-Za-z]{1,3})\$?(\d+)", cell.strip())
    if not match:
        return None
    column = 0
    for char in match.group(1).upper():
        column = column * 26 + ord(char) - 64
    return int(match.group(2)), column


def sheet_window(path: Path, sheet: str | None, cell: str | None, *, rows: int = 14, columns: int = 8) -> SheetWindow:
    """Recorte de la hoja alrededor de la celda (o su esquina si no hay celda)."""
    target = _cell_position(cell) if cell else None
    if path.suffix.lower() == ".csv":
        import csv
        with path.open("r", encoding="utf-8-sig", errors="replace", newline="") as handle:
            sample = handle.read(4096)
            handle.seek(0)
            delimiter = ";" if sample.count(";") >= sample.count(",") else ","
            grid = [row for row in csv.reader(handle, delimiter=delimiter)]
        name = path.name
    else:
        from source_analysis import _tabular_sheets
        sheets = list(_tabular_sheets(path))
        if not sheets:
            raise ValueError("La hoja no contiene filas")
        chosen = next((item for item in sheets if sheet and item[1] == sheet), sheets[0])
        grid, name = [list(row) for row in chosen[0]], chosen[1] or path.name
    first_row = max(1, (target[0] - rows // 2) if target else 1)
    first_column = max(1, (target[1] - columns // 2) if target else 1)
    window = []
    for row in grid[first_row - 1:first_row - 1 + rows]:
        cells = list(row)[first_column - 1:first_column - 1 + columns]
        window.append(tuple("" if value is None else str(value) for value in cells))
    return SheetWindow(str(name), first_row, first_column, tuple(window), target)


# ── Entrada principal ───────────────────────────────────────────────────

def build_preview(path: Path, evidence: Evidence) -> Preview:
    path = Path(path)
    suffix = path.suffix.lower()
    if suffix == ".pdf":
        try:
            words = pdf_words(path)
        except Exception as error:
            preview = Preview("text", evidence)
            preview.notes.append(f"No se pudo abrir el PDF: {error}")
            return preview
        preview = Preview("pdf", evidence, page_count=len(words))
        preview.match = locate_in_pdf(words, evidence)
        if preview.match is not None:
            preview.page = preview.match.page
        elif evidence.page and evidence.page <= len(words):
            preview.page = evidence.page
            preview.notes.append(f"La extracción indicó la página {evidence.page}.")
        if not any(words):
            preview.notes.append("El PDF no tiene capa de texto (escaneado): el dato se leyó por OCR.")
        return preview
    if suffix in {".xlsx", ".xlsm", ".xls", ".csv"}:
        try:
            return Preview("sheet", evidence, sheet=sheet_window(path, evidence.sheet, evidence.cell))
        except Exception as error:
            preview = Preview("text", evidence)
            preview.notes.append(f"No se pudo abrir la hoja: {error}")
            return preview
    return Preview("text", evidence)


__all__ = [
    "Evidence", "PdfMatch", "Preview", "SheetWindow", "build_preview", "evidence_for_issue",
    "locate_in_pdf", "pdf_words", "render_pdf_page", "sheet_window",
]
