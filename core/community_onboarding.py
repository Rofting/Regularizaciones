"""Análisis no destructivo de fuentes para el alta guiada de comunidades."""

from __future__ import annotations

import hashlib
from dataclasses import dataclass
from pathlib import Path

from openpyxl import load_workbook

from lector_pdf import extraer_texto


_EXCEL_SUFFIXES = frozenset({".xls", ".xlsx"})
_SERVICE_LABELS = {
    "ACS": ("acs", "agua caliente"),
    "AGUA": ("agua",),
    "CALEFACCION": ("calefaccion", "calefacción"),
}


@dataclass(frozen=True)
class SourceCandidate:
    path: Path
    kind: str
    sha256: str
    headers: tuple[str, ...] = ()
    text: str = ""


@dataclass(frozen=True)
class OnboardingQuestion:
    key: str
    prompt: str
    candidates: tuple[str, ...]
    required: bool


@dataclass(frozen=True)
class OnboardingDraft:
    community_code: str
    community_name: str
    sources: tuple[SourceCandidate, ...]
    detected_modules: tuple[str, ...]
    questions: tuple[OnboardingQuestion, ...]


def analyse_sources(
    *,
    community_code: str,
    community_name: str,
    owner_list_path: Path,
    reading_paths: tuple[Path, ...],
    invoice_paths: tuple[Path, ...],
    project_root: Path,
) -> OnboardingDraft:
    """Return a draft from explicit source files without persisting any data."""
    del project_root  # Kept in the public API for later confirmation steps.
    if not reading_paths:
        raise ValueError("Se requiere al menos una fuente de lecturas")

    owner_path = _validated_path(owner_list_path, {".csv"}, "listado de propietarios")
    sources = [_source(owner_path, "owner_list")]
    for reading_path in reading_paths:
        path = _validated_path(reading_path, _EXCEL_SUFFIXES | {".pdf"}, "lectura")
        sources.append(_source(path, "meter_reading_pdf" if path.suffix.lower() == ".pdf" else "meter_reading_excel"))
    for invoice_path in invoice_paths:
        path = _validated_path(invoice_path, {".pdf"}, "factura")
        sources.append(_source(path, "invoice_pdf"))

    readings = tuple(source for source in sources if source.kind.startswith("meter_reading_"))
    return OnboardingDraft(
        community_code=community_code,
        community_name=community_name,
        sources=tuple(sources),
        detected_modules=_detected_modules(readings),
        questions=_questions(readings),
    )


def _validated_path(path: Path, allowed_suffixes: frozenset[str] | set[str], label: str) -> Path:
    resolved = Path(path)
    if not resolved.is_file():
        raise ValueError(f"No existe el archivo de {label}: {resolved}")
    if resolved.suffix.lower() not in allowed_suffixes:
        allowed = ", ".join(sorted(allowed_suffixes))
        raise ValueError(f"Extensión no permitida para {label}: {resolved.suffix} (se permite {allowed})")
    return resolved


def _source(path: Path, kind: str) -> SourceCandidate:
    headers = _excel_headers(path) if kind == "meter_reading_excel" else ()
    text = extraer_texto(str(path)) if path.suffix.lower() == ".pdf" else ""
    return SourceCandidate(path=path, kind=kind, sha256=_sha256(path), headers=headers, text=text)


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as source:
        for block in iter(lambda: source.read(65536), b""):
            digest.update(block)
    return digest.hexdigest()


def _excel_headers(path: Path) -> tuple[str, ...]:
    if path.suffix.lower() == ".xls":
        import xlrd

        sheet = xlrd.open_workbook(path).sheet_by_index(0)
        return tuple(str(sheet.cell_value(0, column)).strip() for column in range(sheet.ncols))
    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        sheet = workbook.active
        first_row = next(sheet.iter_rows(min_row=1, max_row=1, values_only=True), ())
        return tuple(str(value or "").strip() for value in first_row)
    finally:
        workbook.close()


def _detected_modules(readings: tuple[SourceCandidate, ...]) -> tuple[str, ...]:
    material = " ".join(
        " ".join(source.headers) + " " + source.text for source in readings
    ).lower()
    return tuple(
        module for module, labels in _SERVICE_LABELS.items()
        if any(label in material for label in labels)
    )


def _questions(readings: tuple[SourceCandidate, ...]) -> tuple[OnboardingQuestion, ...]:
    questions: list[OnboardingQuestion] = []
    reading_columns = tuple(
        header
        for source in readings
        for header in source.headers
        if "lectura" in header.lower() or "contador" in header.lower()
    )
    if len(reading_columns) > 1:
        questions.append(OnboardingQuestion(
            key="reading_column",
            prompt="Selecciona la columna de lectura que se debe usar.",
            candidates=reading_columns,
            required=True,
        ))
    modules = _detected_modules(readings)
    if len(modules) > 1:
        questions.append(OnboardingQuestion(
            key="service",
            prompt="Selecciona el servicio asociado a las lecturas.",
            candidates=modules,
            required=True,
        ))
    return tuple(questions)
