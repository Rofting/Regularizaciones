"""Detect community identities from a mixed initial source folder.

The first import into an empty office database can contain documents for more
than one community.  This module deliberately favours a safe *no result* over
assigning a supplier invoice to the wrong community.
"""

from __future__ import annotations

import re
import sqlite3
from collections import Counter, defaultdict
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

from lector_pdf import extraer_cif_pdf, extraer_texto
from community_cups import LearnedSupplyPoint, cups_in_text, lookup_supply_point
from document_text_service import get_document_text


_CODE_AT_START = re.compile(r"^\s*(\d{3,6})(?=[_\s-]|$)")
_CODE_AT_END = re.compile(r"(?:^|[_\s-])(\d{3,6})\s*$")
_CODE_LABELLED = re.compile(
    r"\b(?:comunidad|cdad\.?|c\.?\s*p\.?)\s*[-#:]*\s*(\d{3,6})\b",
    re.IGNORECASE,
)
_COMMUNITY_NAME = re.compile(
    r"\b(?:comunidad(?:\s+de\s+propietarios)?|cdad\.?\s*prop\.?|c\.?\s*p\.?)"
    r"\s*[:\-]\s*([^\n]{5,110})",
    re.IGNORECASE,
)
_SUPPORTED = frozenset({".pdf", ".xlsx", ".xls", ".csv"})
_DUPLICATE_SUFFIX = re.compile(r"\s*\((?P<index>\d+)\)$")
_MONTHS = {
    "ene": 1, "enero": 1, "feb": 2, "febrero": 2, "mar": 3, "marzo": 3,
    "abr": 4, "abril": 4, "may": 5, "mayo": 5, "jun": 6, "junio": 6,
    "jul": 7, "julio": 7, "ago": 8, "agosto": 8, "sep": 9, "sept": 9,
    "septiembre": 9, "oct": 10, "octubre": 10, "nov": 11, "noviembre": 11,
    "dic": 12, "diciembre": 12,
}
_SERVICE_HINTS = (
    ("LIMPIEZA", ("limpieza",)),
    ("ELECTRICIDAD", ("electricidad", "luz", "endesa", "iberdrola", "naturgy")),
    ("AGUA", ("agua", "canal", "aqualia", "contador agua")),
    ("CALEFACCION", ("calefaccion", "calefacción", "gas", "acs")),
    ("ASCENSOR", ("ascensor", "elevador")),
    ("MANTENIMIENTO", ("mantenimiento", "mant.")),
)


@dataclass(frozen=True)
class DetectedCommunity:
    """One safely grouped community proposed to the user."""

    code: str
    name: str
    cif: str | None
    source_paths: tuple[Path, ...]
    blocked_reason: str | None = None

    @property
    def can_create(self) -> bool:
        return self.blocked_reason is None


@dataclass(frozen=True)
class FilenameEvidence:
    """Señales seguras presentes en el nombre y la carpeta de una fuente.

    No sustituye la extracción de la factura: permite agrupar documentos de
    correo antes de conocer una comunidad o un período seleccionados.
    """

    community_code: str | None
    category: str
    supply_hint: str | None
    month: int | None
    year: int | None
    duplicate_index: int | None


@dataclass(frozen=True)
class GlobalIntakeGroup:
    """Documentos que pueden revisarse juntos sin contexto seleccionado."""

    community_code: str
    source_paths: tuple[Path, ...]
    period_hints: tuple[tuple[int, int], ...]
    supply_hints: tuple[str, ...]
    categories: tuple[str, ...]

    @property
    def display_periods(self) -> str:
        if not self.period_hints:
            return "Período por confirmar"
        return ", ".join(f"{month:02d}/{year}" for month, year in self.period_hints)


@dataclass(frozen=True)
class GlobalIntakeProposal:
    """Resultado puro de la clasificación inicial de una carpeta mixta."""

    groups: tuple[GlobalIntakeGroup, ...]
    unassigned_paths: tuple[Path, ...]


def filename_evidence(path: Path) -> FilenameEvidence:
    """Extrae indicios conservadores de nombres habituales recibidos por correo."""
    path = Path(path)
    stem = path.stem
    duplicate = _DUPLICATE_SUFFIX.search(stem)
    duplicate_index = int(duplicate.group("index")) if duplicate else None
    if duplicate:
        stem = stem[:duplicate.start()].rstrip()
    normalized = stem.casefold()
    category = _category_from_path(path)
    supply_hint = next(
        (hint for hint, tokens in _SERVICE_HINTS if any(token in normalized for token in tokens)),
        None,
    )
    month = next((number for token, number in _MONTHS.items()
                  if re.search(rf"\b{re.escape(token)}\b", normalized)), None)
    year_match = re.search(r"\b(20\d{2})\b", normalized)
    return FilenameEvidence(
        community_code=code_from_path(path), category=category,
        supply_hint=supply_hint, month=month,
        year=int(year_match.group(1)) if year_match else None,
        duplicate_index=duplicate_index,
    )


def _identity_for_source(
    path: Path, connection: sqlite3.Connection | None,
) -> tuple[str | None, LearnedSupplyPoint | None]:
    code = code_from_path(path)
    learned = None
    has_learned_cups = connection is not None and connection.execute(
        "SELECT 1 FROM learned_supply_points LIMIT 1"
    ).fetchone() is not None
    if has_learned_cups and path.suffix.lower() == ".pdf" and path.is_file():
        text = get_document_text(connection, path).text
        values = cups_in_text(text)
        points = [lookup_supply_point(connection, cups) for cups in values]
        known_communities = {point.community_code for point in points if point}
        if known_communities and (
            len(known_communities) != 1 or any(point is None for point in points)
        ):
            return None, next(point for point in points if point)
        learned = points[0] if points else None
    if learned and code and code != learned.community_code:
        return None, learned
    return code or (learned.community_code if learned else None), learned


def build_global_intake(
    paths: Iterable[Path], *, connection: sqlite3.Connection | None = None,
) -> GlobalIntakeProposal:
    """Agrupa fuentes mixtas por código sin usar la comunidad activa.

    Con una conexión, un CUPS ya confirmado también puede identificar PDFs sin
    código. Una contradicción entre CUPS y nombre queda sin asignar. Los meses
    del nombre son sólo indicios, no fechas oficiales de un expediente.
    """
    grouped: dict[str, list[tuple[Path, FilenameEvidence, LearnedSupplyPoint | None]]] = defaultdict(list)
    unassigned: list[Path] = []
    for raw_path in paths:
        path = Path(raw_path)
        evidence = filename_evidence(path)
        code, learned = _identity_for_source(path, connection)
        if code is None or evidence.category == "unassigned":
            unassigned.append(path)
            continue
        grouped[code].append((path, evidence, learned))

    groups: list[GlobalIntakeGroup] = []
    for code in sorted(grouped, key=lambda value: (len(value), value)):
        entries = sorted(grouped[code], key=lambda item: item[0].name.casefold())
        periods = sorted({(item.month, item.year) for _path, item, _learned in entries
                          if item.month is not None and item.year is not None}, key=lambda value: (value[1], value[0]))
        supplies = sorted({hint for _path, item, learned in entries
                           if (hint := item.supply_hint or (learned.supply_type if learned else None))})
        categories = sorted({item.category for _path, item, _learned in entries})
        groups.append(GlobalIntakeGroup(
            community_code=code,
            source_paths=tuple(path for path, _item, _learned in entries),
            period_hints=tuple(periods), supply_hints=tuple(supplies),
            categories=tuple(categories),
        ))
    return GlobalIntakeProposal(
        groups=tuple(groups),
        unassigned_paths=tuple(sorted(unassigned, key=lambda path: path.name.casefold())),
    )


def _category_from_path(path: Path) -> str:
    names = {parent.name.casefold() for parent in Path(path).parents}
    if "sin_comunidad" in names or "sin comunidad" in names:
        return "unassigned"
    if "extraordinarias" in names or "extraordinaria" in names:
        return "extraordinaria"
    return "habitual"


def discover_communities(
    paths: Iterable[Path], *, connection: sqlite3.Connection | None = None,
) -> tuple[DetectedCommunity, ...]:
    """Agrupa archivos con código explícito o CUPS previamente confirmado.

    El llamador decide si crear las comunidades propuestas. Los originales
    nunca se modifican; con conexión se guarda la caché de texto del PDF.
    """
    grouped: dict[str, list[Path]] = defaultdict(list)
    for source_path in paths:
        path = Path(source_path)
        if path.is_file() and path.suffix.lower() in _SUPPORTED:
            code, _learned = _identity_for_source(path, connection)
            if code:
                grouped[code].append(path)

    candidates: list[DetectedCommunity] = []
    for code in sorted(grouped, key=lambda value: (len(value), value)):
        files = tuple(sorted(grouped[code], key=lambda item: item.name.casefold()))
        existing = connection.execute(
            "SELECT nombre,cif FROM comunidades WHERE codigo=?", (code,),
        ).fetchone() if connection is not None else None
        if existing is not None:
            candidates.append(DetectedCommunity(
                code=code, name=existing[0], cif=existing[1], source_paths=files,
            ))
            continue
        names: list[str] = []
        cifs: list[str] = []
        for path in files:
            if path.suffix.lower() != ".pdf":
                continue
            try:
                text = extraer_texto(str(path))
            except (OSError, ValueError):
                continue
            if not text:
                continue
            name = _name_from_text(text)
            if name:
                names.append(name)
            cif = extraer_cif_pdf(text)
            if cif:
                cifs.append(cif.upper())

        distinct_cifs = tuple(sorted(set(cifs)))
        blocked_reason = (
            "Los documentos de este código contienen CIF distintos; no se creará "
            "hasta que se separen o revisen."
            if len(distinct_cifs) > 1 else None
        )
        candidates.append(DetectedCommunity(
            code=code,
            name=_most_common(names) or f"Comunidad {code}",
            cif=distinct_cifs[0] if len(distinct_cifs) == 1 else None,
            source_paths=files,
            blocked_reason=blocked_reason,
        ))
    return tuple(candidates)


def code_from_path(path: Path) -> str | None:
    """Return a code only when the pathname gives an unambiguous signal."""
    name = Path(path).stem
    match = (
        _CODE_AT_START.search(name)
        or _CODE_LABELLED.search(name)
        or _CODE_AT_END.search(name)
    )
    if match:
        return match.group(1)

    # A folder named ``658`` is a clear declaration by the operator and avoids
    # mistaking invoice numbers inside a filename for community codes.
    for parent in Path(path).parents:
        if re.fullmatch(r"\d{3,6}", parent.name):
            return parent.name
    return None


def partition_sources_for_community(
    paths: Iterable[Path], community_code: str,
    *, connection: sqlite3.Connection | None = None,
) -> tuple[tuple[Path, ...], tuple[Path, ...]]:
    """Aparta fuentes que indican otra comunidad por código o CUPS confirmado.

    Las fuentes sin señal siguen revisables en el expediente seleccionado.
    """
    accepted: list[Path] = []
    foreign: list[Path] = []
    expected = str(community_code).strip()
    for raw_path in paths:
        path = Path(raw_path)
        detected, learned = _identity_for_source(path, connection)
        if learned and detected is None:
            foreign.append(path)
            continue
        (foreign if detected and detected != expected else accepted).append(path)
    return tuple(accepted), tuple(foreign)


def _name_from_text(text: str) -> str | None:
    match = _COMMUNITY_NAME.search(text)
    if not match:
        return None
    candidate = " ".join(match.group(1).split())
    candidate = re.split(r"\s{2,}|\b(?:cif|nif|domicilio|factura)\b", candidate,
                           maxsplit=1, flags=re.IGNORECASE)[0].strip(" -:;,.")
    return candidate if len(candidate) >= 5 else None


def _most_common(values: Iterable[str]) -> str | None:
    normalized = [" ".join(value.split()) for value in values if value.strip()]
    if not normalized:
        return None
    return Counter(normalized).most_common(1)[0][0]
