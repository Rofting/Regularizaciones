"""Detect community identities from a mixed initial source folder.

The first import into an empty office database can contain documents for more
than one community.  This module deliberately favours a safe *no result* over
assigning a supplier invoice to the wrong community.
"""

from __future__ import annotations

import re
from collections import Counter, defaultdict
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

from lector_pdf import extraer_cif_pdf, extraer_texto


_CODE_AT_START = re.compile(r"^\s*(\d{3,6})(?=[_\s-]|$)")
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


def discover_communities(paths: Iterable[Path]) -> tuple[DetectedCommunity, ...]:
    """Group files that contain an explicit community code.

    The caller decides whether to create the candidates.  No database or file
    is modified here, so the user always sees the proposed communities first.
    """
    grouped: dict[str, list[Path]] = defaultdict(list)
    for source_path in paths:
        path = Path(source_path)
        if path.is_file() and path.suffix.lower() in _SUPPORTED:
            code = code_from_path(path)
            if code:
                grouped[code].append(path)

    candidates: list[DetectedCommunity] = []
    for code in sorted(grouped, key=lambda value: (len(value), value)):
        files = tuple(sorted(grouped[code], key=lambda item: item.name.casefold()))
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
    match = _CODE_AT_START.search(name) or _CODE_LABELLED.search(name)
    if match:
        return match.group(1)

    # A folder named ``658`` is a clear declaration by the operator and avoids
    # mistaking invoice numbers inside a filename for community codes.
    for parent in Path(path).parents:
        if re.fullmatch(r"\d{3,6}", parent.name):
            return parent.name
    return None


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
