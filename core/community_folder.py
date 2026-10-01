"""Proponer el alta de una comunidad a partir de una carpeta con sus archivos.

El usuario elige la carpeta donde guarda los documentos de una comunidad y
aquí se clasifica cada archivo (listado de propietarios, lecturas, facturas)
y se deducen código, nombre, CIF y período. No se escribe nada: el resultado
rellena el asistente de alta, que sigue pidiendo confirmación.
"""

from __future__ import annotations

import re
from collections import Counter
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable

from community_discovery import code_from_path
from provider_registry import find_tax_ids


SUPPORTED = frozenset({".pdf", ".xlsx", ".xls", ".csv"})
_NAME_PATTERNS = (
    re.compile(
        r"\b(?:comunidad\s+de\s+propietarios|cdad\.?\s*(?:de\s*)?prop(?:ietarios)?\.?|"
        r"com\.?\s*prop\.?|c\.\s?p\.)\s*[:\-]?\s*(?P<name>[A-ZÁÉÍÓÚÑ0-9][^\n]{3,90})",
        re.IGNORECASE,
    ),
)
_NAME_STOP = re.compile(
    r"\s{2,}|\b(?:cif|nif|n\.?i\.?f|domicilio|factura|c/|calle\s+\w+\s+n|fecha|tel)\b",
    re.IGNORECASE,
)


@dataclass
class FolderProposal:
    folder: Path
    code: str | None = None
    name: str | None = None
    cif: str | None = None
    owners: Path | None = None
    readings: list[Path] = field(default_factory=list)
    invoices: list[Path] = field(default_factory=list)
    ignored: list[tuple[Path, str]] = field(default_factory=list)
    start: str | None = None
    end: str | None = None

    @property
    def summary(self) -> str:
        lines = [
            f"Código: {self.code or '— (indícalo)'}",
            f"Nombre: {self.name or '— (indícalo)'}",
            f"CIF: {self.cif or '—'}",
            f"Propietarios: {self.owners.name if self.owners else '— no encontrado'}",
            f"Lecturas: {len(self.readings)} archivo(s)",
            f"Facturas: {len(self.invoices)} archivo(s)",
        ]
        if self.start and self.end:
            lines.append(f"Período de las lecturas: {self.start} → {self.end}")
        if self.ignored:
            lines.append(f"Sin usar: {len(self.ignored)} archivo(s)")
        return "\n".join(lines)


def _files(folder: Path) -> list[Path]:
    return sorted(
        path for path in folder.rglob("*")
        if path.is_file() and path.suffix.lower() in SUPPORTED
        and not path.name.startswith((".", "~$"))
    )


def community_name_from_text(text: str) -> str | None:
    for pattern in _NAME_PATTERNS:
        for match in pattern.finditer(text or ""):
            candidate = _NAME_STOP.split(match.group("name"), maxsplit=1)[0]
            candidate = " ".join(candidate.split()).strip(" -:;,.")
            if len(candidate) >= 4 and not candidate.isdigit():
                return candidate.upper()
    return None


def _default_text(path: Path) -> str:
    from document_text_service import get_document_text
    return get_document_text(None, path, timeout_seconds=60).text


def propose_from_folder(
    folder: str | Path,
    *,
    text_reader: Callable[[Path], str] | None = None,
    progress: Callable[[str], None] | None = None,
) -> FolderProposal:
    from document_classifier import classify_document
    from owner_lists import read_owner_rows
    from source_analysis import analyse_tabular

    folder = Path(folder)
    if not folder.is_dir():
        raise ValueError(f"No existe la carpeta {folder}")
    read_text = text_reader or _default_text
    proposal = FolderProposal(folder)
    owner_candidates: list[tuple[int, Path]] = []
    names: list[str] = []
    tax_ids: list[str] = []
    codes: list[str] = []
    periods: list[tuple[str, str]] = []

    folder_code = folder.name.strip() if re_code(folder.name) else None
    for path in _files(folder):
        if progress:
            progress(path.name)
        code = code_from_path(path)
        if code:
            codes.append(code)
        suffix = path.suffix.lower()
        if suffix in {".csv", ".xlsx", ".xls"}:
            try:
                owners = read_owner_rows(path)
            except (ValueError, OSError, ImportError):
                owners = []
            analysis = analyse_tabular(path)
            if analysis.kind == "reading":
                proposal.readings.append(path)
                if analysis.candidates.get("fecha_inicio") and analysis.candidates.get("fecha_fin"):
                    periods.append((analysis.candidates["fecha_inicio"], analysis.candidates["fecha_fin"]))
            elif owners:
                owner_candidates.append((len(owners), path))
            elif analysis.kind == "other":
                proposal.ignored.append((path, "modelo Excel de referencia"))
            else:
                proposal.ignored.append((path, "tabla no reconocida"))
            continue
        try:
            text = read_text(path)
        except Exception as error:  # OCR/lectura fallida: se informa, no se adivina
            proposal.ignored.append((path, f"no se pudo leer ({error})"))
            continue
        kind = classify_document(text, path.name).kind
        name = community_name_from_text(text)
        if name:
            names.append(name)
        tax_ids.extend(tax_id for tax_id in find_tax_ids(text) if tax_id.startswith("H"))
        if kind in {"invoice", "credit_note"}:
            proposal.invoices.append(path)
        elif kind == "reading":
            proposal.readings.append(path)
        else:
            proposal.ignored.append((path, {
                "quote": "presupuesto", "delivery_note": "albarán",
                "bank_receipt": "justificante bancario", "report": "informe",
                "owners": "listado en PDF (usa Excel o CSV)",
            }.get(kind, "documento no reconocido")))

    if owner_candidates:
        owner_candidates.sort(key=lambda item: -item[0])
        proposal.owners = owner_candidates[0][1]
        for _, path in owner_candidates[1:]:
            proposal.ignored.append((path, "otro listado de propietarios"))
    proposal.code = folder_code or (Counter(codes).most_common(1)[0][0] if codes else None)
    proposal.name = Counter(names).most_common(1)[0][0] if names else None
    proposal.cif = Counter(tax_ids).most_common(1)[0][0] if tax_ids else None
    if periods:
        proposal.start, proposal.end = Counter(periods).most_common(1)[0][0]
    return proposal


def re_code(name: str) -> bool:
    return bool(re.fullmatch(r"\d{3,6}", name.strip()))


__all__ = ["FolderProposal", "community_name_from_text", "propose_from_folder"]
