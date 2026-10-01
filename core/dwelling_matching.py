"""Emparejar la vivienda de un informe con la del listado de propietarios.

Cada empresa escribe los pisos a su manera: «1º B», «1B», «PISO 1 PUERTA B»,
«BAJO IZQ.», «0 IZDA», «BL1-1ºB». Aquí se reducen a una forma canónica y sólo
se empareja cuando el resultado es inequívoco. No se usa parecido aproximado:
«1A» y «1B» se parecen mucho, y asignar una lectura a la vivienda de al lado
es peor que pedir una confirmación.
"""

from __future__ import annotations

import re
import unicodedata
from typing import Iterable, Mapping, TypeVar

T = TypeVar("T")

_REPLACEMENTS = (
    (r"\b(?:PISO|PLANTA|PL|PTA|PUERTA|VIVIENDA|VIV|PORTAL|ESC|ESCALERA|NUM|N)\b", " "),
    (r"\b(?:BAJO|BAJOS|BJ|BAJ|PB|PLANTA BAJA)\b", " 0 "),
    (r"\b(?:ENTRESUELO|ENTLO|ENTL|ENT)\b", " EN "),
    (r"\b(?:PRINCIPAL|PRAL)\b", " PR "),
    (r"\b(?:ATICO|ATIC|AT)\b", " AT "),
    (r"\b(?:SOBREATICO|SAT)\b", " SAT "),
    (r"\b(?:IZQUIERDA|IZQDA|IZQ|IZDA|IZD|IZ)\b", " IZ "),
    (r"\b(?:DERECHA|DCHA|DCH|DRCHA|DER|DE|D)\b", " DE "),
    (r"\b(?:CENTRO|CTRO|CENT|CEN)\b", " CE "),
    (r"\b(?:BLOQUE|BLQ|BL)\s*", " BL"),
)


def canonical_dwelling(value: object) -> str:
    """Forma comparable: sin acentos, ordinales, palabras vacías ni separadores."""
    text = unicodedata.normalize("NFKD", str(value or "").upper())
    text = "".join(char for char in text if not unicodedata.combining(char))
    text = re.sub(r"[ºª°]", " ", text)
    text = re.sub(r"[^A-Z0-9]+", " ", text)
    for pattern, replacement in _REPLACEMENTS:
        text = re.sub(pattern, replacement, text)
    # «1 B», «01B» y «1B» son la misma puerta.
    tokens = [token.lstrip("0") or "0" if token.isdigit() else token for token in text.split()]
    return "".join(tokens)


def _simple(value: object) -> str:
    return re.sub(r"[\s.º°ª-]", "", str(value or "").upper())


def match_dwelling(code: object, candidates: Mapping[str, T]) -> T | None:
    """Devuelve el candidato cuyo código coincide de forma inequívoca.

    ``candidates`` asocia el código de vivienda del listado con lo que haya que
    devolver (p. ej. la fila del propietario).
    """
    if not str(code or "").strip():
        return None
    simple = _simple(code)
    for key, value in candidates.items():
        if _simple(key) == simple:
            return value
    wanted = canonical_dwelling(code)
    if not wanted:
        return None
    canonical: dict[str, list[T]] = {}
    for key, value in candidates.items():
        canonical.setdefault(canonical_dwelling(key), []).append(value)
    exact = canonical.get(wanted, [])
    if len(exact) == 1:
        return exact[0]
    if exact:
        return None
    # El informe omite el bloque o portal («1B» frente a «BL2 1B»): sólo vale
    # si una única vivienda del listado termina así.
    if len(wanted) >= 2:
        endings = [items for key, items in canonical.items() if key.endswith(wanted)]
        if len(endings) == 1 and len(endings[0]) == 1:
            return endings[0][0]
    return None


def unmatched(codes: Iterable[object], candidates: Mapping[str, T]) -> list[str]:
    return [str(code) for code in codes if match_dwelling(code, candidates) is None]


__all__ = ["canonical_dwelling", "match_dwelling", "unmatched"]
