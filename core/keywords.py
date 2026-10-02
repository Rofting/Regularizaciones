"""Palabras clave configurables y comparación aproximada de textos.

Las listas viven en ``config/palabras_clave.json`` (producto) y cada despacho
puede ampliarlas en ``config/palabras_clave_despacho.json`` sin tocar el
archivo del producto: las listas del despacho se suman a las del producto.

La comparación aproximada tolera los errores típicos del OCR y de los PDF
(«LECTURA ANTERI0R», «Lect.Ant.»). Usa ``rapidfuzz`` si está instalado y, si
no, ``difflib`` de la biblioteca estándar.
"""

from __future__ import annotations

import json
import re
import unicodedata
from difflib import SequenceMatcher
from functools import lru_cache
from pathlib import Path
from typing import Iterable, Mapping

try:  # pragma: no cover - depende de la instalación
    from rapidfuzz import fuzz as _fuzz
except ImportError:  # pragma: no cover
    _fuzz = None


from app_paths import ApplicationPaths

CONFIG_DIR = ApplicationPaths.resolve().home / "config"
PRODUCT_FILE = "palabras_clave.json"
OFFICE_FILE = "palabras_clave_despacho.json"


def fold(value: object) -> str:
    """Minúsculas, sin acentos ni puntuación, espacios simples."""
    text = unicodedata.normalize("NFKD", str(value or ""))
    text = "".join(char for char in text if not unicodedata.combining(char)).lower()
    text = re.sub(r"[^a-z0-9%]+", " ", text)
    return " ".join(text.split())


def similarity(left: str, right: str) -> float:
    """Parecido 0-100 entre dos textos ya normalizados."""
    if not left or not right:
        return 0.0
    if _fuzz is not None:
        return float(_fuzz.ratio(left, right))
    return SequenceMatcher(None, left, right).ratio() * 100


def partial_similarity(needle: str, haystack: str) -> float:
    """Mejor parecido de ``needle`` con un fragmento de ``haystack``."""
    if not needle or not haystack:
        return 0.0
    if _fuzz is not None:
        return float(_fuzz.partial_ratio(needle, haystack))
    if needle in haystack:
        return 100.0
    size = len(needle)
    best = 0.0
    step = max(1, size // 4)
    for start in range(0, max(1, len(haystack) - size + 1), step):
        best = max(best, SequenceMatcher(None, needle, haystack[start:start + size]).ratio())
    return best * 100


def _merge(base: Mapping[str, object], extra: Mapping[str, object]) -> dict:
    merged = dict(base)
    for key, value in extra.items():
        if key.startswith("_"):
            continue
        current = merged.get(key)
        if isinstance(current, Mapping) and isinstance(value, Mapping):
            merged[key] = _merge(current, value)
        elif isinstance(current, list) and isinstance(value, list):
            merged[key] = current + [item for item in value if item not in current]
        else:
            merged[key] = value
    return merged


@lru_cache(maxsize=4)
def load_keywords(config_dir: str | None = None) -> dict:
    directory = Path(config_dir) if config_dir else CONFIG_DIR
    payload: dict = {}
    for name in (PRODUCT_FILE, OFFICE_FILE):
        path = directory / name
        if path.is_file():
            try:
                data = json.loads(path.read_text(encoding="utf-8"))
            except (OSError, ValueError):
                continue
            if isinstance(data, Mapping):
                payload = _merge(payload, data)
    return payload


def section(name: str) -> dict[str, tuple[str, ...]]:
    """Sección ``{clave: (palabras normalizadas, ...)}`` del archivo."""
    raw = load_keywords().get(name, {})
    if isinstance(raw, list):
        return {name: tuple(fold(item) for item in raw if fold(item))}
    return {
        str(key): tuple(fold(item) for item in values if fold(item))
        for key, values in raw.items()
        if isinstance(values, list)
    }


def word_list(name: str) -> tuple[str, ...]:
    raw = load_keywords().get(name, [])
    return tuple(fold(item) for item in raw if fold(item)) if isinstance(raw, list) else ()


def best_role(
    header: object,
    roles: Mapping[str, Iterable[str]],
    *,
    threshold: float = 88.0,
) -> tuple[str, float] | None:
    """Rol cuyo sinónimo mejor encaja con la cabecera, si supera el umbral.

    Exacto o «empieza por» puntúa 100; si no, se usa la similitud aproximada.
    """
    text = fold(header)
    if not text:
        return None
    best: tuple[str, float] | None = None
    for role, synonyms in roles.items():
        for synonym in synonyms:
            if not synonym:
                continue
            if text == synonym or text.startswith(synonym + " "):
                score = 100.0 + len(synonym) / 100
            else:
                score = similarity(text, synonym)
            if best is None or score > best[1]:
                best = (role, score)
    return best if best and best[1] >= threshold else None


def count_hits(text: str, words: Iterable[str]) -> int:
    folded = f" {fold(text)} "
    return sum(folded.count(f" {word} ") for word in words if word)


__all__ = [
    "best_role",
    "count_hits",
    "fold",
    "load_keywords",
    "partial_similarity",
    "section",
    "similarity",
    "word_list",
]
