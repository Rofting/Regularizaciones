"""Preparación de páginas escaneadas antes del OCR.

Los escaneos de los despachos llegan torcidos, tenues o con puntos de polvo.
Aquí se mide cada problema y sólo se corrige el que existe:

* enderezado: se busca el giro (±5°) que deja las líneas de texto más nítidas;
* ruido: un filtro de mediana cuando hay muchos puntos sueltos;
* contraste: se estira el histograma cuando la página es gris o tenue.

La imagen original nunca se modifica. Quien llama compara el OCR del original
con el de la versión preparada y se queda con el mejor, de modo que una página
que ya se leía bien no empeora (ver ``ocr_score``).
"""

from __future__ import annotations

import re
from dataclasses import dataclass

import numpy as np
from PIL import Image, ImageFilter, ImageOps

_ANALYSIS_WIDTH = 900
_MAX_SKEW = 5.0
_SKEW_STEP = 0.25
_MIN_SKEW = 0.4          # por debajo no merece la pena girar
_LOW_CONTRAST = 110      # papel menos tinta, en escala de grises
_NOISE_RATIO = 0.04      # proporción de píxeles oscuros aislados


@dataclass(frozen=True)
class ScanQuality:
    skew_degrees: float
    contrast: float
    noise_ratio: float

    @property
    def problems(self) -> tuple[str, ...]:
        found = []
        if abs(self.skew_degrees) >= _MIN_SKEW:
            found.append("torcida")
        if self.noise_ratio > _NOISE_RATIO:
            found.append("ruido")
        if self.contrast < _LOW_CONTRAST:
            found.append("poco contraste")
        return tuple(found)


@dataclass(frozen=True)
class PreparedScan:
    image: Image.Image
    steps: tuple[str, ...]
    quality: ScanQuality

    @property
    def changed(self) -> bool:
        return bool(self.steps)


def _grey(image: Image.Image) -> Image.Image:
    return image.convert("L")


def _small(grey: Image.Image) -> Image.Image:
    if grey.width <= _ANALYSIS_WIDTH:
        return grey
    height = max(1, round(grey.height * _ANALYSIS_WIDTH / grey.width))
    return grey.resize((_ANALYSIS_WIDTH, height), Image.Resampling.BILINEAR)


def _otsu(values: np.ndarray) -> int:
    histogram = np.bincount(values.ravel(), minlength=256).astype(float)
    total = histogram.sum()
    if total == 0:
        return 128
    cumulative = np.cumsum(histogram)
    cumulative_mean = np.cumsum(histogram * np.arange(256))
    mean = cumulative_mean[-1] / total
    with np.errstate(divide="ignore", invalid="ignore"):
        between = (mean * cumulative - cumulative_mean) ** 2 / (cumulative * (total - cumulative))
    between = np.nan_to_num(between)
    return int(np.argmax(between))


def _dark_mask(grey: Image.Image) -> np.ndarray:
    values = np.asarray(grey, dtype=np.uint8)
    return values < _otsu(values)


def estimate_skew(grey: Image.Image) -> float:
    """Giro (grados, positivo antihorario) que maximiza la nitidez de las líneas."""
    small = _small(grey)
    mask = _dark_mask(small)
    if mask.mean() < 0.002:  # página casi en blanco
        return 0.0
    binary = Image.fromarray((mask * 255).astype(np.uint8))
    best_angle, best_score = 0.0, -1.0
    for angle in np.arange(-_MAX_SKEW, _MAX_SKEW + 1e-9, _SKEW_STEP):
        rotated = np.asarray(binary.rotate(float(angle), resample=Image.Resampling.NEAREST, fillcolor=0))
        rows = rotated.sum(axis=1, dtype=np.float64)
        score = float(np.var(rows))
        if score > best_score:
            best_angle, best_score = float(angle), score
    return round(best_angle, 2)


def _noise_ratio(grey: Image.Image) -> float:
    mask = _dark_mask(_small(grey))
    dark = int(mask.sum())
    if dark == 0:
        return 0.0
    padded = np.pad(mask.astype(np.uint8), 1)
    neighbours = sum(
        padded[1 + dy:padded.shape[0] - 1 + dy, 1 + dx:padded.shape[1] - 1 + dx]
        for dy in (-1, 0, 1) for dx in (-1, 0, 1) if dy or dx
    )
    isolated = int(np.logical_and(mask, neighbours <= 1).sum())
    return isolated / dark


def _ink_contrast(grey: Image.Image) -> float:
    """Diferencia entre la tinta y el papel (no entre percentiles de toda la página,
    que en una hoja casi blanca caen todos en el blanco)."""
    values = np.asarray(_small(grey), dtype=np.uint8)
    threshold = _otsu(values)
    ink, paper = values[values < threshold], values[values >= threshold]
    if ink.size == 0 or paper.size == 0:
        return 255.0
    return float(np.median(paper) - np.median(ink))


def assess(image: Image.Image) -> ScanQuality:
    grey = _grey(image)
    return ScanQuality(
        skew_degrees=estimate_skew(grey),
        contrast=_ink_contrast(grey),
        noise_ratio=round(_noise_ratio(grey), 4),
    )


def prepare(image: Image.Image) -> PreparedScan:
    """Devuelve una copia corregida sólo en lo que lo necesita."""
    quality = assess(image)
    grey = _grey(image)
    steps: list[str] = []
    if quality.noise_ratio > _NOISE_RATIO:
        grey = grey.filter(ImageFilter.MedianFilter(3))
        steps.append("reducción de ruido")
    if quality.contrast < _LOW_CONTRAST:
        grey = ImageOps.autocontrast(grey, cutoff=1)
        steps.append("contraste")
    if abs(quality.skew_degrees) >= _MIN_SKEW:
        grey = grey.rotate(quality.skew_degrees, resample=Image.Resampling.BICUBIC,
                           expand=True, fillcolor=255)
        steps.append(f"enderezado {quality.skew_degrees:+.1f}°")
    return PreparedScan(grey if steps else image, tuple(steps), quality)


_DATE = re.compile(r"\b\d{1,2}[/.-]\d{1,2}[/.-]\d{2,4}\b")
_AMOUNT = re.compile(r"\b\d{1,3}(?:[.\s]\d{3})*,\d{2}\b|\b\d+[.,]\d{2}\b")
_TAX_ID = re.compile(r"\b[A-HJ-NP-SUVW]\d{7}[0-9A-J]\b|\b\d{8}[A-Z]\b")
_WORD = re.compile(r"[A-Za-zÁÉÍÓÚÑáéíóúñ]{3,}")


def ocr_score(text: str) -> float:
    """Cuánto dato útil trae un OCR: fechas, importes y CIF pesan más que letras."""
    text = text or ""
    return (5 * len(_DATE.findall(text)) + 4 * len(_AMOUNT.findall(text))
            + 4 * len(_TAX_ID.findall(text)) + 0.2 * len(_WORD.findall(text)))


@dataclass(frozen=True)
class OcrChoice:
    text: str
    steps: tuple[str, ...]
    original_score: float
    prepared_score: float | None

    @property
    def used_prepared(self) -> bool:
        return bool(self.steps)

    @property
    def summary(self) -> str:
        if self.prepared_score is None:
            return "Imagen sin tratamiento (no lo necesitaba)."
        if self.used_prepared:
            return (f"Imagen preparada ({', '.join(self.steps)}): lectura {self.prepared_score:.0f} "
                    f"frente a {self.original_score:.0f} del original.")
        return (f"Se probó la imagen preparada pero el original se leía igual o mejor "
                f"({self.original_score:.0f} frente a {self.prepared_score:.0f}).")


def best_ocr(image: Image.Image, read) -> OcrChoice:
    """Lee el original y, si la página tiene problemas, también la preparada.

    ``read`` recibe una imagen PIL y devuelve el texto del OCR. Sólo se usa la
    versión preparada si obtiene estrictamente más datos útiles.
    """
    original = read(image) or ""
    original_score = ocr_score(original)
    try:
        prepared = prepare(image)
    except Exception:  # la preparación nunca debe impedir la lectura del original
        return OcrChoice(original, (), original_score, None)
    if not prepared.changed:
        return OcrChoice(original, (), original_score, None)
    candidate = read(prepared.image) or ""
    candidate_score = ocr_score(candidate)
    if candidate_score > original_score:
        return OcrChoice(candidate, prepared.steps, original_score, candidate_score)
    return OcrChoice(original, (), original_score, candidate_score)


__all__ = ["OcrChoice", "PreparedScan", "ScanQuality", "assess", "best_ocr", "estimate_skew", "ocr_score", "prepare"]
