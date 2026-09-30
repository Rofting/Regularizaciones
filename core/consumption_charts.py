"""Gráficas compactas de consumo para las cartas de regularización."""

from __future__ import annotations

import math
from pathlib import Path
from collections.abc import Iterable, Sequence


FIVE_BAND_LABELS = ("Muy bajo", "Bajo", "Medio", "Alto", "Muy alto")
_NICE_FACTORS = (1, 1.5, 2, 2.5, 3, 4, 5, 6, 8, 10)

_PRIMARY = "#2E6F95"
_SOFT = "#DDE8F3"
_HISTORY = "#A9C9DB"
_TITLE = "#183B56"
_MUTED = "#52606D"
_FAINT = "#7B8794"
_GRID = "#EEF2F5"
_AVERAGE = "#C0392B"


def consumption_band(consumption: float, *, step: float = 10, maximum: float | None = None) -> tuple[tuple[float, float], int]:
    """Devuelve la franja [n·step, (n+1)·step) que contiene el consumo."""
    if step <= 0:
        raise ValueError("step debe ser positivo")
    value = max(float(consumption or 0), 0.0)
    index = int(math.floor(value / step))
    upper = (index + 1) * step
    if maximum is not None and upper > maximum:
        upper = float(maximum)
    return (float(index * step), float(upper)), index


def consumption_bands(consumptions: Iterable[float], *, step: float = 10) -> tuple[tuple[float, float], ...]:
    """Devuelve cinco franjas de ancho ``step``; la última queda abierta."""
    if step <= 0:
        raise ValueError("step debe ser positivo")
    return tuple(
        (float(index * step), float((index + 1) * step))
        for index in range(4)
    ) + ((float(4 * step), float("inf")),)


def _nice_ceil(value: float) -> float:
    """Redondea hacia arriba a 1; 1,5; 2; 2,5; 3; 4; 5; 6 u 8 × 10ⁿ."""
    if value <= 0:
        return 1.0
    magnitude = 10 ** math.floor(math.log10(value))
    for factor in _NICE_FACTORS:
        if factor * magnitude >= value - 1e-9:
            return float(factor * magnitude)
    return float(10 * magnitude)


def adaptive_step(consumptions: Iterable[float]) -> float:
    """Ancho de franja para que los vecinos se repartan en las cinco franjas.

    Con un ancho fijo, comunidades de consumo alto acababan con casi todas las
    viviendas en «Muy alto». Se usa el percentil 95 para que un contador
    anómalo no aplaste el resto de la escala.
    """
    values = sorted(
        float(value) for value in consumptions
        if value is not None and math.isfinite(float(value)) and float(value) >= 0
    )
    if not values or values[-1] <= 0:
        return 10.0
    p95 = values[min(len(values) - 1, int(round(0.95 * (len(values) - 1))))]
    return _nice_ceil(max(p95, values[-1] * 0.5) / 5)


def _fmt(value: float) -> str:
    if abs(value) >= 100:
        return f"{value:,.0f}".replace(",", ".")
    text = f"{value:,.1f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return text[:-2] if text.endswith(",0") else text


def _fmt_edge(value: float) -> str:
    return _fmt(value) if value % 1 else f"{int(value):,}".replace(",", ".")


def render_consumption_charts(
    output_path: str | Path,
    *,
    owner_consumption: float | None,
    neighbor_consumptions: Sequence[float],
    history: Sequence[tuple[str, float]] = (),
    unit: str = "m³",
    step: float | None = None,
    title: str | None = None,
    current_label: str | None = None,
) -> Path | None:
    """Genera una imagen de dos paneles: vecinos por franjas e histórico propio.

    ``history`` contiene ejercicios anteriores; si se indica ``current_label``
    el periodo actual se añade al final, resaltado, para que la comparación
    con años anteriores sea inmediata.
    """
    try:
        import matplotlib
        matplotlib.use("Agg")
        import matplotlib.pyplot as plt
    except ImportError:
        return None
    valid_neighbors = [
        float(value) for value in neighbor_consumptions
        if value is not None and math.isfinite(float(value)) and float(value) >= 0
    ]
    has_owner = owner_consumption is not None and math.isfinite(float(owner_consumption))
    owner_value = max(float(owner_consumption), 0.0) if has_owner else 0.0
    band_step = step if step else adaptive_step(valid_neighbors + ([owner_value] if has_owner else []))
    bands = consumption_bands(valid_neighbors, step=band_step)
    band_labels = [
        f"{label}  {_fmt_edge(low)}–{_fmt_edge(high)}" if not math.isinf(high)
        else f"{label}  > {_fmt_edge(low)}"
        for label, (low, high) in zip(FIVE_BAND_LABELS, bands)
    ]
    counts = [sum(low <= value < high for value in valid_neighbors) for low, high in bands]
    owner_index = min(int(owner_value // band_step), 4) if has_owner else None
    average = sum(valid_neighbors) / len(valid_neighbors) if valid_neighbors else None

    fig, axes = plt.subplots(1, 2, figsize=(7.2, 2.3), dpi=180, gridspec_kw={"width_ratios": (1.15, 1)})
    fig.patch.set_facecolor("#FFFFFF")
    if title:
        fig.suptitle(title, x=0.01, ha="left", fontsize=9, color=_TITLE, fontweight="bold")

    ax = axes[0]
    colors = [_SOFT] * len(bands)
    if owner_index is not None:
        colors[owner_index] = _PRIMARY
    bars = ax.barh(range(len(bands)), counts, color=colors, edgecolor="none", height=0.7)
    ax.set_yticks(range(len(bands)), band_labels, fontsize=6.5)
    ax.set_xlabel(f"Nº de viviendas · franjas de consumo en {unit}", fontsize=6.5, color=_MUTED)
    subtitle = (
        f"Tu consumo: {_fmt(owner_value)} {unit}" if has_owner else "Sin lectura propia en el periodo"
    )
    if average is not None:
        subtitle += f"  ·  media: {_fmt(average)} {unit}"
    ax.set_title(f"Comparación con tus vecinos\n{subtitle}", fontsize=7.5, loc="left",
                 color=_TITLE, pad=6, fontweight="bold", linespacing=1.5)
    ax.invert_yaxis()
    ax.grid(axis="x", color=_GRID, linewidth=0.7)
    ax.set_axisbelow(True)
    ax.spines[:].set_visible(False)
    ax.tick_params(axis="x", labelsize=6, colors=_FAINT)
    ax.tick_params(axis="y", length=0)
    ax.xaxis.get_major_locator().set_params(integer=True)
    peak = max(counts) if counts and max(counts) else 1
    ax.set_xlim(0, peak * 1.35)
    if not valid_neighbors:
        ax.text(0.5, 0.5, "Comparativa no disponible", ha="center", va="center",
                transform=ax.transAxes, fontsize=7.5, color=_FAINT)
    for index, bar in enumerate(bars):
        width = bar.get_width()
        label = f"{int(width)}" if width else ""
        if index == owner_index:
            label = f"{label}  ◀ tú".strip()
        if label:
            ax.text(width + peak * 0.03, bar.get_y() + bar.get_height() / 2, label, va="center",
                    fontsize=6.3, color=_PRIMARY if index == owner_index else _MUTED,
                    fontweight="bold" if index == owner_index else "normal")

    ax = axes[1]
    series = [(str(label), float(value or 0)) for label, value in history]
    highlight = None
    if current_label and has_owner:
        series.append((str(current_label), owner_value))
        highlight = len(series) - 1
    series = series[-5:]
    if highlight is not None:
        highlight = len(series) - 1
    if series:
        labels = [label for label, _ in series]
        values = [value for _, value in series]
        colors = [_PRIMARY if index == highlight else _HISTORY for index in range(len(series))]
        bars = ax.bar(range(len(labels)), values, color=colors, width=0.6)
        ax.set_xticks(range(len(labels)), labels, fontsize=6.2)
        top = max(values + ([average] if average else [])) or 1
        ax.set_ylim(0, top * 1.22)
        for bar, value in zip(bars, values):
            ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + top * 0.02, _fmt(value),
                    ha="center", va="bottom", fontsize=6, color=_MUTED)
        if average is not None and highlight is not None:
            ax.axhline(average, color=_AVERAGE, linewidth=0.8, linestyle=(0, (4, 3)))
            ax.text(len(labels) - 0.5, average, " media\n comunidad", fontsize=5.5, color=_AVERAGE,
                    va="center", ha="left", clip_on=False)
    else:
        ax.text(0.5, 0.5, "Sin histórico disponible", ha="center", va="center",
                transform=ax.transAxes, fontsize=7.5, color=_FAINT)
        ax.set_xticks([])
    ax.set_title("Tu consumo por ejercicio\n ", fontsize=7.5, loc="left", color=_TITLE, pad=6,
                 fontweight="bold", linespacing=1.5)
    ax.set_ylabel(unit, fontsize=6.5, color=_MUTED)
    ax.grid(axis="y", color=_GRID, linewidth=0.7)
    ax.set_axisbelow(True)
    ax.spines[:].set_visible(False)
    ax.tick_params(axis="y", labelsize=6, colors=_FAINT)
    ax.tick_params(axis="x", length=0)
    fig.tight_layout(pad=0.6, w_pad=2.0)
    destination = Path(output_path)
    destination.parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(destination, facecolor=fig.get_facecolor(), bbox_inches="tight")
    plt.close(fig)
    return destination
