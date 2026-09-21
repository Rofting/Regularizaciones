"""Gráficas compactas de consumo para las cartas de regularización."""

from __future__ import annotations

import math
from pathlib import Path
from collections.abc import Iterable, Sequence


FIVE_BAND_LABELS = ("Muy bajo", "Bajo", "Medio", "Alto", "Muy alto")


def consumption_band(consumption: float, *, step: int = 10, maximum: float | None = None) -> tuple[tuple[float, float], int]:
    """Devuelve la franja [n·step, (n+1)·step) que contiene el consumo."""
    if step <= 0:
        raise ValueError("step debe ser positivo")
    value = max(float(consumption or 0), 0.0)
    index = int(math.floor(value / step))
    upper = (index + 1) * step
    if maximum is not None and upper > maximum:
        upper = float(maximum)
    return (float(index * step), float(upper)), index


def consumption_bands(consumptions: Iterable[float], *, step: int = 10) -> tuple[tuple[float, float], ...]:
    """Devuelve las cinco franjas estables usadas en todas las cartas."""
    if step <= 0:
        raise ValueError("step debe ser positivo")
    return tuple(
        (float(index * step), float((index + 1) * step))
        for index in range(4)
    ) + ((float(4 * step), float("inf")),)


def render_consumption_charts(
    output_path: str | Path,
    *,
    owner_consumption: float,
    neighbor_consumptions: Sequence[float],
    history: Sequence[tuple[str, float]] = (),
    unit: str = "m³",
    step: int = 10,
) -> Path | None:
    """Genera una imagen de dos paneles: vecinos por franjas e histórico anual."""
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
    bands = consumption_bands(valid_neighbors, step=step)
    band_labels = [
        f"{label} · {int(low)}–{'+' if math.isinf(high) else int(high)}"
        for label, (low, high) in zip(FIVE_BAND_LABELS, bands)
    ]
    counts = [sum(low <= value < high for value in valid_neighbors) for low, high in bands]
    owner_index = min(max(int(max(float(owner_consumption or 0), 0) // step), 0), 4)

    fig, axes = plt.subplots(1, 2, figsize=(7.2, 2.15), dpi=180, gridspec_kw={"width_ratios": (1.12, 1)})
    fig.patch.set_facecolor("#FFFFFF")
    ax = axes[0]
    colors = ["#DDE8F3"] * len(bands)
    colors[owner_index] = "#2E6F95"
    bars = ax.barh(range(len(bands)), counts, color=colors, edgecolor="none", height=0.72)
    ax.set_yticks(range(len(bands)), band_labels, fontsize=6.5)
    ax.set_xlabel(f"Viviendas · consumo ({unit})", fontsize=7, color="#52606D")
    ax.set_title("Comparación con vecinos", fontsize=8.5, loc="left", color="#183B56", pad=7, fontweight="bold")
    ax.invert_yaxis()
    ax.grid(axis="x", color="#EEF2F5", linewidth=0.7)
    ax.set_axisbelow(True)
    ax.spines[:].set_visible(False)
    ax.tick_params(axis="x", labelsize=6, colors="#7B8794")
    ax.text(0.98, 0.03, f"Tu consumo: {owner_consumption:.1f} {unit}", transform=ax.transAxes,
            ha="right", va="bottom", fontsize=6.5, color="#2E6F95", fontweight="bold")
    if not valid_neighbors:
        ax.text(0.5, 0.5, "Comparativa no disponible", ha="center", va="center",
                transform=ax.transAxes, fontsize=7.5, color="#7B8794")
    for bar in bars:
        if bar.get_width() > 0:
            ax.text(bar.get_width() + 0.15, bar.get_y() + bar.get_height() / 2, f"{int(bar.get_width())}",
                    va="center", fontsize=6, color="#52606D")

    ax = axes[1]
    labels = [str(label) for label, _ in history]
    values = [float(value or 0) for _, value in history]
    if labels:
        bars = ax.bar(range(len(labels)), values, color="#70A9C5", width=0.62)
        ax.set_xticks(range(len(labels)), labels, fontsize=6.5)
        for bar, value in zip(bars, values):
            ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height(), f"{value:.1f}", ha="center", va="bottom", fontsize=6)
    else:
        ax.text(0.5, 0.5, "Sin histórico", ha="center", va="center", transform=ax.transAxes, fontsize=8, color="#7B8794")
    ax.set_title("Tu consumo por ejercicio", fontsize=8.5, loc="left", color="#183B56", pad=7, fontweight="bold")
    ax.set_ylabel(unit, fontsize=7, color="#52606D")
    ax.grid(axis="y", color="#EEF2F5", linewidth=0.7)
    ax.set_axisbelow(True)
    ax.spines[:].set_visible(False)
    ax.tick_params(axis="y", labelsize=6, colors="#7B8794")
    fig.tight_layout(pad=0.7, w_pad=1.4)
    destination = Path(output_path)
    destination.parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(destination, facecolor=fig.get_facecolor(), bbox_inches="tight")
    plt.close(fig)
    return destination
