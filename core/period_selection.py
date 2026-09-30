"""Ayudas puras para elegir el intervalo de un expediente.

Sin dependencias de interfaz: la ventana de expediente las usa para ofrecer
atajos (ejercicio anterior, año natural, continuar el último expediente...),
aceptar fechas escritas de varias formas y avisar de solapes antes de guardar.
"""

from __future__ import annotations

import calendar
import re
from dataclasses import dataclass
from datetime import date, timedelta
from typing import Iterable, Sequence

MONTH_NAMES = (
    "enero", "febrero", "marzo", "abril", "mayo", "junio", "julio",
    "agosto", "septiembre", "octubre", "noviembre", "diciembre",
)
_MONTHS = {name: index for index, name in enumerate(MONTH_NAMES, start=1)}
_MONTHS.update({name[:3]: index for name, index in list(_MONTHS.items())})
_MONTHS["sept"] = 9
_MONTHS["setiembre"] = 9


@dataclass(frozen=True)
class PeriodPreset:
    key: str
    label: str
    start: date
    end: date

    @property
    def summary(self) -> str:
        return describe_range(self.start, self.end)


@dataclass(frozen=True)
class ExistingCase:
    """Lo mínimo de un expediente existente para detectar solapes."""

    name: str
    start: date
    end: date


def parse_user_date(value: str, *, reference_year: int | None = None) -> date:
    """Interpreta ``31/08/2026``, ``31-8-26``, ``2026-08-31``, ``31.08.2026``,
    ``31 agosto 2026`` o ``31/08`` (con el año de referencia).

    Lanza ``ValueError`` con un mensaje apto para el usuario.
    """
    text = " ".join(str(value or "").strip().lower().split())
    if not text:
        raise ValueError("Falta la fecha.")
    iso = re.fullmatch(r"(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})", text)
    if iso:
        year, month, day = (int(part) for part in iso.groups())
        return _build(year, month, day, text)
    numeric = re.fullmatch(r"(\d{1,2})[-/. ](\d{1,2})(?:[-/. ](\d{2}|\d{4}))?", text)
    if numeric:
        day, month = int(numeric[1]), int(numeric[2])
        year = _year(numeric[3], reference_year, text)
        return _build(year, month, day, text)
    textual = re.fullmatch(
        r"(\d{1,2})\s+(?:de\s+)?([a-záéíóú]+)\.?(?:\s+(?:de\s+)?(\d{2}|\d{4}))?", text
    )
    if textual and textual[2] in _MONTHS:
        year = _year(textual[3], reference_year, text)
        return _build(year, _MONTHS[textual[2]], int(textual[1]), text)
    raise ValueError(f"No se entiende la fecha «{value}». Usa DD/MM/AAAA.")


def _year(raw: str | None, reference_year: int | None, text: str) -> int:
    if raw is None:
        if reference_year is None:
            raise ValueError(f"Falta el año en «{text}».")
        return reference_year
    return int(raw) + (2000 if len(raw) == 2 else 0)


def _build(year: int, month: int, day: int, text: str) -> date:
    try:
        return date(year, month, day)
    except ValueError:
        raise ValueError(f"La fecha «{text}» no existe.") from None


def format_date(value: date) -> str:
    return value.strftime("%d/%m/%Y")


def describe_range(start: date, end: date) -> str:
    days = (end - start).days + 1
    months = _whole_months(start, end)
    extent = f"{months} mes{'es' if months != 1 else ''}" if months else f"{days} días"
    return f"{format_date(start)} → {format_date(end)} · {extent} ({days} días)"


def _whole_months(start: date, end: date) -> int | None:
    """Número de meses si el rango empieza el día 1 y acaba a fin de mes."""
    if start.day != 1 or end.day != calendar.monthrange(end.year, end.month)[1]:
        return None
    return (end.year - start.year) * 12 + end.month - start.month + 1


def default_case_name(start: date, end: date) -> str:
    """Nombre legible: «2025-2026», «Año 2025», «Enero 2026» o «Ene–Mar 2026»."""
    months = _whole_months(start, end)
    if months == 12 and start.month == 1:
        return f"Año {start.year}"
    if months == 12:
        return f"{start.year}-{end.year}"
    if months == 1:
        return f"{MONTH_NAMES[start.month - 1].capitalize()} {start.year}"
    if months and start.year == end.year:
        first = MONTH_NAMES[start.month - 1][:3].capitalize()
        last = MONTH_NAMES[end.month - 1][:3].capitalize()
        return f"{first}–{last} {start.year}"
    if months:
        return f"{MONTH_NAMES[start.month - 1][:3].capitalize()} {start.year} – " \
               f"{MONTH_NAMES[end.month - 1][:3].capitalize()} {end.year}"
    return f"{format_date(start)} – {format_date(end)}"


def _add_months(value: date, months: int) -> date:
    month_index = value.month - 1 + months
    year = value.year + month_index // 12
    month = month_index % 12 + 1
    return date(year, month, min(value.day, calendar.monthrange(year, month)[1]))


def period_presets(
    today: date,
    previous: Sequence[ExistingCase] = (),
    *,
    fiscal_start_month: int = 9,
) -> tuple[PeriodPreset, ...]:
    """Atajos ordenados por probabilidad de uso.

    Si hay expedientes anteriores, el primero continúa justo después del último
    con la misma duración: es el caso habitual de regularizaciones periódicas.
    """
    presets: list[PeriodPreset] = []
    if previous:
        last = max(previous, key=lambda item: item.end)
        start = last.end + timedelta(days=1)
        whole = _whole_months(last.start, last.end)
        end = (
            _add_months(start, whole) - timedelta(days=1)
            if whole else start + (last.end - last.start)
        )
        presets.append(PeriodPreset("continue", f"Continuar «{last.name}»", start, end))

    fiscal_year = today.year if today.month >= fiscal_start_month else today.year - 1
    last_fiscal = date(fiscal_year - 1, fiscal_start_month, 1)
    presets.append(PeriodPreset(
        "last_fiscal",
        f"Ejercicio {last_fiscal.year}-{last_fiscal.year + 1}",
        last_fiscal,
        _add_months(last_fiscal, 12) - timedelta(days=1),
    ))
    current_fiscal = date(fiscal_year, fiscal_start_month, 1)
    presets.append(PeriodPreset(
        "current_fiscal",
        f"Ejercicio {current_fiscal.year}-{current_fiscal.year + 1}",
        current_fiscal,
        _add_months(current_fiscal, 12) - timedelta(days=1),
    ))
    presets.append(PeriodPreset(
        "last_year", f"Año {today.year - 1}", date(today.year - 1, 1, 1), date(today.year - 1, 12, 31),
    ))
    quarter_start = date(today.year, (today.month - 1) // 3 * 3 + 1, 1)
    previous_quarter = _add_months(quarter_start, -3)
    presets.append(PeriodPreset(
        "last_quarter",
        f"Trimestre anterior ({default_case_name(previous_quarter, quarter_start - timedelta(days=1))})",
        previous_quarter,
        quarter_start - timedelta(days=1),
    ))
    month_start = date(today.year, today.month, 1)
    previous_month = _add_months(month_start, -1)
    presets.append(PeriodPreset(
        "last_month",
        f"Mes anterior ({default_case_name(previous_month, month_start - timedelta(days=1))})",
        previous_month,
        month_start - timedelta(days=1),
    ))
    unique: dict[tuple[date, date], PeriodPreset] = {}
    for preset in presets:
        unique.setdefault((preset.start, preset.end), preset)
    return tuple(unique.values())


def validate_range(
    start: date,
    end: date,
    existing: Iterable[ExistingCase] = (),
) -> tuple[list[str], list[str]]:
    """Devuelve ``(errores, avisos)``. Los errores impiden guardar."""
    errors: list[str] = []
    warnings: list[str] = []
    if end < start:
        errors.append("La fecha final es anterior a la inicial.")
        return errors, warnings
    days = (end - start).days + 1
    if days > 400:
        warnings.append(f"El intervalo dura {days} días; revisa que no abarque dos ejercicios.")
    if days < 7:
        warnings.append(f"El intervalo solo dura {days} día(s).")
    for case in existing:
        if start <= case.end and case.start <= end:
            warnings.append(
                f"Se solapa con «{case.name}» ({format_date(case.start)} – {format_date(case.end)}); "
                "las facturas del tramo común podrían contarse dos veces."
            )
    return errors, warnings


__all__ = [
    "ExistingCase",
    "PeriodPreset",
    "default_case_name",
    "describe_range",
    "format_date",
    "parse_user_date",
    "period_presets",
    "validate_range",
]
