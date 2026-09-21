"""Field-level invoice extraction shared by provider families.

Only labelled values are accepted.  Every value carries its rule, confidence
and local text fragment so low-confidence guesses never silently reach a case.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import datetime
from decimal import Decimal, InvalidOperation
from types import MappingProxyType
from typing import Callable, Mapping

from provider_registry import ProviderProfile


EXTRACTOR_VERSION = "invoice-fields-v1"
_DATE = r"(?P<value>\d{1,2}[./-]\d{1,2}[./-]\d{2,4})"
_MONEY = r"(?P<value>-?\d{1,3}(?:\.\d{3})*(?:,\d{2})|-?\d+(?:[.,]\d{2}))"
_QUANTITY = r"(?P<value>-?\d+(?:[.,]\d+)?)"


@dataclass(frozen=True)
class FieldEvidence:
    value: str
    confidence: str
    source: str
    locator: Mapping[str, object]
    rule_id: str
    extractor_version: str = EXTRACTOR_VERSION

    def __post_init__(self) -> None:
        object.__setattr__(self, "locator", MappingProxyType(dict(self.locator)))


@dataclass(frozen=True)
class ExtractionBundle:
    fields: Mapping[str, FieldEvidence]
    diagnostics: tuple[str, ...] = ()

    def __post_init__(self) -> None:
        object.__setattr__(self, "fields", MappingProxyType(dict(self.fields)))
        object.__setattr__(self, "diagnostics", tuple(self.diagnostics))


def _iso_date(value: str) -> str | None:
    cleaned = value.strip().replace(".", "/").replace("-", "/")
    for fmt in ("%d/%m/%Y", "%d/%m/%y"):
        try:
            return datetime.strptime(cleaned, fmt).date().isoformat()
        except ValueError:
            continue
    return None


def _decimal(value: str) -> Decimal | None:
    cleaned = value.strip().replace(" ", "")
    if "," in cleaned:
        cleaned = cleaned.replace(".", "").replace(",", ".")
    try:
        return Decimal(cleaned)
    except InvalidOperation:
        return None


def _quantity(value: str) -> str | None:
    cleaned = value.strip().replace(" ", "")
    if "," in cleaned:
        if "." in cleaned:
            cleaned = cleaned.replace(".", "")
        cleaned = cleaned.replace(",", ".")
    elif re.fullmatch(r"-?\d{1,3}(?:\.\d{3})+", cleaned):
        cleaned = cleaned.replace(".", "")
    try:
        number = Decimal(cleaned)
    except InvalidOperation:
        return None
    return format(number, "f").rstrip("0").rstrip(".") if "." in format(number, "f") else format(number, "f")


def _evidence(
    field_name: str,
    value: str,
    match: re.Match[str] | None,
    *,
    confidence: str = "high",
    family: str = "standard_spanish_invoice",
    source: str = "text",
) -> FieldEvidence:
    fragment = match.group(0).strip()[:500] if match is not None else ""
    return FieldEvidence(
        value=value,
        confidence=confidence,
        source=source,
        locator={"fragment": fragment},
        rule_id=f"{family}:{field_name}:v1",
    )


def _first_match(patterns: tuple[str, ...], text: str) -> re.Match[str] | None:
    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE | re.MULTILINE)
        if match:
            return match
    return None


def _extract_common(profile: ProviderProfile, text: str) -> ExtractionBundle:
    family = profile.extractor_family
    fields: dict[str, FieldEvidence] = {}
    diagnostics: list[str] = []

    period = _first_match(
        (
            rf"\b(?:periodo|per[ií]odo|periodo de facturaci[oó]n|periodo de consumo)"
            rf"[^\d]{{0,35}}{_DATE}\s*(?:a|al|hasta|-)\s*"
            rf"(?P<end>\d{{1,2}}[./-]\d{{1,2}}[./-]\d{{2,4}})",
            rf"\bdesde\s+{_DATE}\s+(?:a|hasta)\s+"
            rf"(?P<end>\d{{1,2}}[./-]\d{{1,2}}[./-]\d{{2,4}})",
        ),
        text,
    )
    if period:
        start = _iso_date(period.group("value"))
        end = _iso_date(period.group("end"))
        if start and end:
            fields["fecha_inicio"] = _evidence(
                "fecha_inicio", start, period, family=family
            )
            fields["fecha_fin"] = _evidence("fecha_fin", end, period, family=family)

    invoice_date = _first_match(
        (
            rf"\bfecha(?:\s+de)?\s+(?:emisi[oó]n(?:\s+de)?\s+)?(?:factura\s*)?[:\-]?\s*{_DATE}",
            rf"\bfactura[^\n]{{0,40}}\bfecha\s*[:\-]?\s*{_DATE}",
        ),
        text,
    )
    if invoice_date:
        value = _iso_date(invoice_date.group("value"))
        if value:
            fields["fecha_factura"] = _evidence(
                "fecha_factura", value, invoice_date, family=family
            )

    number = _first_match(
        (
            r"\b(?:n[º°o.]?\s*(?:de\s*)?)?factura\s*(?:n[º°o.]?\s*)?[:\-]?\s*(?P<value>[A-Z0-9][A-Z0-9/._-]{1,30})",
        ),
        text,
    )
    if number:
        fields["num_factura"] = _evidence(
            "num_factura", number.group("value"), number, family=family
        )

    amount_matches: dict[str, re.Match[str] | None] = {
        "base_imponible": _first_match(
            (rf"\bbase\s+imponible\b[^\d-]{{0,20}}{_MONEY}\s*(?:€|eur)?",), text
        ),
        "iva": _first_match(
            (rf"\biva(?:\s*\(?\d{{1,2}}(?:[.,]\d+)?\s*%\)?)?\b[^\d-]{{0,35}}{_MONEY}\s*(?:€|eur)?",),
            text,
        ),
        "importe_total": _first_match(
            (
                rf"\btotal\s+(?:a\s+pagar|factura|importe)\b[^\d-]{{0,25}}{_MONEY}\s*(?:€|eur)?",
                rf"\bimporte\s+total\b[^\d-]{{0,25}}{_MONEY}\s*(?:€|eur)?",
                rf"\btotal\b[^\d-]{{0,15}}{_MONEY}\s*(?:€|eur)",
            ),
            text,
        ),
    }
    decimal_values: dict[str, Decimal] = {}
    for field_name, match in amount_matches.items():
        if match is None:
            continue
        value = _decimal(match.group("value"))
        if value is None:
            continue
        decimal_values[field_name] = value
        confidence = "medium" if field_name == "importe_total" else "high"
        fields[field_name] = _evidence(
            field_name,
            format(value, ".2f"),
            match,
            confidence=confidence,
            family=family,
        )

    if {"base_imponible", "iva", "importe_total"}.issubset(decimal_values):
        expected = decimal_values["base_imponible"] + decimal_values["iva"]
        total = decimal_values["importe_total"]
        if abs(expected - total) <= Decimal("0.02"):
            current = fields["importe_total"]
            fields["importe_total"] = FieldEvidence(
                current.value,
                "high",
                current.source,
                current.locator,
                f"{family}:importe_total:reconciled:v1",
            )
        else:
            current = fields["importe_total"]
            fields["importe_total"] = FieldEvidence(
                current.value,
                "low",
                current.source,
                current.locator,
                f"{family}:importe_total:not-reconciled:v1",
            )
            diagnostics.append("total_not_reconciled")

    if profile.service_family:
        fields["tipo_suministro"] = _evidence(
            "tipo_suministro",
            profile.service_family,
            None,
            family=family,
            source="provider_profile",
        )
    return ExtractionBundle(fields, tuple(diagnostics))


def _add_quantity(
    bundle: ExtractionBundle,
    text: str,
    profile: ProviderProfile,
    field_name: str,
    unit_pattern: str,
) -> ExtractionBundle:
    match = _first_match(
        (
            rf"\bconsumo(?:\s+total)?\b[^\d-]{{0,35}}{_QUANTITY}\s*{unit_pattern}\b",
        ),
        text,
    )
    if not match:
        return bundle
    value = _quantity(match.group("value"))
    if value is None:
        return bundle
    fields = dict(bundle.fields)
    fields[field_name] = _evidence(
        field_name, value, match, family=profile.extractor_family
    )
    return ExtractionBundle(fields, bundle.diagnostics)


def _extract_standard(profile: ProviderProfile, text: str) -> ExtractionBundle:
    return _extract_common(profile, text)


def _extract_electricity(profile: ProviderProfile, text: str) -> ExtractionBundle:
    return _add_quantity(_extract_common(profile, text), text, profile, "consumo_kwh", r"kwh")


def _extract_gas_fuel(profile: ProviderProfile, text: str) -> ExtractionBundle:
    result = _add_quantity(_extract_common(profile, text), text, profile, "consumo_kwh", r"kwh")
    return _add_quantity(result, text, profile, "consumo_m3", r"m(?:3|³)")


def _extract_water(profile: ProviderProfile, text: str) -> ExtractionBundle:
    return _add_quantity(_extract_common(profile, text), text, profile, "consumo_m3", r"m(?:3|³)")


_FAMILY_EXTRACTORS: Mapping[str, Callable[[ProviderProfile, str], ExtractionBundle]] = {
    "standard_spanish_invoice": _extract_standard,
    "electricity": _extract_electricity,
    "gas_fuel": _extract_gas_fuel,
    "periodic_maintenance": _extract_standard,
    "elevators": _extract_standard,
    "boilers_hvac": _extract_standard,
    "metering_management": _extract_standard,
    "water_public_fees": _extract_water,
}


def extract_invoice_fields(profile: ProviderProfile, text: str) -> ExtractionBundle:
    extractor = _FAMILY_EXTRACTORS.get(profile.extractor_family, _extract_standard)
    return extractor(profile, text or "")


__all__ = [
    "EXTRACTOR_VERSION",
    "ExtractionBundle",
    "FieldEvidence",
    "extract_invoice_fields",
]
