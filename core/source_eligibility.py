"""Pure eligibility rules separating recognition from accounting use."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date


SERVICE_MODULES = {
    "ELECTRICIDAD": ("ELECTRICIDAD",),
    "GAS": ("GAS",),
    "GASOLEO": ("GAS",),
    "COMBUSTIBLE": ("GAS",),
    "AGUA": ("AGUA",),
    "TASAS": ("AGUA",),
    "CANON": ("AGUA",),
    "MANTENIMIENTO": ("OTROS_GASTOS",),
    "LIMPIEZA": ("OTROS_GASTOS",),
    "ASCENSORES": ("OTROS_GASTOS",),
    "CALEFACCION": ("CALEFACCION",),
    "CALDERAS": ("CALEFACCION",),
    "CONTADORES": ("ACS", "CALEFACCION"),
}


@dataclass(frozen=True)
class EligibilityDecision:
    status: str
    reason: str | None
    concept_keys: tuple[str, ...] = ()


def evaluate_invoice_eligibility(
    *,
    community_code: str | None,
    case_start: date,
    case_end: date,
    active_modules: tuple[str, ...] | list[str],
    document_kind: str,
    service_family: str | None,
    period_start: date | None,
    period_end: date | None,
    community_confidence: str,
) -> EligibilityDecision:
    if document_kind.lower() not in {"invoice", "credit_note"}:
        return EligibilityDecision("not_applicable", "document_type_not_billable")
    if not community_code or community_confidence.lower() != "high":
        return EligibilityDecision("review_required", "community_unknown")
    if period_start is None or period_end is None or period_start > period_end:
        return EligibilityDecision("review_required", "period_unknown")
    if period_end < case_start or period_start > case_end:
        return EligibilityDecision("review_required", "period_outside_case")

    family = str(service_family or "").strip().upper()
    concepts = SERVICE_MODULES.get(family)
    if not concepts:
        return EligibilityDecision("review_required", "service_unknown")
    active = {str(item).strip().upper() for item in active_modules}
    applicable = tuple(concept for concept in concepts if concept in active)
    if not applicable:
        return EligibilityDecision("not_applicable", "service_not_active")
    return EligibilityDecision("eligible", None, applicable)


__all__ = [
    "EligibilityDecision",
    "SERVICE_MODULES",
    "evaluate_invoice_eligibility",
]
