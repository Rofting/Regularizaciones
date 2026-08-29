"""Tipos inmutables usados por la importación de libros de referencia."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from decimal import Decimal


@dataclass(frozen=True)
class SourceValue:
    entity_type: str
    entity_key: str
    field_name: str
    sheet_name: str
    cell_address: str
    raw_value: str
    normalized_value: str


@dataclass(frozen=True)
class ConceptReference:
    concept_key: str
    billed_cents: int
    actual_cents: int
    billed_unit_price: Decimal | None
    actual_unit_price: Decimal | None
    source_values: tuple[SourceValue, ...]


@dataclass(frozen=True)
class ExerciseReference:
    community_name: str
    period_name: str
    property_count: int
    start_date: date
    end_date: date
    total_billed_cents: int
    total_actual_cents: int
    concepts: tuple[ConceptReference, ...]
    source_values: tuple[SourceValue, ...]
