"""Deriva costes reales de conceptos directamente de facturas confirmadas."""

from __future__ import annotations

import sqlite3
from dataclasses import dataclass
from datetime import date
from decimal import Decimal, ROUND_HALF_UP
from typing import Mapping


class DerivedPricingError(ValueError):
    """Las facturas no permiten obtener un precio reproducible."""


@dataclass(frozen=True)
class DerivedConceptPrice:
    total_net_cents: int
    total_consumption: float
    unit_price_euros: float | None
    invoice_count: int


def _date(value: object, field: str) -> date:
    try:
        return date.fromisoformat(str(value))
    except (TypeError, ValueError):
        raise DerivedPricingError(f"La {field} de una factura no es válida") from None


def _case_value(case: Mapping[str, object] | sqlite3.Row, key: str) -> object:
    try:
        return case[key]
    except (KeyError, IndexError):
        raise DerivedPricingError(f"Falta {key} en el expediente") from None


def derive_unit_price(
    connection: sqlite3.Connection,
    case: Mapping[str, object] | sqlite3.Row,
    *,
    service: str,
    component: str,
) -> DerivedConceptPrice:
    """Suma y prorratea el componente neto de las facturas que se solapan.

    Los componentes ``fixed`` y ``variable`` son importes netos extraídos de
    la factura. IVA e impuestos permanecen separados y no se vuelven a sumar.
    Si una factura relevante no contiene el componente solicitado, el cálculo
    se detiene: usar el total bruto como sustituto produciría un precio falso.
    """
    component = component.strip().lower()
    if component not in {"fixed", "variable"}:
        raise DerivedPricingError(f"Componente de factura no soportado: {component}")
    services = tuple(
        item.strip().upper() for item in service.split("+") if item.strip()
    )
    if not services:
        raise DerivedPricingError("El origen derivado no indica un suministro")
    start = _date(_case_value(case, "fecha_inicio"), "fecha inicial")
    end = _date(_case_value(case, "fecha_fin"), "fecha final")
    if end < start:
        raise DerivedPricingError("El período del expediente no es válido")

    placeholders = ",".join("?" for _ in services)
    rows = connection.execute(
        f"""SELECT f.id_factura,f.fecha_inicio,f.fecha_fin,f.consumo_total,
                   c.amount AS component_amount
              FROM facturas f
              LEFT JOIN invoice_components c
                ON c.id_factura=f.id_factura AND c.component_key=?
             WHERE f.id_comunidad=?
               AND upper(f.tipo_suministro) IN ({placeholders})
               AND f.fecha_inicio IS NOT NULL AND f.fecha_fin IS NOT NULL
               AND f.fecha_inicio<=? AND f.fecha_fin>=?
             ORDER BY f.fecha_inicio,f.id_factura""",
        (
            component, int(_case_value(case, "id_comunidad")), *services,
            end.isoformat(), start.isoformat(),
        ),
    ).fetchall()
    if not rows:
        raise DerivedPricingError(
            f"No hay facturas de {' + '.join(services)} dentro del período"
        )

    net_total = Decimal(0)
    consumption_total = Decimal(0)
    for row in rows:
        if row["component_amount"] is None:
            raise DerivedPricingError(
                f"La factura {row['id_factura']} no contiene el componente {component}"
            )
        invoice_start = _date(row["fecha_inicio"], "fecha inicial")
        invoice_end = _date(row["fecha_fin"], "fecha final")
        if invoice_end < invoice_start:
            raise DerivedPricingError(
                f"La factura {row['id_factura']} tiene un intervalo invertido"
            )
        overlap_start = max(start, invoice_start)
        overlap_end = min(end, invoice_end)
        overlap_days = (overlap_end - overlap_start).days + 1
        invoice_days = (invoice_end - invoice_start).days + 1
        ratio = Decimal(overlap_days) / Decimal(invoice_days)
        net_total += Decimal(str(row["component_amount"])) * ratio
        if component == "variable":
            consumption_total += Decimal(str(row["consumo_total"] or 0)) * ratio

    cents = int((net_total * 100).quantize(Decimal("1"), ROUND_HALF_UP))
    unit_price = None
    if consumption_total > 0:
        unit_price = float(net_total / consumption_total)
    return DerivedConceptPrice(
        total_net_cents=cents,
        total_consumption=float(consumption_total),
        unit_price_euros=unit_price,
        invoice_count=len(rows),
    )


def derive_source_cents(
    connection: sqlite3.Connection,
    case: Mapping[str, object] | sqlite3.Row,
    source: str,
) -> int:
    """Resuelve ``derived_invoices.<servicios>.<fixed|variable>``."""
    prefix = "derived_invoices."
    if not source.startswith(prefix):
        raise DerivedPricingError(f"Fuente derivada no soportada: {source}")
    parts = source[len(prefix):].rsplit(".", 1)
    if len(parts) != 2 or not all(parts):
        raise DerivedPricingError(
            "La fuente derivada debe indicar suministro y componente"
        )
    return derive_unit_price(
        connection, case, service=parts[0], component=parts[1]
    ).total_net_cents
