"""Resumen de un expediente para la pantalla principal.

Sólo lee: cuenta lo que ya está en la base (facturas, cuotas cobradas, último
reparto y cartas) para que el despacho vea de un vistazo cómo va el ejercicio
sin abrir el Excel. No calcula nada que no haya calculado ya el flujo.
"""

from __future__ import annotations

import sqlite3
from dataclasses import dataclass

SUPPLY_LABELS = {
    "GAS": "Gas",
    "ELECTRICIDAD": "Electricidad",
    "AGUA": "Agua",
    "MANTENIMIENTO": "Mantenimiento",
}


@dataclass(frozen=True)
class SupplyTotal:
    key: str
    label: str
    count: int
    cents: int


@dataclass(frozen=True)
class CaseSummary:
    supplies: tuple[SupplyTotal, ...]
    invoice_count: int
    invoice_cents: int
    charged_fixed_cents: int
    charged_variable_cents: int
    charged_entries: int
    has_distribution: bool
    actual_cents: int
    billed_cents: int
    difference_cents: int
    consumption: float | None
    consumption_unit: str
    owners: int
    owners_to_pay: int
    owners_to_refund: int
    letters: int
    open_issues: int

    @property
    def charged_cents(self) -> int:
        return self.charged_fixed_cents + self.charged_variable_cents


def _cents(value: object) -> int:
    return int(round(float(value or 0) * 100))


def euros(cents: int, *, decimals: bool = True) -> str:
    """1234567 → «12.345,67 €», con el formato que usa el despacho."""
    sign = "-" if cents < 0 else ""
    whole, fraction = divmod(abs(int(cents)), 100)
    grouped = f"{whole:,}".replace(",", ".")
    if not decimals:
        if fraction >= 50:
            grouped = f"{whole + 1:,}".replace(",", ".")
        return f"{sign}{grouped} €"
    return f"{sign}{grouped},{fraction:02d} €"


def quantity(value: float) -> str:
    text = f"{value:,.1f}".replace(",", "·").replace(".", ",").replace("·", ".")
    return text[:-2] if text.endswith(",0") else text


def summarise_case(connection: sqlite3.Connection, case_id: int) -> CaseSummary | None:
    case = connection.execute(
        "SELECT id_case,id_comunidad,id_periodo FROM regularization_cases WHERE id_case=?",
        (case_id,),
    ).fetchone()
    if case is None:
        return None
    community_id, period_id = case["id_comunidad"], case["id_periodo"]

    supplies: list[SupplyTotal] = []
    if period_id is not None:
        for row in connection.execute(
            """SELECT UPPER(COALESCE(tipo_suministro,'')) AS kind, COUNT(*) AS n,
                      COALESCE(SUM(importe_total),0) AS total
               FROM facturas WHERE id_comunidad=? AND id_periodo=?
               GROUP BY kind ORDER BY total DESC""",
            (community_id, period_id),
        ):
            kind = str(row["kind"] or "OTROS")
            supplies.append(SupplyTotal(
                kind, SUPPLY_LABELS.get(kind, kind.capitalize() or "Otros"),
                int(row["n"]), _cents(row["total"]),
            ))

    charged = (0, 0, 0)
    if period_id is not None:
        row = connection.execute(
            """SELECT COALESCE(SUM(CASE WHEN concepto='fija' THEN importe END),0),
                      COALESCE(SUM(CASE WHEN concepto='variable' THEN importe END),0),
                      COUNT(*)
               FROM cuotas_servicio WHERE id_comunidad=? AND id_periodo=?""",
            (community_id, period_id),
        ).fetchone()
        charged = (_cents(row[0]), _cents(row[1]), int(row[2]))

    run = connection.execute(
        """SELECT id_distribution_run FROM distribution_runs
           WHERE id_case=? AND status='completed'
           ORDER BY created_at DESC,id_distribution_run DESC LIMIT 1""",
        (case_id,),
    ).fetchone()
    actual = billed = difference = owners = to_pay = to_refund = 0
    consumption: float | None = None
    unit = ""
    if run is not None:
        totals = connection.execute(
            """SELECT COALESCE(SUM(actual_cents),0),COALESCE(SUM(billed_cents),0),
                      COALESCE(SUM(difference_cents),0)
               FROM owner_distribution_snapshots WHERE id_distribution_run=?""",
            (run[0],),
        ).fetchone()
        actual, billed, difference = (int(value) for value in totals)
        for row in connection.execute(
            """SELECT SUM(difference_cents) AS balance
               FROM owner_distribution_snapshots WHERE id_distribution_run=?
               GROUP BY id_propietario""",
            (run[0],),
        ):
            owners += 1
            balance = int(row["balance"] or 0)
            to_pay += balance > 0
            to_refund += balance < 0
        measured = connection.execute(
            """SELECT concept_key,consumption_unit,SUM(consumption) AS total
               FROM owner_distribution_snapshots
               WHERE id_distribution_run=? AND consumption IS NOT NULL
               GROUP BY concept_key,consumption_unit ORDER BY concept_key LIMIT 1""",
            (run[0],),
        ).fetchone()
        if measured is not None and measured["total"] is not None:
            consumption = float(measured["total"])
            unit = str(measured["consumption_unit"] or "")

    letters = connection.execute(
        """SELECT COUNT(*) FROM generated_letters g
           WHERE g.status='generated' AND g.id_letter_run=(
               SELECT id_letter_run FROM letter_generation_runs
               WHERE id_case=? AND status='completed'
               ORDER BY created_at DESC,id_letter_run DESC LIMIT 1)""",
        (case_id,),
    ).fetchone()[0]
    open_issues = connection.execute(
        "SELECT COUNT(*) FROM review_issues WHERE id_case=? AND status='open'", (case_id,),
    ).fetchone()[0]

    return CaseSummary(
        supplies=tuple(supplies),
        invoice_count=sum(item.count for item in supplies),
        invoice_cents=sum(item.cents for item in supplies),
        charged_fixed_cents=charged[0],
        charged_variable_cents=charged[1],
        charged_entries=charged[2],
        has_distribution=run is not None,
        actual_cents=actual,
        billed_cents=billed,
        difference_cents=difference,
        consumption=consumption,
        consumption_unit=unit,
        owners=owners,
        owners_to_pay=to_pay,
        owners_to_refund=to_refund,
        letters=int(letters),
        open_issues=int(open_issues),
    )


__all__ = ["CaseSummary", "SupplyTotal", "euros", "quantity", "summarise_case"]
