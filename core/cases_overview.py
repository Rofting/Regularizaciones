"""Consulta global de expedientes de todas las comunidades.

La pantalla principal trabaja con una comunidad y un expediente. Este panel
responde a la pregunta del despacho: «¿qué tengo pendiente en todas?». Cada
expediente se clasifica en un grupo (bloqueado, pendiente, listo, calculado,
cartas listas, cerrado) a partir de su estado y de sus incidencias abiertas, y
se muestra su última salida (Excel publicado o lote de cartas).
"""

from __future__ import annotations

import sqlite3
import unicodedata
from dataclasses import dataclass
from datetime import date
from typing import Iterable


BLOCKED = "blocked"
PENDING = "pending"
READY = "ready"
CALCULATED = "calculated"
LETTERS = "letters"
CLOSED = "closed"

BUCKETS = (
    (BLOCKED, "Bloqueados"),
    (PENDING, "Pendientes"),
    (READY, "Listos"),
    (CALCULATED, "Calculados"),
    (LETTERS, "Cartas listas"),
    (CLOSED, "Cerrados"),
)
BUCKET_LABELS = dict(BUCKETS)
BUCKET_ROW_LABELS = {
    BLOCKED: "Bloqueado", PENDING: "Pendiente", READY: "Listo para calcular",
    CALCULATED: "Calculado", LETTERS: "Cartas listas", CLOSED: "Cerrado",
}

STATUS_LABELS = {
    "draft": "Borrador",
    "gathering_sources": "Reuniendo fuentes",
    "under_review": "En revisión",
    "ready_for_calculation": "Listo para cálculo",
    "calculated": "Calculado",
    "reconciled": "Reparto cuadrado",
    "deliveries_generated": "Cartas generadas",
    "closed": "Cerrado",
}


@dataclass(frozen=True)
class CaseOverview:
    id_case: int
    community_id: int
    community_code: str
    community_name: str
    case_name: str
    start_date: date
    end_date: date
    status: str
    period_id: int | None
    open_issues: int
    last_output_kind: str | None   # "excel" | "letters"
    last_output_at: str | None
    last_output_path: str | None

    @property
    def bucket(self) -> str:
        return bucket_for(self.status, self.open_issues)

    @property
    def status_label(self) -> str:
        return STATUS_LABELS.get(self.status, self.status)

    @property
    def community_label(self) -> str:
        return f"{self.community_code} — {self.community_name}"

    @property
    def period_label(self) -> str:
        return f"{self.start_date:%d/%m/%Y} – {self.end_date:%d/%m/%Y}"

    @property
    def last_output_label(self) -> str:
        if not self.last_output_kind:
            return "Sin salidas"
        what = "Excel" if self.last_output_kind == "excel" else "Cartas"
        when = (self.last_output_at or "")[:16].replace("T", " ")
        try:
            day = date.fromisoformat(when[:10])
            when = f"{day:%d/%m/%Y}{when[10:]}"
        except ValueError:
            pass
        return f"{what} · {when}".rstrip(" ·")


def bucket_for(status: str, open_issues: int) -> str:
    if status == "closed":
        return CLOSED
    if status == "deliveries_generated":
        return LETTERS
    if status in {"calculated", "reconciled"}:
        return CALCULATED
    if open_issues:
        return BLOCKED
    if status == "ready_for_calculation":
        return READY
    return PENDING


def _fold(text: object) -> str:
    value = unicodedata.normalize("NFKD", str(text or ""))
    return "".join(char for char in value if not unicodedata.combining(char)).casefold()


def list_cases(connection: sqlite3.Connection) -> tuple[CaseOverview, ...]:
    """Todos los expedientes de comunidades activas, con incidencias y salidas."""
    rows = connection.execute(
        """
        WITH issues AS (
            SELECT id_case, COUNT(*) AS open_issues
            FROM review_issues WHERE status='open' GROUP BY id_case
        ),
        outputs AS (
            SELECT id_case, 'excel' AS kind, COALESCE(published_at, updated_at) AS at, output_path
            FROM excel_export_runs WHERE status='published'
            UNION ALL
            SELECT id_case, 'letters', COALESCE(completed_at, updated_at), output_path
            FROM letter_generation_runs WHERE status='completed'
        ),
        latest AS (
            SELECT id_case, kind, at, output_path,
                   ROW_NUMBER() OVER (PARTITION BY id_case ORDER BY at DESC, kind DESC) AS position
            FROM outputs
        )
        SELECT r.id_case, r.id_comunidad, c.codigo, c.nombre, r.nombre AS case_name,
               r.fecha_inicio, r.fecha_fin, r.estado, r.id_periodo,
               COALESCE(i.open_issues, 0) AS open_issues,
               l.kind, l.at, l.output_path
        FROM regularization_cases r
        JOIN comunidades c ON c.id_comunidad = r.id_comunidad
        LEFT JOIN issues i ON i.id_case = r.id_case
        LEFT JOIN latest l ON l.id_case = r.id_case AND l.position = 1
        WHERE c.activa = 1
        ORDER BY c.codigo, r.fecha_fin DESC, r.id_case DESC
        """
    ).fetchall()
    return tuple(
        CaseOverview(
            id_case=int(row[0]), community_id=int(row[1]), community_code=str(row[2]),
            community_name=str(row[3] or ""), case_name=str(row[4]),
            start_date=date.fromisoformat(row[5]), end_date=date.fromisoformat(row[6]),
            status=str(row[7]), period_id=int(row[8]) if row[8] is not None else None,
            open_issues=int(row[9]), last_output_kind=row[10], last_output_at=row[11],
            last_output_path=row[12],
        )
        for row in rows
    )


def filter_cases(
    cases: Iterable[CaseOverview], *, bucket: str | None = None, text: str = "",
    latest_only: bool = False,
) -> tuple[CaseOverview, ...]:
    """Filtra por grupo y por texto (código, nombre de comunidad o expediente).

    ``latest_only`` deja sólo el expediente más reciente de cada comunidad.
    """
    words = _fold(text).split()
    selected = []
    seen: set[int] = set()
    for case in cases:
        if latest_only:
            if case.community_id in seen:
                continue
            seen.add(case.community_id)
        if bucket and case.bucket != bucket:
            continue
        haystack = _fold(f"{case.community_code} {case.community_name} {case.case_name}")
        if all(word in haystack for word in words):
            selected.append(case)
    return tuple(selected)


def bucket_counts(cases: Iterable[CaseOverview]) -> dict[str, int]:
    counts = {key: 0 for key, _label in BUCKETS}
    for case in cases:
        counts[case.bucket] += 1
    return counts


__all__ = [
    "BLOCKED", "BUCKETS", "BUCKET_LABELS", "BUCKET_ROW_LABELS", "CALCULATED", "CLOSED", "CaseOverview", "LETTERS",
    "PENDING", "READY", "STATUS_LABELS", "bucket_counts", "bucket_for", "filter_cases", "list_cases",
]
