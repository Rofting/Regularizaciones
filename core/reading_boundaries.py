"""Lecturas que abren y cierran un período, con un único criterio.

La lectura inicial es la última tomada antes de que arranque el ejercicio y la
final, la última que cabe dentro. Se buscan primero entre las lecturas del
propio período: el cierre del ejercicio anterior puede estar fechado unos días
después que la apertura que trae el informe de este (31 de julio frente a
«7/2025»), y elegirlo dejaba al período sin lectura inicial propia.
"""

from __future__ import annotations

import sqlite3


def boundary_dates(
    connection: sqlite3.Connection, *, community_id: int, period_id: int | None,
    reading_type: str, start: str, end: str,
) -> tuple[str | None, str | None]:
    def latest(limit: str, own_period: bool) -> str | None:
        query = """SELECT MAX(r.fecha_lectura) FROM period_readings r
                   JOIN propietarios p ON p.id_propietario=r.id_propietario
                   WHERE p.id_comunidad=? AND p.tipo_unidad='vivienda' AND r.tipo=?
                     AND r.fecha_lectura<=?"""
        parameters: list[object] = [community_id, reading_type, limit]
        if own_period:
            query += " AND r.id_periodo=?"
            parameters.append(period_id)
        return connection.execute(query, parameters).fetchone()[0]

    own = period_id is not None
    initial = (latest(start, True) if own else None) or latest(start, False)
    final = (latest(end, True) if own else None) or latest(end, False)
    return initial, final


__all__ = ["boundary_dates"]
