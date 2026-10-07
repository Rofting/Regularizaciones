"""Repartir una carpeta mixta entre los expedientes de cada comunidad.

La bandeja global decide la comunidad de cada archivo (``community_discovery``).
Aquí se elige el expediente de destino (el abierto más reciente de esa
comunidad) y se incorporan los archivos con el mismo proceso que «Añadir
fuentes»: análisis, archivado, incidencias y aplicación automática de lo
completo. Cada archivo se vuelve a comprobar contra su comunidad antes de
entrar, por si cambió algo entre la revisión y el reparto.
"""

from __future__ import annotations

import sqlite3
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Iterable, Mapping, Sequence


@dataclass(frozen=True)
class ManualAssignment:
    path: Path
    code: str
    identity_reason: str
    justification: str = ""


@dataclass(frozen=True)
class RouteTarget:
    code: str
    community_id: int | None
    community_name: str | None
    case_id: int | None
    case_label: str | None
    paths: tuple[Path, ...]
    manual_assignments: tuple[ManualAssignment, ...] = ()

    @property
    def blocked_reason(self) -> str | None:
        if self.community_id is None:
            return "La comunidad no está dada de alta"
        if self.case_id is None:
            return "La comunidad no tiene un expediente abierto"
        return None


@dataclass
class RouteResult:
    target: RouteTarget
    created: int = 0
    duplicates: int = 0
    foreign: tuple[Path, ...] = ()
    errors: list[tuple[str, str]] = field(default_factory=list)
    automatic: int = 0
    ready_for_calculation: bool = False


def open_case_for(connection: sqlite3.Connection, code: str):
    """(id_comunidad, nombre, id_case, etiqueta) del expediente abierto más reciente."""
    community = connection.execute(
        "SELECT id_comunidad, nombre FROM comunidades WHERE codigo=? AND activa=1", (code,),
    ).fetchone()
    if community is None:
        return None, None, None, None
    case = connection.execute(
        """SELECT id_case, nombre, fecha_inicio, fecha_fin FROM regularization_cases
           WHERE id_comunidad=? AND estado<>'closed'
           ORDER BY fecha_fin DESC, id_case DESC LIMIT 1""",
        (community[0],),
    ).fetchone()
    if case is None:
        return int(community[0]), community[1], None, None
    start, end = (f"{value[8:10]}/{value[5:7]}/{value[:4]}" for value in (case[2], case[3]))
    return int(community[0]), community[1], int(case[0]), f"{case[1]} · {start} – {end}"


def plan_routes(
    connection: sqlite3.Connection, assignments: Mapping[str, Sequence[Path]],
    *, manual_assignments: Sequence[ManualAssignment] = (),
) -> tuple[RouteTarget, ...]:
    targets = []
    for code in sorted(assignments, key=lambda value: (len(value), value)):
        paths = tuple(dict.fromkeys(Path(path) for path in assignments[code]))
        if not paths:
            continue
        community_id, name, case_id, label = open_case_for(connection, code)
        decisions = tuple(item for item in manual_assignments if item.code == code and item.path in paths)
        targets.append(RouteTarget(code, community_id, name, case_id, label, paths, decisions))
    return tuple(targets)


def ingest_route(
    database_path: str | Path,
    target: RouteTarget,
    *,
    archive_root: Path,
    progress: Callable[[str], None] | None = None,
    analyser=None,
) -> RouteResult:
    """Incorpora los archivos del grupo a su expediente; un error no corta el resto."""
    import case_ingestion
    import community_discovery
    import gestor_bd
    import source_batch

    result = RouteResult(target)
    if target.blocked_reason:
        result.errors.append((target.code, target.blocked_reason))
        return result
    connection = gestor_bd.conectar(str(database_path))
    try:
        accepted, foreign = community_discovery.partition_sources_for_community(
            target.paths, target.code, connection=connection,
        )
        accepted = list(accepted)
        foreign = list(foreign)
        valid_manual = {}
        for decision in target.manual_assignments:
            if decision.path not in target.paths or decision.code != target.code:
                continue
            identity = community_discovery.identify_source(decision.path, connection)
            if identity.reason != decision.identity_reason:
                archived_in_target = connection.execute(
                    """SELECT 1 FROM source_documents
                       WHERE id_case=? AND sha256=? LIMIT 1""",
                    (target.case_id, community_discovery._sha256(decision.path)),
                ).fetchone() is not None
                original_identity = community_discovery.identify_source(
                    decision.path, connection, include_archived=False,
                )
                if not archived_in_target or original_identity.reason != decision.identity_reason:
                    result.errors.append((decision.path.name, "Las evidencias han cambiado; vuelve a revisar el reparto"))
                    if decision.path in accepted:
                        accepted.remove(decision.path)
                    continue
            if identity.points_elsewhere and not decision.justification.strip():
                result.errors.append((decision.path.name, "Justifica la elección ante evidencias contradictorias"))
                continue
            if decision.path in foreign:
                foreign.remove(decision.path)
                accepted.append(decision.path)
            if decision.path in accepted:
                valid_manual[decision.path] = decision
        connection.commit()
    finally:
        connection.close()
    result.foreign = tuple(foreign)
    options = {"analyser": analyser} if analyser is not None else {}
    items = source_batch.analyse_batch(
        accepted, community_code=target.code, database_path=database_path, max_workers=3, **options,
    )
    for index, item in enumerate(items, start=1):
        if progress:
            progress(f"{target.code}: {index}/{len(items)} {item.path.name}")
        if item.error is not None or item.analysis is None:
            result.errors.append((item.path.name, str(item.error or "Análisis sin resultado")))
            continue
        connection = gestor_bd.conectar(str(database_path))
        try:
            case_ingestion.assert_case_belongs_to_community(connection, target.case_id, target.community_id)
            added = case_ingestion.add_analysed_document_to_case(
                connection, target.case_id, source_path=item.path,
                archive_root=archive_root, analysis=item.analysis,
                manual_routing=(valid_manual[item.path] if item.path in valid_manual else None),
            )
        except Exception as error:  # se informa por archivo y se sigue
            result.errors.append((item.path.name, str(error)))
            continue
        finally:
            connection.close()
        if added.created:
            result.created += 1
        else:
            result.duplicates += 1
    connection = gestor_bd.conectar(str(database_path))
    try:
        result.automatic, result.ready_for_calculation = case_ingestion.finalize_case_source_intake(
            connection, target.case_id,
        )
    finally:
        connection.close()
    return result


def assignments_from_proposal(proposal, manual: Mapping[Path, str] | None = None) -> dict[str, list[Path]]:
    """Grupos detectados más las asignaciones manuales de los sin asignar."""
    assignments: dict[str, list[Path]] = {
        group.community_code: list(group.source_paths) for group in proposal.groups
    }
    for path, code in (manual or {}).items():
        if code:
            assignments.setdefault(code, []).append(Path(path))
    return assignments


__all__ = [
    "ManualAssignment", "RouteResult", "RouteTarget", "assignments_from_proposal", "ingest_route", "open_case_for", "plan_routes",
]
