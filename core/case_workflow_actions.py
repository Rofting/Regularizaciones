"""Controladores sin interfaz para el flujo guiado por expediente.

Cada entrada verifica la comunidad seleccionada antes de delegar en los
servicios de importación, Excel, reparto o cartas.  La interfaz sólo traduce
los eventos recibidos; no contiene reglas de negocio ni rutas implícitas.
"""

from __future__ import annotations

import sqlite3
from pathlib import Path
from typing import Any, Callable

import case_distribution
import case_ingestion
import case_letter_service
import document_review
import excel_bootstrap_importer
import excel_export_service
import expedient_service
from excel_profiles import ExcelProfile, load_profile


class WorkflowBlockedError(ValueError):
    """La acción no puede ejecutarse desde el expediente seleccionado."""


ProgressCallback = Callable[[str, dict[str, Any]], None]


def _emit(progress: ProgressCallback | None, stage: str, **payload: Any) -> None:
    if progress is not None:
        progress(stage, payload)


def _connection(database_path: str | Path) -> sqlite3.Connection:
    connection = sqlite3.connect(str(database_path))
    connection.row_factory = sqlite3.Row
    return connection


def resolve_case_profile(
    connection: sqlite3.Connection,
    *,
    id_case: int,
    active_community_id: int,
    project_root: Path,
) -> ExcelProfile:
    """Confirma pertenencia y resuelve el perfil registrado o de bootstrap."""
    try:
        case = case_ingestion.assert_case_belongs_to_community(
            connection, id_case, active_community_id
        )
    except LookupError as error:
        raise WorkflowBlockedError(str(error)) from error
    rows = connection.execute(
        """SELECT co.codigo,profiles.profile_key,profiles.profile_version
           FROM comunidades co
           LEFT JOIN excel_template_profiles profiles
             ON profiles.id_comunidad=co.id_comunidad AND profiles.status='active'
           WHERE co.id_comunidad=?
           ORDER BY profiles.id_template_profile""",
        (case.community_id,),
    ).fetchall()
    if not rows:
        raise WorkflowBlockedError("La comunidad no existe")
    row = rows[0]
    if not row["profile_key"]:
        root = Path(project_root)
        directory = root / "config" / "excel_profiles"
        candidates: list[ExcelProfile] = []
        if directory.is_dir():
            for path in sorted(directory.glob("*.json")):
                try:
                    configured = load_profile(path.stem, root)
                except (LookupError, ValueError):
                    continue
                if (
                    configured.community_code == str(row["codigo"])
                    and configured.workbook_layout
                ):
                    candidates.append(configured)
        if not candidates:
            raise WorkflowBlockedError(
                "No hay un perfil Excel de configuración compatible para esta comunidad"
            )
        if len(candidates) != 1:
            raise WorkflowBlockedError(
                "La comunidad tiene varios perfiles Excel de configuración compatibles"
            )
        return candidates[0]
    if len(rows) != 1:
        raise WorkflowBlockedError("La comunidad tiene varios perfiles Excel activos")
    try:
        profile = load_profile(str(row["profile_key"]), Path(project_root))
    except (LookupError, ValueError) as error:
        raise WorkflowBlockedError(f"No se puede cargar el perfil Excel activo: {error}") from error
    if profile.community_code != str(row["codigo"]):
        raise WorkflowBlockedError("El perfil Excel activo no corresponde a la comunidad")
    if profile.version != str(row["profile_version"]):
        raise WorkflowBlockedError("La versión del perfil Excel activo no coincide con su registro")
    return profile


def _with_case_profile(
    database_path: str | Path,
    *,
    id_case: int,
    active_community_id: int,
    project_root: Path,
) -> tuple[sqlite3.Connection, ExcelProfile]:
    connection = _connection(database_path)
    try:
        profile = resolve_case_profile(
            connection,
            id_case=id_case,
            active_community_id=active_community_id,
            project_root=project_root,
        )
    except Exception:
        connection.close()
        raise
    return connection, profile


def _forward_progress(
    progress: ProgressCallback | None,
    human_stage: str,
) -> Callable[[str, dict[str, Any]], None]:
    def forward(technical_stage: str, payload: dict[str, Any]) -> None:
        _emit(progress, human_stage, technical_stage=technical_stage, **payload)
    return forward


def run_bootstrap_import(
    database_path: str | Path,
    *,
    id_case: int,
    active_community_id: int,
    project_root: Path,
    master_path: str | Path,
    owner_list_path: str | Path | None = None,
    readings_path: str | Path | None = None,
    actor: str = "usuario_local",
    progress: ProgressCallback | None = None,
):
    """Carga sólo el modelo histórico inicial y sus complementos opcionales."""
    if bool(owner_list_path) != bool(readings_path):
        raise WorkflowBlockedError(
            "Selecciona ambas fuentes complementarias o deja ambas vacías"
        )
    _emit(progress, "validate_case", id_case=id_case)
    connection, profile = _with_case_profile(
        database_path, id_case=id_case, active_community_id=active_community_id,
        project_root=project_root,
    )
    try:
        _emit(progress, "importar_modelo", id_case=id_case)
        result = excel_bootstrap_importer.import_master_excel(
            connection, id_case=id_case, workbook_path=Path(master_path),
            profile=profile, actor=actor, project_root=Path(project_root),
        )
        if owner_list_path and readings_path:
            _emit(progress, "importar_complementarias", id_case=id_case)
            companions = excel_bootstrap_importer.import_companion_sources(
                connection, id_case=id_case, owner_list_path=owner_list_path,
                readings_path=readings_path, profile=profile, actor=actor,
            )
            open_issues = companions.open_issue_count
        else:
            companions = None
            open_issues = result.open_issue_count
        if not open_issues:
            document_review.validate_case_ready(connection, id_case)
        return result, companions
    finally:
        connection.close()


def run_generate_excel(
    database_path: str | Path,
    *,
    id_case: int,
    active_community_id: int,
    project_root: Path,
    output_root: Path,
    progress: ProgressCallback | None = None,
):
    _emit(progress, "validate_case", id_case=id_case)
    connection, _profile = _with_case_profile(
        database_path, id_case=id_case, active_community_id=active_community_id,
        project_root=project_root,
    )
    try:
        _emit(progress, "generar_excel", id_case=id_case)
        return generate_official_excel(
            connection, id_case=id_case, project_root=Path(project_root),
            output_root=Path(output_root),
            progress=_forward_progress(progress, "generar_excel"),
        )
    finally:
        connection.close()


def run_calculate_distribution(
    database_path: str | Path,
    *,
    id_case: int,
    active_community_id: int,
    project_root: Path,
    progress: ProgressCallback | None = None,
):
    _emit(progress, "validate_case", id_case=id_case)
    connection, _profile = _with_case_profile(
        database_path, id_case=id_case, active_community_id=active_community_id,
        project_root=project_root,
    )
    try:
        _emit(progress, "calcular_reparto", id_case=id_case)
        result = calculate_case_distribution(
            connection, id_case=id_case,
            progress=_forward_progress(progress, "calcular_reparto"),
        )
        case = expedient_service.get_case(connection, id_case)
        if case.status == "ready_for_calculation":
            case = expedient_service.set_case_status(connection, id_case, "calculated")
        if case.status == "calculated":
            expedient_service.set_case_status(connection, id_case, "reconciled")
        return result
    finally:
        connection.close()


def run_generate_letters(
    database_path: str | Path,
    *,
    id_case: int,
    active_community_id: int,
    project_root: Path,
    selected_concepts: tuple[str, ...],
    progress: ProgressCallback | None = None,
):
    _emit(progress, "validate_case", id_case=id_case)
    connection, _profile = _with_case_profile(
        database_path, id_case=id_case, active_community_id=active_community_id,
        project_root=project_root,
    )
    connection.close()
    _emit(progress, "generar_cartas", id_case=id_case)
    result = generate_case_letters(
        database_path, id_case=id_case, project_root=Path(project_root),
        selected_concepts=selected_concepts,
        progress=_forward_progress(progress, "generar_cartas"),
    )
    if not result.failures:
        connection = _connection(database_path)
        try:
            case_ingestion.assert_case_belongs_to_community(
                connection, id_case, active_community_id
            )
            case = expedient_service.get_case(connection, id_case)
            if case.status == "reconciled":
                expedient_service.set_case_status(
                    connection, id_case, "deliveries_generated"
                )
        finally:
            connection.close()
    return result


def available_case_letter_concepts(
    database_path: str | Path,
    *,
    id_case: int,
    active_community_id: int,
    project_root: Path,
) -> tuple[tuple[str, str], ...]:
    """Devuelve únicamente partidas calculadas que el perfil mantiene activas."""
    connection, profile = _with_case_profile(
        database_path, id_case=id_case, active_community_id=active_community_id,
        project_root=project_root,
    )
    try:
        case = case_ingestion.assert_case_belongs_to_community(
            connection, id_case, active_community_id
        )
        if case.period_id is None:
            raise WorkflowBlockedError("El expediente no tiene un período ligado")
        keys = tuple(concept.key for concept in profile.concepts)
        if not keys:
            return ()
        rows = connection.execute(
            f"""SELECT rc.concept_key,rc.label
                FROM regularization_concepts rc
                WHERE rc.active=1 AND rc.concept_key IN ({','.join('?' for _ in keys)})
                  AND EXISTS (
                    SELECT 1 FROM owner_concept_results results
                    JOIN propietarios owners ON owners.id_propietario=results.id_propietario
                    WHERE results.id_periodo=? AND results.concept_key=rc.concept_key
                      AND owners.id_comunidad=? AND owners.activo=1
                  )""",
            (*keys, case.period_id, case.community_id),
        ).fetchall()
        by_key = {row["concept_key"]: row["label"] for row in rows}
        return tuple((key, by_key[key]) for key in keys if key in by_key)
    finally:
        connection.close()


# Alias explícitos: permiten a la UI y a las pruebas sustituir sólo el borde
# de integración sin falsificar el resto del controlador.
generate_official_excel = excel_export_service.generate_official_excel
calculate_case_distribution = case_distribution.calculate_case_distribution
generate_case_letters = case_letter_service.generate_case_letters
