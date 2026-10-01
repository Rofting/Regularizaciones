"""Informe único de requisitos para cada etapa de una regularización."""

from __future__ import annotations

import sqlite3
from dataclasses import dataclass
from pathlib import Path

import document_review
import fixed_costs
import case_coherence
from excel_export_service import ExportBlockedError, calculate_case_input_hash
from excel_profiles import (
    ExcelProfile,
    calculate_profile_sha256,
    configured_profile_paths,
    load_profile,
)


STAGE_ORDER = {
    "sources": 0,
    "excel": 1,
    "distribution": 2,
    "letters": 3,
    "email": 4,
}


@dataclass(frozen=True)
class ReadinessBlocker:
    code: str
    stage: str
    message: str
    action: str
    count: int = 1


@dataclass(frozen=True)
class CaseReadinessReport:
    case_id: int
    blockers: tuple[ReadinessBlocker, ...]

    def for_stage(self, stage: str) -> tuple[ReadinessBlocker, ...]:
        try:
            maximum = STAGE_ORDER[stage]
        except KeyError:
            raise ValueError(f"Etapa de preparación no soportada: {stage}") from None
        return tuple(
            blocker for blocker in self.blockers
            if STAGE_ORDER[blocker.stage] <= maximum
        )

    def is_ready(self, stage: str) -> bool:
        return not self.for_stage(stage)

    @property
    def sources_ready(self) -> bool:
        return self.is_ready("sources")

    @property
    def excel_ready(self) -> bool:
        return self.is_ready("excel")

    @property
    def distribution_ready(self) -> bool:
        return self.is_ready("distribution")

    @property
    def letters_ready(self) -> bool:
        return self.is_ready("letters")


class CaseNotReadyError(ValueError):
    """La etapa solicitada tiene requisitos concretos pendientes."""

    def __init__(self, stage: str, blockers: tuple[ReadinessBlocker, ...]):
        self.stage = stage
        self.blockers = blockers
        summary = "; ".join(blocker.message for blocker in blockers)
        super().__init__(summary or f"El expediente no está listo para {stage}")


def _case(connection: sqlite3.Connection, case_id: int) -> sqlite3.Row:
    row = connection.execute(
        """SELECT c.id_case,c.id_comunidad,c.id_periodo,c.fecha_inicio,c.fecha_fin,
                  c.estado,co.codigo
           FROM regularization_cases c
           JOIN comunidades co ON co.id_comunidad=c.id_comunidad
           WHERE c.id_case=?""",
        (case_id,),
    ).fetchone()
    if row is None:
        raise LookupError("El expediente no existe")
    return row


def _active_profile(
    connection: sqlite3.Connection,
    case: sqlite3.Row,
    project_root: Path,
) -> tuple[ExcelProfile | None, ReadinessBlocker | None]:
    rows = connection.execute(
        """SELECT profile_key,profile_version,template_relative_path,
                  template_sha256,profile_sha256
           FROM excel_template_profiles
           WHERE id_comunidad=? AND status='active'
           ORDER BY id_template_profile""",
        (case["id_comunidad"],),
    ).fetchall()
    if not rows:
        configured: list[ExcelProfile] = []
        for profile_path in configured_profile_paths(project_root):
            try:
                candidate = load_profile(profile_path.stem, project_root)
            except (LookupError, ValueError):
                continue
            if candidate.community_code == str(case["codigo"]):
                configured.append(candidate)
        if not configured:
            return None, ReadinessBlocker(
                "MISSING_PROFILE", "excel",
                "Falta preparar el perfil Excel de la comunidad.",
                "prepare_excel",
            )
        if len(configured) != 1:
            return None, ReadinessBlocker(
                "MULTIPLE_ACTIVE_PROFILES", "excel",
                "Hay varios perfiles Excel configurados para la comunidad.",
                "manage_profile", len(configured),
            )
        profile = configured[0]
        template = (project_root / profile.template_relative_path).resolve()
        if not template.is_file():
            return None, ReadinessBlocker(
                "MISSING_TEMPLATE", "excel",
                "No se encuentra la plantilla Excel configurada.",
                "prepare_excel",
            )
        return profile, None
    if len(rows) != 1:
        return None, ReadinessBlocker(
            "MULTIPLE_ACTIVE_PROFILES", "excel",
            "La comunidad debe tener exactamente un perfil Excel activo.",
            "manage_profile", len(rows),
        )
    registered = rows[0]
    try:
        profile = load_profile(str(registered["profile_key"]), project_root)
    except (LookupError, ValueError) as error:
        return None, ReadinessBlocker(
            "INVALID_PROFILE", "excel",
            f"No se puede cargar el perfil Excel activo: {error}",
            "prepare_excel",
        )
    if (
        profile.version != registered["profile_version"]
        or profile.community_code != str(case["codigo"])
        or calculate_profile_sha256(profile, project_root)
        != registered["profile_sha256"]
    ):
        return None, ReadinessBlocker(
            "STALE_PROFILE", "excel",
            "El perfil Excel activo cambió y debe revalidarse.",
            "revalidate_profile",
        )
    template = Path(project_root) / str(registered["template_relative_path"])
    if not template.is_file():
        return None, ReadinessBlocker(
            "MISSING_TEMPLATE", "excel",
            "No se encuentra la plantilla Excel registrada.",
            "prepare_excel",
        )
    return profile, None


def _eligible_owners(
    connection: sqlite3.Connection, community_id: int
) -> tuple[sqlite3.Row, ...]:
    return tuple(connection.execute(
        """SELECT id_propietario,codigo_vivienda,coeficiente
           FROM propietarios
           WHERE id_comunidad=? AND activo=1 AND tipo_unidad='vivienda'
           ORDER BY id_propietario""",
        (community_id,),
    ).fetchall())


def _required_parameter_missing(
    connection: sqlite3.Connection,
    case: sqlite3.Row,
    profile: ExcelProfile,
) -> tuple[str, ...]:
    required: set[str] = set()
    for concept in profile.concepts:
        if not concept.required:
            continue
        for source in (concept.actual_source, concept.billed_source):
            if source and source.startswith("period_parameters."):
                required.add(source.split(".", 1)[1])
    if not required or case["id_periodo"] is None:
        return tuple(sorted(required)) if case["id_periodo"] is None else ()
    present = {
        str(row["parameter_key"])
        for row in connection.execute(
            """SELECT parameter_key FROM period_parameters
               WHERE id_comunidad=? AND id_periodo=? AND numeric_value IS NOT NULL""",
            (case["id_comunidad"], case["id_periodo"]),
        )
    }
    return tuple(sorted(required.difference(present)))


def _reading_service(concept_key: str) -> str | None:
    if concept_key.startswith("acs_"):
        return "ACS"
    if concept_key.startswith("heating_"):
        return "CALEFACCION"
    return None


def _owners_missing_boundary_readings(
    connection: sqlite3.Connection,
    case: sqlite3.Row,
    owners: tuple[sqlite3.Row, ...],
    service: str,
) -> int:
    initial = connection.execute(
        """SELECT MAX(r.fecha_lectura)
           FROM period_readings r
           JOIN propietarios p ON p.id_propietario=r.id_propietario
           WHERE p.id_comunidad=? AND p.tipo_unidad='vivienda'
             AND r.tipo=? AND r.fecha_lectura<=?""",
        (case["id_comunidad"], service, case["fecha_inicio"]),
    ).fetchone()[0]
    final = connection.execute(
        """SELECT MAX(r.fecha_lectura)
           FROM period_readings r
           JOIN propietarios p ON p.id_propietario=r.id_propietario
           WHERE p.id_comunidad=? AND p.tipo_unidad='vivienda'
             AND r.tipo=? AND r.fecha_lectura<=?""",
        (case["id_comunidad"], service, case["fecha_fin"]),
    ).fetchone()[0]
    if initial is None or final is None or initial >= final:
        return len(owners)
    missing = 0
    for owner in owners:
        count = connection.execute(
            """SELECT COUNT(DISTINCT fecha_lectura) FROM period_readings
               WHERE id_propietario=? AND id_periodo=? AND tipo=?
                 AND fecha_lectura IN (?,?)
                 AND estado IN ('real','estimado')""",
            (owner["id_propietario"], case["id_periodo"], service, initial, final),
        ).fetchone()[0]
        if count != 2:
            missing += 1
    return missing


def _validated_export_blocker(
    connection: sqlite3.Connection,
    case: sqlite3.Row,
    project_root: Path,
    profile: ExcelProfile,
) -> ReadinessBlocker | None:
    export = connection.execute(
        """SELECT input_sha256,status,id_periodo FROM excel_export_runs
           WHERE id_case=? ORDER BY created_at DESC,id_export_run DESC LIMIT 1""",
        (case["id_case"],),
    ).fetchone()
    if (
        export is None
        or export["status"] != "validated"
        or export["id_periodo"] != case["id_periodo"]
    ):
        return ReadinessBlocker(
            "MISSING_VALIDATED_EXCEL", "distribution",
            "Falta generar y validar el Excel oficial vigente.",
            "generate_excel",
        )
    try:
        current_hash = calculate_case_input_hash(
            connection,
            id_case=int(case["id_case"]),
            project_root=project_root,
            profile=profile,
        )
    except (ExportBlockedError, ValueError) as error:
        return ReadinessBlocker(
            "INVALID_EXCEL_INPUTS", "distribution",
            f"No se pueden comprobar las entradas del Excel: {error}",
            "generate_excel",
        )
    if export["input_sha256"] != current_hash:
        return ReadinessBlocker(
            "STALE_EXCEL", "distribution",
            "Las entradas cambiaron desde el último Excel validado.",
            "generate_excel",
        )
    return None


def evaluate_case_readiness(
    connection: sqlite3.Connection,
    case_id: int,
    project_root: str | Path,
) -> CaseReadinessReport:
    """Evalúa requisitos sin modificar la base ni producir archivos."""
    case = _case(connection, case_id)
    root = Path(project_root).resolve()
    blockers: list[ReadinessBlocker] = []

    document_count = int(connection.execute(
        "SELECT COUNT(*) FROM source_documents WHERE id_case=?", (case_id,)
    ).fetchone()[0])
    if document_count == 0:
        blockers.append(ReadinessBlocker(
            "NO_SOURCES", "sources", "El expediente todavía no tiene fuentes.",
            "add_sources",
        ))
    open_issues = int(connection.execute(
        "SELECT COUNT(*) FROM review_issues WHERE id_case=? AND status='open'",
        (case_id,),
    ).fetchone()[0])
    if open_issues:
        blockers.append(ReadinessBlocker(
            "OPEN_ISSUES", "sources",
            f"Hay {open_issues} decisión(es) pendiente(s) de revisión.",
            "resolve_issues", open_issues,
        ))
    pending = int(connection.execute(
        """SELECT COUNT(*) FROM source_documents
           WHERE id_case=? AND status NOT IN ('validated','not_applicable')""",
        (case_id,),
    ).fetchone()[0])
    if pending:
        blockers.append(ReadinessBlocker(
            "PENDING_SOURCES", "sources",
            f"Hay {pending} fuente(s) pendientes de confirmar.",
            "confirm_sources", pending,
        ))
    if document_review.case_has_unapplied_sources(connection, case_id):
        blockers.append(ReadinessBlocker(
            "UNAPPLIED_SOURCES", "sources",
            "Hay fuentes reconocidas pendientes de aplicar al expediente.",
            "confirm_sources",
        ))

    if case["id_periodo"] is None:
        blockers.append(ReadinessBlocker(
            "MISSING_PERIOD", "excel", "El expediente no tiene un período ligado.",
            "manage_periods",
        ))
    owners = _eligible_owners(connection, int(case["id_comunidad"]))
    if not owners:
        blockers.append(ReadinessBlocker(
            "MISSING_OWNERS", "excel",
            "Faltan propietarios activos para preparar el reparto.",
            "import_owners",
        ))

    profile, profile_blocker = _active_profile(connection, case, root)
    if profile_blocker is not None:
        blockers.append(profile_blocker)
    if profile is not None:
        if case['id_periodo'] is not None:
            try:
                coherence = case_coherence.evaluate(connection, case_id, profile)
            except ValueError as error:
                blockers.append(ReadinessBlocker('INVALID_COHERENCE_SETTINGS', 'excel', str(error), 'review_coherence'))
            else:
                for finding in coherence.pending:
                    blockers.append(ReadinessBlocker('COHERENCE_' + finding.key.split(':')[0].upper(),
                                                     'excel', finding.message, 'review_coherence'))
        template = root / profile.template_relative_path
        if "OTROS_GASTOS" in profile.active_modules and template.is_file() and case["id_periodo"] is not None:
            try:
                findings = fixed_costs.unconfirmed_template_costs(
                    template, fixed_costs.load_fixed_costs(connection, case["id_comunidad"], case["id_periodo"]),
                )
            except fixed_costs.TemplateCostReviewError as error:
                blockers.append(ReadinessBlocker("UNREADABLE_TEMPLATE", "excel", str(error), "prepare_excel"))
            else:
                if findings:
                    blockers.append(ReadinessBlocker(
                        "UNCONFIRMED_FIXED_COSTS", "excel",
                        fixed_costs.inherited_cost_message(findings), "review_fixed_costs", len(findings),
                    ))
        missing_parameters = _required_parameter_missing(connection, case, profile)
        if missing_parameters:
            blockers.append(ReadinessBlocker(
                "MISSING_CONCEPT_VALUES", "excel",
                "Faltan importes requeridos para: " + ", ".join(missing_parameters),
                "review_sources", len(missing_parameters),
            ))
        if owners:
            coefficient_concepts = [
                concept for concept in profile.concepts
                if concept.required and concept.allocation_method == "coefficient"
            ]
            if coefficient_concepts:
                missing_coefficients = sum(
                    1 for owner in owners if float(owner["coeficiente"] or 0) <= 0
                )
                if missing_coefficients:
                    blockers.append(ReadinessBlocker(
                        "MISSING_COEFFICIENTS", "distribution",
                        f"Falta un coeficiente válido en {missing_coefficients} propiedad(es).",
                        "review_owners", missing_coefficients,
                    ))
            services = {
                service for concept in profile.concepts
                if concept.required and concept.allocation_method == "consumption"
                if (service := _reading_service(concept.key)) is not None
            }
            missing_readings = max(
                (_owners_missing_boundary_readings(
                    connection, case, owners, service
                ) for service in services),
                default=0,
            )
            if missing_readings:
                noun = "propiedad" if missing_readings == 1 else "propiedades"
                blockers.append(ReadinessBlocker(
                    "MISSING_READINGS", "distribution",
                    f"Faltan lecturas iniciales o finales en {missing_readings} {noun}.",
                    "review_readings", missing_readings,
                ))
        export_blocker = _validated_export_blocker(connection, case, root, profile)
        if export_blocker is not None:
            blockers.append(export_blocker)

    distribution = connection.execute(
        """SELECT input_sha256,status FROM distribution_runs
           WHERE id_case=? ORDER BY created_at DESC,id_distribution_run DESC LIMIT 1""",
        (case_id,),
    ).fetchone()
    if distribution is None or distribution["status"] != "completed":
        blockers.append(ReadinessBlocker(
            "MISSING_DISTRIBUTION", "letters",
            "Falta un reparto final completado y conciliado.",
            "calculate_distribution",
        ))
    else:
        if profile is not None:
            try:
                current_distribution_hash = calculate_case_input_hash(
                    connection,
                    id_case=int(case["id_case"]),
                    project_root=root,
                    profile=profile,
                )
            except (ExportBlockedError, ValueError) as error:
                blockers.append(ReadinessBlocker(
                    "INVALID_DISTRIBUTION_INPUTS", "letters",
                    f"No se pueden comprobar las entradas del reparto: {error}",
                    "calculate_distribution",
                ))
            else:
                if distribution["input_sha256"] != current_distribution_hash:
                    blockers.append(ReadinessBlocker(
                        "STALE_DISTRIBUTION", "letters",
                        "Las entradas cambiaron desde el último reparto final.",
                        "calculate_distribution",
                    ))
        if connection.execute(
            """SELECT 1 FROM reconciliations
               WHERE id_periodo=? AND status<>'cuadrado' LIMIT 1""",
            (case["id_periodo"],),
        ).fetchone() is not None:
            blockers.append(ReadinessBlocker(
                "UNRECONCILED_DISTRIBUTION", "letters",
                "El reparto contiene conceptos que no cuadran al céntimo.",
                "calculate_distribution",
            ))

    return CaseReadinessReport(int(case["id_case"]), tuple(blockers))


def require_stage(
    connection: sqlite3.Connection,
    case_id: int,
    stage: str,
    project_root: str | Path,
) -> CaseReadinessReport:
    report = evaluate_case_readiness(connection, case_id, project_root)
    blockers = report.for_stage(stage)
    if blockers:
        raise CaseNotReadyError(stage, blockers)
    return report
