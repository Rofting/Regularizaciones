"""Análisis no destructivo de fuentes para el alta guiada de comunidades."""

from __future__ import annotations

import hashlib
import json
import os
import sqlite3
import tempfile
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Mapping

from openpyxl import load_workbook

import excel_generator
import expedient_service
import gestor_bd
from excel_profiles import (
    MODULE_REQUIRED_SHEETS,
    runtime_profile_path,
    validate_profile_payload,
)
from lector_pdf import extraer_texto


_EXCEL_SUFFIXES = frozenset({".xls", ".xlsx"})
_METER_MODULES = ("ACS", "CALEFACCION")
_SERVICE_LABELS = {
    "ACS": ("acs", "agua caliente"),
    "CALEFACCION": ("calefaccion", "calefacción"),
}
_SERVICE_DECISIONS = {
    "ACS": ("ACS",),
    "CALEFACCION": ("CALEFACCION",),
    "ACS+CALEFACCION": ("ACS", "CALEFACCION"),
    "NO_APLICA": (),
}
_WINDOWS_RESERVED_NAMES = frozenset({
    "CON", "PRN", "AUX", "NUL",
    *(f"COM{number}" for number in range(1, 10)),
    *(f"LPT{number}" for number in range(1, 10)),
})
_WINDOWS_RESERVED_CHARACTERS = frozenset('<>:"/\\|?*')


@dataclass(frozen=True)
class SourceCandidate:
    path: Path
    kind: str
    sha256: str
    headers: tuple[str, ...] = ()
    text: str = ""


@dataclass(frozen=True)
class OnboardingQuestion:
    key: str
    prompt: str
    candidates: tuple[str, ...]
    required: bool


@dataclass(frozen=True)
class ReadingBinding:
    module: str
    column: str
    source_sha256s: tuple[str, ...]


@dataclass(frozen=True)
class OnboardingConfiguration:
    """Única decisión normalizada que comparten resumen, perfil y publicación."""

    service_decision: str
    active_modules: tuple[str, ...]
    reading_column: str
    reading_bindings: tuple[ReadingBinding, ...]
    source_traces: tuple[tuple[str, str, str], ...]


@dataclass(frozen=True)
class OnboardingDraft:
    community_code: str
    community_name: str
    sources: tuple[SourceCandidate, ...]
    detected_modules: tuple[str, ...]
    questions: tuple[OnboardingQuestion, ...]


@dataclass(frozen=True)
class OnboardingResult:
    community_id: int
    period_id: int | None
    case_id: int | None
    profile_path: Path
    template_path: Path
    source_document_ids: tuple[int, ...]


def build_profile_payload(
    draft: OnboardingDraft,
    answers: Mapping[str, str | bool] | OnboardingConfiguration,
) -> dict[str, object]:
    """Build an in-memory, validated-shape profile from explicit answers."""
    configuration = (
        answers
        if isinstance(answers, OnboardingConfiguration)
        else resolve_onboarding_configuration(draft, answers)
    )
    active_modules = list(configuration.active_modules)
    _validate_community_code(draft.community_code)
    key = f"{draft.community_code}_v1"

    required_sheets = ["DATOS"] + [
        sheet
        for module in active_modules
        for sheet in MODULE_REQUIRED_SHEETS[module]
    ]
    return {
        "key": key,
        "version": "1",
        "community_code": draft.community_code,
        "template_relative_path": (
            f"plantillas/comunidades/{draft.community_code}/{key}.xlsx"
        ),
        "active_modules": active_modules,
        "required_sheets": required_sheets,
        "required_formula_cells": [],
        "concepts": [
            concept
            for module in active_modules
            for concept in _concepts_for(module)
        ],
        "onboarding_configuration": _configuration_payload(configuration),
    }


def resolve_onboarding_configuration(
    draft: OnboardingDraft,
    answers: Mapping[str, str | bool],
) -> OnboardingConfiguration:
    """Normaliza respuestas humanas en un contrato único y trazable."""
    _require_answers(draft.questions, answers)
    service_answer = answers.get("service")
    if isinstance(service_answer, str) and service_answer in _SERVICE_DECISIONS:
        service_decision = service_answer
        active_modules = _SERVICE_DECISIONS[service_answer]
    elif len(draft.detected_modules) == 1:
        # Compatibilidad con borradores construidos antes de que el análisis
        # incorporase siempre la decisión explícita de servicio.
        service_decision = draft.detected_modules[0]
        active_modules = (draft.detected_modules[0],)
    else:
        raise ValueError("Debe confirmarse el servicio o indicar que no aplica")

    unsupported = sorted(set(active_modules).difference(_METER_MODULES))
    if unsupported:
        raise ValueError("Servicio de contador no soportado: " + ", ".join(unsupported))

    reading_sources = tuple(
        source for source in draft.sources
        if source.kind.startswith("meter_reading_")
    )
    reading_columns = _reading_columns(reading_sources)
    selected_column = answers.get("reading_column")
    if isinstance(selected_column, str) and selected_column.strip():
        reading_column = selected_column.strip()
    elif len(reading_columns) == 1:
        reading_column = reading_columns[0]
    elif active_modules:
        raise ValueError("Debe responderse la columna de lectura obligatoria")
    else:
        reading_column = ""

    if active_modules and not reading_sources:
        raise ValueError("Cada servicio activo necesita una fuente de lecturas")
    fingerprints = tuple(source.sha256 for source in reading_sources)
    bindings = tuple(
        ReadingBinding(module, reading_column, fingerprints)
        for module in active_modules
    )
    traces = tuple(
        (source.kind, source.path.name, source.sha256) for source in draft.sources
    )
    return OnboardingConfiguration(
        service_decision=service_decision,
        active_modules=active_modules,
        reading_column=reading_column,
        reading_bindings=bindings,
        source_traces=traces,
    )


def _configuration_payload(configuration: OnboardingConfiguration) -> dict[str, object]:
    return {
        "service_decision": configuration.service_decision,
        "active_modules": list(configuration.active_modules),
        "reading_column": configuration.reading_column,
        "reading_bindings": [
            {
                "module": binding.module,
                "column": binding.column,
                "source_sha256s": list(binding.source_sha256s),
            }
            for binding in configuration.reading_bindings
        ],
        "sources": [
            {"kind": kind, "name": name, "sha256": sha256}
            for kind, name, sha256 in configuration.source_traces
        ],
    }


def confirm_onboarding(
    connection: sqlite3.Connection,
    *,
    draft: OnboardingDraft,
    answers: Mapping[str, str | bool],
    project_root: Path,
    archive_root: Path,
    period_name: str | None,
    start_date: date | None,
    end_date: date | None,
    actor: str,
) -> OnboardingResult:
    """Publish one onboarding or compensate every artifact created by it."""
    _validate_period(period_name, start_date, end_date)
    if connection.in_transaction:
        raise RuntimeError("No confirme un alta dentro de una transacción externa")
    if not isinstance(actor, str) or not actor.strip():
        raise ValueError("Debe indicarse quién confirma el alta")

    root = Path(project_root).resolve()
    archives = Path(archive_root).resolve()
    payload = build_profile_payload(draft, answers)
    _verify_source_fingerprints(draft.sources)
    key = str(payload["key"])
    profile_path = runtime_profile_path(key, root)
    template_path = root / str(payload["template_relative_path"])
    if profile_path.exists():
        raise FileExistsError(f"Ya existe el perfil {profile_path}")
    if template_path.exists():
        raise FileExistsError(f"Ya existe la plantilla {template_path}")

    profile_temporary: Path | None = None
    template_temporary: Path | None = None
    profile_created = False
    template_created = False
    community_created = False
    case_id: int | None = None
    period_id: int | None = None
    period_created = False
    document_ids: list[int] = []
    archived_paths: list[Path] = []
    created_directories: list[Path] = []
    existing_community = connection.execute(
        "SELECT id_comunidad FROM comunidades WHERE codigo=?",
        (draft.community_code,),
    ).fetchone()
    community_id: int | None = (
        int(existing_community[0]) if existing_community is not None else None
    )
    period_ids_before: set[int] = set()
    if community_id is not None:
        period_ids_before = {
            int(row[0]) for row in connection.execute(
                "SELECT id_periodo FROM periodos WHERE id_comunidad=?", (community_id,)
            )
        }

    try:
        created_directories.extend(_ensure_directory(profile_path.parent, root))
        created_directories.extend(_ensure_directory(template_path.parent, root))
        with tempfile.NamedTemporaryFile(
            dir=template_path.parent, prefix=f".{template_path.name}.",
            suffix=".tmp.xlsx", delete=False,
        ) as temporary:
            template_temporary = Path(temporary.name)
        layout = excel_generator.create_canonical_community_template(
            template_temporary,
            community_code=draft.community_code,
            community_name=draft.community_name,
            active_modules=tuple(payload["active_modules"]),
        )
        payload["workbook_layout"] = layout
        profile = validate_profile_payload(payload, root)
        _validate_generated_template(template_temporary, profile)

        with tempfile.NamedTemporaryFile(
            mode="w", encoding="utf-8", newline="\n",
            dir=profile_path.parent, prefix=f".{profile_path.name}.",
            suffix=".tmp", delete=False,
        ) as temporary:
            json.dump(payload, temporary, ensure_ascii=False, indent=2)
            temporary.write("\n")
            profile_temporary = Path(temporary.name)
        reloaded_payload = json.loads(profile_temporary.read_text(encoding="utf-8"))
        validate_profile_payload(reloaded_payload, root)

        _publish_exclusive(template_temporary, template_path)
        template_created = True
        _unlink_created_file(template_temporary)
        template_temporary = None
        _publish_exclusive(profile_temporary, profile_path)
        profile_created = True
        _unlink_created_file(profile_temporary)
        profile_temporary = None

        community_id = gestor_bd.obtener_o_crear_comunidad(
            connection, draft.community_code, draft.community_name
        )
        community_created = existing_community is None
        if start_date is not None and end_date is not None:
            case = expedient_service.create_case(
                connection,
                community_id,
                name=str(period_name),
                start_date=start_date,
                end_date=end_date,
            )
            case_id = case.id_case
            period_id = expedient_service.link_case_to_period(connection, case_id)
            period_created = period_id not in period_ids_before
            for source in draft.sources:
                document_ids_before = _source_document_ids(connection, case_id)
                try:
                    document, created = expedient_service.register_source_document(
                        connection,
                        case_id,
                        source_path=source.path,
                        archive_root=archives,
                        document_kind=source.kind,
                    )
                except Exception:
                    for recovered in _new_source_document_rows(
                        connection, case_id, document_ids_before
                    ):
                        document_ids.append(int(recovered[0]))
                        archived_paths.append(Path(recovered[1]))
                    raise
                if created:
                    document_ids.append(document.id_document)
                    archived_paths.append(document.archived_path)

        return OnboardingResult(
            community_id=community_id,
            period_id=period_id,
            case_id=case_id,
            profile_path=profile_path,
            template_path=template_path,
            source_document_ids=tuple(document_ids),
        )
    except Exception:
        _compensate_onboarding(
            connection,
            community_id=community_id if community_created else None,
            period_id=period_id if period_created else None,
            case_id=case_id,
            document_ids=tuple(document_ids),
        )
        for archived_path in reversed(archived_paths):
            _unlink_created_file(archived_path)
            _remove_empty_parents(archived_path.parent, archives)
        if profile_created:
            _unlink_created_file(profile_path)
        if template_created:
            _unlink_created_file(template_path)
        _remove_created_directories(created_directories)
        raise
    finally:
        if profile_temporary is not None:
            _unlink_created_file(profile_temporary)
        if template_temporary is not None:
            _unlink_created_file(template_temporary)


def _validate_period(
    period_name: str | None, start_date: date | None, end_date: date | None
) -> None:
    supplied_dates = (start_date is not None, end_date is not None)
    if any(supplied_dates) and not all(supplied_dates):
        raise ValueError("Las fechas del período deben estar completas o no indicarse")
    if all(supplied_dates) and (not isinstance(period_name, str) or not period_name.strip()):
        raise ValueError("Debe indicarse el nombre del período cuando hay fechas")


def _verify_source_fingerprints(sources: tuple[SourceCandidate, ...]) -> None:
    for source in sources:
        if _sha256(source.path) != source.sha256:
            raise ValueError(f"La fuente cambió desde el análisis: {source.path.name}")


def _validate_generated_template(path: Path, profile) -> None:
    workbook = load_workbook(path, read_only=False, data_only=False)
    try:
        if workbook.sheetnames != list(profile.required_sheets):
            raise ValueError("La plantilla canónica no coincide con las hojas del perfil")
        for module, table in profile.workbook_layout["tables"].items():
            if module not in profile.active_modules or table["sheet"] not in workbook.sheetnames:
                raise ValueError("El layout de la plantilla no coincide con sus módulos")
    finally:
        workbook.close()


def _compensate_onboarding(
    connection: sqlite3.Connection,
    *,
    community_id: int | None,
    period_id: int | None,
    case_id: int | None,
    document_ids: tuple[int, ...],
) -> None:
    connection.rollback()
    connection.execute("BEGIN IMMEDIATE")
    try:
        if document_ids:
            placeholders = ",".join("?" for _ in document_ids)
            connection.execute(
                f"DELETE FROM source_documents WHERE id_document IN ({placeholders})",
                document_ids,
            )
        if case_id is not None:
            connection.execute("DELETE FROM regularization_cases WHERE id_case=?", (case_id,))
        if period_id is not None:
            connection.execute("DELETE FROM periodos WHERE id_periodo=?", (period_id,))
        if community_id is not None:
            connection.execute("DELETE FROM comunidades WHERE id_comunidad=?", (community_id,))
        connection.commit()
    except Exception:
        connection.rollback()
        raise


def _unlink_created_file(path: Path) -> None:
    try:
        path.unlink()
    except FileNotFoundError:
        pass


def _publish_exclusive(source: Path, destination: Path) -> None:
    """Publish a same-filesystem temporary without ever replacing a destination."""
    os.link(source, destination)


def _source_document_ids(
    connection: sqlite3.Connection, case_id: int
) -> frozenset[int]:
    return frozenset(
        int(row[0]) for row in connection.execute(
            "SELECT id_document FROM source_documents WHERE id_case=?", (case_id,)
        )
    )


def _new_source_document_rows(
    connection: sqlite3.Connection, case_id: int, ids_before: frozenset[int]
):
    return tuple(
        row for row in connection.execute(
            """SELECT id_document, archived_path FROM source_documents
               WHERE id_case=?""",
            (case_id,),
        )
        if int(row[0]) not in ids_before
    )


def _remove_empty_parents(directory: Path, boundary: Path) -> None:
    boundary = boundary.resolve()
    current = directory.resolve()
    while current == boundary or boundary in current.parents:
        try:
            current.rmdir()
        except OSError:
            break
        if current == boundary:
            break
        current = current.parent


def _ensure_directory(directory: Path, boundary: Path) -> tuple[Path, ...]:
    missing: list[Path] = []
    current = directory
    boundary = boundary.resolve()
    while not current.exists() and current != boundary:
        missing.append(current)
        current = current.parent
    directory.mkdir(parents=True, exist_ok=True)
    return tuple(missing)


def _remove_created_directories(directories: list[Path]) -> None:
    for directory in directories:
        try:
            directory.rmdir()
        except FileNotFoundError:
            continue
        except OSError:
            continue


def _require_answers(
    questions: tuple[OnboardingQuestion, ...], answers: Mapping[str, str | bool]
) -> None:
    for question in questions:
        if not question.required:
            continue
        answer = answers.get(question.key)
        if (
            not isinstance(answer, str)
            or not answer.strip()
            or (question.candidates and answer not in question.candidates)
        ):
            if question.key == "reading_column":
                raise ValueError("Debe responderse la columna de lectura obligatoria")
            if question.key == "service":
                raise ValueError("Debe responderse el servicio obligatorio")
            raise ValueError(f"Debe responderse la pregunta obligatoria: {question.prompt}")


def _validate_community_code(community_code: str) -> None:
    if (
        not isinstance(community_code, str)
        or not community_code
        or community_code[-1] in {".", " "}
        or any(character in _WINDOWS_RESERVED_CHARACTERS for character in community_code)
        or any(ord(character) <= 0x1F for character in community_code)
        or community_code.split(".", 1)[0].upper() in _WINDOWS_RESERVED_NAMES
    ):
        raise ValueError("El código de comunidad no permite crear una clave segura")


def _concepts_for(module: str) -> tuple[dict[str, object], ...]:
    prefix = "acs" if module == "ACS" else "heating"
    return (
        {
            "key": f"{prefix}_fixed",
            "allocation_method": "equal",
            "actual_source": f"period_parameters.{prefix}_fixed_actual",
            "billed_source": f"period_parameters.{prefix}_fixed_billed",
            "required": True,
        },
        {
            "key": f"{prefix}_variable",
            "allocation_method": "consumption",
            "actual_source": f"period_parameters.{prefix}_variable_actual",
            "billed_source": f"period_parameters.{prefix}_variable_billed",
            "required": True,
        },
    )


def analyse_sources(
    *,
    community_code: str,
    community_name: str,
    owner_list_path: Path,
    reading_paths: tuple[Path, ...],
    invoice_paths: tuple[Path, ...],
    project_root: Path,
) -> OnboardingDraft:
    """Return a draft from explicit source files without persisting any data."""
    del project_root  # Kept in the public API for later confirmation steps.
    if not reading_paths:
        raise ValueError("Se requiere al menos una fuente de lecturas")

    owner_path = _validated_path(owner_list_path, {".csv"}, "listado de propietarios")
    sources = [_source(owner_path, "owner_list")]
    for reading_path in reading_paths:
        path = _validated_path(reading_path, _EXCEL_SUFFIXES | {".pdf"}, "lectura")
        sources.append(_source(path, "meter_reading_pdf" if path.suffix.lower() == ".pdf" else "meter_reading_excel"))
    for invoice_path in invoice_paths:
        path = _validated_path(invoice_path, {".pdf"}, "factura")
        sources.append(_source(path, "invoice_pdf"))

    readings = tuple(source for source in sources if source.kind.startswith("meter_reading_"))
    detected_modules = _detected_modules(tuple(sources))
    return OnboardingDraft(
        community_code=community_code,
        community_name=community_name,
        sources=tuple(sources),
        detected_modules=detected_modules,
        questions=_questions(readings, detected_modules),
    )


def _validated_path(path: Path, allowed_suffixes: frozenset[str] | set[str], label: str) -> Path:
    resolved = Path(path)
    if not resolved.is_file():
        raise ValueError(f"No existe el archivo de {label}: {resolved}")
    if resolved.suffix.lower() not in allowed_suffixes:
        allowed = ", ".join(sorted(allowed_suffixes))
        raise ValueError(f"Extensión no permitida para {label}: {resolved.suffix} (se permite {allowed})")
    return resolved


def _source(path: Path, kind: str) -> SourceCandidate:
    headers = _excel_headers(path) if kind == "meter_reading_excel" else ()
    text = extraer_texto(str(path)) if path.suffix.lower() == ".pdf" else ""
    return SourceCandidate(path=path, kind=kind, sha256=_sha256(path), headers=headers, text=text)


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as source:
        for block in iter(lambda: source.read(65536), b""):
            digest.update(block)
    return digest.hexdigest()


def _excel_headers(path: Path) -> tuple[str, ...]:
    if path.suffix.lower() == ".xls":
        import xlrd

        sheet = xlrd.open_workbook(path).sheet_by_index(0)
        return tuple(str(sheet.cell_value(0, column)).strip() for column in range(sheet.ncols))
    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        sheet = workbook.active
        first_row = next(sheet.iter_rows(min_row=1, max_row=1, values_only=True), ())
        return tuple(str(value or "").strip() for value in first_row)
    finally:
        workbook.close()


def _detected_modules(readings: tuple[SourceCandidate, ...]) -> tuple[str, ...]:
    material = " ".join(
        " ".join(source.headers) + " " + source.text for source in readings
    ).lower()
    modules = []
    acs_labels = _SERVICE_LABELS["ACS"]
    if any(label in material for label in acs_labels):
        modules.append("ACS")
        for label in acs_labels:
            material = material.replace(label, " ")
    if any(label in material for label in _SERVICE_LABELS["CALEFACCION"]):
        modules.append("CALEFACCION")
    return tuple(modules)


def _reading_columns(readings: tuple[SourceCandidate, ...]) -> tuple[str, ...]:
    return _unique_labels(
        header
        for source in readings
        for header in source.headers
        if "lectura" in header.lower() or "contador" in header.lower()
    )


def _service_candidates(detected_modules: tuple[str, ...]) -> tuple[str, ...]:
    candidates = list(detected_modules)
    candidates.extend(
        module for module in _METER_MODULES if module not in candidates
    )
    candidates.extend(("ACS+CALEFACCION", "NO_APLICA"))
    return tuple(candidates)


def _questions(
    readings: tuple[SourceCandidate, ...],
    detected_modules: tuple[str, ...],
) -> tuple[OnboardingQuestion, ...]:
    questions: list[OnboardingQuestion] = []
    reading_columns = _reading_columns(readings)
    if len(reading_columns) != 1:
        questions.append(OnboardingQuestion(
            key="reading_column",
            prompt=(
                "Selecciona la columna de lectura que se debe usar."
                if reading_columns
                else "Indica la columna o campo de lectura que se debe usar."
            ),
            candidates=reading_columns,
            required=True,
        ))
    questions.append(OnboardingQuestion(
        key="service",
        prompt="Confirma el servicio asociado a las lecturas o indica que no aplica.",
        candidates=_service_candidates(detected_modules),
        required=True,
    ))
    return tuple(questions)


def _unique_labels(labels) -> tuple[str, ...]:
    unique: list[str] = []
    seen: set[str] = set()
    for label in labels:
        normalized = " ".join(label.split()).casefold()
        if normalized not in seen:
            seen.add(normalized)
            unique.append(label)
    return tuple(unique)
