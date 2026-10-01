"""Alta masiva: una comunidad por subcarpeta, con aprobación previa.

Une la detección desde carpeta (``community_folder``) con el alta guiada
(``community_onboarding``). Para cada subcarpeta se propone código, nombre,
período, listado de propietarios, lecturas y facturas; sólo se dan por listas
las comunidades cuyas respuestas son inequívocas. Las demás quedan con su
diagnóstico para abrirlas en el alta guiada individual.

Crear el lote es repetible: una comunidad ya registrada no se vuelve a crear,
y el fallo de una comunidad se compensa (lo hace ``confirm_onboarding``) sin
detener las siguientes.
"""

from __future__ import annotations

import re
import sqlite3
from dataclasses import dataclass, field
from datetime import date
from pathlib import Path
from typing import Callable, Iterable

READY = "ready"            # se puede crear tal cual
REVIEW = "review"          # faltan decisiones: alta guiada individual
EXISTING = "existing"      # ya registrada: repetir el lote no la duplica
FAILED = "failed"          # no se pudo analizar

_FOLDER_CODE_NAME = re.compile(r"^\s*(\d{3,6})\s*(?:[-_.–—]\s*|\s+)(.+?)\s*$")


@dataclass
class BatchEntry:
    folder: Path
    status: str
    code: str | None = None
    name: str | None = None
    start: date | None = None
    end: date | None = None
    owners: Path | None = None
    readings: tuple[Path, ...] = ()
    invoices: tuple[Path, ...] = ()
    invoices_for_inbox: tuple[Path, ...] = ()
    reasons: list[str] = field(default_factory=list)
    draft: object | None = None
    answers: dict[str, str] = field(default_factory=dict)

    @property
    def can_create(self) -> bool:
        return self.status == READY

    @property
    def period_name(self) -> str | None:
        if self.start is None or self.end is None:
            return None
        from period_selection import default_case_name
        return default_case_name(self.start, self.end)

    @property
    def summary(self) -> str:
        parts = [f"{len(self.readings)} lectura(s)", f"{len(self.invoices)} factura(s)"]
        if self.owners:
            parts.insert(0, "propietarios")
        if self.start and self.end:
            parts.append(f"{self.start:%d/%m/%Y} → {self.end:%d/%m/%Y}")
        if self.answers:
            parts.append(sample_text(self.answers) or self.answers.get("service", ""))
        return " · ".join(parts)


@dataclass(frozen=True)
class BatchOutcome:
    folder: Path
    code: str | None
    status: str  # created | skipped | failed
    message: str
    case_id: int | None = None
    owner_count: int = 0
    reading_count: int = 0


def _subfolders(root: Path) -> list[Path]:
    return sorted(
        (path for path in root.iterdir() if path.is_dir() and not path.name.startswith((".", "~"))),
        key=lambda path: path.name.casefold(),
    )


def code_and_name_from_folder(name: str) -> tuple[str | None, str | None]:
    """«658 - CP Las Flores» → («658», «CP LAS FLORES»); «658» → («658», None)."""
    text = name.strip()
    if re.fullmatch(r"\d{3,6}", text):
        return text, None
    match = _FOLDER_CODE_NAME.match(text)
    if match:
        return match.group(1), " ".join(match.group(2).split()).upper()
    return None, None


def _registered_codes(connection: sqlite3.Connection) -> set[str]:
    return {str(row[0]).strip() for row in connection.execute("SELECT codigo FROM comunidades")}


def _profile_exists(code: str, project_root: Path) -> bool:
    from excel_profiles import runtime_profile_path
    return runtime_profile_path(f"{code}_v1", project_root).exists() or (
        project_root / "plantillas" / "comunidades" / code / f"{code}_v1.xlsx"
    ).exists()


def _parse_iso(value: str | None) -> date | None:
    try:
        return date.fromisoformat(str(value)[:10]) if value else None
    except ValueError:
        return None


def _ambiguous_invoices(draft) -> set[str]:
    return {
        evidence.source_sha256 for evidence in draft.invoice_evidence
        if any(len(getattr(evidence, attribute)) != 1
               for attribute in ("providers", "periods", "amounts", "concepts"))
    }


_SERVICE_LABELS = {"ACS": "ACS", "CALEFACCION": "Calefacción"}


def _onboarding_readable(path: Path) -> bool:
    """El alta guiada entiende una hoja con vivienda, contador, fecha y lectura."""
    if path.suffix.lower() not in {".xlsx", ".xls"}:
        return False
    from meter_reading_sources import excel_observations
    try:
        observations = excel_observations(path)
    except Exception:
        return False
    return bool(observations) and all(item.meter and item.date and item.property_code
                                      for item in observations)


def _reading_table(path: Path, text_reader: Callable[[Path], str] | None):
    from reading_tables import parse_reading_document, parse_reading_rows, tables_from_pdf
    if path.suffix.lower() == ".pdf":
        text = (text_reader or _default_text)(path)
        return parse_reading_document(text=text, table_rows=tables_from_pdf(path))
    from source_analysis import _tabular_sheets
    for rows, sheet in _tabular_sheets(path):
        title = " ".join(str(cell) for row in rows[:6] for cell in row if cell)
        table = parse_reading_rows(rows, f"{sheet or ''} {title}")
        if table is not None:
            return table
    return None


def _default_text(path: Path) -> str:
    from document_text_service import get_document_text
    return get_document_text(None, path, timeout_seconds=60).text


def normalise_reading_source(
    path: Path,
    destination_dir: Path,
    *,
    start: date,
    end: date,
    text_reader: Callable[[Path], str] | None = None,
) -> Path:
    """Devuelve una hoja que el alta guiada entiende: el original o una conversión.

    Los informes de las empresas traen una fila por vivienda con la lectura
    anterior y la actual; el alta trabaja con una fila por lectura (vivienda,
    contador, fecha, valor). La conversión no cambia ningún valor. Si el
    informe no trae número de contador se usa el código de la vivienda.
    """
    path = Path(path)
    if _onboarding_readable(path):
        return path
    table = _reading_table(path, text_reader)
    if table is None:
        raise ValueError(f"{path.name}: no se reconoce la tabla de lecturas")
    services: list[str] = []
    readings: dict[tuple[str, str], dict[str, tuple[str, object]]] = {}
    for row in table.rows:
        service = row.get("tipo") or table.service
        if service not in _SERVICE_LABELS:
            raise ValueError(f"{path.name}: no se sabe si las lecturas son de ACS o de calefacción")
        if service not in services:
            services.append(service)
        meter = row.get("contador") or row["vivienda"]
        for when, value in ((row.get("fecha_ant") or table.start or start.isoformat(), row.get("val_ant")),
                            (row.get("fecha_act") or table.end or end.isoformat(), row.get("val_act"))):
            if value is not None:
                readings.setdefault((row["vivienda"], str(when)[:10]), {})[service] = (meter, value)
    from openpyxl import Workbook
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Lecturas"
    header = ["Vivienda", "Fecha"]
    for service in services:
        header += [f"Contador {_SERVICE_LABELS[service]}", f"Lectura {_SERVICE_LABELS[service]}"]
    sheet.append(header)
    for (dwelling, when), values in readings.items():
        line: list[object] = [dwelling, when]
        for service in services:
            meter, value = values.get(service, (None, None))
            line += [meter, value]
        sheet.append(line)
    destination = Path(destination_dir) / f"{path.stem}_lecturas.xlsx"
    destination.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(destination)
    workbook.close()
    return destination


def automatic_answers(draft, end: date | None = None) -> tuple[dict[str, str], list[str]]:
    """Respuestas que se pueden dar sin preguntar y lo que queda por decidir.

    El alta pide confirmar una lectura de ejemplo por servicio; en el lote se
    toma la de la primera vivienda en la fecha final y se muestra en el
    resumen para que el usuario la apruebe.
    """
    modules = tuple(draft.detected_modules)
    if set(modules) == {"ACS", "CALEFACCION"}:
        service, active = "ACS+CALEFACCION", ("ACS", "CALEFACCION")
    elif len(modules) == 1:
        service, active = modules[0], modules
    else:
        return {}, ["Servicio de las lecturas por confirmar (no se ha detectado ninguno)."]
    answers = {"service": service}
    pending = []
    for module in active:
        observations = [
            item for evidence in draft.reading_evidence
            if not evidence.services or module in evidence.services
            for item in evidence.observations
            if item.meter and item.date and item.value
        ]
        columns = {item.column for item in observations}
        if len(columns) != 1:
            pending.append(f"Columna de lectura de {module} por confirmar.")
            continue
        final = [item for item in observations if end is None or item.date == end.isoformat()]
        sample = (final or observations)[0]
        answers.update({
            f"reading_column:{module}": sample.column,
            f"reading_meter:{module}": sample.meter,
            f"reading_date:{module}": sample.date,
            f"reading_value:{module}": sample.value,
        })
    for question in draft.questions:
        if question.key.startswith(("invoice_",)) and question.required:
            pending.append(question.prompt)
    return answers, pending


def _display_date(value: str) -> str:
    parsed = _parse_iso(value)
    return f"{parsed:%d/%m/%Y}" if parsed else value


def sample_text(answers: dict[str, str]) -> str:
    """«ACS: contador C-1, 31/12/2026, lectura 130» para el resumen."""
    parts = []
    for module in answers.get("service", "").split("+"):
        if f"reading_value:{module}" in answers:
            parts.append(
                f"{_SERVICE_LABELS.get(module, module)}: contador {answers[f'reading_meter:{module}']}, "
                f"{_display_date(answers[f'reading_date:{module}'])}, lectura {answers[f'reading_value:{module}']}"
            )
    return " · ".join(parts)


def scan_batch(
    root: str | Path,
    *,
    connection: sqlite3.Connection,
    project_root: Path,
    work_directory: Path,
    text_reader: Callable[[Path], str] | None = None,
    progress: Callable[[str], None] | None = None,
) -> tuple[BatchEntry, ...]:
    """Analiza cada subcarpeta sin escribir en la base ni en el proyecto.

    ``work_directory`` recibe sólo los listados de propietarios convertidos al
    formato estándar, como hace el alta guiada.
    """
    from community_folder import propose_from_folder
    from community_onboarding import analyse_sources
    from owner_lists import normalise_owner_list

    root = Path(root)
    if not root.is_dir():
        raise ValueError(f"No existe la carpeta {root}")
    registered = _registered_codes(connection)
    entries: list[BatchEntry] = []
    for folder in _subfolders(root):
        if progress:
            progress(folder.name)
        entry = BatchEntry(folder, FAILED)
        entries.append(entry)
        try:
            proposal = propose_from_folder(folder, text_reader=text_reader)
        except Exception as error:  # el diagnóstico se conserva y el lote sigue
            entry.reasons.append(f"No se pudo analizar la carpeta: {error}")
            continue
        folder_code, folder_name = code_and_name_from_folder(folder.name)
        entry.code = folder_code or proposal.code
        entry.name = proposal.name or folder_name
        entry.start, entry.end = _parse_iso(proposal.start), _parse_iso(proposal.end)
        entry.owners = proposal.owners
        entry.readings = tuple(proposal.readings)
        entry.invoices = tuple(proposal.invoices)
        entry.status = REVIEW
        if entry.code and (entry.code in registered or _profile_exists(entry.code, project_root)):
            entry.status = EXISTING
            entry.reasons.append("Ya está registrada; no se volverá a crear.")
            continue
        missing = [label for label, value in (
            ("código", entry.code), ("nombre", entry.name), ("listado de propietarios", entry.owners),
            ("lecturas", entry.readings), ("período de las lecturas", entry.start and entry.end),
        ) if not value]
        if not (entry.owners or entry.readings or entry.invoices):
            entry.reasons.append("La carpeta no contiene documentos de una comunidad.")
            continue
        if missing:
            entry.reasons.append("Falta " + ", ".join(missing) + ".")
            continue
        # Cada comunidad convierte en su carpeta: dos listados llamados igual
        # en subcarpetas distintas no deben pisarse.
        converted = Path(work_directory) / f"lote_{entry.code}"
        try:
            owners = normalise_owner_list(entry.owners, converted)
            entry.readings = tuple(
                normalise_reading_source(path, converted, start=entry.start, end=entry.end,
                                         text_reader=text_reader)
                for path in entry.readings
            )
            arguments = dict(
                community_code=entry.code, community_name=entry.name, owner_list_path=owners,
                reading_paths=entry.readings, project_root=project_root,
            )
            draft = analyse_sources(invoice_paths=entry.invoices, **arguments)
            ambiguous = _ambiguous_invoices(draft)
            if ambiguous:
                # Una factura dudosa no frena el alta: se deja para la bandeja
                # del expediente, donde se revisa con sus evidencias.
                by_hash = {source.sha256: source.path for source in draft.sources}
                entry.invoices_for_inbox = tuple(by_hash[sha] for sha in sorted(ambiguous))
                entry.invoices = tuple(path for path in entry.invoices
                                       if path not in entry.invoices_for_inbox)
                draft = analyse_sources(invoice_paths=entry.invoices, **arguments)
        except Exception as error:
            entry.status = FAILED
            entry.reasons.append(f"No se pudieron analizar las fuentes: {error}")
            continue
        answers, pending = automatic_answers(draft, entry.end)
        entry.draft, entry.answers = draft, answers
        if pending:
            entry.reasons.extend(pending)
            continue
        try:
            from community_onboarding import resolve_onboarding_configuration
            resolve_onboarding_configuration(draft, answers)
        except ValueError as error:
            entry.reasons.append(str(error))
            continue
        entry.status = READY

    _block_repeated_codes(entries)
    return tuple(entries)


def _block_repeated_codes(entries: Iterable[BatchEntry]) -> None:
    by_code: dict[str, list[BatchEntry]] = {}
    for entry in entries:
        if entry.code and entry.status in {READY, REVIEW}:
            by_code.setdefault(entry.code, []).append(entry)
    for code, items in by_code.items():
        if len(items) > 1:
            names = ", ".join(item.folder.name for item in items)
            for item in items:
                item.status = REVIEW
                item.reasons.append(f"El código {code} aparece en varias carpetas ({names}).")


def create_batch(
    connection: sqlite3.Connection,
    entries: Iterable[BatchEntry],
    *,
    project_root: Path,
    archive_root: Path,
    actor: str,
    progress: Callable[[str], None] | None = None,
) -> tuple[BatchOutcome, ...]:
    """Crea las comunidades aprobadas, una a una y de forma independiente."""
    from case_workflow_actions import resolve_case_profile
    from community_onboarding import confirm_onboarding
    from excel_bootstrap_importer import import_companion_sources

    outcomes: list[BatchOutcome] = []
    for entry in entries:
        if not entry.can_create:
            continue
        if progress:
            progress(f"{entry.code} — {entry.name}")
        if entry.code in _registered_codes(connection) or _profile_exists(entry.code, project_root):
            outcomes.append(BatchOutcome(entry.folder, entry.code, "skipped",
                                         "Ya estaba registrada; no se ha duplicado."))
            continue
        try:
            result = confirm_onboarding(
                connection, draft=entry.draft, answers=entry.answers,
                project_root=project_root, archive_root=archive_root,
                period_name=entry.period_name, start_date=entry.start, end_date=entry.end,
                actor=actor,
            )
        except Exception as error:
            if connection.in_transaction:
                connection.rollback()
            outcomes.append(BatchOutcome(entry.folder, entry.code, "failed", f"No se pudo crear: {error}"))
            continue
        owners = readings = 0
        notes = []
        try:
            owner_source = next(source.path for source in entry.draft.sources if source.kind == "owner_list")
            profile = resolve_case_profile(
                connection, id_case=result.case_id, active_community_id=result.community_id,
                project_root=project_root,
            )
            companions = import_companion_sources(
                connection, id_case=result.case_id, owner_list_path=owner_source,
                readings_path=entry.readings, profile=profile, actor=actor,
            )
            owners, readings = companions.imported_owner_count, companions.imported_reading_count
            if companions.open_issue_count:
                notes.append(f"{companions.open_issue_count} incidencia(s) por revisar")
        except Exception as error:
            # La comunidad y su expediente ya existen; los propietarios y
            # lecturas se pueden cargar después desde el expediente.
            if connection.in_transaction:
                connection.rollback()
            notes.append(f"propietarios y lecturas pendientes de cargar ({error})")
        if entry.invoices_for_inbox:
            notes.append(f"{len(entry.invoices_for_inbox)} factura(s) dudosa(s) para añadir desde la bandeja")
        message = f"Creada con {owners} propietario(s) y {readings} lectura(s)"
        outcomes.append(BatchOutcome(
            entry.folder, entry.code, "created",
            message + (": " + "; ".join(notes) if notes else "."),
            case_id=result.case_id, owner_count=owners, reading_count=readings,
        ))
    return tuple(outcomes)


__all__ = [
    "BatchEntry", "BatchOutcome", "EXISTING", "FAILED", "READY", "REVIEW",
    "automatic_answers", "code_and_name_from_folder", "create_batch", "normalise_reading_source",
    "sample_text", "scan_batch",
]
