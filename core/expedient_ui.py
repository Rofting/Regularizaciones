import os
import json
import sqlite3
import stat
import sys
import threading
import tkinter as tk
from collections import Counter
from dataclasses import dataclass
import re
from datetime import date, datetime
from pathlib import Path
from tkinter import filedialog, messagebox
from typing import TYPE_CHECKING, Any, Mapping, Sequence, TypeVar

import customtkinter as ctk

import case_ingestion
import case_readiness
import cases_overview
import community_batch
import cuotas_servicio
import community_discovery
import community_folder
import community_onboarding
import document_review
import expedient_service
import fixed_costs
import case_coherence
import gestor_bd
import office_settings
import owner_lists
import period_selection
import source_batch
import source_preview
import ui_moderna as UIM
from expedient_models import ReviewIssue
from ui_moderna import C


if TYPE_CHECKING:
    from app import AppGestionFincas


_SOURCE_SUFFIXES = frozenset({".pdf", ".xlsx", ".xls", ".csv"})
_Issue = TypeVar("_Issue")


def issue_page(
    issues: Sequence[_Issue], page: int, page_size: int = 10,
) -> tuple[tuple[_Issue, ...], int, int]:
    """Return one bounded, one-based page without losing the total page count."""
    if page_size < 1:
        raise ValueError("El tamaño de página debe ser positivo")
    items = tuple(issues)
    total_pages = max(1, (len(items) + page_size - 1) // page_size)
    normalised_page = min(max(int(page), 1), total_pages)
    start = (normalised_page - 1) * page_size
    return items[start:start + page_size], normalised_page, total_pages


def source_summary(kinds) -> str:
    """Agrupa los resultados del análisis para la revisión guiada."""
    counts = Counter(kinds)
    labels = (
        ("invoice", "factura detectada", "facturas detectadas"),
        ("reading", "lectura", "lecturas"),
        ("owners", "listado de propietarios", "listados de propietarios"),
        ("other", "otro documento", "otros documentos"),
        ("unknown", "documento por revisar", "documentos por revisar"),
    )
    return " · ".join(
        f"{counts[kind]} {singular if counts[kind] == 1 else plural}"
        for kind, singular, plural in labels if counts[kind]
    ) or "Sin documentos analizados"


def group_review_issues(issues: Sequence[ReviewIssue]) -> tuple[dict[str, object], ...]:
    """Agrupa revisiones repetidas del mismo informe en una sola decisión."""
    groups: list[dict[str, object]] = []
    repeatable_groups: dict[tuple[int, str], dict[str, object]] = {}
    for issue in issues:
        if issue.code not in {"COUNTER_RESET", "READING_ZERO_REVIEW"}:
            groups.append({"document_id": issue.id_document, "count": 1, "representative": issue})
            continue
        key = (issue.id_document, issue.code)
        group = repeatable_groups.get(key)
        if group is None:
            group = {"document_id": issue.id_document, "count": 0, "representative": issue}
            repeatable_groups[key] = group
            groups.append(group)
        group["count"] = int(group["count"]) + 1
    return tuple(groups)


def issue_context_label(source_context: str | None) -> str:
    """Describe metadatos nuevos y fragmentos históricos sin exigir un visor."""
    fallback = "No hay contexto guardado. Usa Abrir archivo para consultar la fuente."
    if not source_context:
        return fallback
    try:
        context = json.loads(source_context)
    except (ValueError, TypeError):
        return str(source_context)
    if not isinstance(context, dict):
        return fallback
    parts = []
    if context.get("page"):
        parts.append(f"Página {context['page']}")
    position = "!".join(str(context[key]) for key in ("sheet", "cell") if context.get(key))
    if position:
        parts.append(position)
    fragment = context.get("fragment") or context.get("excerpt")
    if fragment:
        parts.append(str(fragment))
    return "\n\n".join(parts) or fallback


def show_issue_context(app: "AppGestionFincas", issue: ReviewIssue) -> None:
    connection = gestor_bd.conectar(str(app.ruta_bd_expedientes))
    try:
        # Missing fields may have no candidate of their own; prefer their field,
        # then a fragment extracted from another field of the same document.
        row = connection.execute(
            """SELECT source_context FROM extraction_candidates
               WHERE id_document = ? AND source_context IS NOT NULL
                 AND trim(source_context) <> ''
               ORDER BY (field_name = ?) DESC, field_name LIMIT 1""",
            (issue.id_document, issue.field_name),
        ).fetchone()
        context = row["source_context"] if row else None
        if not context:
            row = connection.execute("SELECT source_context FROM source_documents WHERE id_document=?",
                                     (issue.id_document,)).fetchone()
            context = row['source_context'] if row else None
    finally:
        connection.close()
    messagebox.showinfo(
        "Contexto de la fuente",
        f"{issue.archived_path.name}\n\n{issue_context_label(context)}",
        parent=app,
    )


def _history_event_label(event_type: str, details_json: str) -> str:
    """Turn durable audit events into concise labels for the period review."""
    try:
        details = json.loads(details_json or "{}")
    except (TypeError, ValueError):
        details = {}
    labels = {
        "source_registered": "Fuente archivada",
        "source_status_changed": "Estado de fuente actualizado",
        "manual_correction": "Decisión registrada",
        "case_status_changed": "Estado del expediente actualizado",
        "excel_export_recorded": "Excel oficial registrado",
        "letters_recorded": "Lote de cartas registrado",
    }
    title = labels.get(event_type, event_type.replace("_", " ").capitalize())
    meaningful = [
        str(details[key]) for key in ("name", "field", "reason", "to", "status", "output_path")
        if details.get(key) not in (None, "")
    ]
    return title + (" · " + " · ".join(meaningful) if meaningful else "")


def open_case_history_dialog(app: "AppGestionFincas", case_id: int) -> None:
    """Lets a manager audit any selected historical period without reprocessing it."""
    dialog = _dialog(app, "Historial del período", 820, 620)
    panel = ctk.CTkScrollableFrame(dialog)
    panel.pack(fill="both", expand=True, padx=16, pady=16)
    connection = gestor_bd.conectar(str(app.ruta_bd_expedientes))
    try:
        rows = connection.execute(
            """SELECT h.event_type,h.details_json,h.created_at,d.original_name
                 FROM case_history_events h
                 LEFT JOIN source_documents d ON d.id_document=h.id_document
                WHERE h.id_case=?
                ORDER BY h.created_at DESC,h.id_event DESC""",
            (case_id,),
        ).fetchall()
    finally:
        connection.close()
    if not rows:
        ctk.CTkLabel(
            panel,
            text=("Todavía no hay eventos del período. Las fuentes, lecturas y resultados "
                  "se conservarán aquí a medida que se incorporen o validen."),
            justify="left", wraplength=720, text_color=C["texto_sec"],
        ).pack(anchor="w", padx=14, pady=18)
        return
    ctk.CTkLabel(
        panel,
        text="Historial inmutable del expediente",
        font=UIM.fuente(17, "bold"), text_color=C["texto"],
    ).pack(anchor="w", padx=12, pady=(10, 3))
    ctk.CTkLabel(
        panel,
        text=("Incluye fuentes archivadas, decisiones manuales y salidas. "
              "Las lecturas originales no se sustituyen."),
        wraplength=720, justify="left", text_color=C["texto_sec"],
    ).pack(anchor="w", padx=12, pady=(0, 12))
    for row in rows:
        card = ctk.CTkFrame(panel, fg_color=C["panel"], corner_radius=10,
                            border_width=1, border_color=C["borde"])
        card.pack(fill="x", padx=8, pady=4)
        ctk.CTkLabel(
            card, text=_history_event_label(row["event_type"], row["details_json"]),
            font=UIM.fuente(11, "bold"), text_color=C["texto"], anchor="w",
            wraplength=670,
        ).pack(fill="x", padx=12, pady=(9, 2))
        source = f" · {row['original_name']}" if row["original_name"] else ""
        ctk.CTkLabel(
            card, text=f"{row['created_at']}{source}", font=UIM.fuente(10),
            text_color=C["texto_sec"], anchor="w", wraplength=670,
        ).pack(fill="x", padx=12, pady=(0, 9))


@dataclass(frozen=True)
class GuidedStep:
    """One visible step in the guided regularization workspace."""

    key: str
    label: str
    status: str


@dataclass(frozen=True)
class GuidedWorkspaceState:
    """UI-only projection of a regularization case and its allowed next action."""

    active_step: str
    next_action: str
    headline: str
    detail: str
    steps: tuple[GuidedStep, ...]
    blockers: tuple[case_readiness.ReadinessBlocker, ...] = ()
    # Acciones que rehacen una etapa ya superada (reevaluar fuentes, regenerar
    # Excel...). Cada repetición conserva las ejecuciones previas en el historial.
    repeat_actions: tuple[str, ...] = ()


@dataclass(frozen=True)
class ReviewSummary:
    """Separates stored technical checks from the decisions a manager must take."""

    technical_count: int
    groups: tuple[dict[str, object], ...]

    @property
    def actionable_count(self) -> int:
        return len(self.groups)


def review_summary(issues: Sequence[ReviewIssue]) -> ReviewSummary:
    """Collapse repeatable checks without hiding their underlying audit rows."""
    items = tuple(issues)
    return ReviewSummary(
        technical_count=len(items),
        groups=group_review_issues(items),
    )


_GUIDED_STEP_LABELS = (
    ("fuentes", "Fuentes"),
    ("validar", "Validar"),
    ("reparto", "Reparto"),
    ("cartas", "Cartas"),
)


def guided_workspace_state(
    *, has_case: bool, document_count: int, open_issue_count: int, case_status: str,
    has_registered_template: bool = True, profile_issue: str | None = None,
    profile_missing: bool = False,
    readiness_report: case_readiness.CaseReadinessReport | None = None,
) -> GuidedWorkspaceState:
    """Returns the next safe user action without replacing workflow service gates."""
    blockers = tuple(readiness_report.blockers) if readiness_report is not None else ()
    if not has_case:
        active_step, next_action = "fuentes", "crear_expediente"
        headline = "Crear expediente"
        detail = "Elige el intervalo que vas a regularizar antes de incorporar fuentes."
    elif blockers:
        blocker = blockers[0]
        routes = {
            "add_sources": ("fuentes", "anadir_fuentes", "Añadir fuentes"),
            "resolve_issues": ("validar", "resolver_incidencias", "Resolver incidencias"),
            "confirm_sources": ("validar", "confirmar_fuentes", "Confirmar fuentes"),
            "manage_periods": ("fuentes", "gestionar_periodos", "Definir período"),
            "import_owners": ("fuentes", "importar_propietarios", "Importar propietarios"),
            "prepare_excel": ("reparto", "generar_excel", "Preparar Excel oficial"),
            "manage_profile": ("validar", "revalidar_perfil", "Revisar perfil Excel"),
            "revalidate_profile": ("validar", "revalidar_perfil", "Revalidar perfil Excel"),
            "review_sources": ("validar", "confirmar_fuentes", "Completar datos de fuentes"),
            "review_owners": ("validar", "importar_propietarios", "Revisar coeficientes"),
            "review_readings": ("validar", "anadir_fuentes", "Completar lecturas"),
            "review_fixed_costs": ("reparto", "revisar_gastos_fijos", "Revisar gastos fijos"),
            "review_coherence": ("reparto", "revisar_coherencia", "Revisar coherencia"),
            "generate_excel": ("reparto", "generar_excel", "Generar Excel oficial"),
            "calculate_distribution": ("reparto", "calcular_reparto", "Calcular reparto"),
        }
        active_step, next_action, headline = routes.get(
            blocker.action,
            ("validar", "confirmar_fuentes", "Completar expediente"),
        )
        detail = blocker.message
        if len(blockers) > 1:
            detail += f" Después quedan {len(blockers) - 1} requisito(s) más."
    elif open_issue_count:
        active_step, next_action = "validar", "resolver_incidencias"
        headline = "Resuelve las incidencias"
        detail = f"Hay {open_issue_count} decisión(es) pendiente(s) antes de continuar."
    elif profile_missing:
        active_step, next_action = "reparto", "generar_excel"
        headline = "Prepara el Excel oficial"
        detail = (
            "La comunidad aún no tiene un perfil Excel registrado. Se creará "
            "automáticamente desde el modelo común al generar el Excel oficial."
        )
    elif profile_issue:
        active_step, next_action = "validar", "revalidar_perfil"
        headline = "Revalidar configuración del Excel"
        detail = (
            f"{profile_issue}. Comprueba la plantilla y actualiza su registro "
            "para volver a generar el Excel oficial."
        )
    elif case_status in {"ready_for_calculation"}:
        # Ya no se pide el modelo inicial: si la comunidad no tiene plantilla,
        # generar la crea desde el modelo canónico del despacho. «Importar
        # modelo inicial» sigue disponible para partir de un libro propio.
        active_step, next_action = "reparto", "generar_excel"
        headline = "Genera el Excel oficial"
        detail = (
            "Las fuentes están validadas; prepara el modelo antes del reparto final."
            if has_registered_template else
            "Las fuentes están validadas. Se creará la plantilla de esta comunidad "
            "a partir del modelo del despacho."
        )
    elif case_status in {"calculated"}:
        active_step, next_action = "reparto", "calcular_reparto"
        headline = "Calcula el reparto final"
        detail = "El Excel está preparado; calcula y concilia el reparto por propietario."
    elif case_status == "reconciled":
        active_step, next_action = "cartas", "generar_cartas"
        headline = "Genera las cartas"
        detail = "El reparto está conciliado y listo para comunicar a cada propietario."
    elif case_status in {"deliveries_generated", "closed"}:
        active_step, next_action = "cartas", "abrir_salidas"
        headline = "Cartas generadas"
        detail = (
            "Puedes abrir las salidas o repetir cualquier etapa: reevaluar fuentes, "
            "Excel, reparto o cartas. Cada repetición conserva las anteriores en el historial."
        )
    elif document_count <= 0 or case_status in {"draft", "gathering_sources", ""}:
        active_step, next_action = "fuentes", "anadir_fuentes"
        headline = "Incorpora las fuentes"
        detail = "Añade facturas, lecturas y Excel del período para iniciar la revisión."
    elif case_status == "under_review":
        active_step, next_action = "validar", "confirmar_fuentes"
        headline = "Confirmar fuentes"
        detail = (
            "No quedan incidencias. Confirma y aplica las fuentes al expediente "
            "antes de generar el Excel oficial."
        )
    else:
        active_step, next_action = "fuentes", "anadir_fuentes"
        headline = "Completa las fuentes"
        detail = "El estado del expediente requiere revisar sus fuentes antes de avanzar."

    active_index = next(
        index for index, (key, _label) in enumerate(_GUIDED_STEP_LABELS)
        if key == active_step
    )
    complete = case_status in {"deliveries_generated", "closed"}
    steps = tuple(
        GuidedStep(
            key=key,
            label=label,
            status=(
                "done" if complete or index < active_index else
                "active" if index == active_index else
                "blocked"
            ),
        )
        for index, (key, label) in enumerate(_GUIDED_STEP_LABELS)
    )
    return GuidedWorkspaceState(
        active_step=active_step,
        next_action=next_action,
        headline=headline,
        detail=detail,
        steps=steps,
        blockers=blockers,
        repeat_actions=(
            () if not has_case
            # Con requisitos pendientes, lo ya generado no es coherente con las
            # fuentes: sólo tiene sentido volver a leerlas.
            else tuple(
                item for item in repeat_actions_for_status(case_status)
                if not blockers or item in {"reevaluar_fuentes"}
            )
        ),
    )


# Etapas ya superadas que pueden repetirse desde cada estado. El servicio de
# flujo sigue validando cada acción; esto sólo decide qué botones mostrar.
_REPEATABLE_BY_STATUS = {
    "ready_for_calculation": ("reevaluar_fuentes", "anadir_fuentes"),
    "calculated": ("reevaluar_fuentes", "anadir_fuentes", "generar_excel"),
    "reconciled": ("reevaluar_fuentes", "anadir_fuentes", "generar_excel", "calcular_reparto"),
    "deliveries_generated": (
        "reevaluar_fuentes", "anadir_fuentes", "generar_excel", "calcular_reparto", "generar_cartas",
    ),
    "closed": (
        "reevaluar_fuentes", "anadir_fuentes", "generar_excel", "calcular_reparto", "generar_cartas",
    ),
}


def repeat_actions_for_status(case_status: str) -> tuple[str, ...]:
    return _REPEATABLE_BY_STATUS.get(case_status, ())


def _is_hidden_source_path(path: Path) -> bool:
    """Identifica ocultos de Windows y temporales de Office sin depender del tema."""
    if path.name.startswith((".", "~$")):
        return True
    try:
        attributes = getattr(path.stat(), "st_file_attributes", 0)
    except OSError:
        return True
    hidden_flag = getattr(stat, "FILE_ATTRIBUTE_HIDDEN", 0x2)
    return bool(attributes & hidden_flag)


def is_supported_source_file(path: Path) -> bool:
    """Indica si un documento se puede incorporar de forma segura al expediente."""
    candidate = Path(path)
    return (
        candidate.is_file()
        and not _is_hidden_source_path(candidate)
        and candidate.suffix.lower() in _SOURCE_SUFFIXES
    )


def source_files_in_folder(folder: Path) -> list[Path]:
    """Devuelve las fuentes compatibles, visibles y ordenadas de una carpeta.

    Incluye sus subcarpetas para que un expediente completo se pueda cargar
    en una sola acción, ignorando los directorios y documentos ocultos.
    """
    root = Path(folder)
    if not root.is_dir():
        return []
    return sorted(
        (
            path for path in root.rglob("*")
            if is_supported_source_file(path)
            and not any(
                _is_hidden_source_path(root.joinpath(*path.relative_to(root).parts[:index]))
                for index in range(1, len(path.relative_to(root).parts))
            )
        ),
        key=lambda path: str(path.relative_to(root)).casefold(),
    )


def open_detect_communities_dialog(app: "AppGestionFincas") -> None:
    """Offer automatic creation for clear community groups in a source folder."""
    if getattr(app, "_procesando", False):
        messagebox.showwarning("Espera", "Termina la operación actual antes de analizar una carpeta.", parent=app)
        return
    folder = filedialog.askdirectory(parent=app, title="Selecciona la carpeta con documentos de comunidades")
    if not folder:
        return
    paths = source_files_in_folder(Path(folder))
    if not paths:
        messagebox.showwarning(
            "Sin fuentes compatibles",
            "La carpeta no contiene PDF, XLSX, XLS o CSV visibles.",
            parent=app,
        )
        return

    app._estado("Detectando comunidades en la carpeta…", procesando=True)

    def work():
        try:
            proposal = community_discovery.build_global_intake(paths)
            candidates = community_discovery.discover_communities(paths)
        except Exception as error:
            message = str(error)

            def failed():
                app._estado("No se pudo analizar la carpeta", procesando=False)
                messagebox.showerror("No se pudo analizar la carpeta", message, parent=app)
            app.after(0, failed)
            return
        app.after(0, lambda: _show_detected_communities(
            app, Path(folder), candidates, len(paths), proposal,
        ))

    app._en_hilo(work)


def _show_detected_communities(
    app: "AppGestionFincas", folder: Path, candidates, source_count: int, proposal,
) -> None:
    app._estado("Detección de comunidades lista", procesando=False)
    dialog = _dialog(app, "Comunidades detectadas", 800, 620)
    panel = ctk.CTkFrame(dialog, fg_color=C["panel"], corner_radius=16,
                         border_width=1, border_color=C["borde"])
    panel.pack(fill="both", expand=True, padx=18, pady=18)
    ctk.CTkLabel(panel, text="Comunidades detectadas", font=UIM.fuente(21, "bold"),
                 text_color=C["texto"]).pack(anchor="w", padx=22, pady=(20, 2))
    ctk.CTkLabel(
        panel,
        text=(f"Se han revisado {source_count} archivo(s) de {folder.name}. "
              "Se han agrupado sin usar la comunidad o el período seleccionados. "
              "Los originales no se moverán ni se modificarán."),
        font=UIM.fuente(11), text_color=C["texto_sec"], wraplength=720, justify="left",
    ).pack(anchor="w", padx=22, pady=(0, 14))

    body = ctk.CTkScrollableFrame(panel, fg_color=C["panel_2"], corner_radius=11)
    body.pack(fill="both", expand=True, padx=22, pady=(0, 12))
    approved = []
    groups_by_code = {group.community_code: group for group in proposal.groups}
    for candidate in candidates:
        row = ctk.CTkFrame(body, fg_color=C["panel"], corner_radius=10,
                           border_width=1, border_color=C["borde"])
        row.pack(fill="x", padx=8, pady=6)
        variable = tk.BooleanVar(value=candidate.can_create, master=dialog)
        if candidate.can_create:
            approved.append((candidate, variable))
        ctk.CTkCheckBox(row, text="", variable=variable, width=26,
                         state="normal" if candidate.can_create else "disabled",
                         fg_color=C["primario"], hover_color=C["primario_hover"]).pack(side="left", padx=(12, 4), pady=12)
        details = ctk.CTkFrame(row, fg_color="transparent")
        details.pack(side="left", fill="x", expand=True, padx=(0, 12), pady=11)
        ctk.CTkLabel(details, text=f"{candidate.code} — {candidate.name}",
                     font=UIM.fuente(12, "bold"), text_color=C["texto"]).pack(anchor="w")
        group = groups_by_code.get(candidate.code)
        cif_text = candidate.cif or "CIF no localizado"
        group_detail = ""
        if group is not None:
            services = ", ".join(group.supply_hints) or "tipo por revisar"
            group_detail = (
                f"{len(group.source_paths)} documento(s) · {group.display_periods} · "
                f"{services}"
            )
        ctk.CTkLabel(
            details,
            text=(candidate.blocked_reason or
                  f"{group_detail or f'{len(candidate.source_paths)} documento(s)'} · {cif_text}"),
            font=UIM.fuente(10),
            text_color=C["alerta"] if candidate.blocked_reason else C["texto_sec"],
            wraplength=560, justify="left",
        ).pack(anchor="w", pady=(2, 0))

    if proposal.unassigned_paths:
        ctk.CTkLabel(
            body,
            text=(f"{len(proposal.unassigned_paths)} documento(s) quedan sin comunidad asignada: "
                  "no se asociarán a la comunidad activa. Renómbralos con el código o revísalos desde la bandeja."),
            font=UIM.fuente(10), text_color=C["alerta"], wraplength=680, justify="left",
        ).pack(anchor="w", padx=14, pady=(10, 14))

    if not candidates:
        ctk.CTkLabel(
            body,
            text="No se ha encontrado un código seguro. Nombra los archivos con el código al inicio (por ejemplo, 658_factura.pdf) o guárdalos dentro de una carpeta llamada 658.",
            font=UIM.fuente(11), text_color=C["texto_sec"], wraplength=650, justify="left",
        ).pack(anchor="w", padx=14, pady=14)

    footer = ctk.CTkFrame(panel, fg_color="transparent")
    footer.pack(fill="x", padx=22, pady=(0, 18))

    def create_selected():
        selected = tuple(candidate for candidate, variable in approved if variable.get())
        if not selected:
            messagebox.showwarning("Sin comunidades seleccionadas", "Marca al menos una comunidad detectada.", parent=dialog)
            return
        for widget in footer.winfo_children():
            widget.configure(state="disabled")
        app._estado("Creando comunidades detectadas…", procesando=True)

        def create_work():
            connection = None
            try:
                connection = gestor_bd.conectar(str(app.ruta_bd_expedientes))
                created = []
                for candidate in selected:
                    community_id = gestor_bd.obtener_o_crear_comunidad(
                        connection, candidate.code, candidate.name, cif=candidate.cif,
                    )
                    created.append((community_id, candidate))
            except Exception as error:
                message = str(error)

                def failed():
                    app._estado("No se pudieron crear las comunidades", procesando=False)
                    for widget in footer.winfo_children():
                        widget.configure(state="normal")
                    messagebox.showerror("No se pudieron crear las comunidades", message, parent=dialog)
                app.after(0, failed)
                return
            finally:
                if connection is not None:
                    connection.close()

            def completed():
                for _community_id, candidate in created:
                    app._crear_excel_si_no_existe(candidate.code, candidate.name)
                app._cargar_comunidades()
                first = created[0][1]
                option = f"{first.code} — {first.name}"
                if option in getattr(app, "_ids_comunidad", {}):
                    app.comunidad_actual.set(option)
                    app.cb_comunidad.set(option)
                    app._on_comunidad_seleccionada()
                app._estado("Comunidades detectadas creadas", procesando=False)
                dialog.destroy()
                messagebox.showinfo(
                    "Comunidades creadas",
                    f"Se han creado o actualizado {len(created)} comunidad(es).\n\n"
                    "Los documentos originales permanecen en su carpeta. Selecciona una comunidad y crea su expediente antes de añadir sus fuentes al flujo guiado.",
                    parent=app,
                )
            app.after(0, completed)

        app._en_hilo(create_work)

    ctk.CTkButton(footer, text="Crear comunidades seleccionadas", command=create_selected,
                  height=38, corner_radius=9, font=UIM.fuente(11, "bold"),
                  fg_color=C["primario"], hover_color=C["primario_hover"]).pack(side="right")
    ctk.CTkButton(footer, text="Cancelar", command=dialog.destroy, height=38,
                  corner_radius=9, font=UIM.fuente(11), fg_color="transparent",
                  border_width=1, border_color=C["borde"], text_color=C["primario"],
                  hover_color=C["acento_suave"]).pack(side="right", padx=(0, 8))


_BATCH_STATUS = {
    community_batch.READY: ("Lista para crear", "primario"),
    community_batch.REVIEW: ("Necesita revisión", "alerta"),
    community_batch.EXISTING: ("Ya registrada", "texto_sec"),
    community_batch.FAILED: ("No se pudo analizar", "alerta"),
}


def open_batch_onboarding_dialog(app: "AppGestionFincas") -> None:
    """Alta masiva: una comunidad por subcarpeta, con aprobación previa."""
    if getattr(app, "_procesando", False):
        messagebox.showwarning("Espera", "Termina la operación actual antes de analizar un lote.", parent=app)
        return
    folder = filedialog.askdirectory(
        parent=app, title="Carpeta con una subcarpeta por comunidad (p. ej. «658 - CP Las Flores»)",
    )
    if not folder:
        return
    project_root = Path(__file__).resolve().parent.parent
    work_directory = Path(app.ruta_bd_expedientes).parent / "importaciones"
    app._estado("Analizando el lote de comunidades…", procesando=True)

    def work():
        connection = None
        try:
            connection = gestor_bd.conectar(str(app.ruta_bd_expedientes))
            entries = community_batch.scan_batch(
                Path(folder), connection=connection, project_root=project_root,
                work_directory=work_directory,
                progress=lambda name: app.after(0, lambda: app._estado(f"Analizando {name}…", procesando=True)),
            )
        except Exception as error:
            message = str(error)

            def failed():
                app._estado("No se pudo analizar el lote", procesando=False)
                messagebox.showerror("No se pudo analizar el lote", message, parent=app)
            app.after(0, failed)
            return
        finally:
            if connection is not None:
                connection.close()
        app.after(0, lambda: _show_batch_entries(app, Path(folder), entries, project_root))

    app._en_hilo(work)


def _show_batch_entries(app: "AppGestionFincas", folder: Path, entries, project_root: Path) -> None:
    app._estado("Lote analizado", procesando=False)
    dialog = _dialog(app, "Alta masiva de comunidades", 860, 660)
    panel = ctk.CTkFrame(dialog, fg_color=C["panel"], corner_radius=16,
                         border_width=1, border_color=C["borde"])
    panel.pack(fill="both", expand=True, padx=18, pady=18)
    ctk.CTkLabel(panel, text="Alta masiva de comunidades", font=UIM.fuente(21, "bold"),
                 text_color=C["texto"]).pack(anchor="w", padx=22, pady=(20, 2))
    ready = sum(entry.can_create for entry in entries)
    ctk.CTkLabel(
        panel,
        text=(f"{len(entries)} subcarpeta(s) en {folder.name}; {ready} lista(s) para crear. "
              "Revisa cada comunidad y la lectura de ejemplo antes de aprobarla. Se crearán "
              "comunidad, perfil, expediente, propietarios y lecturas; los originales no se tocan. "
              "Las que necesitan revisión se completan con «Alta guiada desde fuentes» → «Rellenar desde carpeta»."),
        font=UIM.fuente(11), text_color=C["texto_sec"], wraplength=780, justify="left",
    ).pack(anchor="w", padx=22, pady=(0, 12))

    body = ctk.CTkScrollableFrame(panel, fg_color=C["panel_2"], corner_radius=11)
    body.pack(fill="both", expand=True, padx=22, pady=(0, 12))
    approvals = []
    for entry in entries:
        row = ctk.CTkFrame(body, fg_color=C["panel"], corner_radius=10,
                           border_width=1, border_color=C["borde"])
        row.pack(fill="x", padx=8, pady=5)
        variable = tk.BooleanVar(value=entry.can_create, master=dialog)
        if entry.can_create:
            approvals.append((entry, variable))
        ctk.CTkCheckBox(row, text="", variable=variable, width=26,
                         state="normal" if entry.can_create else "disabled",
                         fg_color=C["primario"], hover_color=C["primario_hover"]).pack(side="left", padx=(12, 4), pady=12)
        details = ctk.CTkFrame(row, fg_color="transparent")
        details.pack(side="left", fill="x", expand=True, padx=(0, 12), pady=10)
        status, colour = _BATCH_STATUS[entry.status]
        title = f"{entry.code or '¿código?'} — {entry.name or entry.folder.name}"
        ctk.CTkLabel(details, text=f"{title}   ·   {status}", font=UIM.fuente(12, "bold"),
                     text_color=C["texto"] if entry.can_create else C[colour]).pack(anchor="w")
        lines = [f"Carpeta: {entry.folder.name} · {entry.summary}"]
        if entry.invoices_for_inbox:
            lines.append(f"{len(entry.invoices_for_inbox)} factura(s) dudosa(s) quedarán para la bandeja del expediente.")
        lines.extend(entry.reasons)
        ctk.CTkLabel(details, text="\n".join(lines), font=UIM.fuente(10),
                     text_color=C["texto_sec"], wraplength=680, justify="left").pack(anchor="w", pady=(2, 0))

    footer = ctk.CTkFrame(panel, fg_color="transparent")
    footer.pack(fill="x", padx=22, pady=(0, 18))

    def create_selected():
        selected = tuple(entry for entry, variable in approvals if variable.get())
        if not selected:
            messagebox.showwarning("Sin comunidades aprobadas", "Marca al menos una comunidad lista.", parent=dialog)
            return
        for widget in footer.winfo_children():
            widget.configure(state="disabled")
        app._estado("Creando comunidades del lote…", procesando=True)

        def create_work():
            connection = None
            try:
                connection = gestor_bd.conectar(str(app.ruta_bd_expedientes))
                outcomes = community_batch.create_batch(
                    connection, selected, project_root=project_root,
                    archive_root=Path(app.ruta_archivo_expedientes), actor="usuario_local",
                    progress=lambda name: app.after(0, lambda: app._estado(f"Creando {name}…", procesando=True)),
                )
            except Exception as error:
                message = str(error)

                def failed():
                    app._estado("No se pudo crear el lote", procesando=False)
                    for widget in footer.winfo_children():
                        widget.configure(state="normal")
                    messagebox.showerror("No se pudo crear el lote", message, parent=dialog)
                app.after(0, failed)
                return
            finally:
                if connection is not None:
                    connection.close()

            def completed():
                app._estado("Lote de comunidades terminado", procesando=False)
                app._cargar_comunidades()
                created = [item for item in outcomes if item.status == "created"]
                for item in outcomes:
                    app.log(f"{item.code}: {item.message}", "ok" if item.status == "created" else "aviso")
                dialog.destroy()
                labels = {"created": "✓", "skipped": "=", "failed": "✗"}
                messagebox.showinfo(
                    "Alta masiva terminada",
                    f"Creadas {len(created)} de {len(outcomes)} comunidad(es).\n\n"
                    + "\n".join(f"{labels[item.status]} {item.code}: {item.message}" for item in outcomes)
                    + "\n\nPuedes repetir el lote: las comunidades ya creadas no se duplican.",
                    parent=app,
                )
            app.after(0, completed)

        app._en_hilo(create_work)

    ctk.CTkButton(footer, text="Crear comunidades aprobadas", command=create_selected,
                  state="normal" if approvals else "disabled",
                  height=38, corner_radius=9, font=UIM.fuente(11, "bold"),
                  fg_color=C["primario"], hover_color=C["primario_hover"]).pack(side="right")
    ctk.CTkButton(footer, text="Cerrar", command=dialog.destroy, height=38,
                  corner_radius=9, font=UIM.fuente(11), fg_color="transparent",
                  border_width=1, border_color=C["borde"], text_color=C["primario"],
                  hover_color=C["acento_suave"]).pack(side="right", padx=(0, 8))


_BUCKET_COLOURS = {
    cases_overview.BLOCKED: "alerta",
    cases_overview.PENDING: "texto_sec",
    cases_overview.READY: "primario",
    cases_overview.CALCULATED: "primario",
    cases_overview.LETTERS: "exito",
    cases_overview.CLOSED: "texto_sec",
}


def open_cases_overview_dialog(app: "AppGestionFincas") -> None:
    """Panel no modal con los expedientes de todas las comunidades.

    Se puede dejar abierto mientras se trabaja: se actualiza tras cada acción
    del expediente y «Abrir» activa la comunidad y el expediente exactos.
    """
    existing = getattr(app, "_panel_comunidades", None)
    if existing is not None:
        try:
            if existing.winfo_exists():
                existing.deiconify()
                existing.lift()
                existing.focus_force()
                return
        except tk.TclError:
            pass
    window = ctk.CTkToplevel(app, fg_color=C["fondo"])
    window.title("Todas las comunidades")
    window.geometry("1040x700")
    window.minsize(860, 520)
    window.transient(app)
    app._panel_comunidades = window
    state = {"bucket": None, "cases": (), "rows": []}
    search = tk.StringVar(master=window)
    latest_only = tk.BooleanVar(master=window, value=True)

    panel = ctk.CTkFrame(window, fg_color=C["panel"], corner_radius=16,
                         border_width=1, border_color=C["borde"])
    panel.pack(fill="both", expand=True, padx=14, pady=14)
    top = ctk.CTkFrame(panel, fg_color="transparent")
    top.pack(fill="x", padx=18, pady=(16, 6))
    ctk.CTkLabel(top, text="Todas las comunidades", font=UIM.fuente(20, "bold"),
                 text_color=C["texto"]).pack(side="left")
    ctk.CTkButton(top, text="Actualizar", command=lambda: reload(), width=100, height=32,
                  corner_radius=8, **UIM.secondary_button_kwargs()).pack(side="right")

    chips = ctk.CTkFrame(panel, fg_color="transparent")
    chips.pack(fill="x", padx=18, pady=(0, 6))
    chip_buttons: dict[str | None, Any] = {}

    filters = ctk.CTkFrame(panel, fg_color="transparent")
    filters.pack(fill="x", padx=18, pady=(0, 8))
    ctk.CTkLabel(filters, text="Buscar", font=UIM.fuente(11, "bold"),
                 text_color=C["texto_sec"]).pack(side="left", padx=(0, 8))
    entry = ctk.CTkEntry(filters, textvariable=search, width=300, height=34, corner_radius=8,
                         border_color=C["borde"], fg_color=C["panel_2"], text_color=C["texto"])
    entry.pack(side="left")
    ctk.CTkCheckBox(filters, text="Sólo el expediente más reciente de cada comunidad",
                    variable=latest_only, command=lambda: render(), font=UIM.fuente(11),
                    fg_color=C["primario"], hover_color=C["primario_hover"]).pack(side="left", padx=14)
    summary = ctk.CTkLabel(filters, text="", font=UIM.fuente(11), text_color=C["texto_sec"])
    summary.pack(side="right")

    header = ctk.CTkFrame(panel, fg_color="transparent")
    header.pack(fill="x", padx=26)
    columns = (("Comunidad", 270), ("Expediente", 180), ("Estado", 170), ("Incidencias", 90),
               ("Última salida", 150))
    for title, width in columns:
        ctk.CTkLabel(header, text=title.upper(), width=width, anchor="w", font=UIM.fuente(9, "bold"),
                     text_color=C["texto_sec"]).pack(side="left")
    body = ctk.CTkScrollableFrame(panel, fg_color=C["panel_2"], corner_radius=11)
    body.pack(fill="both", expand=True, padx=18, pady=(4, 16))

    def set_bucket(bucket):
        state["bucket"] = None if state["bucket"] == bucket else bucket
        render()

    def open_case(case):
        if app.abrir_expediente(case.community_id, case.id_case):
            app.log(f"Abierto {case.community_label} · {case.case_name}", "info")
            app.lift()
            render()

    def render(*_args):
        cases = state["cases"]
        scoped = cases_overview.filter_cases(cases, latest_only=latest_only.get())
        counts = cases_overview.bucket_counts(scoped)
        for widget in chips.winfo_children():
            widget.destroy()
        chip_buttons.clear()
        for key, label in ((None, "Todos"), *cases_overview.BUCKETS):
            active = state["bucket"] == key
            count = len(scoped) if key is None else counts[key]
            button = ctk.CTkButton(
                chips, text=f"{label} · {count}", height=30, width=0, corner_radius=15,
                command=lambda value=key: set_bucket(value), font=UIM.fuente(11, "bold" if active else "normal"),
                **({"fg_color": C["primario"], "hover_color": C["primario_hover"]} if active
                   else UIM.secondary_button_kwargs()),
            )
            button.pack(side="left", padx=(0, 6))
            chip_buttons[key] = button
        visible = cases_overview.filter_cases(
            cases, bucket=state["bucket"], text=search.get(), latest_only=latest_only.get())
        summary.configure(text=f"{len(visible)} expediente(s)")
        for widget in body.winfo_children():
            widget.destroy()
        active_case = getattr(app, "id_expediente", None)
        for case in visible[:300]:
            current = case.id_case == active_case
            row = ctk.CTkFrame(body, fg_color=C["acento_suave"] if current else C["panel"], corner_radius=8,
                               border_width=1, border_color=C["primario"] if current else C["borde"])
            row.pack(fill="x", padx=6, pady=3)
            values = (
                (case.community_label, 270, "texto", "bold"),
                (f"{case.case_name}\n{case.period_label}", 180, "texto", "normal"),
                (f"{cases_overview.BUCKET_ROW_LABELS[case.bucket]}\n{case.status_label}", 170,
                 _BUCKET_COLOURS[case.bucket], "bold"),
                (str(case.open_issues) if case.open_issues else "—", 90,
                 "alerta" if case.open_issues else "texto_sec", "bold"),
                (case.last_output_label, 150, "texto_sec", "normal"),
            )
            ctk.CTkButton(row, text="Activo" if current else "Abrir", width=80, height=30, corner_radius=8,
                          state="disabled" if current else "normal",
                          command=lambda selected=case: open_case(selected),
                          fg_color=C["primario"], hover_color=C["primario_hover"]).pack(side="right", padx=10)
            for text, width, colour, weight in values:
                ctk.CTkLabel(row, text=text, width=width, anchor="w", justify="left",
                             font=UIM.fuente(11, weight), text_color=C.get(colour, C["texto"]),
                             wraplength=width - 12).pack(side="left", padx=(8, 0), pady=6)
        if not visible:
            ctk.CTkLabel(body, text="No hay expedientes con estos filtros.", font=UIM.fuente(11),
                         text_color=C["texto_sec"]).pack(pady=20)

    def reload():
        connection = gestor_bd.conectar(str(app.ruta_bd_expedientes))
        try:
            state["cases"] = cases_overview.list_cases(connection)
        finally:
            connection.close()
        render()

    observers = getattr(app, "_observadores_expedientes", None)
    if not isinstance(observers, set):
        observers = set()
        app._observadores_expedientes = observers

    def on_change():
        if window.winfo_exists():
            reload()

    observers.add(on_change)

    def close():
        observers.discard(on_change)
        app._panel_comunidades = None
        window.destroy()

    window.protocol("WM_DELETE_WINDOW", close)
    search.trace_add("write", render)
    reload()


def issue_guidance(field_name: str) -> dict[str, str]:
    """Traduce campos técnicos a instrucciones accionables de revisión."""
    guides = {
        "document_kind": {
            "label": "Tipo de documento",
            "what_to_find": "Consulta el original y confirma si es una factura, una lectura o un listado de propietarios.",
            "format": "Elige el tipo de fuente y explica la clasificación.",
            "why": "Solo se solicita cuando el análisis automático no reconoce el documento.",
        },
        "fecha_inicio": {
            "label": "Fecha de inicio del período facturado",
            "what_to_find": "Busca en la factura el inicio del período de suministro o consumo, normalmente junto a «Período facturado», «Desde» o «Fecha inicial».",
            "format": "Escribe la fecha como dd/mm/aaaa. Ejemplo: 01/08/2025.",
            "why": "Se necesita para decidir qué parte del consumo corresponde al ejercicio y para comprobar que el cálculo usa el período correcto.",
        },
        "fecha_fin": {
            "label": "Fecha de fin del período facturado",
            "what_to_find": "Busca el último día de suministro o consumo, junto a «Período facturado», «Hasta» o «Fecha final».",
            "format": "Escribe la fecha como dd/mm/aaaa. Ejemplo: 31/08/2025.",
            "why": "Cierra el intervalo de la factura y permite asignar el consumo al período y a su temporada correcta.",
        },
        "importe_total": {
            "label": "Importe total de la factura",
            "what_to_find": "Busca «Total factura», «Importe total» o el total final con impuestos. No uses una línea parcial ni el importe sin impuestos.",
            "format": "Escribe solo el número con dos decimales. Ejemplo: 245,70.",
            "why": "Es el importe que debe conciliar con los conceptos fijo y variable antes de repartir.",
        },
        "consumo_total": {
            "label": "Consumo total facturado",
            "what_to_find": "Busca el consumo del período y su unidad, por ejemplo kWh para energía o m³ para agua.",
            "format": "Escribe solo el valor numérico. Ejemplo: 1.245,50.",
            "why": "Permite comprobar el consumo contra las lecturas y calcular la parte variable.",
        },
        "proveedor": {
            "label": "Proveedor emisor de la factura",
            "what_to_find": "Busca la razón social que figura como emisor o comercializadora en la cabecera de la factura.",
            "format": "Escribe el nombre tal como aparece en el documento.",
            "why": "Identifica la fuente y evita mezclar facturas de suministros distintos.",
        },
        "concepto": {
            "label": "Concepto o suministro de la factura",
            "what_to_find": "Indica si corresponde a gas, electricidad, agua, lectura de contadores, mantenimiento u otro gasto.",
            "format": "Elige o escribe el concepto que describe el servicio real.",
            "why": "Determina la hoja del Excel y la regla de reparto que se aplicará.",
        },
    }
    guides.update({
        "tipo_suministro": {
            "label": "Tipo de suministro",
            "what_to_find": "Indica qué servicio factura el documento: GAS, ELECTRICIDAD, AGUA, ACS, CALEFACCION, MANTENIMIENTO…",
            "format": "Una palabra en mayúsculas. Ejemplo: GAS.",
            "why": "Decide en qué concepto del reparto entra la factura.",
        },
        "fecha_factura": {
            "label": "Fecha de emisión de la factura",
            "what_to_find": "Busca «Fecha de factura», «Fecha de emisión» o «Fecha» junto al número de factura.",
            "format": "Escribe la fecha como dd/mm/aaaa. Ejemplo: 05/04/2026.",
            "why": "Identifica la factura y ayuda a detectar duplicados.",
        },
        "num_factura": {
            "label": "Número de factura",
            "what_to_find": "Busca «Nº factura», «Número de factura» o «Factura nº» en la cabecera.",
            "format": "Cópialo tal cual, con letras y guiones. Ejemplo: FE-2026/00123.",
            "why": "Evita incorporar dos veces la misma factura.",
        },
        "consumo_kwh": {
            "label": "Consumo facturado (kWh)",
            "what_to_find": "Busca «Consumo total», «Energía consumida» o el total de kWh del período.",
            "format": "Solo el número. Ejemplo: 1.245,50.",
            "why": "Permite contrastar el consumo con las lecturas.",
        },
        "consumo_m3": {
            "label": "Consumo facturado (m³)",
            "what_to_find": "Busca «Consumo» seguido de m³ en el detalle de lecturas de la factura.",
            "format": "Solo el número. Ejemplo: 45.",
            "why": "Permite contrastar el consumo con las lecturas.",
        },
        "document.provider": {
            "label": "Proveedor de la factura",
            "what_to_find": "Mira en la cabecera la empresa que emite la factura (nombre y CIF).",
            "format": "Elige o escribe el proveedor tal como figura en el documento.",
            "why": "Con el proveedor se aplican sus reglas de lectura y se releen fechas e importes automáticamente.",
        },
        "document.eligibility": {
            "label": "¿La factura entra en este expediente?",
            "what_to_find": "Compara el período facturado y el servicio con las fechas y conceptos del expediente.",
            "format": "Ciérrala como no aplicable u omítela si pertenece a otro período o servicio.",
            "why": "Evita que una factura de otro período o servicio se sume al reparto.",
        },
        "tipo": {
            "label": "Servicio de las lecturas",
            "what_to_find": "Indica si las lecturas son de agua caliente (ACS) o de calefacción.",
            "format": "ACS o CALEFACCION.",
            "why": "Cada servicio se reparte en su propia hoja.",
        },
    })
    default = {
        "label": field_name.replace("_", " ").replace(".", " ").capitalize(),
        "what_to_find": "Busca este dato en el documento original antes de confirmarlo.",
        "format": "Copia el valor con el formato que aparece en la fuente.",
        "why": "Es necesario para mantener trazabilidad y evitar un cálculo con datos incompletos.",
    }
    return guides.get(field_name, default)


_ISSUE_TITLES = {
    "PROVIDER_UNKNOWN": "Proveedor sin identificar",
    "INVOICE_OUTSIDE_PERIOD": "Factura fuera del período",
    "ELIGIBILITY_REVIEW_REQUIRED": "Factura por confirmar",
    "DOCUMENT_CLASSIFICATION_REQUIRED": "Tipo de documento desconocido",
    "INVOICE_CONFLICT": "Posible factura duplicada",
    "INVOICE_PERIOD_CONFLICT": "Facturas con períodos solapados",
    "READING_CONFLICT": "Lecturas contradictorias",
    "OWNER_COEFFICIENT_CONFLICT": "Coeficiente distinto al registrado",
    "COUNTER_RESET": "Contador que baja de valor",
    "READING_ZERO_REVIEW": "Lectura a cero",
    "ARCHIVED_SOURCE_MISSING": "Archivo original no encontrado",
    "ARCHIVED_SOURCE_DUPLICATE": "Varias copias del archivo",
}


def issue_title(issue: ReviewIssue) -> str:
    """Título legible para la bandeja: qué pasa, no el nombre técnico del campo."""
    if issue.code in _ISSUE_TITLES:
        return _ISSUE_TITLES[issue.code]
    label = issue_guidance(issue.field_name)["label"]
    if issue.code == "MISSING_REQUIRED_FIELD" or not issue.detected_value:
        return f"Falta: {label[0].lower()}{label[1:]}"
    return f"Revisar: {label[0].lower()}{label[1:]}"


def source_display_name(path: Path) -> str:
    """Nombre original de una fuente archivada (sin el prefijo de huella)."""
    name = Path(path).name
    return re.sub(r"^[0-9a-f]{12}_", "", name)


def issue_message(issue: ReviewIssue) -> str:
    """Explicación para el usuario; los mensajes genéricos se sustituyen por la guía."""
    message = (issue.message or "").strip()
    if not message or message.startswith("Falta el campo requerido"):
        return issue_guidance(issue.field_name)["what_to_find"]
    return message


_DATE_FIELDS = frozenset({"fecha_inicio", "fecha_fin", "fecha_factura"})
_CONFIDENCE_RANK = {"high": 0, "medium": 1, "low": 2}


def issue_suggestion(connection, issue: ReviewIssue) -> str | None:
    """Valor propuesto para rellenar el formulario: el detectado o la mejor evidencia.

    Las evidencias de confianza baja no se aplican solas, pero son un buen punto
    de partida: el usuario sólo tiene que comprobarlas en el documento.
    """
    value = issue.detected_value
    if not value:
        rows = connection.execute(
            """SELECT value,confidence FROM source_field_evidence
               WHERE id_document=? AND field_name=? AND value IS NOT NULL AND trim(value)<>''""",
            (issue.id_document, issue.field_name),
        ).fetchall()
        if rows:
            value = min(rows, key=lambda row: _CONFIDENCE_RANK.get(row["confidence"], 3))["value"]
    if not value:
        return None
    value = str(value).strip()
    if issue.field_name in _DATE_FIELDS:
        try:
            return datetime.strptime(value[:10], "%Y-%m-%d").strftime("%d/%m/%Y")
        except ValueError:
            return value
    return value


def onboarding_summary_data(
    configuration: community_onboarding.OnboardingConfiguration,
) -> dict[str, object]:
    """Project the confirmed configuration into a Tk-independent summary."""
    modules = {
        binding.module: {
            "reading_column": binding.column,
            "meter": binding.meter,
            "date": binding.date,
            "value": binding.value,
            "source_sha256s": tuple(binding.source_sha256s),
        }
        for binding in configuration.reading_bindings
    }
    invoices = [
        {
            "source_sha256": decision.source_sha256,
            "provider": decision.provider,
            "period": decision.period,
            "amount": decision.amount,
            "concept": decision.concept,
        }
        for decision in configuration.invoice_decisions
    ]
    active_modules = tuple(configuration.active_modules)
    bindings_complete = (
        len(modules) == len(configuration.reading_bindings) == len(active_modules)
        and set(modules) == set(active_modules)
        and all(
            all(str(module_data[field]).strip() for field in (
                "reading_column", "meter", "date", "value"
            ))
            and bool(module_data["source_sha256s"])
            and all(
                str(source_sha256).strip()
                for source_sha256 in module_data["source_sha256s"]
            )
            for module_data in modules.values()
        )
    )
    invoice_source_sha256s = tuple(
        sha256
        for kind, _name, sha256 in configuration.source_traces
        if kind == "invoice_pdf"
    )
    invoices_complete = (
        len(invoices) == len(invoice_source_sha256s)
        and sorted(invoice["source_sha256"] for invoice in invoices)
        == sorted(invoice_source_sha256s)
        and all(
            all(str(invoice[field]).strip() for field in (
                "source_sha256", "provider", "period", "amount", "concept"
            ))
            for invoice in invoices
        )
    )
    return {
        "service_decision": configuration.service_decision,
        "modules": modules,
        "invoices": invoices,
        "ready_to_publish": bool(configuration.service_decision.strip())
        and bindings_complete
        and invoices_complete,
    }


def onboarding_step_route(state: Mapping[str, Any]) -> str:
    """Return the first onboarding stage that is safe to show next."""
    if not state.get("identity_valid", False):
        return "identity"
    if not state.get("has_sources", False):
        return "sources"
    if not state.get("analysis_complete", False):
        return "sources"

    step = state.get("step", "identity")
    required_answers_missing = (
        state.get("has_required_questions", False)
        and not state.get("answers_complete", False)
    )
    if step in {"detection", "detected", "confirmations", "summary"}:
        if required_answers_missing:
            return "confirmations"
        if step in {"confirmations", "summary"}:
            return "summary"
        return "confirmations"
    return "detection"


def open_community_onboarding_dialog(app: "AppGestionFincas") -> None:
    """Open the five-stage v3 assistant and delegate all domain work."""
    dialog = _dialog(app, "Alta guiada de comunidad", 860, 680)
    dialog.minsize(760, 620)
    dialog.resizable(True, True)

    identity = {
        "code": tk.StringVar(master=dialog),
        "name": tk.StringVar(master=dialog),
        "period_name": tk.StringVar(master=dialog),
        "start_date": tk.StringVar(master=dialog),
        "end_date": tk.StringVar(master=dialog),
    }
    selected: dict[str, Any] = {
        "owners": None,
        "readings": [],
        "invoices": [],
    }
    source_labels = {
        "owners": tk.StringVar(master=dialog, value="Ningún archivo seleccionado"),
        "readings": tk.StringVar(master=dialog, value="Ningún archivo seleccionado"),
        "invoices": tk.StringVar(master=dialog, value="Opcional · ningún archivo"),
    }
    state: dict[str, Any] = {
        "step": "identity",
        "identity_valid": False,
        "has_sources": False,
        "analysis_complete": False,
        "has_required_questions": False,
        "answers_complete": False,
        "draft": None,
        "answers": {},
        "configuration": None,
        "analysis_generation": 0,
    }
    answer_variables: dict[str, tk.Variable] = {}
    busy = {"active": False}

    def invalidate_analysis():
        state["analysis_generation"] += 1
        state["draft"] = None
        state["analysis_complete"] = False
        state["has_required_questions"] = False
        state["answers_complete"] = False
        state["answers"] = {}
        state["configuration"] = None
        answer_variables.clear()
        state["has_sources"] = bool(selected["owners"] and selected["readings"])

    def identity_changed(*_args):
        state["identity_valid"] = False
        invalidate_analysis()

    for variable in identity.values():
        variable.trace_add("write", identity_changed)

    shell = ctk.CTkFrame(
        dialog,
        fg_color=C["panel"],
        corner_radius=18,
        border_width=1,
        border_color=C["borde"],
    )
    shell.pack(fill="both", expand=True, padx=18, pady=18)

    heading = ctk.CTkFrame(shell, fg_color="transparent")
    heading.pack(fill="x", padx=26, pady=(24, 12))
    ctk.CTkLabel(
        heading,
        text="Alta guiada desde fuentes",
        font=UIM.fuente(22, "bold"),
        text_color=C["texto"],
    ).pack(anchor="w")
    ctk.CTkLabel(
        heading,
        text="Crea la comunidad sin preparar antes un Excel maestro.",
        font=UIM.fuente(11),
        text_color=C["texto_sec"],
    ).pack(anchor="w", pady=(2, 0))

    tracker = ctk.CTkFrame(shell, fg_color=C["panel_2"], corner_radius=12)
    tracker.pack(fill="x", padx=26, pady=(0, 14))
    stage_order = (
        ("identity", "Identidad"),
        ("sources", "Fuentes"),
        ("detection", "Detección"),
        ("confirmations", "Confirmaciones"),
        ("summary", "Resumen"),
    )
    stage_labels = {}
    for index, (key, label) in enumerate(stage_order, start=1):
        item = ctk.CTkLabel(
            tracker,
            text=f"{index:02d}  {label}",
            font=UIM.fuente(10, "bold"),
            text_color=C["texto_sec"],
        )
        item.pack(side="left", expand=True, padx=6, pady=11)
        stage_labels[key] = item

    divider = ctk.CTkFrame(shell, height=1, fg_color=C["borde"])
    divider.pack(fill="x", padx=26)
    content = ctk.CTkScrollableFrame(
        shell, fg_color="transparent", corner_radius=0
    )
    content.pack(fill="both", expand=True, padx=26, pady=(12, 4))
    footer = ctk.CTkFrame(shell, fg_color="transparent", height=54)
    footer.pack(fill="x", padx=26, pady=(4, 20))

    def close_dialog():
        if busy["active"]:
            messagebox.showwarning(
                "Operación en curso",
                "Espera a que termine la operación antes de cerrar el asistente.",
                parent=dialog,
            )
            return
        dialog.destroy()

    dialog.protocol("WM_DELETE_WINDOW", close_dialog)

    def clear(frame):
        for child in frame.winfo_children():
            child.destroy()

    def section_title(title: str, description: str):
        ctk.CTkLabel(
            content,
            text=title,
            font=UIM.fuente(18, "bold"),
            text_color=C["texto"],
        ).pack(anchor="w", pady=(2, 2))
        ctk.CTkLabel(
            content,
            text=description,
            font=UIM.fuente(11),
            text_color=C["texto_sec"],
            justify="left",
            wraplength=720,
        ).pack(anchor="w", pady=(0, 14))

    def entry_field(label: str, variable: tk.StringVar, placeholder: str):
        ctk.CTkLabel(
            content,
            text=label,
            font=UIM.fuente(11, "bold"),
            text_color=C["texto_sec"],
        ).pack(anchor="w", pady=(8, 4))
        entry = ctk.CTkEntry(
            content,
            textvariable=variable,
            placeholder_text=placeholder,
            height=38,
            corner_radius=9,
            border_color=C["borde"],
            fg_color=C["panel_2"],
            text_color=C["texto"],
            font=UIM.fuente(12),
        )
        entry.pack(fill="x")
        return entry

    def footer_buttons(*, back=None, next_text=None, next_command=None):
        clear(footer)
        ctk.CTkButton(
            footer,
            text="Cancelar",
            command=close_dialog,
            height=38,
            corner_radius=9,
            fg_color="transparent",
            border_width=1,
            border_color=C["borde"],
            hover_color=C["acento_suave"],
            text_color=C["texto_sec"],
        ).pack(side="left")
        if next_command is not None:
            ctk.CTkButton(
                footer,
                text=next_text,
                command=next_command,
                height=38,
                corner_radius=9,
                font=UIM.fuente(12, "bold"),
                fg_color=C["primario"],
                hover_color=C["primario_hover"],
            ).pack(side="right")
        if back is not None:
            ctk.CTkButton(
                footer,
                text="Atrás",
                command=back,
                height=38,
                corner_radius=9,
                fg_color=C["acento_suave"],
                hover_color=C["acento_suave_hover"],
                text_color=C["primario"],
            ).pack(side="right", padx=(0, 8))

    def parse_period():
        values = tuple(identity[key].get().strip() for key in (
            "period_name", "start_date", "end_date"
        ))
        if not any(values):
            return None, None, None
        if not all(values):
            raise ValueError(
                "Para crear el expediente inicial, completa nombre y ambas fechas."
            )
        try:
            start = datetime.strptime(values[1], "%d/%m/%Y").date()
            end = datetime.strptime(values[2], "%d/%m/%Y").date()
        except ValueError as exc:
            raise ValueError("Usa DD/MM/AAAA en las dos fechas.") from exc
        if end < start:
            raise ValueError("La fecha final no puede ser anterior a la inicial.")
        return values[0], start, end

    def validate_identity():
        if busy["active"]:
            return
        if not identity["code"].get().strip() or not identity["name"].get().strip():
            messagebox.showwarning(
                "Identidad incompleta",
                "Indica el código y el nombre de la comunidad.",
                parent=dialog,
            )
            return
        try:
            parse_period()
        except ValueError as exc:
            messagebox.showwarning("Período incompleto", str(exc), parent=dialog)
            return
        state["identity_valid"] = True
        render("sources")

    def select_owners():
        if busy["active"]:
            return
        path = filedialog.askopenfilename(
            parent=dialog,
            title="Selecciona la lista de propietarios",
            filetypes=(("Listado Excel o CSV", "*.csv *.xlsx *.xls"),),
        )
        if path:
            set_owner_list(Path(path))

    def set_owner_list(path: Path) -> bool:
        """Acepta listados en Excel o CSV de cualquier programa (se normalizan)."""
        try:
            canonical = owner_lists.normalise_owner_list(
                path, Path(app.ruta_bd_expedientes).parent / "importaciones",
            )
        except (ValueError, OSError, ImportError) as error:
            messagebox.showwarning("Listado de propietarios", str(error), parent=dialog)
            return False
        selected["owners"] = canonical
        source_labels["owners"].set(
            path.name if canonical == path else f"{path.name} (convertido a formato estándar)"
        )
        invalidate_analysis()
        return True

    def fill_from_folder():
        if busy["active"]:
            return
        folder = filedialog.askdirectory(parent=dialog, title="Carpeta con los documentos de la comunidad")
        if not folder:
            return
        busy["active"] = True
        app._estado("Analizando la carpeta de la comunidad…", procesando=True)

        def work():
            try:
                proposal = community_folder.propose_from_folder(
                    Path(folder),
                    progress=lambda name: app.after(0, lambda: app._estado(f"Analizando {name}…", procesando=True)),
                )
            except Exception as error:
                message = str(error)

                def failed():
                    busy["active"] = False
                    app._estado("Listo", procesando=False)
                    messagebox.showerror("No se pudo analizar la carpeta", message, parent=dialog)
                app.after(0, failed)
                return

            def apply():
                busy["active"] = False
                app._estado("Listo", procesando=False)
                if proposal.code:
                    identity["code"].set(proposal.code)
                if proposal.name:
                    identity["name"].set(proposal.name)
                if proposal.start and proposal.end:
                    start = datetime.strptime(proposal.start, "%Y-%m-%d").date()
                    end = datetime.strptime(proposal.end, "%Y-%m-%d").date()
                    identity["start_date"].set(start.strftime("%d/%m/%Y"))
                    identity["end_date"].set(end.strftime("%d/%m/%Y"))
                    identity["period_name"].set(period_selection.default_case_name(start, end))
                if proposal.owners:
                    set_owner_list(proposal.owners)
                if proposal.readings:
                    selected["readings"] = list(proposal.readings)
                    source_labels["readings"].set(
                        f"{len(proposal.readings)} archivo(s): "
                        + ", ".join(path.name for path in proposal.readings)
                    )
                if proposal.invoices:
                    selected["invoices"] = list(proposal.invoices)
                    source_labels["invoices"].set(
                        f"{len(proposal.invoices)} factura(s): "
                        + ", ".join(path.name for path in proposal.invoices)
                    )
                invalidate_analysis()
                render("identity")
                detail = proposal.summary
                if proposal.ignored:
                    detail += "\n\n" + "\n".join(
                        f"· {path.name}: {reason}" for path, reason in proposal.ignored[:12]
                    )
                messagebox.showinfo(
                    "Datos detectados en la carpeta",
                    detail + "\n\nRevisa los datos y continúa: nada se guarda hasta el final.",
                    parent=dialog,
                )
            app.after(0, apply)

        threading.Thread(target=work, daemon=True).start()

    def select_readings():
        if busy["active"]:
            return
        paths = filedialog.askopenfilenames(
            parent=dialog,
            title="Selecciona una o más lecturas",
            filetypes=(
                ("Lecturas Excel o PDF", "*.xlsx *.xls *.pdf"),
                ("Excel", "*.xlsx *.xls"),
                ("PDF", "*.pdf"),
            ),
        )
        if paths:
            selected["readings"] = [Path(path) for path in paths]
            source_labels["readings"].set(
                f"{len(paths)} archivo(s): " + ", ".join(Path(path).name for path in paths)
            )
            invalidate_analysis()

    def select_invoices():
        if busy["active"]:
            return
        paths = filedialog.askopenfilenames(
            parent=dialog,
            title="Selecciona facturas PDF opcionales",
            filetypes=(("Facturas PDF", "*.pdf"),),
        )
        if paths:
            selected["invoices"] = [Path(path) for path in paths]
            source_labels["invoices"].set(
                f"{len(paths)} archivo(s): " + ", ".join(Path(path).name for path in paths)
            )
            invalidate_analysis()

    def clear_invoices():
        if busy["active"]:
            return
        selected["invoices"] = []
        source_labels["invoices"].set("Opcional · ningún archivo")
        invalidate_analysis()

    def source_row(title: str, help_text: str, variable, command, button_text: str):
        row = ctk.CTkFrame(
            content,
            fg_color=C["panel_2"],
            corner_radius=10,
            border_width=1,
            border_color=C["borde"],
        )
        row.pack(fill="x", pady=6)
        text = ctk.CTkFrame(row, fg_color="transparent")
        text.pack(side="left", fill="both", expand=True, padx=14, pady=11)
        ctk.CTkLabel(
            text, text=title, font=UIM.fuente(12, "bold"), text_color=C["texto"]
        ).pack(anchor="w")
        ctk.CTkLabel(
            text,
            text=help_text,
            font=UIM.fuente(10),
            text_color=C["texto_sec"],
        ).pack(anchor="w")
        ctk.CTkLabel(
            text,
            textvariable=variable,
            font=UIM.fuente(10),
            text_color=C["primario"],
            wraplength=490,
            justify="left",
        ).pack(anchor="w", pady=(3, 0))
        ctk.CTkButton(
            row,
            text=button_text,
            command=command,
            width=132,
            height=34,
            corner_radius=8,
            fg_color=C["acento_suave"],
            hover_color=C["acento_suave_hover"],
            text_color=C["primario"],
        ).pack(side="right", padx=14)

    def begin_analysis():
        if busy["active"]:
            return
        if not state["identity_valid"]:
            render("identity")
            return
        state["has_sources"] = bool(selected["owners"] and selected["readings"])
        if not state["has_sources"]:
            messagebox.showwarning(
                "Faltan fuentes mínimas",
                "Selecciona la lista de propietarios y al menos una lectura Excel o PDF.",
                parent=dialog,
            )
            return
        if getattr(app, "_procesando", False):
            messagebox.showwarning(
                "Espera", "Ya hay una operación en curso.", parent=dialog
            )
            return
        invalidate_analysis()
        generation = state["analysis_generation"]
        busy["active"] = True
        community_code = identity["code"].get().strip()
        community_name = identity["name"].get().strip()
        owner_path = selected["owners"]
        reading_paths = tuple(selected["readings"])
        invoice_paths = tuple(selected["invoices"])
        render("detection")
        app._estado("Analizando fuentes del alta…", procesando=True)

        def work():
            try:
                draft = community_onboarding.analyse_sources(
                    community_code=community_code,
                    community_name=community_name,
                    owner_list_path=owner_path,
                    reading_paths=reading_paths,
                    invoice_paths=invoice_paths,
                    project_root=Path(__file__).resolve().parent.parent,
                )
            except Exception as exc:
                def failed(error=exc):
                    busy["active"] = False
                    render("sources" if state["identity_valid"] else "identity")
                    if generation != state["analysis_generation"]:
                        return
                    messagebox.showerror(
                        "No se pudieron analizar las fuentes", str(error), parent=dialog
                    )
                app.after(0, failed)
                return

            def analysed():
                busy["active"] = False
                if generation != state["analysis_generation"]:
                    render("sources" if state["identity_valid"] else "identity")
                    return
                state["draft"] = draft
                state["analysis_complete"] = True
                state["has_required_questions"] = any(
                    question.required for question in draft.questions
                )
                answer_variables.clear()
                for question in draft.questions:
                    answer_variables[question.key] = tk.StringVar(master=dialog)
                state["answers_complete"] = not state["has_required_questions"]
                render("detection")
            app.after(0, analysed)

        app._en_hilo(work)

    def collect_answers():
        if busy["active"]:
            return
        if not state["analysis_complete"] or state["draft"] is None:
            render(onboarding_step_route(state))
            return
        draft = state["draft"]
        service_variable = answer_variables.get("service")
        active_modules = service_variable.get().split("+") if service_variable else draft.detected_modules
        answers = {
            key: variable.get() for key, variable in answer_variables.items()
            if not (
                key.startswith("reading_") and ":" in key
                and key.rsplit(":", 1)[1] not in active_modules
            )
        }
        state["step"] = "confirmations"
        try:
            configuration = community_onboarding.resolve_onboarding_configuration(
                draft, answers
            )
        except ValueError as error:
            state["answers"] = {}
            state["configuration"] = None
            state["answers_complete"] = False
            messagebox.showwarning(
                "Confirmaciones pendientes",
                str(error),
                parent=dialog,
            )
            return
        state["answers"] = configuration
        state["configuration"] = configuration
        state["answers_complete"] = True
        render(onboarding_step_route(state))

    def publish():
        if busy["active"]:
            return
        state["step"] = "summary"
        if onboarding_step_route(state) != "summary":
            render(onboarding_step_route(state))
            return
        if getattr(app, "_procesando", False):
            messagebox.showwarning(
                "Espera", "Ya hay una operación en curso.", parent=dialog
            )
            return
        if not onboarding_summary_data(state["configuration"])["ready_to_publish"]:
            state["answers_complete"] = False
            messagebox.showwarning(
                "Confirmaciones pendientes",
                "Completa las decisiones de lectura y factura antes de publicar.",
                parent=dialog,
            )
            render("confirmations")
            return
        try:
            period_name, start_date, end_date = parse_period()
        except ValueError as exc:
            messagebox.showwarning("Período incompleto", str(exc), parent=dialog)
            return
        busy["active"] = True
        draft = state["draft"]
        answers = state["configuration"]
        clear(footer)
        ctk.CTkLabel(
            footer,
            text=("Creando comunidad y archivando las fuentes…" if period_name
                  else "Creando comunidad y perfil local…"),
            font=UIM.fuente(11, "bold"),
            text_color=C["primario"],
        ).pack(side="right", pady=10)
        app._estado("Creando comunidad y perfil local…", procesando=True)

        def work():
            connection = None
            try:
                connection = gestor_bd.conectar(str(app.ruta_bd_expedientes))
                result = community_onboarding.confirm_onboarding(
                    connection,
                    draft=draft,
                    answers=answers,
                    project_root=Path(__file__).resolve().parent.parent,
                    archive_root=Path(app.ruta_archivo_expedientes),
                    period_name=period_name,
                    start_date=start_date,
                    end_date=end_date,
                    actor="usuario_local",
                )
            except Exception as exc:
                def failed(error=exc):
                    busy["active"] = False
                    render("summary")
                    messagebox.showerror(
                        "No se pudo crear la comunidad", str(error), parent=dialog
                    )
                app.after(0, failed)
                return
            finally:
                if connection is not None:
                    connection.close()

            def completed():
                busy["active"] = False
                code = draft.community_code
                name = draft.community_name
                dialog.destroy()
                app._cargar_comunidades()
                option = f"{code} — {name}"
                if option in getattr(app, "_ids_comunidad", {}):
                    app.comunidad_actual.set(option)
                    app.cb_comunidad.set(option)
                    app._on_comunidad_seleccionada()
                if result.case_id is not None:
                    app._refrescar_lista_expedientes(select_case_id=result.case_id)
                    app._refrescar_expediente()
                app.log(f"Comunidad '{code}' creada desde fuentes verificadas", "ok")
                messagebox.showinfo(
                    "Comunidad creada",
                    "La comunidad y el perfil local se han registrado. "
                    + (f"Se han archivado {len(result.source_document_ids)} fuente(s)."
                       if result.source_document_ids else
                       "No se han archivado fuentes en un expediente."),
                    parent=app,
                )
            app.after(0, completed)

        app._en_hilo(work)

    def render(step: str):
        if busy["active"] and step != "detection":
            return
        if step in {"detection", "confirmations", "summary"} and not busy["active"]:
            if not state["identity_valid"]:
                step = "identity"
            elif not state["analysis_complete"] or state["draft"] is None:
                step = "sources"
        state["step"] = step
        clear(content)
        clear(footer)
        for key, label in stage_labels.items():
            label.configure(
                text_color=C["primario"] if key == step else C["texto_sec"]
            )

        if step == "identity":
            section_title(
                "Identifica la comunidad",
                "El período inicial es opcional; si lo indicas, se creará también su expediente.",
            )
            shortcut = ctk.CTkFrame(content, fg_color=C["acento_suave"], corner_radius=10)
            shortcut.pack(fill="x", pady=(0, 6))
            ctk.CTkLabel(
                shortcut,
                text=(
                    "¿Tienes los documentos de la comunidad en una carpeta? Elígela y se "
                    "rellenarán código, nombre, período, propietarios, lecturas y facturas."
                ),
                font=UIM.fuente(11), text_color=C["primario"], wraplength=520, justify="left",
            ).pack(side="left", padx=12, pady=10)
            ctk.CTkButton(
                shortcut, text="Rellenar desde carpeta…", command=fill_from_folder,
                height=34, corner_radius=8, font=UIM.fuente(11, "bold"),
                fg_color=C["primario"], hover_color=C["primario_hover"],
            ).pack(side="right", padx=12, pady=10)
            first = entry_field("Código de comunidad", identity["code"], "Ej. 101")
            entry_field("Nombre", identity["name"], "Nombre completo de la comunidad")
            entry_field("Nombre del período (opcional)", identity["period_name"], "Ej. 2026")
            dates = ctk.CTkFrame(content, fg_color="transparent")
            dates.pack(fill="x")
            for column, (key, label) in enumerate((
                ("start_date", "Fecha inicial"), ("end_date", "Fecha final")
            )):
                dates.grid_columnconfigure(column, weight=1)
                field = ctk.CTkFrame(dates, fg_color="transparent")
                field.grid(row=0, column=column, sticky="ew", padx=(0, 6) if column == 0 else (6, 0))
                ctk.CTkLabel(
                    field, text=label, font=UIM.fuente(11, "bold"), text_color=C["texto_sec"]
                ).pack(anchor="w", pady=(8, 4))
                ctk.CTkEntry(
                    field,
                    textvariable=identity[key],
                    placeholder_text="DD/MM/AAAA",
                    height=38,
                    corner_radius=9,
                    border_color=C["borde"],
                    fg_color=C["panel_2"],
                ).pack(fill="x")
            footer_buttons(next_text="Continuar a fuentes", next_command=validate_identity)
            first.focus_set()
            return

        if step == "sources":
            section_title(
                "Selecciona las fuentes",
                "Propietarios y lecturas son obligatorios. Las facturas ayudan a detectar conceptos, pero son opcionales.",
            )
            source_row(
                "Lista de propietarios",
                "Excel o CSV con viviendas, propietarios y coeficientes (cualquier programa).",
                source_labels["owners"], select_owners, "Elegir archivo",
            )
            source_row(
                "Lecturas",
                "Una o más lecturas; Excel y PDF son alternativas válidas.",
                source_labels["readings"], select_readings, "Elegir lecturas",
            )
            source_row(
                "Facturas PDF",
                "Opcionales. Puedes seleccionar varias facturas.",
                source_labels["invoices"], select_invoices, "Elegir facturas",
            )
            ctk.CTkButton(
                content, text="Quitar facturas", command=clear_invoices,
                height=30, corner_radius=8, fg_color=C["acento_suave"],
                hover_color=C["acento_suave_hover"], text_color=C["primario"],
            ).pack(anchor="e", pady=(0, 6))
            footer_buttons(
                back=lambda: render("identity"),
                next_text="Analizar fuentes",
                next_command=begin_analysis,
            )
            return

        if step == "detection":
            draft = state.get("draft")
            if draft is None:
                section_title(
                    "Analizando fuentes",
                    "Se revisan estructura, servicios y posibles ambigüedades sin guardar nada todavía.",
                )
                ctk.CTkLabel(
                    content,
                    text="La operación continúa en segundo plano.",
                    font=UIM.fuente(12, "bold"),
                    text_color=C["primario"],
                ).pack(anchor="w", pady=18)
                return
            section_title(
                "Propuesta detectada",
                "Revisa lo encontrado antes de decidir qué formará parte del perfil local.",
            )
            modules = ", ".join(draft.detected_modules) or "Ningún servicio detectado"
            facts = (
                ("Fuentes revisadas", str(len(draft.sources))),
                ("Servicios detectados", modules),
                ("Confirmaciones obligatorias", str(sum(q.required for q in draft.questions))),
            )
            for label, value in facts:
                row = ctk.CTkFrame(content, fg_color="transparent")
                row.pack(fill="x", pady=7)
                ctk.CTkLabel(
                    row, text=label, font=UIM.fuente(11), text_color=C["texto_sec"]
                ).pack(side="left")
                ctk.CTkLabel(
                    row, text=value, font=UIM.fuente(11, "bold"), text_color=C["texto"]
                ).pack(side="right")
            footer_buttons(
                back=lambda: render("sources"),
                next_text="Revisar confirmaciones",
                next_command=lambda: render(onboarding_step_route(state)),
            )
            return

        if step == "confirmations":
            draft = state["draft"]
            section_title(
                "Confirma lo detectado",
                "Confirma el servicio y cada dato incierto. No se inventa ningún valor.",
            )
            service = answer_variables.get("service")
            active = service.get().split("+") if service and service.get() else None
            for question in sorted(draft.questions, key=lambda item: item.key != "service"):
                if (active is not None and question.key.startswith("reading_")
                        and ":" in question.key
                        and question.key.rsplit(":", 1)[1] not in active):
                    continue
                ctk.CTkLabel(
                    content,
                    text=question.prompt,
                    font=UIM.fuente(11, "bold"),
                    text_color=C["texto"],
                ).pack(anchor="w", pady=(12, 4))
                if question.candidates:
                    ctk.CTkComboBox(
                        content,
                        variable=answer_variables[question.key],
                        values=list(question.candidates),
                        state="readonly",
                        height=38,
                        border_color=C["borde"],
                        fg_color=C["panel_2"],
                        button_color=C["primario"],
                        command=(lambda _value: render("confirmations"))
                        if question.key == "service" else None,
                    ).pack(fill="x")
                else:
                    ctk.CTkEntry(
                        content,
                        textvariable=answer_variables[question.key],
                        height=38,
                        border_color=C["borde"],
                        fg_color=C["panel_2"],
                    ).pack(fill="x")
            footer_buttons(
                back=lambda: render("detection"),
                next_text="Preparar resumen",
                next_command=collect_answers,
            )
            return

        draft = state["draft"]
        configuration = state["configuration"]
        summary_data = onboarding_summary_data(configuration)
        active_modules = list(summary_data["modules"])
        period_name, _start, _end = parse_period()
        section_title(
            "Revisa y crea la comunidad",
            "Al confirmar se guardarán el perfil y la plantilla locales. "
            + ("Se archivarán copias de las fuentes en el expediente inicial."
               if period_name else
               "Sin período inicial no se archivarán las fuentes; podrás añadirlas a un expediente después."),
        )
        summary = (
            ("Comunidad", f"{draft.community_code} — {draft.community_name}"),
            ("Fuentes que se archivarán", str(len(draft.sources)) if period_name else "0 · sin expediente inicial"),
            ("Servicios confirmados", ", ".join(active_modules) or "Ninguno"),
            ("Expediente inicial", period_name or "No se creará todavía"),
        )
        for label, value in summary:
            row = ctk.CTkFrame(
                content, fg_color=C["panel_2"], corner_radius=9,
                border_width=1, border_color=C["borde"],
            )
            row.pack(fill="x", pady=5)
            ctk.CTkLabel(
                row, text=label, font=UIM.fuente(11), text_color=C["texto_sec"]
            ).pack(side="left", padx=13, pady=10)
            ctk.CTkLabel(
                row, text=value, font=UIM.fuente(11, "bold"), text_color=C["texto"]
            ).pack(side="right", padx=13, pady=10)

        for module, reading in summary_data["modules"].items():
            ctk.CTkLabel(
                content,
                text=f"Lectura {module}",
                font=UIM.fuente(13, "bold"),
                text_color=C["texto"],
            ).pack(anchor="w", pady=(15, 4))
            reading_row = ctk.CTkFrame(
                content, fg_color=C["panel_2"], corner_radius=9,
                border_width=1, border_color=C["borde"],
            )
            reading_row.pack(fill="x", pady=5)
            reading_text = " · ".join((
                f"Contador: {reading['meter']}",
                f"Fecha: {reading['date']}",
                f"Valor: {reading['value']}",
                f"Columna: {reading['reading_column']}",
            ))
            ctk.CTkLabel(
                reading_row,
                text=reading_text,
                font=UIM.fuente(10, "bold"),
                text_color=C["texto"],
                justify="left",
                wraplength=670,
            ).pack(anchor="w", padx=13, pady=10)

        if summary_data["invoices"]:
            ctk.CTkLabel(
                content,
                text="Facturas confirmadas",
                font=UIM.fuente(13, "bold"),
                text_color=C["texto"],
            ).pack(anchor="w", pady=(15, 4))
        for index, invoice in enumerate(summary_data["invoices"], start=1):
            invoice_row = ctk.CTkFrame(
                content, fg_color=C["panel_2"], corner_radius=9,
                border_width=1, border_color=C["borde"],
            )
            invoice_row.pack(fill="x", pady=5)
            invoice_text = " · ".join((
                f"Factura {index}",
                f"Proveedor: {invoice['provider']}",
                f"Período: {invoice['period']}",
                f"Importe: {invoice['amount']}",
                f"Concepto: {invoice['concept']}",
            ))
            ctk.CTkLabel(
                invoice_row,
                text=invoice_text,
                font=UIM.fuente(10, "bold"),
                text_color=C["texto"],
                justify="left",
                wraplength=670,
            ).pack(anchor="w", padx=13, pady=10)

        ready_to_publish = summary_data["ready_to_publish"]
        ctk.CTkLabel(
            content,
            text=("Confirmaciones completas" if ready_to_publish
                  else "Faltan confirmaciones obligatorias"),
            font=UIM.fuente(11, "bold"),
            text_color=C["primario"] if ready_to_publish else C["alerta"],
        ).pack(anchor="w", pady=(12, 2))
        footer_buttons(
            back=lambda: render("confirmations"),
            next_text="Crear comunidad" if ready_to_publish else None,
            next_command=publish if ready_to_publish else None,
        )

    render("identity")


def _dialog(app: "AppGestionFincas", title: str, width: int, height: int):
    prepare = getattr(app, "_preparar_dialogo", None)
    if prepare is not None:
        return prepare(title, width, height)
    dialog = ctk.CTkToplevel(app, fg_color=C["fondo"])
    dialog.title(title)
    dialog.geometry(f"{width}x{height}")
    dialog.resizable(False, False)
    dialog.transient(app)
    dialog.grab_set()
    return dialog


def _field(parent, label: str, row: int, *, placeholder: str = ""):
    ctk.CTkLabel(
        parent,
        text=label.upper(),
        font=UIM.fuente(10, "bold"),
        text_color=C["texto_sec"],
    ).grid(row=row, column=0, sticky="w", padx=22, pady=(13, 4))
    entry = ctk.CTkEntry(
        parent,
        height=38,
        corner_radius=9,
        border_color=C["borde"],
        fg_color=C["panel_2"],
        text_color=C["texto"],
        placeholder_text=placeholder,
        font=UIM.fuente(12),
    )
    entry.grid(row=row + 1, column=0, sticky="ew", padx=22)
    return entry


def _packed_field(parent, label: str, *, placeholder: str = ""):
    """Campo apilado para paneles que ya distribuyen sus hijos con ``pack``."""
    ctk.CTkLabel(
        parent,
        text=label.upper(),
        font=UIM.fuente(10, "bold"),
        text_color=C["texto_sec"],
    ).pack(anchor="w", padx=22, pady=(13, 4))
    entry = ctk.CTkEntry(
        parent,
        height=38,
        corner_radius=9,
        border_color=C["borde"],
        fg_color=C["panel_2"],
        text_color=C["texto"],
        placeholder_text=placeholder,
        font=UIM.fuente(12),
    )
    entry.pack(fill="x", padx=22)
    return entry


def open_office_settings_dialog(app: "AppGestionFincas", *, first_run: bool = False) -> None:
    """Datos del despacho: identidad de las cartas, CIF propio y ejercicio."""
    connection = gestor_bd.conectar(str(app.ruta_bd_expedientes))
    try:
        current = office_settings.suggested_settings(connection)
    finally:
        connection.close()
    project_root = Path(__file__).resolve().parents[1]

    dialog = _dialog(app, "Datos del despacho", 640, 720)
    panel = ctk.CTkScrollableFrame(
        dialog, fg_color=C["panel"], corner_radius=16, border_width=1, border_color=C["borde"],
    )
    panel.pack(fill="both", expand=True, padx=18, pady=(18, 8))
    ctk.CTkLabel(
        panel, text="Bienvenido: configura tu despacho" if first_run else "Datos del despacho",
        font=UIM.fuente(20, "bold"), text_color=C["texto"],
    ).pack(anchor="w", padx=18, pady=(14, 2))
    ctk.CTkLabel(
        panel,
        text=(
            "Se usan en la cabecera, el pie y la firma de las cartas, y para no confundir "
            "el CIF del despacho con el de un proveedor. Se guardan en la base de datos "
            "de esta instalación y puedes cambiarlos en Ajustes."
        ),
        font=UIM.fuente(11), text_color=C["texto_sec"], wraplength=520, justify="left",
    ).pack(anchor="w", padx=18, pady=(0, 8))

    variables: dict[str, ctk.CTkEntry] = {}
    for key, label, placeholder in (
        ("name", "Nombre del despacho *", "Ej. Administración de Fincas García"),
        ("tax_id", "CIF / NIF", "Ej. B12345674"),
        ("city", "Ciudad (aparece junto a la fecha)", "Ej. Valencia"),
        ("address", "Dirección", "Calle, número, código postal"),
        ("phone", "Teléfono", ""),
        ("email", "Correo electrónico", ""),
        ("signature", "Firma de las cartas", "Si lo dejas vacío se usa el nombre del despacho"),
        ("footer", "Pie de las cartas", "Si lo dejas vacío: dirección · teléfono · correo"),
    ):
        ctk.CTkLabel(panel, text=label.upper(), font=UIM.fuente(10, "bold"),
                     text_color=C["texto_sec"]).pack(anchor="w", padx=18, pady=(10, 3))
        # Sin textvariable para que el ejemplo (placeholder) sea visible.
        variables[key] = ctk.CTkEntry(
            panel, height=36, corner_radius=9,
            border_color=C["borde"], fg_color=C["panel_2"], text_color=C["texto"],
            placeholder_text=placeholder, font=UIM.fuente(12),
        )
        variables[key].pack(fill="x", padx=18)
        value = str(getattr(current, key) or "")
        if value:
            variables[key].insert(0, value)

    month_labels = [name.capitalize() for name in period_selection.MONTH_NAMES]
    ctk.CTkLabel(panel, text="EL EJERCICIO EMPIEZA EN", font=UIM.fuente(10, "bold"),
                 text_color=C["texto_sec"]).pack(anchor="w", padx=18, pady=(12, 3))
    month = ctk.CTkComboBox(panel, values=month_labels, state="readonly", width=200)
    month.set(month_labels[current.fiscal_start_month - 1])
    month.pack(anchor="w", padx=18)
    ctk.CTkLabel(
        panel, text="Se usa para proponer las fechas de los expedientes nuevos.",
        font=UIM.fuente(10), text_color=C["texto_sec"],
    ).pack(anchor="w", padx=18, pady=(2, 0))

    logo_state = {"path": current.logo_path, "source": None}
    ctk.CTkLabel(panel, text="LOGO (OPCIONAL)", font=UIM.fuente(10, "bold"),
                 text_color=C["texto_sec"]).pack(anchor="w", padx=18, pady=(12, 3))
    logo_row = ctk.CTkFrame(panel, fg_color="transparent")
    logo_row.pack(fill="x", padx=18, pady=(0, 12))
    logo_label = ctk.CTkLabel(
        logo_row, text=Path(current.logo_path).name if current.logo_path else "Sin logo",
        font=UIM.fuente(11), text_color=C["texto"],
    )

    def choose_logo():
        selected = filedialog.askopenfilename(
            parent=dialog, title="Logo del despacho",
            filetypes=[("Imágenes", "*.png *.jpg *.jpeg")],
        )
        if selected:
            logo_state["source"] = Path(selected)
            logo_label.configure(text=Path(selected).name)

    def remove_logo():
        logo_state.update(path="", source=None)
        logo_label.configure(text="Sin logo")

    ctk.CTkButton(logo_row, text="Elegir imagen…", width=120, height=30, corner_radius=8,
                  command=choose_logo, **UIM.secondary_button_kwargs()).pack(side="left")
    ctk.CTkButton(logo_row, text="Quitar", width=70, height=30, corner_radius=8,
                  command=remove_logo, **UIM.secondary_button_kwargs()).pack(side="left", padx=6)
    logo_label.pack(side="left", padx=8)

    def save():
        try:
            logo_path = logo_state["path"]
            if logo_state["source"] is not None:
                logo_path = office_settings.install_logo(project_root, logo_state["source"])
            settings = office_settings.OfficeSettings(
                **{key: variable.get() for key, variable in variables.items()},
                logo_path=logo_path or "",
                fiscal_start_month=month_labels.index(month.get()) + 1,
            )
            connection = gestor_bd.conectar(str(app.ruta_bd_expedientes))
            try:
                saved = office_settings.save_office_settings(connection, settings)
            finally:
                connection.close()
        except (ValueError, OSError) as error:
            messagebox.showwarning("Datos del despacho", str(error), parent=dialog)
            return
        dialog.destroy()
        app.log(f"Datos del despacho guardados: {saved.name}", "ok")
        refresh = getattr(app, "_actualizar_titulo_despacho", None)
        if refresh:
            refresh()

    buttons = ctk.CTkFrame(dialog, fg_color="transparent")
    buttons.pack(fill="x", padx=20, pady=(0, 14))
    ctk.CTkButton(
        buttons, text="Guardar", command=save, height=36, width=120, corner_radius=8,
        font=UIM.fuente(12, "bold"), fg_color=C["primario"], hover_color=C["primario_hover"],
    ).pack(side="right")
    ctk.CTkButton(
        buttons, text="Más tarde" if first_run else "Cancelar", command=dialog.destroy,
        height=36, width=110, corner_radius=8, **UIM.secondary_button_kwargs(),
    ).pack(side="right", padx=8)


def _existing_period_cases(connection, community_id: int, exclude_case_id: int | None):
    return [
        period_selection.ExistingCase(case.name, case.start_date, case.end_date)
        for case in expedient_service.list_cases(connection, community_id)
        if case.id_case != exclude_case_id
    ]


def open_create_case_dialog(app: "AppGestionFincas") -> None:
    open_case_period_dialog(app)


def open_case_period_dialog(app: "AppGestionFincas", case_id: int | None = None) -> None:
    """Crea un expediente o cambia sus fechas desde una única ventana.

    Ofrece atajos (continuar el último expediente, ejercicio, año, trimestre,
    mes), acepta fechas escritas de varias formas, propone un nombre y avisa
    en vivo de solapes o duraciones sospechosas antes de guardar.
    """
    if not getattr(app, "id_comunidad", None):
        messagebox.showwarning(
            "Comunidad requerida",
            "Selecciona una comunidad antes de crear un expediente.",
        )
        return
    connection = gestor_bd.conectar(str(app.ruta_bd_expedientes))
    try:
        current = expedient_service.get_case(connection, case_id) if case_id else None
        existing = _existing_period_cases(connection, app.id_comunidad, case_id)
        office = office_settings.suggested_settings(connection)
    finally:
        connection.close()
    editing = current is not None

    dialog = _dialog(app, "Cambiar fechas" if editing else "Nuevo expediente", 560, 640)
    content = ctk.CTkFrame(
        dialog, fg_color=C["panel"], corner_radius=16, border_width=1, border_color=C["borde"],
    )
    content.pack(fill="both", expand=True, padx=18, pady=18)
    ctk.CTkLabel(
        content,
        text="Cambiar fechas del expediente" if editing else "Nuevo expediente",
        font=UIM.fuente(20, "bold"), text_color=C["texto"],
    ).pack(anchor="w", padx=22, pady=(20, 1))
    ctk.CTkLabel(
        content,
        text=(
            "Si cambian las fechas, el expediente vuelve a revisión para comprobar qué "
            "facturas y lecturas entran en el nuevo intervalo."
            if editing else
            "Elige un atajo o escribe las fechas del intervalo que vas a regularizar."
        ),
        font=UIM.fuente(11), text_color=C["texto_sec"], wraplength=490, justify="left",
    ).pack(anchor="w", padx=22, pady=(0, 8))

    start_var = tk.StringVar(value=period_selection.format_date(current.start_date) if editing else "")
    end_var = tk.StringVar(value=period_selection.format_date(current.end_date) if editing else "")
    name_var = tk.StringVar(value=current.name if editing else "")
    name_state = {"auto": not editing}

    presets = period_selection.period_presets(
        date.today(), existing, fiscal_start_month=office.fiscal_start_month,
    )
    ctk.CTkLabel(
        content, text="ATAJOS", font=UIM.fuente(10, "bold"), text_color=C["texto_sec"],
    ).pack(anchor="w", padx=22, pady=(4, 4))
    chips = ctk.CTkFrame(content, fg_color="transparent")
    chips.pack(fill="x", padx=18)
    chips.grid_columnconfigure((0, 1), weight=1)

    def apply_preset(preset):
        start_var.set(period_selection.format_date(preset.start))
        end_var.set(period_selection.format_date(preset.end))

    for index, preset in enumerate(presets):
        ctk.CTkButton(
            chips, text=preset.label, command=lambda item=preset: apply_preset(item),
            height=30, corner_radius=8, font=UIM.fuente(10),
            **UIM.secondary_button_kwargs(),
        ).grid(row=index // 2, column=index % 2, sticky="ew", padx=4, pady=3)

    dates = ctk.CTkFrame(content, fg_color="transparent")
    dates.pack(fill="x", padx=22, pady=(10, 0))
    dates.grid_columnconfigure((0, 1), weight=1)
    entries = {}
    for column, (label, variable) in enumerate((("Fecha inicial", start_var), ("Fecha final", end_var))):
        ctk.CTkLabel(
            dates, text=label.upper(), font=UIM.fuente(10, "bold"), text_color=C["texto_sec"],
        ).grid(row=0, column=column, sticky="w", padx=(0 if column == 0 else 8, 0), pady=(0, 4))
        entry = ctk.CTkEntry(
            dates, textvariable=variable, height=38, corner_radius=9,
            border_color=C["borde"], fg_color=C["panel_2"], text_color=C["texto"],
            placeholder_text="DD/MM/AAAA", font=UIM.fuente(12),
        )
        entry.grid(row=1, column=column, sticky="ew", padx=(0 if column == 0 else 8, 0))
        entries[label] = entry
    ctk.CTkLabel(
        content, text="También vale 1/9/25, 2025-09-01 o «1 septiembre 2025».",
        font=UIM.fuente(10), text_color=C["texto_sec"],
    ).pack(anchor="w", padx=22, pady=(3, 0))

    ctk.CTkLabel(
        content, text="NOMBRE", font=UIM.fuente(10, "bold"), text_color=C["texto_sec"],
    ).pack(anchor="w", padx=22, pady=(12, 4))
    name_entry = ctk.CTkEntry(
        content, textvariable=name_var, height=38, corner_radius=9,
        border_color=C["borde"], fg_color=C["panel_2"], text_color=C["texto"],
        placeholder_text="Se propone a partir de las fechas", font=UIM.fuente(12),
    )
    name_entry.pack(fill="x", padx=22)
    name_entry.bind("<Key>", lambda _event: name_state.update(auto=False))

    summary = ctk.CTkLabel(
        content, text="", font=UIM.fuente(11, "bold"), text_color=C["primario"],
        wraplength=490, justify="left", anchor="w",
    )
    summary.pack(fill="x", padx=22, pady=(12, 0))
    notes = ctk.CTkLabel(
        content, text="", font=UIM.fuente(10), text_color=C["aviso"],
        wraplength=490, justify="left", anchor="w",
    )
    notes.pack(fill="x", padx=22, pady=(2, 0))
    parsed: dict[str, object] = {}

    def refresh(*_args):
        parsed.clear()
        try:
            end_date = period_selection.parse_user_date(end_var.get(), reference_year=date.today().year)
            try:
                start_date = period_selection.parse_user_date(start_var.get())
            except ValueError:
                # «1/9» sin año: el del final, o el anterior si quedaría después.
                start_date = period_selection.parse_user_date(
                    start_var.get(), reference_year=end_date.year,
                )
                if start_date > end_date:
                    start_date = start_date.replace(year=start_date.year - 1)
        except ValueError as error:
            has_text = start_var.get().strip() or end_var.get().strip()
            summary.configure(text=str(error) if has_text else "", text_color=C["texto_sec"])
            notes.configure(text="")
            save_button.configure(state="disabled")
            return
        errors, warnings = period_selection.validate_range(start_date, end_date, existing)
        if errors:
            summary.configure(text=errors[0], text_color=C["alerta"])
            notes.configure(text="")
            save_button.configure(state="disabled")
            return
        parsed.update(start=start_date, end=end_date)
        summary.configure(
            text=period_selection.describe_range(start_date, end_date), text_color=C["primario"],
        )
        notes.configure(text="\n".join(f"⚠ {item}" for item in warnings))
        if name_state["auto"]:
            name_var.set(period_selection.default_case_name(start_date, end_date))
        save_button.configure(state="normal")

    actions = ctk.CTkFrame(content, fg_color="transparent")
    actions.pack(side="bottom", fill="x", padx=22, pady=(12, 18))
    ctk.CTkButton(
        actions, text="Cancelar", command=dialog.destroy, height=38, corner_radius=9,
        fg_color="transparent", border_width=1, border_color=C["borde"],
        hover_color=C["acento_suave"], text_color=C["texto_sec"],
    ).pack(side="right")

    def save():
        if "start" not in parsed:
            return
        start_date, end_date = parsed["start"], parsed["end"]
        name = name_var.get().strip() or period_selection.default_case_name(start_date, end_date)
        connection = gestor_bd.conectar(str(app.ruta_bd_expedientes))
        try:
            if editing:
                case = expedient_service.update_case_dates(
                    connection, current.id_case, name=name,
                    start_date=start_date, end_date=end_date,
                )
            else:
                case = expedient_service.create_case(
                    connection, app.id_comunidad, name=name,
                    start_date=start_date, end_date=end_date,
                )
        except ValueError as error:
            messagebox.showwarning("No se pudo guardar", str(error), parent=dialog)
            return
        finally:
            connection.close()
        app.id_expediente = case.id_case
        app.expediente_actual.set(case.name)
        dialog.destroy()
        app._refrescar_lista_expedientes(select_case_id=case.id_case)
        app._refrescar_expediente()
        duration = (end_date - start_date).days + 1
        if not editing:
            app.log(f"Expediente '{case.name}' creado · {duration} día(s) incluidos", "ok")
            return
        app.log(f"Fechas de '{case.name}' actualizadas · {duration} día(s)", "ok")
        dates_changed = (current.start_date, current.end_date) != (start_date, end_date)
        if dates_changed and getattr(app, "_accion_reanalizar_fuentes", None) and messagebox.askyesno(
            "Fechas actualizadas",
            "¿Quieres reevaluar ahora las fuentes para comprobar qué facturas y "
            "lecturas entran en el nuevo intervalo?",
            parent=app,
        ):
            app._accion_reanalizar_fuentes()

    save_button = ctk.CTkButton(
        actions, text="Guardar fechas" if editing else "Crear expediente", command=save,
        height=38, corner_radius=9, font=UIM.fuente(12, "bold"),
        fg_color=C["primario"], hover_color=C["primario_hover"], state="disabled",
    )
    save_button.pack(side="right", padx=(0, 8))
    start_var.trace_add("write", refresh)
    end_var.trace_add("write", refresh)
    refresh()
    entries["Fecha inicial"].focus_set()


def open_add_sources_dialog(app: "AppGestionFincas", case_id: int) -> None:
    if not case_id:
        messagebox.showwarning(
            "Expediente requerido",
            "Crea o selecciona un expediente antes de añadir fuentes.",
        )
        return

    dialog = _dialog(app, "Añadir fuentes", 600, 350)
    panel = ctk.CTkFrame(
        dialog,
        fg_color=C["panel"],
        corner_radius=16,
        border_width=1,
        border_color=C["borde"],
    )
    panel.pack(fill="both", expand=True, padx=18, pady=18)
    ctk.CTkLabel(
        panel,
        text="Añade documentos al expediente",
        font=UIM.fuente(20, "bold"),
        text_color=C["texto"],
    ).pack(anchor="w", padx=22, pady=(22, 2))
    ctk.CTkLabel(
        panel,
        text="Los originales se archivan sin cambios y quedan vinculados a la revisión.",
        font=UIM.fuente(11),
        text_color=C["texto_sec"],
    ).pack(anchor="w", padx=22, pady=(0, 18))

    ctk.CTkLabel(
        panel, text="Detectamos facturas, lecturas y propietarios automáticamente.\nSolo tendrás que clasificar las fuentes que no se reconozcan.",
        font=UIM.fuente(11), text_color=C["texto_sec"], justify="left",
    ).pack(fill="x", padx=22)

    def select_files(paths=None):
        if app._procesando:
            return
        if paths is None:
            paths = filedialog.askopenfilenames(
                parent=dialog,
                title="Seleccionar fuentes",
                filetypes=(
                    ("Documentos compatibles", "*.pdf *.xlsx *.xls *.csv"),
                    ("PDF", "*.pdf"),
                    ("Excel", "*.xlsx *.xls"),
                    ("CSV", "*.csv"),
                    ("Todos los archivos", "*.*"),
                ),
            )
        selected_paths = tuple(Path(path) for path in paths)
        paths = tuple(str(path) for path in selected_paths if is_supported_source_file(path))
        ignored_count = len(selected_paths) - len(paths)
        if not paths:
            messagebox.showwarning(
                "Sin fuentes compatibles",
                "Selecciona PDF, XLSX, XLS o CSV que no sean ocultos ni temporales.",
                parent=dialog,
            )
            return
        if ignored_count:
            messagebox.showinfo(
                "Archivos omitidos",
                f"Se han omitido {ignored_count} archivo(s) no compatible(s), oculto(s) o temporal(es).",
                parent=dialog,
            )
        dialog.destroy()
        database_path = str(app.ruta_bd_expedientes)
        archive_root = app.ruta_archivo_expedientes
        community_id = app.id_comunidad

        def add_all():
            created_count = 0
            duplicate_count = 0
            errors = []
            kinds = []
            lookup = gestor_bd.conectar(database_path)
            try:
                case_ingestion.assert_case_belongs_to_community(
                    lookup, case_id, community_id,
                )
                community = lookup.execute(
                    "SELECT codigo FROM comunidades WHERE id_comunidad = ?", (community_id,),
                ).fetchone()
                community_code = community["codigo"]
            finally:
                lookup.close()
            accepted_paths, foreign_paths = community_discovery.partition_sources_for_community(
                (Path(path) for path in paths), community_code,
            )
            if foreign_paths:
                app.log(
                    f"Se han apartado {len(foreign_paths)} fuente(s) que indican otra comunidad. "
                    "Impórtalas desde Bandeja global.",
                    "aviso",
                )
            total = len(accepted_paths)
            phase_labels = {
                "hashing": "Calculando huella",
                "text": "Leyendo texto",
                "classification": "Clasificando",
                "provider": "Identificando proveedor",
                "fields": "Extrayendo campos",
                "completed": "Finalizado",
            }

            def batch_progress(event):
                label = phase_labels.get(event.phase, event.phase)
                filename = f": {event.filename}" if event.filename else ""
                app.after(
                    0,
                    lambda text=f"{label} {event.completed}/{event.total}{filename}":
                    app._estado(text, procesando=True),
                )

            batch_items = source_batch.analyse_batch(
                accepted_paths,
                community_code=community_code,
                database_path=database_path,
                max_workers=3,
                progress=batch_progress,
            )
            for index, item in enumerate(batch_items, start=1):
                source_path = item.path
                path = str(source_path)
                filename = Path(path).name
                connection = None
                app._estado(
                    f"Incorporando {index} de {total}: {filename}",
                    procesando=True,
                )
                app.log(f"Fuente {index} de {total}: {filename}", "info")
                if item.error is not None or item.analysis is None:
                    errors.append((filename, item.error or "Análisis sin resultado"))
                    app.log(
                        f"No se pudo analizar {filename}: {item.error or 'sin resultado'}",
                        "error",
                    )
                    continue
                try:
                    connection = gestor_bd.conectar(database_path)
                    case_ingestion.assert_case_belongs_to_community(connection, case_id, community_id)
                    result = case_ingestion.add_analysed_document_to_case(
                        connection,
                        case_id,
                        source_path=path,
                        archive_root=archive_root,
                        analysis=item.analysis,
                    )
                except Exception as exc:
                    errors.append((filename, str(exc)))
                    app.log(
                        f"No se pudo incorporar {filename}: {exc}",
                        "error",
                    )
                    continue
                finally:
                    if connection is not None:
                        connection.close()
                if result.created:
                    created_count += 1
                else:
                    duplicate_count += 1
                kinds.append(result.document.document_kind)
            final_connection = gestor_bd.conectar(database_path)
            try:
                automatic, ready_for_calculation = case_ingestion.finalize_case_source_intake(
                    final_connection, case_id,
                )
                readiness = case_readiness.evaluate_case_readiness(
                    final_connection,
                    case_id,
                    Path(__file__).resolve().parents[1],
                )
            finally:
                final_connection.close()
            app.log(
                f"Fuentes añadidas: {created_count} nueva(s), "
                f"{duplicate_count} duplicada(s), {len(errors)} con error",
                "ok" if created_count and not errors else "aviso",
            )
            def refresh_case():
                app._refrescar_lista_expedientes(select_case_id=case_id)
                app._refrescar_expediente()
                next_requirement = (
                    readiness.blockers[0].message
                    if readiness.blockers else
                    "El expediente no tiene decisiones pendientes."
                )
                messagebox.showinfo(
                    "Fuentes analizadas",
                    f"{source_summary(kinds)}\n\n{created_count} nuevas · {duplicate_count} duplicadas · {len(errors)} con error\n"
                    + (
                        f"Se aplicaron automáticamente {automatic} fuente(s) completas. "
                        "El expediente está listo para generar el Excel oficial."
                        if ready_for_calculation else
                        f"Siguiente paso: {next_requirement}"
                    )
                    + (
                        f"\n\n{len(foreign_paths)} fuente(s) de otras comunidades se apartaron. "
                        "Usa Bandeja global para clasificarlas."
                        if foreign_paths else ""
                    )
                    + ("\n\n" + "\n".join(f"{name}: {error}" for name, error in errors) if errors else ""),
                    parent=app,
                )

            app.after(0, refresh_case)

        app._estado(f"Incorporando 0 de {len(paths)} fuentes", procesando=True)
        app.after(0, lambda: app._en_hilo(add_all))

    def select_folder():
        folder = filedialog.askdirectory(
            parent=dialog,
            title="Selecciona una carpeta de fuentes",
        )
        if not folder:
            return
        paths = tuple(str(path) for path in source_files_in_folder(Path(folder)))
        if not paths:
            messagebox.showwarning(
                "Sin fuentes compatibles",
                "La carpeta no contiene PDF, XLSX, XLS o CSV visibles.",
                parent=dialog,
            )
            return
        select_files(paths)

    source_actions = ctk.CTkFrame(panel, fg_color="transparent")
    source_actions.pack(anchor="w", padx=22, pady=(22, 8))
    ctk.CTkButton(
        source_actions, text="Añadir archivos", command=select_files,
        height=42, corner_radius=10, font=UIM.fuente(12, "bold"),
        fg_color=C["primario"], hover_color=C["primario_hover"],
        border_width=1, border_color=C["primario"],
    ).pack(side="left", padx=(0, 8))
    ctk.CTkButton(
        source_actions, text="Añadir carpeta", command=select_folder,
        height=42, corner_radius=10, font=UIM.fuente(12, "bold"),
        fg_color="transparent", hover_color=C["acento_suave"],
        border_width=1, border_color=C["borde"], text_color=C["primario"],
    ).pack(side="left")
    ctk.CTkLabel(
        panel,
        text="Puedes seleccionar varios archivos o cargar una carpeta completa · PDF · XLSX · XLS · CSV",
        font=UIM.fuente(10),
        text_color=C["texto_sec"],
    ).pack(anchor="w", padx=22)


def open_confirm_sources_dialog(app: "AppGestionFincas", case_id: int) -> None:
    """Muestra sólo fuentes con dudas después de aplicar las seguras en lote."""
    dialog = _dialog(app, "Confirmar fuentes", 850, 720)
    panel = ctk.CTkScrollableFrame(dialog)
    panel.pack(fill="both", expand=True, padx=16, pady=16)

    def confirm_pending_source(document_id: int) -> None:
        error = None
        connection = gestor_bd.conectar(str(app.ruta_bd_expedientes))
        try:
            result = case_ingestion.confirm_source_candidates(
                connection, case_id, document_id, confirmed_by="usuario_local",
            )
            if result.status != "validated":
                error = "La fuente necesita resolver las nuevas incidencias detectadas antes de continuar."
        except (LookupError, ValueError, RuntimeError, sqlite3.Error) as exc:
            connection.rollback()
            error = str(exc)
            case_ingestion.ensure_pending_source_issues(connection, case_id)
        finally:
            connection.close()
        if error:
            messagebox.showwarning(
                "No se pudo confirmar la fuente", error, parent=dialog,
            )
        else:
            app.log("Fuente confirmada y aplicada al expediente.", "ok")
        render()

    def render():
        ready_for_calculation = False
        for widget in panel.winfo_children():
            widget.destroy()
        connection = gestor_bd.conectar(str(app.ruta_bd_expedientes))
        try:
            automatic, _pending = case_ingestion.auto_apply_case_sources(connection, case_id)
            case_ingestion.ensure_pending_source_issues(connection, case_id)
            issues = document_review.list_open_issues(connection, case_id)
            issue_document_ids = {issue.id_document for issue in issues}
            pending_sources = tuple(
                source for source in case_ingestion.list_pending_sources(connection, case_id)
                if source.id_document not in issue_document_ids
            )
            if not issues and not pending_sources:
                document_review.validate_case_ready(connection, case_id)
                ready_for_calculation = True
            for grouped in group_review_issues(issues):
                issue = grouped["representative"]
                reset_count = int(grouped["count"])
                group = ctk.CTkFrame(panel)
                group.pack(fill="x", pady=8)
                ctk.CTkLabel(
                    group, text=Path(issue.archived_path).name,
                    font=UIM.fuente(12, "bold"), wraplength=740,
                ).pack(anchor="w", padx=12, pady=(10, 2))
                if issue.code == "COUNTER_RESET" and reset_count > 1:
                    message = (
                        f"Se detectaron {reset_count} contadores reiniciados en el mismo informe de período. "
                        "Puedes conservar de una vez la última lectura válida de cada vivienda hasta que "
                        "lleguen lecturas posteriores fiables."
                    )
                    action_text = f"Mantener lecturas previas ({reset_count})"
                    action = lambda selected=issue, count=reset_count: (
                        dialog.destroy(),
                        open_counter_reset_carry_forward_dialog(
                            app, selected.id_case, selected.id_document, count,
                            Path(selected.archived_path).name,
                        ),
                    )
                elif issue.code == "READING_ZERO_REVIEW":
                    message = (
                        f"Se detectaron {reset_count} lecturas a 0 en el mismo informe. "
                        "Se conserva una lectura canónica ya existente; si no la hay, el cero se confirma "
                        "como lectura inicial. Cada decisión queda auditada."
                    )
                    action_text = f"Aplicar criterio a ceros ({reset_count})"
                    action = lambda selected=issue, count=reset_count: (
                        dialog.destroy(),
                        open_initial_zero_confirmation_dialog(
                            app, selected.id_case, selected.id_document, count,
                            Path(selected.archived_path).name,
                        ),
                    )
                else:
                    message = issue.message
                    action_text = "Resolver incidencia"
                    action = lambda selected=issue: (
                        dialog.destroy(), open_issue_dialog(app, selected),
                    )
                ctk.CTkLabel(
                    group, text=message, wraplength=740, justify="left",
                    text_color=C["texto_sec"],
                ).pack(anchor="w", padx=12, pady=(0, 6))
                ctk.CTkButton(
                    group, text=action_text, height=32, corner_radius=8,
                    fg_color=C["primario"], hover_color=C["primario_hover"],
                    command=action,
                ).pack(anchor="e", padx=12, pady=(0, 10))
            kind_labels = {
                "invoice": "Factura", "reading": "Lecturas", "owners": "Propietarios",
            }
            for source in pending_sources:
                group = ctk.CTkFrame(panel)
                group.pack(fill="x", pady=8)
                ctk.CTkLabel(
                    group, text=source.original_name,
                    font=UIM.fuente(12, "bold"), wraplength=740,
                ).pack(anchor="w", padx=12, pady=(10, 2))
                ctk.CTkLabel(
                    group,
                    text=(
                        f"Tipo: {kind_labels.get(source.document_kind, source.document_kind)}. "
                        "Los datos están completos, pero necesitan tu confirmación antes de aplicarse."
                    ),
                    wraplength=740, justify="left", text_color=C["texto_sec"],
                ).pack(anchor="w", padx=12, pady=(0, 6))
                ctk.CTkButton(
                    group, text="Confirmar fuente", height=32, corner_radius=8,
                    fg_color=C["primario"], hover_color=C["primario_hover"],
                    command=lambda document_id=source.id_document: confirm_pending_source(document_id),
                ).pack(anchor="e", padx=12, pady=(0, 10))
            if not issues and not pending_sources:
                text = (
                    f"Se han aplicado automáticamente {automatic} fuente(s) completas.\n"
                    "El expediente ya está listo para generar el Excel oficial."
                    if automatic else "El expediente ya está listo para generar el Excel oficial."
                )
                ctk.CTkLabel(panel, text=text, justify="left").pack(pady=16)
        finally:
            connection.close()

        if ready_for_calculation:
            def completed():
                if dialog.winfo_exists():
                    dialog.destroy()
                app._refrescar_lista_expedientes(select_case_id=case_id)
                app._refrescar_expediente()
                app.log(
                    "Fuentes aplicadas y validadas. El expediente está listo para generar el Excel oficial.",
                    "ok",
                )
                messagebox.showinfo(
                    "Fuentes listas",
                    "Las fuentes se han aplicado automáticamente. Ya puedes generar el Excel oficial.",
                    parent=app,
                )
            app.after(0, completed)

    render()


def open_counter_reset_carry_forward_dialog(
    app: "AppGestionFincas", case_id: int, document_id: int, count: int, source_name: str,
) -> None:
    """Una sola autorización para un reinicio generalizado en una fuente."""
    dialog = _dialog(app, "Reinicio masivo de contadores", 700, 430)
    panel = ctk.CTkFrame(dialog, fg_color=C["panel"], corner_radius=16)
    panel.pack(fill="both", expand=True, padx=18, pady=18)
    ctk.CTkLabel(
        panel, text="Conservar las últimas lecturas válidas",
        font=UIM.fuente(20, "bold"), text_color=C["texto"],
    ).pack(anchor="w", padx=22, pady=(22, 6))
    ctk.CTkLabel(
        panel,
        text=(
            f"{source_name} contiene {count} contadores cuyo valor disminuye. "
            "Se conservará la lectura previa como cierre temporal y se anotará un consumo provisional de 0. "
            "Cuando llegue una lectura posterior fiable, podrás recalcular el período."
        ),
        wraplength=600, justify="left", text_color=C["texto_sec"],
    ).pack(anchor="w", padx=22, pady=(0, 14))
    reason = _packed_field(panel, "Motivo y fuente consultada")
    reason.insert(0, "Reinicio masivo; se conserva la última lectura válida hasta la siguiente lectura fiable.")

    actions = ctk.CTkFrame(panel, fg_color="transparent")
    actions.pack(fill="x", padx=22, pady=(22, 18))
    ctk.CTkButton(
        actions, text="Cancelar", command=dialog.destroy, height=38, corner_radius=9,
        **UIM.secondary_button_kwargs(),
    ).pack(side="right")

    def apply_carry_forward():
        if app._procesando:
            return
        rationale = reason.get().strip()
        if not rationale:
            messagebox.showwarning("Motivo requerido", "Indica el motivo del criterio temporal.", parent=dialog)
            return
        connection = gestor_bd.conectar(str(app.ruta_bd_expedientes))
        try:
            resolved = document_review.carry_forward_counter_resets_for_source(
                connection, case_id=case_id, document_id=document_id,
                reason=rationale, approved_by="usuario_local",
            )
            remaining = len(document_review.list_open_issues(connection, case_id))
            ready = None
            if not remaining and not document_review.case_has_unapplied_sources(connection, case_id):
                ready = document_review.validate_case_ready(connection, case_id)
        except Exception as error:
            connection.rollback()
            messagebox.showwarning("No se pudo aplicar el criterio", str(error), parent=dialog)
            return
        finally:
            connection.close()
        dialog.destroy()
        app._refrescar_lista_expedientes(select_case_id=case_id)
        app._refrescar_expediente()
        if ready is not None and ready.status == "ready_for_calculation":
            app.log(f"Criterio temporal aplicado a {resolved} contador(es). Listo para cálculo.", "ok")
        else:
            app.log(f"Criterio temporal aplicado a {resolved} contador(es). Quedan {remaining} incidencia(s).", "aviso")

    ctk.CTkButton(
        actions, text=f"Aplicar a los {count} contadores", command=apply_carry_forward,
        height=38, corner_radius=9, fg_color=C["exito"], hover_color=C["exito_hover"],
        font=UIM.fuente(12, "bold"),
    ).pack(side="right", padx=(0, 8))


def open_initial_zero_confirmation_dialog(
    app: "AppGestionFincas", case_id: int, document_id: int, count: int, source_name: str,
) -> None:
    """Pide una única confirmación humana para ceros iniciales de una fuente."""
    dialog = _dialog(app, "Resolver lecturas a 0", 700, 430)
    panel = ctk.CTkFrame(dialog, fg_color=C["panel"], corner_radius=16)
    panel.pack(fill="both", expand=True, padx=18, pady=18)
    ctk.CTkLabel(
        panel, text="Aplicar criterio a las lecturas a 0", font=UIM.fuente(20, "bold"), text_color=C["texto"],
    ).pack(anchor="w", padx=22, pady=(22, 6))
    ctk.CTkLabel(
        panel,
        text=(
            f"{source_name} contiene {count} lecturas a 0. Si ya hay una lectura canónica fiable "
            "para esa fecha, se conservará; si no existe, el cero se guardará como lectura inicial real. "
            "Las lecturas posteriores del archivo podrán continuar normalmente."
        ),
        wraplength=600, justify="left", text_color=C["texto_sec"],
    ).pack(anchor="w", padx=22, pady=(0, 14))
    reason = _packed_field(panel, "Motivo y fuente consultada")
    reason.insert(0, "Ceros contrastados; se conserva la lectura canónica existente cuando la haya.")

    actions = ctk.CTkFrame(panel, fg_color="transparent")
    actions.pack(fill="x", padx=22, pady=(22, 18))
    ctk.CTkButton(
        actions, text="Cancelar", command=dialog.destroy, height=38, corner_radius=9,
        **UIM.secondary_button_kwargs(),
    ).pack(side="right")

    def confirm_zeroes():
        if app._procesando:
            return
        rationale = reason.get().strip()
        if not rationale:
            messagebox.showwarning("Motivo requerido", "Indica el criterio aplicado a estos ceros.", parent=dialog)
            return
        connection = gestor_bd.conectar(str(app.ruta_bd_expedientes))
        try:
            resolved = document_review.confirm_initial_zero_readings_for_source(
                connection, case_id=case_id, document_id=document_id,
                reason=rationale, approved_by="usuario_local",
            )
            remaining = len(document_review.list_open_issues(connection, case_id))
            ready = None
            if not remaining and not document_review.case_has_unapplied_sources(connection, case_id):
                ready = document_review.validate_case_ready(connection, case_id)
        except Exception as error:
            connection.rollback()
            messagebox.showwarning("No se pudo aplicar el criterio", str(error), parent=dialog)
            return
        finally:
            connection.close()
        dialog.destroy()
        app._refrescar_lista_expedientes(select_case_id=case_id)
        app._refrescar_expediente()
        if ready is not None and ready.status == "ready_for_calculation":
            app.log(f"Criterio aplicado a {resolved} lecturas a 0. Listo para cálculo.", "ok")
        else:
            app.log(f"Criterio aplicado a {resolved} lecturas a 0. Quedan {remaining} incidencia(s).", "aviso")

    ctk.CTkButton(
        actions, text=f"Aplicar a los {count} ceros", command=confirm_zeroes,
        height=38, corner_radius=9, fg_color=C["exito"], hover_color=C["exito_hover"],
        font=UIM.fuente(12, "bold"),
    ).pack(side="right", padx=(0, 8))


def open_archived_path_resolution_dialog(app: "AppGestionFincas", case_id: int) -> None:
    """Permite elegir una copia ya verificada sin mover ni borrar fuentes."""
    dialog = _dialog(app, "Resolver copias archivadas", 860, 680)
    ctk.CTkLabel(
        dialog,
        text="Elige la copia correcta",
        font=UIM.fuente(20, "bold"),
        text_color=C["texto"],
    ).pack(anchor="w", padx=22, pady=(20, 3))
    ctk.CTkLabel(
        dialog,
        text=(
            "Se muestran solo archivos con la misma huella que el documento registrado. "
            "Elegir una copia actualiza la referencia del expediente; no mueve ni borra archivos."
        ),
        font=UIM.fuente(11), text_color=C["texto_sec"], justify="left", wraplength=790,
    ).pack(anchor="w", padx=22, pady=(0, 12))
    panel = ctk.CTkScrollableFrame(dialog, fg_color=C["fondo"])
    panel.pack(fill="both", expand=True, padx=16, pady=(0, 16))

    def open_candidate(path: Path) -> None:
        try:
            if sys.platform != "win32":
                raise OSError("Abrir el archivo solo está disponible en Windows.")
            os.startfile(str(path))
        except Exception as exc:
            messagebox.showerror("No se pudo abrir la copia", str(exc), parent=dialog)

    def render() -> None:
        for widget in panel.winfo_children():
            widget.destroy()
        connection = gestor_bd.conectar(str(app.ruta_bd_expedientes))
        try:
            candidates = expedient_service.list_archived_path_candidates(
                connection, archive_root=app.ruta_archivo_expedientes, case_id=case_id,
            )
        finally:
            connection.close()

        if not candidates:
            ctk.CTkLabel(
                panel,
                text="No hay copias duplicadas pendientes en este expediente.",
                font=UIM.fuente(12), text_color=C["texto_sec"],
            ).pack(anchor="w", padx=14, pady=20)
            return

        for item in candidates:
            group = ctk.CTkFrame(
                panel, fg_color=C["panel"], corner_radius=11,
                border_width=1, border_color=C["borde"],
            )
            group.pack(fill="x", padx=3, pady=6)
            ctk.CTkLabel(
                group, text=item.original_name, font=UIM.fuente(12, "bold"),
                text_color=C["texto"], anchor="w", wraplength=720,
            ).pack(fill="x", padx=14, pady=(12, 2))
            ctk.CTkLabel(
                group,
                text=f"Hay {len(item.candidate_paths)} copias idénticas. Abre una si necesitas comprobarla.",
                font=UIM.fuente(10), text_color=C["texto_sec"], anchor="w",
            ).pack(fill="x", padx=14, pady=(0, 7))
            for candidate in item.candidate_paths:
                row = ctk.CTkFrame(group, fg_color=C["panel_2"], corner_radius=8)
                row.pack(fill="x", padx=14, pady=3)
                ctk.CTkLabel(
                    row, text=candidate.name, font=UIM.fuente(10), text_color=C["texto"],
                    anchor="w", wraplength=420,
                ).pack(side="left", fill="x", expand=True, padx=10, pady=8)
                ctk.CTkButton(
                    row, text="Abrir", width=68, height=29, corner_radius=7,
                    **UIM.secondary_button_kwargs(),
                    command=lambda path=candidate: open_candidate(path),
                ).pack(side="right", padx=(4, 7), pady=5)
                ctk.CTkButton(
                    row, text="Usar esta copia", width=126, height=29, corner_radius=7,
                    fg_color=C["primario"], hover_color=C["primario_hover"],
                    command=lambda document_id=item.document_id, path=candidate: select(document_id, path),
                ).pack(side="right", padx=(0, 4), pady=5)

    def select(document_id: int, path: Path) -> None:
        if not messagebox.askyesno(
            "Confirmar copia",
            f"Se usará esta copia para el expediente:\n\n{path.name}\n\nNo se moverá ni borrará ningún archivo.",
            parent=dialog,
        ):
            return
        connection = gestor_bd.conectar(str(app.ruta_bd_expedientes))
        try:
            selected = expedient_service.select_archived_source_path(
                connection, document_id, path, archive_root=app.ruta_archivo_expedientes,
            )
            document_review.resolve_archived_source_issue(
                connection, case_id, document_id,
            )
        except (LookupError, ValueError, RuntimeError, OSError) as exc:
            messagebox.showerror("No se pudo actualizar la ruta", str(exc), parent=dialog)
            return
        finally:
            connection.close()
        app.log(f"Ruta archivada recuperada: {selected.name}", "ok")
        app._refrescar_lista_expedientes(select_case_id=case_id)
        app._refrescar_expediente()
        render()

    render()


def open_source_preview(app: "AppGestionFincas", issue: ReviewIssue) -> None:
    """Muestra la página o la hoja de la fuente con la evidencia resaltada.

    Sólo se resalta lo que se localiza de verdad en el archivo; si no, se
    enseña el fragmento guardado. Sin visor posible, se usa el contexto.
    """
    connection = gestor_bd.conectar(str(app.ruta_bd_expedientes))
    try:
        evidence = source_preview.evidence_for_issue(
            connection, issue.id_document, issue.field_name, issue.detected_value,
        )
    finally:
        connection.close()
    path = Path(issue.archived_path)
    if not path.is_file() or path.suffix.lower() not in {".pdf", ".xlsx", ".xlsm", ".xls", ".csv"}:
        show_issue_context(app, issue)
        return

    # Las facturas tienen pocas páginas: leer sus palabras es inmediato y así
    # el visor se abre encima del diálogo de la incidencia, que es modal.
    result = source_preview.build_preview(path, evidence)
    if result.kind == "text":
        show_issue_context(app, issue)
    else:
        _show_source_preview(app, issue, path, result)


def _show_source_preview(app: "AppGestionFincas", issue: ReviewIssue, path: Path, result) -> None:
    dialog = _dialog(app, "Evidencia en la fuente", 920, 780)
    dialog.resizable(True, True)
    panel = ctk.CTkFrame(dialog, fg_color=C["panel"], corner_radius=16,
                         border_width=1, border_color=C["borde"])
    panel.pack(fill="both", expand=True, padx=14, pady=14)
    ctk.CTkLabel(panel, text=source_display_name(path), font=UIM.fuente(17, "bold"),
                 text_color=C["texto"]).pack(anchor="w", padx=18, pady=(16, 0))
    located = result.match is not None or (result.sheet is not None and result.sheet.target is not None)
    ctk.CTkLabel(panel, text=result.headline, font=UIM.fuente(11, "bold"),
                 text_color=C["primario"] if located else C["alerta"],
                 wraplength=860, justify="left").pack(anchor="w", padx=18, pady=(2, 0))
    evidence = result.evidence
    details = [f"Dato: {issue_guidance(issue.field_name)['label']}"]
    if evidence.value:
        details.append(f"Valor detectado: {evidence.value}")
    if evidence.fragment:
        details.append(f"Fragmento guardado: «{evidence.fragment}»")
    details.extend(result.notes)
    ctk.CTkLabel(panel, text="\n".join(details), font=UIM.fuente(10), text_color=C["texto_sec"],
                 wraplength=860, justify="left").pack(anchor="w", padx=18, pady=(4, 8))

    body = ctk.CTkFrame(panel, fg_color=C["panel_2"], corner_radius=11)
    body.pack(fill="both", expand=True, padx=18, pady=(0, 8))
    footer = ctk.CTkFrame(panel, fg_color="transparent")
    footer.pack(fill="x", padx=18, pady=(0, 14))

    if result.kind == "pdf":
        _pdf_viewer(body, footer, path, result)
    else:
        _sheet_viewer(body, result.sheet)

    ctk.CTkButton(footer, text="Cerrar", command=dialog.destroy, height=34, corner_radius=8,
                  fg_color=C["primario"], hover_color=C["primario_hover"]).pack(side="right")
    ctk.CTkButton(footer, text="Abrir archivo", command=lambda: open_archived_file(app, issue),
                  height=34, corner_radius=8, **UIM.secondary_button_kwargs()).pack(side="right", padx=(0, 8))


def _pdf_viewer(body, footer, path: Path, result) -> None:
    from PIL import ImageTk

    canvas = tk.Canvas(body, highlightthickness=0, background="#E5E7EB")
    vertical = tk.Scrollbar(body, orient="vertical", command=canvas.yview)
    canvas.configure(yscrollcommand=vertical.set)
    vertical.pack(side="right", fill="y")
    canvas.pack(side="left", fill="both", expand=True)
    state = {"page": result.page, "image": None}
    label = tk.StringVar(master=body)

    def show(page: int):
        page = max(1, min(result.page_count, page))
        state["page"] = page
        boxes = result.match.boxes if result.match is not None and result.match.page == page else ()
        try:
            image = source_preview.render_pdf_page(path, page, boxes)
        except Exception as error:  # una página dañada no cierra el visor
            canvas.delete("all")
            canvas.create_text(20, 20, anchor="nw", text=f"No se pudo mostrar la página: {error}")
            return
        state["image"] = ImageTk.PhotoImage(image, master=canvas)
        canvas.delete("all")
        canvas.create_image(0, 0, anchor="nw", image=state["image"])
        canvas.configure(scrollregion=(0, 0, image.width, image.height))
        if boxes:
            top = min(box[1] for box in boxes) * 110 / 72
            canvas.yview_moveto(max(0.0, (top - 120) / image.height))
        else:
            canvas.yview_moveto(0)
        label.set(f"Página {page} de {result.page_count}")

    def wheel(event):
        canvas.yview_scroll(-1 if (getattr(event, "delta", 0) > 0 or getattr(event, "num", 0) == 4) else 1, "units")
    for sequence in ("<MouseWheel>", "<Button-4>", "<Button-5>"):
        canvas.bind(sequence, wheel)

    navigation = ctk.CTkFrame(footer, fg_color="transparent")
    navigation.pack(side="left")
    for text, command in (("◀ Anterior", lambda: show(state["page"] - 1)),
                          ("Siguiente ▶", lambda: show(state["page"] + 1))):
        ctk.CTkButton(navigation, text=text, command=command, width=100, height=34, corner_radius=8,
                      **UIM.secondary_button_kwargs()).pack(side="left", padx=(0, 6))
    ctk.CTkLabel(navigation, textvariable=label, font=UIM.fuente(11),
                 text_color=C["texto_sec"]).pack(side="left", padx=8)
    if result.match is not None:
        ctk.CTkButton(navigation, text="Ir al resaltado", command=lambda: show(result.match.page),
                      height=34, corner_radius=8, **UIM.secondary_button_kwargs()).pack(side="left", padx=6)
    show(result.page)


def _sheet_viewer(body, window) -> None:
    grid = ctk.CTkScrollableFrame(body, fg_color="transparent", orientation="horizontal")
    grid.pack(fill="both", expand=True, padx=8, pady=8)
    ctk.CTkLabel(grid, text=f"Hoja «{window.sheet}»", font=UIM.fuente(11, "bold"),
                 text_color=C["texto"]).grid(row=0, column=0, columnspan=4, sticky="w", pady=(0, 6))
    width = max((len(row) for row in window.rows), default=0)
    for offset in range(width):
        ctk.CTkLabel(grid, text=source_preview._column_letter(window.first_column + offset),
                     font=UIM.fuente(10, "bold"), text_color=C["texto_sec"]).grid(row=1, column=offset + 1)
    for row_offset, values in enumerate(window.rows):
        number = window.first_row + row_offset
        ctk.CTkLabel(grid, text=str(number), font=UIM.fuente(10, "bold"),
                     text_color=C["texto_sec"]).grid(row=row_offset + 2, column=0, padx=(0, 6))
        for column_offset in range(width):
            value = values[column_offset] if column_offset < len(values) else ""
            target = window.target == (number, window.first_column + column_offset)
            ctk.CTkLabel(
                grid, text=value[:28], width=104, height=26, corner_radius=4, anchor="w",
                font=UIM.fuente(10, "bold" if target else "normal"),
                fg_color="#FDE68A" if target else C["panel"],
                text_color="#7C2D12" if target else C["texto"],
            ).grid(row=row_offset + 2, column=column_offset + 1, padx=1, pady=1, sticky="ew")


def open_archived_file(app: "AppGestionFincas", issue: ReviewIssue) -> None:
    try:
        if sys.platform != "win32":
            raise OSError("Abrir el archivo solo está disponible en Windows.")
        os.startfile(str(issue.archived_path))
    except Exception as exc:
        messagebox.showerror("No se pudo abrir el archivo", str(exc))


def resolution_route_for_issue(issue: ReviewIssue) -> str:
    """Selecciona la ruta de revisión sin crear controles de Tk."""
    if issue.code == "COUNTER_RESET":
        return "counter_reset_estimate"
    if issue.code == "INVOICE_OUTSIDE_PERIOD":
        return "dismiss_invoice_outside_period"
    if issue.code == "DOCUMENT_CLASSIFICATION_REQUIRED":
        return "classify_unknown"
    if issue.code == "ARCHIVED_SOURCE_DUPLICATE":
        return "archived_source_duplicate"
    if issue.code == "ARCHIVED_SOURCE_MISSING":
        return "archived_source_missing"
    return "generic_correction"


def _euros(value) -> str:
    return f"{float(value):,.2f} €".replace(",", "X").replace(".", ",").replace("X", ".")


_FEE_SERVICES = {"ACS": "Agua caliente (ACS)", "CALEFACCION": "Calefacción"}


def open_service_fees_dialog(
    app: "AppGestionFincas", servicio: str = "ACS", period_id: int | None = None,
) -> None:
    """Cuotas que la comunidad cobra a los vecinos por ACS o calefacción.

    Es el lado de los ingresos del estudio y el único dato que no sale de
    ninguna factura ni de ningún contador: lo decide la comunidad. Sin él, el
    análisis no puede comparar lo cobrado con lo que ha costado el servicio.
    """
    period_id = period_id or getattr(app, "id_periodo", None)
    if not app.id_comunidad or not period_id:
        messagebox.showinfo(
            "Cuotas cobradas", "Elige antes una comunidad y un expediente.", parent=app,
        )
        return
    connection = gestor_bd.conectar(str(app.ruta_bd_expedientes))
    try:
        periodo = connection.execute(
            "SELECT nombre,fecha_inicio,fecha_fin FROM periodos WHERE id_periodo=?", (period_id,),
        ).fetchone()
        dwellings = cuotas_servicio.viviendas_facturables(connection, app.id_comunidad)
    finally:
        connection.close()
    if periodo is None:
        messagebox.showinfo("Cuotas cobradas", "El período del expediente no existe.", parent=app)
        return
    months = cuotas_servicio.meses_del_periodo(periodo["fecha_inicio"], periodo["fecha_fin"])
    state = {"service": servicio if servicio in _FEE_SERVICES else "ACS"}

    dialog = _dialog(app, "Cuotas cobradas", 760, 720)
    panel = ctk.CTkScrollableFrame(
        dialog, fg_color=C["panel"], corner_radius=16, border_width=1, border_color=C["borde"],
    )
    panel.pack(fill="both", expand=True, padx=18, pady=(18, 8))

    ctk.CTkLabel(
        panel, text="Cuotas cobradas a los vecinos", font=UIM.fuente(20, "bold"),
        text_color=C["texto"],
    ).pack(anchor="w", padx=18, pady=(14, 2))
    ctk.CTkLabel(
        panel,
        text=(
            f"{periodo['nombre']} · {len(months)} mes(es). Es lo que la comunidad gira en los "
            "recibos; el análisis lo compara con el coste de las facturas. Si cambias algo con "
            "el Excel ya generado, vuelve a generarlo."
        ),
        font=UIM.fuente(11), text_color=C["texto_sec"], wraplength=660, justify="left",
    ).pack(anchor="w", padx=18, pady=(0, 10))

    selector = ctk.CTkSegmentedButton(
        panel, values=list(_FEE_SERVICES.values()),
        command=lambda label: change_service(label), font=UIM.fuente(11),
        selected_color=C["primario"], selected_hover_color=C["primario_hover"],
    )
    selector.pack(anchor="w", padx=18, pady=(0, 12))

    # --- cuota fija mensual --------------------------------------------------
    fixed = ctk.CTkFrame(panel, fg_color=C["panel_2"], corner_radius=10)
    fixed.pack(fill="x", padx=18)
    ctk.CTkLabel(
        fixed, text="Cuota fija mensual", font=UIM.fuente(13, "bold"), text_color=C["texto"],
    ).pack(anchor="w", padx=14, pady=(12, 0))
    ctk.CTkLabel(
        fixed,
        text=(
            "Un tramo por cada cambio de importe. Al generar se sustituyen las cuotas fijas "
            "anteriores de este servicio. Mes: 09/2025, 2025-09 o «septiembre 2025»."
        ),
        font=UIM.fuente(10), text_color=C["texto_sec"], wraplength=640, justify="left",
    ).pack(anchor="w", padx=14, pady=(0, 6))
    rows_frame = ctk.CTkFrame(fixed, fg_color="transparent")
    rows_frame.pack(fill="x", padx=14)
    per_dwelling = tk.BooleanVar(value=False)
    tramos: list[tuple[tk.StringVar, tk.StringVar]] = []
    preview = ctk.CTkLabel(
        fixed, text="", font=UIM.fuente(11, "bold"), text_color=C["primario"],
        wraplength=640, justify="left", anchor="w",
    )

    def parsed_tramos(strict: bool):
        pairs = []
        for month_var, amount_var in tramos:
            month_text, amount_text = month_var.get().strip(), amount_var.get().strip()
            if not month_text and not amount_text:
                continue
            if not month_text or not amount_text:
                if strict:
                    raise ValueError("Cada tramo necesita su mes de inicio y su importe.")
                continue
            amount = cuotas_servicio._importe(amount_text, "El importe mensual")
            if per_dwelling.get():
                amount *= dwellings
            pairs.append((cuotas_servicio.parse_mes(month_text), amount))
        return sorted(pairs)

    def update_preview(*_args):
        try:
            pairs = parsed_tramos(strict=False)
        except ValueError as error:
            preview.configure(text=f"⚠ {error}", text_color=C["aviso"])
            return
        if not pairs:
            preview.configure(text="Indica al menos un importe mensual.", text_color=C["texto_sec"])
            return
        amounts = []
        for month in months:
            current = [amount for start, amount in pairs if start[:7] <= month[:7]]
            if current:
                amounts.append(current[-1])
        groups: dict[object, int] = {}
        for amount in amounts:
            groups[amount] = groups.get(amount, 0) + 1
        detail = " + ".join(f"{count} × {_euros(amount)}" for amount, count in groups.items())
        uncovered = len(months) - len(amounts)
        text = f"{len(amounts)} cuota(s): {detail} = {_euros(sum(amounts))}"
        if per_dwelling.get():
            text += f"  (importe por vivienda × {dwellings} viviendas)"
        if uncovered:
            text += f"\n⚠ {uncovered} mes(es) al inicio del período sin cuota: el primer tramo empieza más tarde."
        preview.configure(text=text, text_color=C["aviso"] if uncovered else C["primario"])

    def add_tramo(month: str = "", amount: str = "") -> None:
        index = len(tramos)
        month_var, amount_var = tk.StringVar(value=month), tk.StringVar(value=amount)
        ctk.CTkLabel(rows_frame, text="Desde", font=UIM.fuente(11), text_color=C["texto_sec"]).grid(
            row=index, column=0, sticky="w", pady=3)
        ctk.CTkEntry(rows_frame, textvariable=month_var, height=30, corner_radius=8, width=130,
                     placeholder_text="MM/AAAA").grid(row=index, column=1, sticky="w", padx=(6, 16), pady=3)
        ctk.CTkLabel(rows_frame, text="Importe al mes", font=UIM.fuente(11), text_color=C["texto_sec"]).grid(
            row=index, column=2, sticky="w", pady=3)
        ctk.CTkEntry(rows_frame, textvariable=amount_var, height=30, corner_radius=8, width=120,
                     placeholder_text="0,00").grid(row=index, column=3, sticky="w", padx=6, pady=3)
        month_var.trace_add("write", update_preview)
        amount_var.trace_add("write", update_preview)
        tramos.append((month_var, amount_var))

    first_month = periodo["fecha_inicio"][:7]
    add_tramo(f"{first_month[5:7]}/{first_month[:4]}")
    add_tramo()
    ctk.CTkCheckBox(
        fixed,
        text=(
            f"El importe es por vivienda (se multiplica por {dwellings} viviendas activas)"
            if dwellings else "El importe es por vivienda (no hay viviendas activas registradas)"
        ),
        variable=per_dwelling, command=update_preview, font=UIM.fuente(11),
        state="normal" if dwellings else "disabled",
    ).pack(anchor="w", padx=14, pady=(6, 2))
    preview.pack(fill="x", padx=14, pady=(4, 4))
    fixed_buttons = ctk.CTkFrame(fixed, fg_color="transparent")
    fixed_buttons.pack(fill="x", padx=14, pady=(2, 12))

    # --- liquidación por consumo --------------------------------------------
    variable = ctk.CTkFrame(panel, fg_color=C["panel_2"], corner_radius=10)
    variable.pack(fill="x", padx=18, pady=(12, 0))
    ctk.CTkLabel(
        variable, text="Liquidación por consumo", font=UIM.fuente(13, "bold"), text_color=C["texto"],
    ).pack(anchor="w", padx=14, pady=(12, 0))
    ctk.CTkLabel(
        variable, text="Lo cobrado tras cada lectura de contadores (una fila por recibo).",
        font=UIM.fuente(10), text_color=C["texto_sec"],
    ).pack(anchor="w", padx=14, pady=(0, 6))
    variable_row = ctk.CTkFrame(variable, fg_color="transparent")
    variable_row.pack(fill="x", padx=14)
    entries = {}
    for column, (key, label, placeholder, width) in enumerate((
        ("fecha", "Fecha", "DD/MM/AAAA", 120),
        ("importe", "Importe", "0,00", 110),
        ("consumo", "Consumo (opcional)", "m³ o kWh", 120),
    )):
        ctk.CTkLabel(variable_row, text=label, font=UIM.fuente(11), text_color=C["texto_sec"]).grid(
            row=0, column=column, sticky="w", padx=(0, 12))
        entries[key] = ctk.CTkEntry(variable_row, height=30, corner_radius=8, width=width,
                                    placeholder_text=placeholder)
        entries[key].grid(row=1, column=column, sticky="w", padx=(0, 12))

    # --- apuntes ---------------------------------------------------------------
    summary = ctk.CTkLabel(
        panel, text="", font=UIM.fuente(12, "bold"), text_color=C["texto"], justify="left", anchor="w",
    )
    summary.pack(fill="x", padx=18, pady=(14, 4))
    ledger = ctk.CTkFrame(panel, fg_color=C["panel_2"], corner_radius=10)
    ledger.pack(fill="x", padx=18, pady=(0, 12))

    def with_connection(action):
        conexion = gestor_bd.conectar(str(app.ruta_bd_expedientes))
        try:
            result = action(conexion)
            conexion.commit()
            return result
        finally:
            conexion.close()

    def delete_entry(entry_id: int) -> None:
        with_connection(lambda conexion: cuotas_servicio.borrar(conexion, entry_id))
        refresh()

    def refresh() -> None:
        service = state["service"]
        apuntes, total = with_connection(lambda conexion: (
            cuotas_servicio.listar(conexion, community_id=app.id_comunidad,
                                   period_id=period_id, servicio=service),
            cuotas_servicio.resumen(conexion, community_id=app.id_comunidad,
                                    period_id=period_id, servicio=service),
        ))
        summary.configure(text=(
            f"Cobrado de {_FEE_SERVICES[service]}: {_euros(total.fija)} fijo + "
            f"{_euros(total.variable)} por consumo = {_euros(total.total)}"
        ))
        for child in ledger.winfo_children():
            child.destroy()
        if not apuntes:
            ctk.CTkLabel(
                ledger, text="Todavía no hay cuotas anotadas para este servicio.",
                font=UIM.fuente(11), text_color=C["texto_sec"],
            ).pack(anchor="w", padx=12, pady=10)
            return
        for apunte in apuntes:
            row = ctk.CTkFrame(ledger, fg_color="transparent")
            row.pack(fill="x", padx=10, pady=1)
            fecha = datetime.strptime(apunte["fecha"][:10], "%Y-%m-%d").strftime("%d/%m/%Y")
            kind = "Cuota fija" if apunte["concepto"] == "fija" else "Liquidación"
            extra = f" · consumo {apunte['consumo']:g}" if apunte["consumo"] is not None else ""
            ctk.CTkLabel(row, text=f"{fecha}   {kind}{extra}", font=UIM.fuente(11),
                         text_color=C["texto"], anchor="w").pack(side="left")
            ctk.CTkButton(
                row, text="Borrar", width=60, height=22, corner_radius=6, font=UIM.fuente(10),
                command=lambda entry_id=apunte["id_cuota"]: delete_entry(entry_id),
                **UIM.secondary_button_kwargs(),
            ).pack(side="right")
            ctk.CTkLabel(row, text=_euros(apunte["importe"]), font=UIM.fuente(11, "bold"),
                         text_color=C["texto"], width=110, anchor="e").pack(side="right", padx=10)

    def generate_fixed() -> None:
        try:
            pairs = parsed_tramos(strict=True)
            if not pairs:
                raise ValueError("Indica al menos un importe mensual.")
            created = with_connection(lambda conexion: cuotas_servicio.generar_cuotas_mensuales(
                conexion, community_id=app.id_comunidad, period_id=period_id,
                servicio=state["service"], tramos=[(start, amount) for start, amount in pairs],
                notas="Cuota mensual del recibo",
            ))
        except (ValueError, LookupError) as error:
            messagebox.showwarning("No se pudo guardar", str(error), parent=dialog)
            return
        app.log(f"Cuota fija de {_FEE_SERVICES[state['service']]}: {created} mes(es) anotados.", "ok")
        refresh()

    def clear_fixed() -> None:
        if not messagebox.askyesno(
            "Borrar cuota fija",
            f"¿Borrar todas las cuotas fijas de {_FEE_SERVICES[state['service']]} de este período?",
            parent=dialog,
        ):
            return
        with_connection(lambda conexion: cuotas_servicio.borrar_fijas(
            conexion, community_id=app.id_comunidad, period_id=period_id, servicio=state["service"],
        ))
        refresh()

    def add_variable() -> None:
        try:
            with_connection(lambda conexion: cuotas_servicio.registrar_cuota(
                conexion, community_id=app.id_comunidad, period_id=period_id,
                servicio=state["service"], concepto="variable",
                fecha=entries["fecha"].get().strip(), importe=entries["importe"].get().strip(),
                consumo=entries["consumo"].get().strip() or None,
                notas="Liquidación por consumo",
            ))
        except (ValueError, LookupError) as error:
            messagebox.showwarning("No se pudo guardar", str(error), parent=dialog)
            return
        for entry in entries.values():
            entry.delete(0, "end")
        refresh()

    def change_service(label: str) -> None:
        state["service"] = next(key for key, text in _FEE_SERVICES.items() if text == label)
        refresh()

    ctk.CTkButton(
        fixed_buttons, text="Añadir tramo", width=110, height=30, corner_radius=8,
        font=UIM.fuente(11), command=lambda: add_tramo(), **UIM.secondary_button_kwargs(),
    ).pack(side="left")
    ctk.CTkButton(
        fixed_buttons, text="Borrar cuota fija", width=130, height=30, corner_radius=8,
        font=UIM.fuente(11), command=clear_fixed, **UIM.secondary_button_kwargs(),
    ).pack(side="left", padx=8)
    ctk.CTkButton(
        fixed_buttons, text="Generar cuota fija del período", height=30, corner_radius=8,
        font=UIM.fuente(11, "bold"), fg_color=C["primario"], hover_color=C["primario_hover"],
        command=generate_fixed,
    ).pack(side="right")
    ctk.CTkButton(
        variable, text="Añadir liquidación", width=160, height=30, corner_radius=8,
        font=UIM.fuente(11, "bold"), fg_color=C["primario"], hover_color=C["primario_hover"],
        command=add_variable,
    ).pack(anchor="e", padx=14, pady=(8, 12))
    ctk.CTkButton(
        dialog, text="Cerrar", width=110, height=34, corner_radius=8,
        command=lambda: (dialog.destroy(), app._refrescar_expediente()
                         if getattr(app, "id_expediente", None) else None),
        **UIM.secondary_button_kwargs(),
    ).pack(anchor="e", padx=20, pady=(0, 14))

    selector.set(_FEE_SERVICES[state["service"]])
    update_preview()
    refresh()


def open_coherence_dialog(app: "AppGestionFincas") -> None:
    """Datos comparables, correcciones y justificaciones antes del Excel."""
    import case_workflow_actions
    from contextlib import closing
    case_id, community_id = app.id_expediente, app.id_comunidad
    root = Path(__file__).resolve().parents[1]
    try:
        with closing(gestor_bd.conectar(str(app.ruta_bd_expedientes))) as con:
            profile = case_workflow_actions.resolve_case_profile(
                con, id_case=case_id, active_community_id=community_id, project_root=root)
            case_coherence.evaluate(con, case_id, profile)
    except (ValueError, LookupError) as error:
        messagebox.showwarning('Coherencia', str(error), parent=app)
        return
    dialog = _dialog(app, 'Coherencia antes del Excel', 820, 720)
    content = ctk.CTkScrollableFrame(dialog, fg_color=C['panel'])
    content.pack(fill='both', expand=True, padx=16, pady=16)

    def refresh():
        for child in content.winfo_children():
            child.destroy()
        with closing(gestor_bd.conectar(str(app.ruta_bd_expedientes))) as con:
            settings = case_coherence.load_settings(con, case_id)
            report = case_coherence.evaluate(con, case_id, profile)
            owners = con.execute("SELECT id_propietario,codigo_vivienda,coeficiente FROM propietarios WHERE id_comunidad=? AND activo=1 AND tipo_unidad='vivienda' ORDER BY codigo_vivienda", (community_id,)).fetchall()
        def label(text):
            ctk.CTkLabel(content, text=text, wraplength=740, justify='left',
                        text_color=C['texto'], font=UIM.fuente(12)).pack(anchor='w', padx=12, pady=6)
        label('Coherencia del expediente · Los avisos impiden generar el Excel hasta corregirlos o justificar una diferencia legítima.')
        ctk.CTkButton(content, text='Actualizar revisión', command=refresh).pack(anchor='w', padx=12, pady=6)
        for note in report.notes:
            label(note)
        label('Coeficientes: los pesos relativos se normalizan. Los porcentajes deben sumar 100 %; no se convierten automáticamente entre ambas formas.')
        mode = ctk.CTkComboBox(content, values=['Pesos relativos', 'Porcentajes'])
        mode.set('Porcentajes' if settings['coefficient_mode'] == 'percent' else 'Pesos relativos')
        mode.pack(anchor='w', padx=12, pady=4)
        entries = {}
        for owner in owners:
            row = ctk.CTkFrame(content, fg_color='transparent')
            row.pack(fill='x', padx=12, pady=2)
            ctk.CTkLabel(row, text=owner['codigo_vivienda'], width=180).pack(side='left')
            entry = ctk.CTkEntry(row)
            entry.insert(0, str(owner['coeficiente']))
            entry.pack(side='left')
            entries[owner['id_propietario']] = entry
        label('Tolerancia de diferencia de consumo (%)')
        tolerance = ctk.CTkEntry(content)
        tolerance.insert(0, settings['consumption_tolerance_percent'])
        tolerance.pack(anchor='w', padx=12)

        label('Ejercicio anterior · Avisar cuando un importe o consumo suba o baje por este factor (×)')
        historical_factor = ctk.CTkEntry(content)
        historical_factor.insert(0, settings['historical_change_factor'])
        historical_factor.pack(anchor='w', padx=12)
        label('Diferencia máxima de duración entre intervalos históricos (%)')
        historical_duration = ctk.CTkEntry(content)
        historical_duration.insert(0, settings['historical_duration_tolerance_percent'])
        historical_duration.pack(anchor='w', padx=12)
        label('Confirma las unidades de contador de ambos ejercicios. Si cambiaron, selecciona «Sin confirmar».')
        historical_units = {}
        for kind, values in (('ACS', ['Sin confirmar', 'm³']), ('CALEFACCION', ['Sin confirmar', 'kWh', 'unidades'])):
            label(kind)
            field = ctk.CTkComboBox(content, values=values)
            unit = settings['historical_reading_units'][kind]
            field.set({'m3': 'm³', 'kwh': 'kWh'}.get(unit, unit or 'Sin confirmar'))
            field.pack(anchor='w', padx=12)
            historical_units[kind] = field

        def save(comparisons):
            try:
                with closing(gestor_bd.conectar(str(app.ruta_bd_expedientes))) as con:
                    case_coherence.save_settings(con, case_id, dict(
                        coefficient_mode='percent' if mode.get() == 'Porcentajes' else 'weights',
                        consumption_tolerance_percent=tolerance.get(), comparisons=comparisons,
                        historical_change_factor=historical_factor.get(),
                        historical_duration_tolerance_percent=historical_duration.get(),
                        historical_reading_units={kind: '' if field.get() == 'Sin confirmar' else field.get()
                                                  for kind, field in historical_units.items()}),
                        {owner: entry.get() for owner, entry in entries.items()})
            except (ValueError, RuntimeError) as error:
                messagebox.showwarning('Coherencia', str(error), parent=dialog)
                return
            refresh()
            app._refrescar_expediente()

        ctk.CTkButton(content, text='Guardar configuración y coeficientes',
                      command=lambda: save(settings['comparisons'])).pack(anchor='w', padx=12, pady=10)
        label('Comparaciones declaradas: sólo declara suministros que miden el mismo consumo. En calefacción confirma que las lecturas son kWh y no unidades de repartidor.')
        for index, rule in enumerate(settings['comparisons']):
            label(f"{rule['invoice_type']} [{rule['cups'] or 'un único CUPS'}], {rule['invoice_unit']} → {rule['reading_type']}, {rule['reading_unit']}")
            ctk.CTkButton(content, text='Quitar comparación', command=lambda i=index:
                          save([r for j, r in enumerate(settings['comparisons']) if j != i])).pack(anchor='w', padx=12)
        label('Añadir una equivalencia (no se guarda hasta pulsar el botón):')
        fields = {}
        for key, title, values in (
            ('invoice_type', 'Suministro de las facturas', ['AGUA', 'GAS', 'ELECTRICIDAD', 'ACS', 'CALEFACCION']),
            ('reading_type', 'Contadores de viviendas', ['ACS', 'CALEFACCION']),
            ('invoice_unit', 'Unidad de las facturas', ['m³', 'kWh', 'unidades']),
            ('reading_unit', 'Unidad de las lecturas', ['m³', 'kWh', 'unidades']),
        ):
            label(title)
            fields[key] = ctk.CTkComboBox(content, values=values)
            fields[key].set(values[0])
            fields[key].pack(anchor='w', padx=12)
        label('CUPS o referencia (necesario para seleccionar uno entre varios suministros)')
        cups = ctk.CTkEntry(content)
        cups.pack(anchor='w', padx=12)
        ctk.CTkButton(content, text='Añadir comparación', command=lambda:
                      save(settings['comparisons'] + [dict(
                          **{key: entry.get() for key, entry in fields.items()}, cups=cups.get())])).pack(anchor='w', padx=12, pady=10)
        for finding in report.findings:
            label(('Justificado · ' if finding.accepted else 'Pendiente · ') + finding.message)
            if finding.accepted or not finding.can_accept:
                continue
            reason = ctk.CTkEntry(content, width=680, placeholder_text='Motivo de la diferencia legítima; se conservará en el registro')
            reason.pack(anchor='w', padx=12)
            def accept(key=finding.key, entry=reason):
                try:
                    with closing(gestor_bd.conectar(str(app.ruta_bd_expedientes))) as con:
                        case_coherence.accept_finding(con, case_id, profile, key, entry.get(),
                                                      expected_signature=report.signature)
                except (ValueError, RuntimeError) as error:
                    messagebox.showwarning('Coherencia', str(error), parent=dialog)
                    return
                refresh()
                app._refrescar_expediente()
            ctk.CTkButton(content, text='Aceptar con motivo', command=accept).pack(anchor='w', padx=12, pady=6)
        if not report.pending:
            label('Sin avisos de coherencia pendientes. Los demás requisitos del expediente siguen comprobándose antes del Excel.')
    refresh()


def open_fixed_costs_dialog(app: "AppGestionFincas", period_id: int) -> None:
    """Gastos fijos mensuales del estudio (lecturas y mantenimientos)."""
    connection = gestor_bd.conectar(str(app.ruta_bd_expedientes))
    try:
        current = fixed_costs.load_fixed_costs(connection, app.id_comunidad, period_id)
        periodo = connection.execute(
            "SELECT nombre FROM periodos WHERE id_periodo=?", (period_id,),
        ).fetchone()
        inherited = ()
        if getattr(app, "id_expediente", None):
            import case_workflow_actions
            project_root = Path(__file__).resolve().parents[1]
            try:
                profile = case_workflow_actions.resolve_case_profile(
                    connection, id_case=app.id_expediente, active_community_id=app.id_comunidad,
                    project_root=project_root,
                )
                template = project_root / profile.template_relative_path
                if template.is_file() and "OTROS_GASTOS" in profile.active_modules:
                    inherited = fixed_costs.unconfirmed_template_costs(template, current)
            except (LookupError, ValueError):
                pass  # Los gastos pueden prepararse antes de instalar el perfil.
    finally:
        connection.close()
    dialog = _dialog(app, "Gastos fijos", 560, 700 if inherited else 520)
    panel = ctk.CTkFrame(
        dialog, fg_color=C["panel"], corner_radius=16, border_width=1, border_color=C["borde"],
    )
    panel.pack(fill="both", expand=True, padx=18, pady=(18, 8))
    ctk.CTkLabel(panel, text="Gastos fijos del estudio", font=UIM.fuente(20, "bold"),
                 text_color=C["texto"]).pack(anchor="w", padx=20, pady=(18, 2))
    ctk.CTkLabel(
        panel,
        text=(
            f"{periodo['nombre'] if periodo else ''} · Importe mensual de cada servicio para esta "
            "comunidad. Se escriben en la hoja OTROS GASTOS al generar el Excel (el modelo los "
            "multiplica por 12). Déjalo vacío si la comunidad no tiene ese gasto."
        ),
        font=UIM.fuente(11), text_color=C["texto_sec"], wraplength=480, justify="left",
    ).pack(anchor="w", padx=20, pady=(0, 10))
    if inherited:
        ctk.CTkLabel(
            panel, text=fixed_costs.inherited_cost_message(inherited),
            font=UIM.fuente(11), text_color=C["texto"], wraplength=480, justify="left",
        ).pack(anchor="w", padx=20, pady=(0, 10))
    entries = {}
    for item in fixed_costs.FIXED_COSTS:
        row = ctk.CTkFrame(panel, fg_color="transparent")
        row.pack(fill="x", padx=20, pady=4)
        ctk.CTkLabel(row, text=item.label, font=UIM.fuente(12), text_color=C["texto"],
                     anchor="w").pack(side="left")
        entries[item.key] = ctk.CTkEntry(row, width=120, height=32, corner_radius=8,
                                         placeholder_text="€/mes")
        entries[item.key].pack(side="right")
        if item.key in current:
            entries[item.key].insert(0, f"{current[item.key]:.2f}".replace(".", ","))

    def save():
        connection = gestor_bd.conectar(str(app.ruta_bd_expedientes))
        try:
            saved = fixed_costs.save_fixed_costs(
                connection, app.id_comunidad, period_id,
                {key: entry.get() for key, entry in entries.items()},
            )
        except ValueError as error:
            messagebox.showwarning("Gastos fijos", str(error), parent=dialog)
            return
        finally:
            connection.close()
        dialog.destroy()
        yearly = sum(saved.values()) * 12
        app.log(f"Gastos fijos guardados: {_euros(yearly)} al año. Regenera el Excel si ya existía.", "ok")
        if getattr(app, "id_expediente", None):
            app._refrescar_expediente()

    buttons = ctk.CTkFrame(dialog, fg_color="transparent")
    buttons.pack(fill="x", padx=20, pady=(0, 14))
    ctk.CTkButton(buttons, text="Guardar", command=save, width=120, height=36, corner_radius=8,
                  font=UIM.fuente(12, "bold"), fg_color=C["primario"],
                  hover_color=C["primario_hover"]).pack(side="right")
    ctk.CTkButton(buttons, text="Cancelar", command=dialog.destroy, width=110, height=36,
                  corner_radius=8, **UIM.secondary_button_kwargs()).pack(side="right", padx=8)


def open_skip_source_dialog(app: "AppGestionFincas", issue: ReviewIssue) -> None:
    """Deja fuera del expediente la fuente de una incidencia, con su motivo.

    Hay documentos que sencillamente no deben entrar (un Excel de consulta, un
    PDF ilegible, un listado repetido). Sin esta salida, su incidencia era
    obligatoria y bloqueaba el expediente entero.
    """
    dialog = _dialog(app, "Omitir fuente", 560, 360)
    panel = ctk.CTkFrame(
        dialog, fg_color=C["panel"], corner_radius=16,
        border_width=1, border_color=C["borde"],
    )
    panel.pack(fill="both", expand=True, padx=18, pady=18)
    panel.grid_columnconfigure(0, weight=1)
    ctk.CTkLabel(
        panel, text="Omitir esta fuente", font=UIM.fuente(20, "bold"),
        text_color=C["texto"],
    ).grid(row=0, column=0, sticky="w", padx=22, pady=(22, 3))
    ctk.CTkLabel(
        panel,
        text=(
            f"{Path(issue.archived_path).name}\n\n"
            "El documento no se borra: se queda registrado como omitido, se "
            "cierran sus incidencias y deja de bloquear el expediente. Puedes "
            "recuperarlo más adelante reanalizando las fuentes."
        ),
        font=UIM.fuente(12), text_color=C["texto_sec"],
        wraplength=470, justify="left",
    ).grid(row=1, column=0, sticky="w", padx=22, pady=(0, 14))
    ctk.CTkLabel(
        panel, text="Motivo (queda registrado)", font=UIM.fuente(12, "bold"),
        text_color=C["texto"],
    ).grid(row=2, column=0, sticky="w", padx=22)
    motivo = ctk.CTkEntry(panel, height=34, corner_radius=8)
    motivo.grid(row=3, column=0, sticky="ew", padx=22, pady=(4, 16))
    motivo.insert(0, "No corresponde a este expediente")

    acciones = ctk.CTkFrame(panel, fg_color="transparent")
    acciones.grid(row=4, column=0, sticky="e", padx=22, pady=(0, 20))

    def confirmar():
        texto = motivo.get().strip()
        if not texto:
            messagebox.showwarning(
                "Falta el motivo", "Indica por qué se omite esta fuente.", parent=dialog,
            )
            return
        connection = gestor_bd.conectar(str(app.ruta_bd_expedientes))
        try:
            cerradas = document_review.ignore_source_document(
                connection, issue.id_case, issue.id_document,
                reason=texto, dismissed_by="usuario_local",
            )
        except (LookupError, ValueError) as error:
            messagebox.showerror("No se pudo omitir", str(error), parent=dialog)
            return
        finally:
            connection.close()
        dialog.destroy()
        app.log(
            f"Fuente omitida: {Path(issue.archived_path).name} "
            f"({cerradas} incidencia(s) cerradas) — {texto}",
            "aviso",
        )
        app._refrescar_expediente()

    ctk.CTkButton(
        acciones, text="Cancelar", width=100, height=32, corner_radius=8,
        fg_color="transparent", border_width=1, border_color=C["borde"],
        text_color=C["texto"], hover_color=C["acento_suave"],
        command=dialog.destroy,
    ).pack(side="left", padx=(0, 8))
    ctk.CTkButton(
        acciones, text="Omitir fuente", width=130, height=32, corner_radius=8,
        fg_color=C["primario"], hover_color=C["primario_hover"],
        command=confirmar,
    ).pack(side="left")



def _catalog_provider_labels() -> dict[str, str]:
    """Nombre visible → clave de cada proveedor del catálogo global."""
    import provider_registry

    try:
        registry = provider_registry.load_provider_registry(
            Path(__file__).resolve().parents[1] / "config" / "proveedores.json"
        )
    except (OSError, ValueError):
        return {}
    labels: dict[str, str] = {}
    for profile in registry.values():
        label = profile.display_name
        if label in labels:
            label = f"{label} ({profile.key})"
        labels[label] = profile.key
    return labels


def open_issue_dialog(app: "AppGestionFincas", issue: ReviewIssue) -> None:
    route = resolution_route_for_issue(issue)
    if route == "archived_source_duplicate":
        open_archived_path_resolution_dialog(app, issue.id_case)
        return
    if route == "archived_source_missing":
        messagebox.showwarning(
            "Archivo archivado no disponible",
            "No se encuentra una copia verificada de esta fuente. Vuelve a añadir el archivo original y después pulsa «Reevaluar fuentes».",
            parent=app,
        )
        return
    dialog = _dialog(app, "Resolver incidencia", 680, 720)
    panel = ctk.CTkFrame(
        dialog,
        fg_color=C["panel"],
        corner_radius=16,
        border_width=1,
        border_color=C["borde"],
    )
    panel.pack(fill="both", expand=True, padx=18, pady=18)
    panel.grid_columnconfigure(0, weight=1)
    ctk.CTkLabel(
        panel,
        text=(
            "Confirma la estimación" if route == "counter_reset_estimate"
            else "Revisa el dato pendiente"
        ),
        font=UIM.fuente(20, "bold"),
        text_color=C["texto"],
    ).grid(row=0, column=0, sticky="w", padx=22, pady=(22, 3))

    guidance = issue_guidance(issue.field_name)
    facts = [
        ("ARCHIVO", source_display_name(issue.archived_path)),
        ("DATO QUE NECESITAMOS", guidance["label"]),
        ("QUÉ BUSCAR", guidance["what_to_find"]),
        ("FORMATO", guidance["format"]),
        ("POR QUÉ SE PIDE", guidance["why"]),
    ]
    if issue.detected_value:
        facts.insert(2, ("VALOR DETECTADO", issue.detected_value))
    if issue_message(issue) != guidance["what_to_find"]:
        facts.append(("RESULTADO DEL ANÁLISIS", issue.message))
    facts_frame = ctk.CTkFrame(panel, fg_color=C["panel_2"], corner_radius=11)
    facts_frame.grid(row=1, column=0, sticky="ew", padx=22, pady=(8, 5))
    facts_frame.grid_columnconfigure(1, weight=1)
    for row, (label, value) in enumerate(facts):
        ctk.CTkLabel(
            facts_frame,
            text=label,
            font=UIM.fuente(9, "bold"),
            text_color=C["texto_sec"],
        ).grid(row=row, column=0, sticky="nw", padx=(12, 10), pady=7)
        ctk.CTkLabel(
            facts_frame,
            text=value,
            font=UIM.fuente(11),
            text_color=C["texto"],
            anchor="w",
            justify="left",
            wraplength=430,
        ).grid(row=row, column=1, sticky="w", padx=(0, 12), pady=5)

    source_actions = ctk.CTkFrame(panel, fg_color="transparent")
    source_actions.grid(row=2, column=0, sticky="w", padx=22, pady=(6, 0))
    ctk.CTkButton(
        source_actions, text="Ver en la fuente", command=lambda: open_source_preview(app, issue),
        height=34, corner_radius=8, **UIM.secondary_button_kwargs(),
    ).pack(side="left", padx=(0, 8))
    ctk.CTkButton(
        source_actions,
        text="Abrir archivo",
        command=lambda: open_archived_file(app, issue),
        height=34,
        corner_radius=8,
        fg_color="transparent",
        border_width=1,
        border_color=C["borde"],
        text_color=C["primario"],
        hover_color=C["acento_suave"],
    ).pack(side="left")

    value = None
    kind_by_label = {"Factura": "invoice", "Lectura": "reading", "Propietarios": "owners", "Otro documento": "other"}
    if route == "classify_unknown":
        value = tk.StringVar(value="Selecciona el tipo")
        ctk.CTkComboBox(
            panel, values=list(kind_by_label), variable=value, state="readonly",
        ).grid(row=3, column=0, sticky="ew", padx=22, pady=(14, 0))
        reason = _field(panel, "Motivo de la clasificación y fuente consultada", 4)
        action_row = 6
        action_text = "Guardar clasificación"
    elif route == "counter_reset_estimate":
        ctk.CTkLabel(
            panel,
            text=(
                "El contador bajó de valor. No se usará hasta que confirmes "
                "un consumo estimado del período; se guardará como una "
                "estimación aprobada y trazable."
            ),
            font=UIM.fuente(11),
            text_color=C["texto_sec"],
            justify="left",
            wraplength=520,
        ).grid(row=3, column=0, sticky="w", padx=22, pady=(14, 0))
        value = _field(panel, "Consumo estimado del período (m³)", 4)
        reason = _field(panel, "Motivo/soporte de la estimación", 6)
        action_row = 8
        action_text = "Aprobar estimación"
    elif route == "dismiss_invoice_outside_period":
        ctk.CTkLabel(
            panel,
            text=(
                "Esta factura queda fuera del período del expediente. Puedes "
                "cerrarla como no aplicable sin modificar la factura ni "
                "incorporarla al cálculo."
            ),
            font=UIM.fuente(11),
            text_color=C["texto_sec"],
            justify="left",
            wraplength=520,
        ).grid(row=3, column=0, sticky="w", padx=22, pady=(14, 0))
        reason = _field(panel, "Motivo del cierre", 4)
        action_row = 6
        action_text = "Cerrar: no corresponde al período"
    elif issue.field_name == "document.provider":
        provider_by_label = _catalog_provider_labels()
        ctk.CTkLabel(
            panel, text="PROVEEDOR DEL CATÁLOGO", font=UIM.fuente(10, "bold"),
            text_color=C["texto_sec"],
        ).grid(row=3, column=0, sticky="w", padx=22, pady=(13, 4))
        value = ctk.CTkComboBox(
            panel, values=sorted(provider_by_label), height=36, font=UIM.fuente(12),
        )
        value.set("")
        value.grid(row=4, column=0, sticky="ew", padx=22)
        ctk.CTkLabel(
            panel,
            text=(
                "Elige el emisor de la lista (puedes escribir para buscar). Se releerán sus "
                "fechas e importes con las reglas de ese proveedor y su CIF se recordará "
                "para reconocer sus próximas facturas."
            ),
            font=UIM.fuente(10), text_color=C["texto_sec"], wraplength=560, justify="left",
        ).grid(row=5, column=0, sticky="w", padx=22, pady=(4, 0))
        reason = _field(panel, "Nota (opcional)", 6)
        action_row = 8
        action_text = "Guardar proveedor"
    else:
        value = _field(panel, f"Valor confirmado · {guidance['label']}", 3)
        suggestion_connection = gestor_bd.conectar(str(app.ruta_bd_expedientes))
        try:
            suggestion = issue_suggestion(suggestion_connection, issue)
        finally:
            suggestion_connection.close()
        if suggestion:
            value.insert(0, suggestion)
            if not issue.detected_value:
                ctk.CTkLabel(
                    panel,
                    text="Valor propuesto por el análisis con poca seguridad: compruébalo en el archivo.",
                    font=UIM.fuente(10), text_color=C["aviso"],
                ).grid(row=5, column=0, sticky="w", padx=22, pady=(4, 0))
        reason = _field(
            panel, "Nota (opcional)", 6,
            placeholder="Si lo dejas vacío: «Comprobado en el documento original»",
        )
        action_row = 8
        action_text = "Guardar"

    actions = ctk.CTkFrame(panel, fg_color="transparent")
    actions.grid(row=action_row, column=0, sticky="ew", padx=22, pady=(20, 16))
    ctk.CTkButton(
        actions,
        text="Cancelar",
        command=dialog.destroy,
        height=38,
        corner_radius=9,
        fg_color="transparent",
        border_width=1,
        border_color=C["borde"],
        hover_color=C["acento_suave"],
        text_color=C["texto_sec"],
    ).pack(side="right")

    optional_reason = route in {"generic_correction", "dismiss_invoice_outside_period"}
    reanalyse_after = route == "classify_unknown" or issue.field_name == "document.provider"

    def save(open_next: bool = False):
        if app._procesando:
            return
        correction_reason = reason.get().strip()
        if not correction_reason and optional_reason:
            correction_reason = (
                "No corresponde al período del expediente"
                if route == "dismiss_invoice_outside_period"
                else "Comprobado en el documento original"
            )
        confirmed = value.get().strip() if value is not None else ""
        if route == "classify_unknown":
            confirmed = kind_by_label.get(confirmed, "")
        elif issue.field_name == "document.provider":
            confirmed = _catalog_provider_labels().get(confirmed, confirmed)
        if (route != "dismiss_invoice_outside_period" and not confirmed) or not correction_reason:
            messagebox.showwarning(
                "Datos requeridos",
                (
                    "Indica el consumo estimado y el soporte de la estimación."
                    if route == "counter_reset_estimate"
                    else "Indica el valor confirmado."
                ),
            )
            return

        connection = gestor_bd.conectar(str(app.ruta_bd_expedientes))
        try:
            if reanalyse_after:
                connection.execute("BEGIN IMMEDIATE")
            if route == "counter_reset_estimate":
                document_review.approve_counter_reset_estimate(
                    connection, issue.id_issue, consumption=confirmed,
                    reason=correction_reason, approved_by="usuario_local",
                )
            elif route == "dismiss_invoice_outside_period":
                document_review.dismiss_invoice_outside_period(
                    connection, issue.id_issue, reason=correction_reason,
                    dismissed_by="usuario_local",
                )
            else:
                document_review.resolve_issue(
                    connection, issue.id_issue, value=confirmed,
                    reason=correction_reason,
                )
                if reanalyse_after:
                    # Reuse the archived source and retain the manual decision.
                    # Missing invoice fields must exist before readiness is checked;
                    # a confirmed provider rereads dates and amounts with its rules.
                    case_ingestion.reanalyze_case_documents(connection, issue.id_case)
            remaining = len(document_review.list_open_issues(connection, issue.id_case))
            ready = None
            if not remaining and not document_review.case_has_unapplied_sources(connection, issue.id_case):
                ready = document_review.validate_case_ready(connection, issue.id_case)
            if reanalyse_after:
                connection.commit()
        except Exception as error:
            connection.rollback()
            messagebox.showwarning("No se pudo guardar", str(error))
            return
        finally:
            connection.close()
        dialog.destroy()
        app._refrescar_lista_expedientes(select_case_id=issue.id_case)
        app._refrescar_expediente()
        if ready is not None and ready.status == "ready_for_calculation":
            app.log("Listo para cálculo", "ok")
        else:
            app.log(f"{remaining} incidencia(s) por resolver", "aviso")
        if open_next and remaining:
            app.after(50, app._accion_resolver_incidencias)

    ctk.CTkButton(
        actions,
        text=action_text,
        command=save,
        height=38,
        corner_radius=9,
        font=UIM.fuente(12, "bold"),
        fg_color=C["exito"],
        hover_color=C["exito_hover"],
        border_width=1,
        border_color=C["exito"],
    ).pack(side="right", padx=(0, 8))
    if route != "counter_reset_estimate":
        ctk.CTkButton(
            actions, text="Guardar y siguiente", command=lambda: save(open_next=True),
            height=38, corner_radius=9, font=UIM.fuente(12),
            **UIM.secondary_button_kwargs(),
        ).pack(side="right", padx=(0, 8))
    dialog.bind("<Return>", lambda _event: save(open_next=True))
    (reason if route == "classify_unknown" else (value or reason)).focus_set()
