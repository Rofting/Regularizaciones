import os
import json
import sqlite3
import stat
import sys
import tkinter as tk
from collections import Counter
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox
from typing import TYPE_CHECKING, Any, Mapping, Sequence, TypeVar

import customtkinter as ctk

import case_ingestion
import cuotas_servicio
import community_discovery
import community_onboarding
import document_review
import expedient_service
import gestor_bd
import source_batch
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
    has_registered_template: bool = True,
) -> GuidedWorkspaceState:
    """Returns the next safe user action without replacing workflow service gates."""
    if not has_case:
        active_step, next_action = "fuentes", "crear_expediente"
        headline = "Crear expediente"
        detail = "Elige el intervalo que vas a regularizar antes de incorporar fuentes."
    elif open_issue_count:
        active_step, next_action = "validar", "resolver_incidencias"
        headline = "Resuelve las incidencias"
        detail = f"Hay {open_issue_count} decisión(es) pendiente(s) antes de continuar."
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
            "Puedes abrir las salidas o repetir Excel, reparto o cartas. "
            "Cada repetición conserva las ejecuciones anteriores en el historial."
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
    )


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
            def failed():
                app._estado("No se pudo analizar la carpeta", procesando=False)
                messagebox.showerror("No se pudo analizar la carpeta", str(error), parent=app)
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
                def failed():
                    app._estado("No se pudieron crear las comunidades", procesando=False)
                    for widget in footer.winfo_children():
                        widget.configure(state="normal")
                    messagebox.showerror("No se pudieron crear las comunidades", str(error), parent=dialog)
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
    default = {
        "label": field_name.replace("_", " ").capitalize(),
        "what_to_find": "Busca este dato en el documento original antes de confirmarlo.",
        "format": "Copia el valor con el formato que aparece en la fuente.",
        "why": "Es necesario para mantener trazabilidad y evitar un cálculo con datos incompletos.",
    }
    return guides.get(field_name, default)


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
            filetypes=(("Listado CSV", "*.csv"),),
        )
        if path:
            selected["owners"] = Path(path)
            source_labels["owners"].set(Path(path).name)
            invalidate_analysis()

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
            first = entry_field("Código de comunidad", identity["code"], "Ej. 644")
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
                "Un archivo CSV con las personas y sus viviendas.",
                source_labels["owners"], select_owners, "Elegir CSV",
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


def open_create_case_dialog(app: "AppGestionFincas") -> None:
    if not getattr(app, "id_comunidad", None):
        messagebox.showwarning(
            "Comunidad requerida",
            "Selecciona una comunidad antes de crear un expediente.",
        )
        return

    dialog = _dialog(app, "Nuevo expediente", 500, 470)
    content = ctk.CTkFrame(
        dialog,
        fg_color=C["panel"],
        corner_radius=16,
        border_width=1,
        border_color=C["borde"],
    )
    content.pack(fill="both", expand=True, padx=18, pady=18)
    content.grid_columnconfigure(0, weight=1)
    ctk.CTkLabel(
        content,
        text="Abre un expediente",
        font=UIM.fuente(20, "bold"),
        text_color=C["texto"],
    ).grid(row=0, column=0, sticky="w", padx=22, pady=(22, 1))
    ctk.CTkLabel(
        content,
        text="Define el intervalo real que vas a revisar.",
        font=UIM.fuente(11),
        text_color=C["texto_sec"],
    ).grid(row=1, column=0, sticky="w", padx=22, pady=(0, 4))

    name = _field(content, "Nombre", 2, placeholder="Ej. Regularización enero")
    start = _field(content, "Fecha inicial", 4, placeholder="DD/MM/AAAA")
    end = _field(content, "Fecha final", 6, placeholder="DD/MM/AAAA")

    actions = ctk.CTkFrame(content, fg_color="transparent")
    actions.grid(row=8, column=0, sticky="ew", padx=22, pady=(22, 18))
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

    def save():
        try:
            start_date = datetime.strptime(start.get().strip(), "%d/%m/%Y").date()
            end_date = datetime.strptime(end.get().strip(), "%d/%m/%Y").date()
        except ValueError:
            messagebox.showwarning(
                "Fecha no válida", "Usa el formato DD/MM/AAAA."
            )
            return
        if end_date < start_date:
            messagebox.showwarning(
                "Rango no válido",
                "La fecha de fin debe ser posterior o igual a la de inicio.",
            )
            return

        connection = gestor_bd.conectar(str(app.ruta_bd_expedientes))
        try:
            case = expedient_service.create_case(
                connection,
                app.id_comunidad,
                name=name.get().strip(),
                start_date=start_date,
                end_date=end_date,
            )
        finally:
            connection.close()
        app.id_expediente = case.id_case
        app.expediente_actual.set(case.name)
        dialog.destroy()
        app._refrescar_lista_expedientes(select_case_id=case.id_case)
        app._refrescar_expediente()
        duration = (end_date - start_date).days + 1
        app.log(
            f"Expediente '{case.name}' creado · {duration} día(s) incluidos",
            "ok",
        )

    ctk.CTkButton(
        actions,
        text="Crear expediente",
        command=save,
        height=38,
        corner_radius=9,
        font=UIM.fuente(12, "bold"),
        fg_color=C["primario"],
        hover_color=C["primario_hover"],
    ).pack(side="right", padx=(0, 8))
    name.focus_set()


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
            app.log(
                f"Fuentes añadidas: {created_count} nueva(s), "
                f"{duplicate_count} duplicada(s), {len(errors)} con error",
                "ok" if created_count and not errors else "aviso",
            )
            def refresh_case():
                app._refrescar_lista_expedientes(select_case_id=case_id)
                app._refrescar_expediente()
                messagebox.showinfo(
                    "Fuentes analizadas",
                    f"{source_summary(kinds)}\n\n{created_count} nuevas · {duplicate_count} duplicadas · {len(errors)} con error\n"
                    "Revisa las incidencias del expediente para confirmar los datos pendientes."
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


def open_service_fees_dialog(app: "AppGestionFincas", servicio: str = "ACS") -> None:
    """Cuotas que la comunidad cobra a los vecinos por ACS o calefacción.

    Es el lado de los ingresos del estudio y el único dato que no sale de
    ninguna factura ni de ningún contador: lo decide la comunidad. Sin él, el
    análisis no puede comparar lo cobrado con lo que ha costado el servicio.
    """
    if not app.id_comunidad or not app.id_periodo:
        messagebox.showinfo(
            "Cuotas cobradas",
            "Elige antes una comunidad y un período.",
            parent=app,
        )
        return

    dialog = _dialog(app, f"Cuotas cobradas de {servicio}", 720, 640)
    panel = ctk.CTkFrame(
        dialog, fg_color=C["panel"], corner_radius=16,
        border_width=1, border_color=C["borde"],
    )
    panel.pack(fill="both", expand=True, padx=18, pady=18)
    panel.grid_columnconfigure(0, weight=1)
    panel.grid_rowconfigure(6, weight=1)

    ctk.CTkLabel(
        panel, text=f"Cuotas cobradas de {servicio}", font=UIM.fuente(20, "bold"),
        text_color=C["texto"],
    ).grid(row=0, column=0, sticky="w", padx=22, pady=(22, 2))
    ctk.CTkLabel(
        panel,
        text=(
            "Lo que la comunidad gira a los vecinos durante el período. La cuota "
            "fija mensual se genera de una vez a partir de sus tramos; las "
            "liquidaciones por consumo se añaden una a una."
        ),
        font=UIM.fuente(12), text_color=C["texto_sec"], wraplength=640, justify="left",
    ).grid(row=1, column=0, sticky="w", padx=22, pady=(0, 14))

    # --- cuota fija mensual --------------------------------------------------
    fijas = ctk.CTkFrame(panel, fg_color=C["fondo"], corner_radius=10)
    fijas.grid(row=2, column=0, sticky="ew", padx=22)
    fijas.grid_columnconfigure(1, weight=1)
    ctk.CTkLabel(
        fijas, text="Cuota fija mensual", font=UIM.fuente(13, "bold"), text_color=C["texto"],
    ).grid(row=0, column=0, columnspan=4, sticky="w", padx=14, pady=(12, 2))
    ctk.CTkLabel(
        fijas,
        text="Un tramo por cada cambio de importe: mes de inicio e importe del recibo.",
        font=UIM.fuente(11), text_color=C["texto_sec"],
    ).grid(row=1, column=0, columnspan=4, sticky="w", padx=14, pady=(0, 8))

    tramos: list[tuple[ctk.CTkEntry, ctk.CTkEntry]] = []
    contenedor_tramos = ctk.CTkFrame(fijas, fg_color="transparent")
    contenedor_tramos.grid(row=2, column=0, columnspan=4, sticky="ew", padx=14)
    contenedor_tramos.grid_columnconfigure((1, 3), weight=1)

    def anadir_tramo(desde: str = "", importe: str = "") -> None:
        fila = len(tramos)
        ctk.CTkLabel(
            contenedor_tramos, text="Desde (aaaa-mm)", font=UIM.fuente(11),
            text_color=C["texto_sec"],
        ).grid(row=fila, column=0, sticky="w", pady=3)
        mes = ctk.CTkEntry(contenedor_tramos, height=30, corner_radius=8, width=120)
        mes.grid(row=fila, column=1, sticky="w", padx=(6, 18), pady=3)
        mes.insert(0, desde)
        ctk.CTkLabel(
            contenedor_tramos, text="Importe (€/mes)", font=UIM.fuente(11),
            text_color=C["texto_sec"],
        ).grid(row=fila, column=2, sticky="w", pady=3)
        valor = ctk.CTkEntry(contenedor_tramos, height=30, corner_radius=8, width=120)
        valor.grid(row=fila, column=3, sticky="w", padx=6, pady=3)
        valor.insert(0, importe)
        tramos.append((mes, valor))

    connection = gestor_bd.conectar(str(app.ruta_bd_expedientes))
    try:
        periodo = connection.execute(
            "SELECT fecha_inicio,fecha_fin FROM periodos WHERE id_periodo=?",
            (app.id_periodo,),
        ).fetchone()
    finally:
        connection.close()
    anadir_tramo(str(periodo["fecha_inicio"])[:7] if periodo else "", "")
    anadir_tramo()

    ctk.CTkButton(
        fijas, text="Añadir tramo", width=120, height=28, corner_radius=8,
        font=UIM.fuente(11), command=lambda: anadir_tramo(),
        **UIM.secondary_button_kwargs(),
    ).grid(row=3, column=0, sticky="w", padx=14, pady=(8, 12))

    # --- liquidación por consumo --------------------------------------------
    variables = ctk.CTkFrame(panel, fg_color=C["fondo"], corner_radius=10)
    variables.grid(row=3, column=0, sticky="ew", padx=22, pady=(12, 0))
    variables.grid_columnconfigure((1, 3, 5), weight=1)
    ctk.CTkLabel(
        variables, text="Liquidación por consumo", font=UIM.fuente(13, "bold"),
        text_color=C["texto"],
    ).grid(row=0, column=0, columnspan=6, sticky="w", padx=14, pady=(12, 2))
    ctk.CTkLabel(
        variables, text="Lo cobrado tras cada lectura de contadores.",
        font=UIM.fuente(11), text_color=C["texto_sec"],
    ).grid(row=1, column=0, columnspan=6, sticky="w", padx=14, pady=(0, 8))
    for columna, etiqueta in ((0, "Fecha"), (2, "Importe (€)"), (4, "Consumo")):
        ctk.CTkLabel(
            variables, text=etiqueta, font=UIM.fuente(11), text_color=C["texto_sec"],
        ).grid(row=2, column=columna, sticky="w", padx=(14 if columna == 0 else 0, 4))
    fecha_var = ctk.CTkEntry(variables, height=30, corner_radius=8, width=120)
    fecha_var.grid(row=2, column=1, sticky="w", padx=(0, 14))
    importe_var = ctk.CTkEntry(variables, height=30, corner_radius=8, width=110)
    importe_var.grid(row=2, column=3, sticky="w", padx=(0, 14))
    consumo_var = ctk.CTkEntry(variables, height=30, corner_radius=8, width=110)
    consumo_var.grid(row=2, column=5, sticky="w", padx=(0, 14))

    resumen_texto = ctk.CTkLabel(
        panel, text="", font=UIM.fuente(12, "bold"), text_color=C["texto"], justify="left",
    )
    resumen_texto.grid(row=4, column=0, sticky="w", padx=22, pady=(14, 4))

    listado = ctk.CTkTextbox(panel, height=170, corner_radius=8, font=UIM.fuente(11))
    listado.grid(row=6, column=0, sticky="nsew", padx=22, pady=(0, 10))

    def refrescar() -> None:
        conexion = gestor_bd.conectar(str(app.ruta_bd_expedientes))
        try:
            apuntes = cuotas_servicio.listar(
                conexion, community_id=app.id_comunidad,
                period_id=app.id_periodo, servicio=servicio,
            )
            total = cuotas_servicio.resumen(
                conexion, community_id=app.id_comunidad,
                period_id=app.id_periodo, servicio=servicio,
            )
        finally:
            conexion.close()
        resumen_texto.configure(
            text=(
                f"Cobrado: {total.fija:,.2f} € de cuota fija + {total.variable:,.2f} € "
                f"por consumo = {total.total:,.2f} €   ({total.apuntes} apunte(s))"
            ).replace(",", " ")
        )
        listado.configure(state="normal")
        listado.delete("1.0", "end")
        for apunte in apuntes:
            extra = f"  ·  consumo {apunte['consumo']}" if apunte["consumo"] is not None else ""
            listado.insert(
                "end",
                f"{apunte['fecha']}   {apunte['concepto']:<9} {apunte['importe']:>10.2f} €{extra}\n",
            )
        if not apuntes:
            listado.insert("end", "Todavía no hay cuotas anotadas para este período.\n")
        listado.configure(state="disabled")

    def generar_mensuales() -> None:
        pares = []
        for mes, valor in tramos:
            texto_mes, texto_valor = mes.get().strip(), valor.get().strip()
            if not texto_mes and not texto_valor:
                continue
            if not texto_mes or not texto_valor:
                messagebox.showwarning(
                    "Tramo incompleto",
                    "Cada tramo necesita su mes de inicio y su importe.",
                    parent=dialog,
                )
                return
            pares.append((f"{texto_mes}-01" if len(texto_mes) == 7 else texto_mes, texto_valor))
        if not pares:
            messagebox.showwarning(
                "Sin tramos", "Indica al menos un importe mensual.", parent=dialog,
            )
            return
        conexion = gestor_bd.conectar(str(app.ruta_bd_expedientes))
        try:
            creadas = cuotas_servicio.generar_cuotas_mensuales(
                conexion, community_id=app.id_comunidad, period_id=app.id_periodo,
                servicio=servicio, tramos=pares, notas="Cuota mensual del recibo",
            )
            conexion.commit()
        except (ValueError, LookupError) as error:
            messagebox.showerror("No se pudo guardar", str(error), parent=dialog)
            return
        finally:
            conexion.close()
        app.log(f"Cuotas mensuales de {servicio}: {creadas} mes(es) anotados.", "ok")
        refrescar()

    def anadir_liquidacion() -> None:
        conexion = gestor_bd.conectar(str(app.ruta_bd_expedientes))
        try:
            cuotas_servicio.registrar_cuota(
                conexion, community_id=app.id_comunidad, period_id=app.id_periodo,
                servicio=servicio, concepto="variable",
                fecha=fecha_var.get().strip(), importe=importe_var.get().strip(),
                consumo=consumo_var.get().strip() or None,
                notas="Liquidación por consumo",
            )
            conexion.commit()
        except (ValueError, LookupError) as error:
            messagebox.showerror("No se pudo guardar", str(error), parent=dialog)
            return
        finally:
            conexion.close()
        for campo in (fecha_var, importe_var, consumo_var):
            campo.delete(0, "end")
        refrescar()

    ctk.CTkButton(
        fijas, text="Generar cuotas del período", width=200, height=30, corner_radius=8,
        font=UIM.fuente(11), fg_color=C["primario"], hover_color=C["primario_hover"],
        command=generar_mensuales,
    ).grid(row=3, column=1, columnspan=3, sticky="e", padx=14, pady=(8, 12))
    ctk.CTkButton(
        variables, text="Añadir liquidación", width=160, height=30, corner_radius=8,
        font=UIM.fuente(11), fg_color=C["primario"], hover_color=C["primario_hover"],
        command=anadir_liquidacion,
    ).grid(row=3, column=0, columnspan=6, sticky="e", padx=14, pady=(8, 12))

    ctk.CTkButton(
        panel, text="Cerrar", width=110, height=32, corner_radius=8,
        command=dialog.destroy, **UIM.secondary_button_kwargs(),
    ).grid(row=7, column=0, sticky="e", padx=22, pady=(0, 18))

    refrescar()



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



def open_issue_dialog(app: "AppGestionFincas", issue: ReviewIssue) -> None:
    route = resolution_route_for_issue(issue)
    if route == "archived_source_duplicate":
        open_archived_path_resolution_dialog(app, issue.id_case)
        return
    if route == "archived_source_missing":
        messagebox.showwarning(
            "Archivo archivado no disponible",
            "No se encuentra una copia verificada de esta fuente. Vuelve a añadir el archivo original y después pulsa Reanalizar fuentes.",
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
    facts = (
        ("ARCHIVO", issue.archived_path.name),
        ("DATO QUE NECESITAMOS", guidance["label"]),
        ("VALOR DETECTADO", issue.detected_value or "—"),
        ("QUÉ BUSCAR", guidance["what_to_find"]),
        ("FORMATO", guidance["format"]),
        ("POR QUÉ SE PIDE", guidance["why"]),
        ("RESULTADO DEL ANÁLISIS", issue.message),
    )
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
            wraplength=390,
        ).grid(row=row, column=1, sticky="w", padx=(0, 12), pady=7)

    source_actions = ctk.CTkFrame(panel, fg_color="transparent")
    source_actions.grid(row=2, column=0, sticky="w", padx=22, pady=(6, 0))
    ctk.CTkButton(
        source_actions, text="Ver contexto", command=lambda: show_issue_context(app, issue),
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
    else:
        value = _field(panel, f"Valor confirmado · {guidance['label']}", 3)
        reason = _field(panel, "Motivo de la corrección y fuente consultada", 5)
        action_row = 7
        action_text = "Guardar corrección"

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

    def save():
        if app._procesando:
            return
        correction_reason = reason.get().strip()
        confirmed = value.get().strip() if value is not None else ""
        if route == "classify_unknown":
            confirmed = kind_by_label.get(confirmed, "")
        if (route != "dismiss_invoice_outside_period" and not confirmed) or not correction_reason:
            messagebox.showwarning(
                "Datos requeridos",
                (
                    "Indica el consumo estimado y el soporte de la estimación."
                    if route == "counter_reset_estimate"
                    else "Indica el valor confirmado y el motivo de la corrección."
                ),
            )
            return

        connection = gestor_bd.conectar(str(app.ruta_bd_expedientes))
        try:
            if route == "classify_unknown":
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
                if route == "classify_unknown":
                    # Reuse the archived source and retain the manual decision.
                    # Missing invoice fields must exist before readiness is checked.
                    case_ingestion.reanalyze_case_documents(connection, issue.id_case)
            remaining = len(document_review.list_open_issues(connection, issue.id_case))
            ready = None
            if not remaining and not document_review.case_has_unapplied_sources(connection, issue.id_case):
                ready = document_review.validate_case_ready(connection, issue.id_case)
            if route == "classify_unknown":
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
    (reason if route == "classify_unknown" else (value or reason)).focus_set()
