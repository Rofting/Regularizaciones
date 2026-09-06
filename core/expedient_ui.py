import os
import sys
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox
from typing import TYPE_CHECKING, Any, Mapping

import customtkinter as ctk

import case_ingestion
import community_onboarding
import document_review
import expedient_service
import gestor_bd
import ui_moderna as UIM
from expedient_models import ReviewIssue
from ui_moderna import C


if TYPE_CHECKING:
    from app import AppGestionFincas


def onboarding_step_route(state: Mapping[str, Any]) -> str:
    """Return the first onboarding stage that is safe to show next."""
    if not state.get("identity_valid", False):
        return "identity"
    if not state.get("has_sources", False):
        return "sources"

    step = state.get("step", "identity")
    required_answers_missing = (
        state.get("has_required_questions", False)
        and not state.get("answers_complete", False)
    )
    if step in {"detected", "confirmations", "summary"}:
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
        "has_required_questions": False,
        "answers_complete": False,
        "draft": None,
        "answers": {},
    }
    answer_variables: dict[str, tk.Variable] = {}
    busy = {"active": False}

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
        draft = state.get("draft")
        if draft is not None and (
            draft.community_code != identity["code"].get().strip()
            or draft.community_name != identity["name"].get().strip()
        ):
            state["draft"] = None
            state["has_sources"] = False
        state["identity_valid"] = True
        state["step"] = "identity"
        render(onboarding_step_route(state))

    def select_owners():
        path = filedialog.askopenfilename(
            parent=dialog,
            title="Selecciona la lista de propietarios",
            filetypes=(("Listado CSV", "*.csv"),),
        )
        if path:
            selected["owners"] = Path(path)
            source_labels["owners"].set(Path(path).name)
            state["draft"] = None
            state["has_sources"] = False

    def select_readings():
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
            state["draft"] = None
            state["has_sources"] = False

    def select_invoices():
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
            state["draft"] = None
            state["has_sources"] = False

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
        state["step"] = "detection"
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
                    state["step"] = "sources"
                    render("sources")
                    messagebox.showerror(
                        "No se pudieron analizar las fuentes", str(error), parent=dialog
                    )
                app.after(0, failed)
                return

            def analysed():
                busy["active"] = False
                state["draft"] = draft
                state["step"] = "detected"
                state["has_required_questions"] = any(
                    question.required for question in draft.questions
                )
                answer_variables.clear()
                for module in draft.detected_modules:
                    answer_variables[f"module:{module}"] = tk.BooleanVar(
                        master=dialog, value=True
                    )
                for question in draft.questions:
                    answer_variables[question.key] = tk.StringVar(master=dialog)
                state["answers_complete"] = not state["has_required_questions"]
                render("detection")
            app.after(0, analysed)

        app._en_hilo(work)

    def collect_answers():
        draft = state["draft"]
        answers = {
            key: variable.get() for key, variable in answer_variables.items()
        }
        missing = [
            question.prompt
            for question in draft.questions
            if question.required and not str(answers.get(question.key, "")).strip()
        ]
        state["answers"] = answers
        state["answers_complete"] = not missing
        state["step"] = "confirmations"
        if missing:
            messagebox.showwarning(
                "Confirmaciones pendientes",
                "Responde todas las preguntas obligatorias antes de continuar.",
                parent=dialog,
            )
            return
        render(onboarding_step_route(state))

    def publish():
        state["step"] = "summary"
        if onboarding_step_route(state) != "summary":
            render(onboarding_step_route(state))
            return
        if getattr(app, "_procesando", False):
            messagebox.showwarning(
                "Espera", "Ya hay una operación en curso.", parent=dialog
            )
            return
        try:
            period_name, start_date, end_date = parse_period()
        except ValueError as exc:
            messagebox.showwarning("Período incompleto", str(exc), parent=dialog)
            return
        busy["active"] = True
        clear(footer)
        ctk.CTkLabel(
            footer,
            text="Creando comunidad y archivando las fuentes…",
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
                    draft=state["draft"],
                    answers=dict(state["answers"]),
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
                code = state["draft"].community_code
                name = state["draft"].community_name
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
                    "La comunidad, el perfil local y sus fuentes se han registrado.",
                    parent=app,
                )
            app.after(0, completed)

        app._en_hilo(work)

    def render(step: str):
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
                "Desactiva los servicios que no correspondan y responde cada dato incierto. No se inventa ningún valor.",
            )
            if draft.detected_modules:
                ctk.CTkLabel(
                    content,
                    text="Servicios que se incluirán",
                    font=UIM.fuente(11, "bold"),
                    text_color=C["texto_sec"],
                ).pack(anchor="w", pady=(2, 6))
                for module in draft.detected_modules:
                    ctk.CTkCheckBox(
                        content,
                        text=module.title().replace("Calefaccion", "Calefacción"),
                        variable=answer_variables[f"module:{module}"],
                        font=UIM.fuente(12),
                        fg_color=C["primario"],
                        hover_color=C["primario_hover"],
                        border_color=C["borde"],
                    ).pack(anchor="w", pady=4)
            else:
                ctk.CTkLabel(
                    content,
                    text="No se detectaron servicios. El perfil se creará sin conceptos activos.",
                    font=UIM.fuente(11),
                    text_color=C["aviso"],
                ).pack(anchor="w", pady=(2, 10))
            for question in draft.questions:
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
        active_modules = [
            module for module in draft.detected_modules
            if state["answers"].get(f"module:{module}") is True
        ]
        period_name, _start, _end = parse_period()
        section_title(
            "Revisa y crea la comunidad",
            "Al confirmar se guardarán el perfil y la plantilla locales, y se archivarán copias de las fuentes.",
        )
        summary = (
            ("Comunidad", f"{draft.community_code} — {draft.community_name}"),
            ("Fuentes que se archivarán", str(len(draft.sources))),
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
        footer_buttons(
            back=lambda: render("confirmations"),
            next_text="Crear comunidad",
            next_command=publish,
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

    labels = ("Facturas", "Lecturas", "Propietarios", "Otros")
    selected_kind = tk.StringVar(value=labels[0])
    ctk.CTkSegmentedButton(
        panel,
        values=list(labels),
        variable=selected_kind,
        height=38,
        font=UIM.fuente(11, "bold"),
        selected_color=C["primario"],
        selected_hover_color=C["primario_hover"],
        unselected_color=C["panel_2"],
        unselected_hover_color=C["acento_suave_hover"],
    ).pack(fill="x", padx=22)

    def select_files():
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
        if not paths:
            return
        kind_by_label = {
            "Facturas": "invoice",
            "Lecturas": "reading",
            "Propietarios": "owners",
            "Otros": "other",
        }
        document_kind = kind_by_label[selected_kind.get()]
        dialog.destroy()

        def add_all():
            created_count = 0
            duplicate_count = 0
            errors = []
            total = len(paths)
            for index, path in enumerate(paths, start=1):
                filename = Path(path).name
                connection = None
                app._estado(
                    f"Incorporando {index} de {total}: {filename}",
                    procesando=True,
                )
                app.log(f"Fuente {index} de {total}: {filename}", "info")
                try:
                    connection = gestor_bd.conectar(str(app.ruta_bd_expedientes))
                    result = case_ingestion.add_document_to_case(
                        connection,
                        case_id,
                        source_path=path,
                        archive_root=app.ruta_archivo_expedientes,
                        document_kind=document_kind,
                        candidates={
                            "nombre_archivo": Path(path).name,
                            "tipo_documento": document_kind,
                        },
                        required_fields=(
                            ("fecha_inicio", "fecha_fin", "importe_total")
                            if document_kind == "invoice"
                            else ()
                        ),
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
            app.log(
                f"Fuentes añadidas: {created_count} nueva(s), "
                f"{duplicate_count} duplicada(s), {len(errors)} con error",
                "ok" if created_count and not errors else "aviso",
            )
            def refresh_case():
                app._refrescar_lista_expedientes(select_case_id=case_id)
                app._refrescar_expediente()

            app.after(0, refresh_case)

        app._estado(f"Incorporando 0 de {len(paths)} fuentes", procesando=True)
        app.after(0, lambda: app._en_hilo(add_all))

    ctk.CTkButton(
        panel,
        text="Seleccionar archivos",
        command=select_files,
        height=46,
        corner_radius=11,
        font=UIM.fuente(13, "bold"),
        fg_color=C["primario"],
        hover_color=C["primario_hover"],
    ).pack(anchor="w", padx=22, pady=(22, 8))
    ctk.CTkLabel(
        panel,
        text="PDF · XLSX · XLS · CSV",
        font=UIM.fuente(10),
        text_color=C["texto_sec"],
    ).pack(anchor="w", padx=22)


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
    return "generic_correction"


def open_issue_dialog(app: "AppGestionFincas", issue: ReviewIssue) -> None:
    route = resolution_route_for_issue(issue)
    dialog = _dialog(app, "Resolver incidencia", 620, 620)
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

    facts = (
        ("ARCHIVO", issue.archived_path.name),
        ("CAMPO", issue.field_name),
        ("VALOR DETECTADO", issue.detected_value or "—"),
        ("EXPLICACIÓN", issue.message),
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

    ctk.CTkButton(
        panel,
        text="Abrir archivo",
        command=lambda: open_archived_file(app, issue),
        height=34,
        corner_radius=8,
        fg_color="transparent",
        border_width=1,
        border_color=C["borde"],
        text_color=C["primario"],
        hover_color=C["acento_suave"],
    ).grid(row=2, column=0, sticky="w", padx=22, pady=(6, 0))

    value = None
    if route == "counter_reset_estimate":
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
        value = _field(panel, "Valor confirmado", 3)
        reason = _field(panel, "Motivo de la corrección", 5)
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
        correction_reason = reason.get().strip()
        confirmed = value.get().strip() if value is not None else ""
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
            remaining = len(document_review.list_open_issues(connection, issue.id_case))
            ready = None
            if not remaining:
                ready = document_review.validate_case_ready(connection, issue.id_case)
        except (LookupError, ValueError) as error:
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
        fg_color=C["primario"],
        hover_color=C["primario_hover"],
    ).pack(side="right", padx=(0, 8))
    (value or reason).focus_set()
