import os
import sys
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox
from typing import TYPE_CHECKING

import customtkinter as ctk

import case_ingestion
import document_review
import expedient_service
import gestor_bd
import ui_moderna as UIM
from expedient_models import ReviewIssue
from ui_moderna import C


if TYPE_CHECKING:
    from app import AppGestionFincas


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
            for path in paths:
                connection = gestor_bd.conectar(str(app.ruta_bd_expedientes))
                try:
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
                finally:
                    connection.close()
                if result.created:
                    created_count += 1
                else:
                    duplicate_count += 1
            app.log(
                f"Fuentes añadidas: {created_count} nueva(s), "
                f"{duplicate_count} duplicada(s)",
                "ok" if created_count else "aviso",
            )
            app.after(0, app._refrescar_expediente)

        app._en_hilo(add_all)

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


def open_issue_dialog(app: "AppGestionFincas", issue: ReviewIssue) -> None:
    dialog = _dialog(app, "Resolver incidencia", 620, 590)
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
        text="Revisa el dato pendiente",
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

    def open_file():
        try:
            if sys.platform != "win32":
                raise OSError("Abrir el archivo solo está disponible en Windows.")
            os.startfile(str(issue.archived_path))
        except Exception as exc:
            messagebox.showerror("No se pudo abrir el archivo", str(exc))

    ctk.CTkButton(
        panel,
        text="Abrir archivo",
        command=open_file,
        height=34,
        corner_radius=8,
        fg_color="transparent",
        border_width=1,
        border_color=C["borde"],
        text_color=C["primario"],
        hover_color=C["acento_suave"],
    ).grid(row=2, column=0, sticky="w", padx=22, pady=(6, 0))

    value = _field(panel, "Valor confirmado", 3)
    reason = _field(panel, "Motivo de la corrección", 5)

    actions = ctk.CTkFrame(panel, fg_color="transparent")
    actions.grid(row=7, column=0, sticky="ew", padx=22, pady=(20, 16))
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
        confirmed = value.get().strip()
        correction_reason = reason.get().strip()
        if not confirmed or not correction_reason:
            messagebox.showwarning(
                "Datos requeridos",
                "Indica el valor confirmado y el motivo de la corrección.",
            )
            return

        connection = gestor_bd.conectar(str(app.ruta_bd_expedientes))
        try:
            document_review.resolve_issue(
                connection,
                issue.id_issue,
                value=confirmed,
                reason=correction_reason,
            )
            remaining = len(document_review.list_open_issues(connection, issue.id_case))
            ready = None
            if not remaining:
                ready = document_review.validate_case_ready(connection, issue.id_case)
        finally:
            connection.close()
        dialog.destroy()
        app._refrescar_expediente()
        if ready is not None and ready.status == "ready_for_calculation":
            app.log("Listo para cálculo", "ok")
        else:
            app.log(f"{remaining} incidencia(s) por resolver", "aviso")

    ctk.CTkButton(
        actions,
        text="Guardar corrección",
        command=save,
        height=38,
        corner_radius=9,
        font=UIM.fuente(12, "bold"),
        fg_color=C["primario"],
        hover_color=C["primario_hover"],
    ).pack(side="right", padx=(0, 8))
    value.focus_set()
