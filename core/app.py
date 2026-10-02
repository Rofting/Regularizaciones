"""
app.py
======
Ventana principal de Regularizaciones (CustomTkinter, modo claro/oscuro).

Flujo guiado por expediente:
    1. Elige la comunidad y crea un expediente con el intervalo a regularizar.
    2. Fuentes: añade facturas, lecturas y listados (PDF, Excel o CSV).
    3. Validar: resuelve las incidencias y confirma las fuentes.
    4. Reparto: genera el Excel oficial y calcula el reparto por vecino.
    5. Cartas: genera una carta por propietario con sus gráficas de consumo.

Cualquier etapa ya superada puede repetirse; las ejecuciones anteriores se
conservan en el historial del expediente.

USO:
    python app.py
"""

import os
import sys
import sqlite3
from contextlib import closing
import threading
import traceback
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog

# Si se lanza desde una consola normal de Windows (doble clic en un .bat,
# o "python app.py" desde cmd), la consola usa cp1252 y cualquier emoji en
# un print() de los módulos internos (lector_pdf, carta_writer...) revienta
# con UnicodeEncodeError incluso con la ventana ya abierta. Ver el mismo
# arreglo en pipeline.py para más detalle.
for _stream in (sys.stdout, sys.stderr):
    if hasattr(_stream, "reconfigure"):
        try:
            _stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass

# ---------------------------------------------------------------------------
# CUSTOMTKINTER (incluido en la instalación)
# ---------------------------------------------------------------------------
try:
    import customtkinter as ctk
except ImportError:
    raise SystemExit("Falta customtkinter en esta instalación. Reinstala el programa.")

from app_paths import ApplicationPaths
import ui_moderna as UIM
from ui_moderna import C

# ---------------------------------------------------------------------------
# RUTAS POR DEFECTO
# ---------------------------------------------------------------------------
APP_PATHS = ApplicationPaths.resolve()
BASE_DIR = APP_PATHS.home

RUTA_BD         = BASE_DIR / "data" / "gestion.db"
RUTA_PLANTILLA  = BASE_DIR / "plantillas" / "Plantilla_Cartas.docx"
RUTA_CARTAS     = BASE_DIR / "salidas" / "cartas"
RUTA_PROVEEDORES= BASE_DIR / "config" / "proveedores.json"
RUTA_PROVEEDORES_DESPACHO = BASE_DIR / "config" / "proveedores_despacho.json"


# ---------------------------------------------------------------------------
# IMPORTAR MÓDULOS DEL SISTEMA (con manejo de errores si faltan)
# ---------------------------------------------------------------------------
def _importar_modulos():
    """Intenta importar los módulos core. Devuelve dict con los que están disponibles."""
    core_dir = Path(__file__).parent
    if str(core_dir) not in sys.path:
        sys.path.insert(0, str(core_dir))

    modulos = {}
    for nombre in ["gestor_bd", "lector_pdf", "motor_reparto",
                   "excel_writer", "carta_writer", "importar_lecturas_metrigest",
                   "importar_excel_maestro", "letter_settings", "regularization_flow",
                   "expedient_service", "document_review", "case_ingestion",
                   "expedient_ui", "case_workflow_actions", "case_readiness",
                   "database_reset", "database_backup", "office_settings"]:
        try:
            modulos[nombre] = __import__(nombre)
        except ImportError:
            modulos[nombre] = None
    return modulos


MOD = _importar_modulos()


# ---------------------------------------------------------------------------
# MOVIMIENTO SEGURO DE ARCHIVOS (evita WinError 32 si el archivo está abierto)
# ---------------------------------------------------------------------------


# ---------------------------------------------------------------------------
# VENTANA PRINCIPAL
# ---------------------------------------------------------------------------
class AppGestionFincas(ctk.CTk):
    def __init__(self):
        APP_PATHS.prepare()
        super().__init__(fg_color=C["fondo"])
        self.title("Regularización de facturas")
        self.geometry("1180x760")
        self.minsize(1020, 680)
        self.resizable(True, True)

        # Estado
        self.comunidad_actual  = tk.StringVar(value="")
        self.expediente_actual  = tk.StringVar(value="")
        self.expediente_seleccionado = tk.StringVar(value="")
        self.id_comunidad       = None
        self.id_periodo         = None
        self.id_expediente      = None
        self.ruta_bd_expedientes = RUTA_BD
        self.ruta_archivo_expedientes = BASE_DIR / "data" / "expedientes"
        self._procesando        = False
        self._issues_page       = 1
        self._issues_case_id    = None
        self._issues_current    = ()
        self._workspace_command = self._accion_crear_expediente
        self._workspace_action_label = "Crear expediente"

        self._crear_ui_guiada()
        try:
            self._verificar_estructura()
        except Exception as error:
            messagebox.showerror("No se pudo proteger la base", str(error), parent=self)
            self.destroy()
            raise
        self._cargar_comunidades()
        self.log("", "bienvenida")
        self.log("  👋 Bienvenido. Selecciona comunidad y periodo,", "bienvenida")
        self.log("     deja los PDFs/Excels en entrada/ y pulsa PROCESAR TODO.", "bienvenida")
        self.log("  ✦ Flujo guiado activo · selección de fuentes y conceptos disponible", "bienvenida")
        UIM.aparecer(self)
        self._actualizar_titulo_despacho()
        # Una instalación nueva (o anterior a los datos del despacho) pide su
        # identidad una sola vez; «Más tarde» deja usar la aplicación.
        self.after(600, self._comprobar_datos_despacho)

    # -----------------------------------------------------------------------
    # CONSTRUCCIÓN DE LA UI
    # -----------------------------------------------------------------------
    def _crear_ui_guiada(self):
        """Shell compacto: contexto arriba, cuatro pasos y una tarea principal."""
        header = ctk.CTkFrame(self, fg_color=C["panel"], corner_radius=0, height=76)
        header.pack(fill="x")
        header.pack_propagate(False)
        ctk.CTkLabel(header, text="Regularizaciones", font=UIM.fuente(22, "bold"), text_color=C["texto"]).pack(side="left", padx=(28, 8), pady=18)
        self.lbl_despacho = ctk.CTkLabel(header, text="Flujo guiado de facturas, lecturas y cartas", font=UIM.fuente(11), text_color=C["texto_sec"])
        self.lbl_despacho.pack(side="left", pady=18)
        self.interruptor = UIM.InterruptorTema(header)
        self.interruptor.pack(side="right", padx=(8, 22))
        for text, command in (
            ("Ajustes", self._configurar_rutas),
            ("Nueva comunidad", self._nueva_comunidad),
            ("Bandeja global", self._accion_bandeja_global),
            ("Todas las comunidades", self._accion_panel_comunidades),
        ):
            ctk.CTkButton(header, text=text, command=command, height=31, corner_radius=8, font=UIM.fuente(11), **UIM.secondary_button_kwargs()).pack(side="right", padx=(0, 8))
        self.linea = UIM.LineaGradiente(self, altura=2)
        self.linea.pack(fill="x")
        self.interruptor.al_cambiar(self.linea.refrescar)

        context = ctk.CTkFrame(self, fg_color="transparent")
        context.pack(fill="x", padx=28, pady=(16, 8))
        context.grid_columnconfigure(0, weight=3)
        context.grid_columnconfigure(1, weight=3)
        # Las fechas pertenecen al expediente: su período contable se crea y
        # enlaza automáticamente, así que no hay un selector de período aparte.
        for column, label in enumerate(("Comunidad", "Expediente (intervalo que se regulariza)")):
            ctk.CTkLabel(context, text=label, font=UIM.fuente(10, "bold"), text_color=C["texto_sec"]).grid(row=0, column=column, sticky="w", padx=(0 if column == 0 else 10, 0), pady=(0, 4))
        self.cb_comunidad = UIM.ComboModerno(context, variable=self.comunidad_actual, values=[], width=380, height=38, command=lambda _v: self._on_comunidad_seleccionada())
        self.cb_comunidad.grid(row=1, column=0, sticky="ew", padx=(0, 10))
        self.cb_expediente = UIM.ComboModerno(context, variable=self.expediente_seleccionado, values=[], width=340, height=38, command=lambda _v: self._on_expediente_seleccionado())
        self.cb_expediente.grid(row=1, column=1, sticky="ew", padx=10)
        actions = ctk.CTkFrame(context, fg_color="transparent")
        actions.grid(row=1, column=2, padx=(4, 0), sticky="e")
        ctk.CTkButton(actions, text="Cambiar fechas", command=self._accion_cambiar_fechas, height=38, width=118, corner_radius=9, font=UIM.fuente(11), **UIM.secondary_button_kwargs()).pack(side="left", padx=(0, 7))
        ctk.CTkButton(actions, text="Nuevo expediente", command=self._accion_crear_expediente, height=38, corner_radius=9, font=UIM.fuente(11, "bold"), fg_color=C["primario"], hover_color=C["primario_hover"]).pack(side="left")
        self.banner_periodo = ctk.CTkFrame(self, fg_color=C["acento_suave"], corner_radius=10)
        self.banner_periodo.pack(fill="x", padx=28, pady=(0, 12))
        self.lbl_banner = ctk.CTkLabel(self.banner_periodo, text="Selecciona una comunidad y un período para empezar.", font=UIM.fuente(11), text_color=C["primario"], anchor="w")
        self.lbl_banner.pack(fill="x", padx=14, pady=8)

        body = ctk.CTkFrame(self, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=28, pady=(0, 12))
        body.grid_columnconfigure(1, weight=1)
        body.grid_rowconfigure(0, weight=1)
        nav = ctk.CTkFrame(body, width=190, fg_color=C["panel"], corner_radius=14, border_width=1, border_color=C["borde"])
        nav.grid(row=0, column=0, sticky="ns", padx=(0, 12))
        nav.grid_propagate(False)
        ctk.CTkLabel(nav, text="Expediente", font=UIM.fuente(11, "bold"), text_color=C["texto_sec"]).pack(anchor="w", padx=16, pady=(16, 8))
        self.workflow_step_rows, self.etapas, self.botones = {}, {}, {}
        rows = (("fuentes", "1", "Fuentes", self._accion_anadir_fuentes), ("validar", "2", "Validar", self._accion_resolver_incidencias), ("reparto", "3", "Reparto", self._accion_siguiente_paso), ("cartas", "4", "Cartas", self._accion_generar_cartas_expediente))
        for key, number, label, command in rows:
            row = UIM.WorkflowStepRow(nav, number=number, label=label, status="pending", command=command)
            row.pack(fill="x", padx=8, pady=3)
            self.workflow_step_rows[key] = row
            self.etapas[key] = row._marker
        # Herramientas que no son un paso del flujo. «Confirmar fuentes» y
        # «Resolver copias» se alcanzan desde «Validar» y sus incidencias, y las
        # fechas se cambian junto al selector de expediente.
        ctk.CTkLabel(nav, text="Herramientas", font=UIM.fuente(11, "bold"), text_color=C["texto_sec"]).pack(anchor="w", padx=16, pady=(18, 6))
        # Lo cobrado a los vecinos no está en ninguna factura ni en ningún
        # contador: sin este paso el análisis no tiene con qué comparar el coste.
        for text, command in (
            ("Cuotas cobradas", self._accion_cuotas_cobradas),
            ("Gastos fijos", self._accion_gastos_fijos),
            ("Coherencia", self._accion_coherencia),
            ("Reevaluar fuentes", self._accion_reanalizar_fuentes),
            ("Historial", self._accion_ver_historial_periodo),
            ("Abrir salidas", self._abrir_salidas),
        ):
            self.botones[text] = ctk.CTkButton(
                nav, text=text, command=command, height=31, corner_radius=8,
                font=UIM.fuente(10), **UIM.secondary_button_kwargs(),
            )
            self.botones[text].pack(fill="x", padx=12, pady=(0, 5))

        work = ctk.CTkFrame(body, fg_color=C["panel"], corner_radius=14, border_width=1, border_color=C["borde"])
        work.grid(row=0, column=1, sticky="nsew")
        self.workspace_headline = ctk.CTkLabel(work, text="Crea un expediente", font=UIM.fuente(23, "bold"), text_color=C["texto"])
        self.workspace_headline.pack(anchor="w", padx=24, pady=(24, 4))
        self.workspace_detail = ctk.CTkLabel(work, text="Elige el intervalo de trabajo para comenzar.", font=UIM.fuente(12), text_color=C["texto_sec"], anchor="w")
        self.workspace_detail.pack(fill="x", padx=24)
        self.workspace_primary = ctk.CTkButton(work, text="Crear expediente", command=self._accion_crear_expediente, height=44, corner_radius=10, font=UIM.fuente(13, "bold"), fg_color=C["primario"], hover_color=C["primario_hover"])
        self.workspace_primary.pack(anchor="w", padx=24, pady=(18, 16))
        self.botones["Crear expediente · principal"] = self.workspace_primary
        # Se rellena en _actualizar_workspace con las etapas repetibles del estado.
        self.workspace_repeat_actions = ctk.CTkFrame(work, fg_color="transparent")
        panel = ctk.CTkFrame(work, fg_color=C["panel_2"], corner_radius=11)
        panel.pack(fill="both", expand=True, padx=24, pady=(0, 20))
        self.lbl_incidencias_bandeja = ctk.CTkLabel(panel, text="Incidencias pendientes", font=UIM.fuente(10, "bold"), text_color=C["texto_sec"])
        self.lbl_incidencias_bandeja.pack(anchor="w", padx=14, pady=(12, 3))
        self.bandeja_incidencias = ctk.CTkScrollableFrame(panel, fg_color="transparent")
        self.bandeja_incidencias.pack(fill="both", expand=True, padx=7, pady=(0, 7))
        self.expediente_metricas = {}
        self.carril_estado_expediente = ctk.CTkFrame(work, width=1, height=1, fg_color=C["primario"])
        self.lbl_estado_resumen = self.workspace_detail
        self.log_area = tk.Text(work, height=1, state="disabled")
        self.barra_estado = ctk.CTkFrame(self, fg_color=C["panel"], corner_radius=0, height=34)
        self.barra_estado.pack(fill="x", side="bottom")
        self.punto_estado = UIM.PuntoEstado(self.barra_estado); self.punto_estado.pack(side="left", padx=(20, 6), pady=7)
        self.lbl_estado = ctk.CTkLabel(self.barra_estado, text="Listo", font=UIM.fuente(10), text_color=C["texto"]); self.lbl_estado.pack(side="left")
        self.lbl_bd = ctk.CTkLabel(self.barra_estado, text="Datos locales", font=UIM.fuente(10), text_color=C["texto_sec"]); self.lbl_bd.pack(side="right", padx=20)
        self.progreso = ctk.CTkProgressBar(self.barra_estado, mode="indeterminate", width=130, height=5, progress_color=C["primario"])



    # -----------------------------------------------------------------------
    # LOG
    # -----------------------------------------------------------------------
    def log(self, mensaje: str, tipo: str = "neutro"):
        """Añade una línea al log en el hilo principal."""
        def _escribir():
            self.log_area.configure(state="normal")
            hora = datetime.now().strftime("%H:%M")
            if tipo == "titulo":
                # Cabecera de sección: limpia los adornos ━━━ y da aire
                limpio = mensaje.replace("━", "").strip()
                self.log_area.insert("end", f"\n  ◆  {limpio}\n", "titulo")
                self.log_area.insert("end", "  " + "─" * 54 + "\n", "sep")
            elif tipo == "bienvenida":
                self.log_area.insert("end", mensaje + "\n", "bienvenida")
            else:
                self.log_area.insert("end", f" {hora}  ", "hora")
                self.log_area.insert("end", mensaje + "\n", tipo)
            self.log_area.see("end")
            self.log_area.configure(state="disabled")
        self.after(0, _escribir)


    def _estado(self, texto: str, procesando: bool = False):
        def _act():
            self.lbl_estado.configure(text=texto)
            if hasattr(self, "lbl_estado_resumen"):
                self.lbl_estado_resumen.configure(text=texto if procesando else "Listo para el siguiente paso.")
            self.punto_estado.procesando(procesando)
            if procesando:
                self.progreso.pack(side="left", padx=8, pady=13)
                self.progreso.start()
            else:
                self.progreso.stop()
                self.progreso.pack_forget()
        self.after(0, _act)

    def _resumen_ejercicio(self, texto: str):
        if hasattr(self, "lbl_estado_resumen"):
            self.after(0, lambda: self.lbl_estado_resumen.configure(text=texto))

    # -----------------------------------------------------------------------
    # CARGA DE DATOS
    # -----------------------------------------------------------------------
    def _verificar_estructura(self):
        """Crea las carpetas necesarias si no existen."""
        for carpeta in [RUTA_CARTAS,
                        self.ruta_archivo_expedientes,
                        BASE_DIR / "data",
                        BASE_DIR / "config"]:
            carpeta.mkdir(parents=True, exist_ok=True)

        if not MOD.get("gestor_bd"):
            self.log("⚠️  Módulo gestor_bd.py no encontrado en core/", "error")
            return

        if not RUTA_BD.exists():
            try:
                MOD["gestor_bd"].crear_bd(str(RUTA_BD))
                self.log("✅ Base de datos creada en data/gestion.db", "ok")
            except Exception as e:
                self.log(f"❌ Error creando BD: {e}", "error")
                raise
        else:
            result = MOD["database_backup"].backup_database(RUTA_BD, reason="startup")
            self.log(f"Copia de seguridad verificada: {result.backup_path.name}", "ok")
            for warning in result.cleanup_warnings:
                self.log(warning, "aviso")

    def _cargar_comunidades(self):
        """Rellena el desplegable de comunidades desde la BD."""
        if not RUTA_BD.exists() or not MOD.get("gestor_bd"):
            return
        try:
            con = MOD["gestor_bd"].conectar(str(RUTA_BD))
            rows = con.execute(
                "SELECT id_comunidad, codigo, nombre FROM comunidades WHERE activa=1 ORDER BY codigo"
            ).fetchall()
            con.close()

            opciones = [f"{r['codigo']} — {r['nombre']}" for r in rows]
            self.cb_comunidad.configure(values=opciones)
            self._ids_comunidad = {f"{r['codigo']} — {r['nombre']}": r["id_comunidad"]
                                    for r in rows}

            if opciones:
                self.cb_comunidad.set(opciones[0])
                self._on_comunidad_seleccionada()
        except Exception:
            pass

    def _on_comunidad_seleccionada(self, event=None):
        self._limpiar_contexto_expediente()
        self.id_comunidad = self._ids_comunidad.get(self.comunidad_actual.get())
        if not self.id_comunidad:
            return
        self._refrescar_lista_expedientes()
        if self.id_expediente:
            self._refrescar_expediente()

    def _refrescar_lista_expedientes(self, select_case_id=None):
        service = MOD.get("expedient_service")
        database = MOD.get("gestor_bd")
        if not self.id_comunidad or not service or not database:
            self._limpiar_contexto_expediente()
            return

        connection = database.conectar(str(self.ruta_bd_expedientes))
        try:
            cases = service.list_cases(connection, self.id_comunidad)
        finally:
            connection.close()

        options = []
        ids_by_option = {}
        options_by_id = {}
        names_by_id = {}
        for case in cases:
            option = (
                f"{case.name} · {case.start_date.strftime('%d/%m/%Y')} – "
                f"{case.end_date.strftime('%d/%m/%Y')}"
            )
            options.append(option)
            ids_by_option[option] = case.id_case
            options_by_id[case.id_case] = option
            names_by_id[case.id_case] = case.name

        self._ids_expediente = ids_by_option
        self.cb_expediente.configure(values=options)
        requested_id = select_case_id or self.id_expediente
        if requested_id not in options_by_id:
            requested_id = cases[0].id_case if cases else None
        if requested_id is None:
            self._limpiar_contexto_expediente()
            return

        self.id_expediente = requested_id
        self.expediente_seleccionado.set(options_by_id[requested_id])
        self.expediente_actual.set(names_by_id[requested_id])

    def _on_expediente_seleccionado(self, event=None):
        selected_id = getattr(self, "_ids_expediente", {}).get(
            self.expediente_seleccionado.get()
        )
        if selected_id is None:
            self._limpiar_contexto_expediente()
            return

        ingestion = MOD.get("case_ingestion")
        database = MOD.get("gestor_bd")
        if not ingestion or not database:
            return
        connection = database.conectar(str(self.ruta_bd_expedientes))
        try:
            ingestion.assert_case_belongs_to_community(
                connection, selected_id, self.id_comunidad
            )
        except LookupError as exc:
            self._limpiar_contexto_expediente()
            self._refrescar_lista_expedientes()
            self.log(str(exc), "aviso")
            return
        finally:
            connection.close()

        self.id_expediente = selected_id
        self._refrescar_expediente()

    def _actualizar_banner(self, case=None, period_label: str | None = None):
        """Resume el intervalo del expediente activo; el período se deriva de él."""
        if not hasattr(self, "lbl_banner"):
            return
        if case is None:
            self.lbl_banner.configure(
                text="Selecciona una comunidad y un expediente para empezar.",
                text_color=C["primario"])
            self.banner_periodo.configure(fg_color=C["banner_abierto"])
            return
        days = (case.end_date - case.start_date).days + 1
        closed = case.status == "closed"
        text = (
            f"{'🔒' if closed else '📅'}  {case.name}  ·  "
            f"{case.start_date.strftime('%d/%m/%Y')} → {case.end_date.strftime('%d/%m/%Y')}"
            f"  ({days} días)"
        )
        if period_label:
            text += f"  ·  {period_label}"
        self.lbl_banner.configure(
            text=text, text_color=C["texto_sec"] if closed else C["primario"])
        self.banner_periodo.configure(
            fg_color=C["banner_cerrado"] if closed else C["banner_abierto"])

    # -----------------------------------------------------------------------
    def _en_hilo(self, fn):
        """Ejecuta fn en un hilo secundario para no bloquear la UI."""
        if self._procesando:
            messagebox.showwarning("Espera", "Ya hay una operación en curso.")
            return
        self._procesando = True
        self._deshabilitar_botones()
        threading.Thread(target=self._ejecutar_hilo, args=(fn,), daemon=True).start()

    def _ejecutar_hilo(self, fn):
        try:
            fn()
        except PermissionError as e:
            archivo = Path(getattr(e, "filename", "") or "").name or "un archivo"
            self.log(f"📛 «{archivo}» está abierto en otro programa (normalmente Excel).", "error")
            self.log("   👉 Ciérralo y pulsa de nuevo el botón. No se ha perdido nada.", "aviso")
        except sqlite3.OperationalError as e:
            if "locked" in str(e).lower():
                self.log("📛 La base de datos está siendo usada por otro proceso.", "error")
                self.log("   👉 Cierra otras ventanas del programa y reinténtalo.", "aviso")
            else:
                self.log(f"📛 Error en la base de datos: {e}", "error")
                self.log("   👉 Prueba con «Sincronizar comunidades» o avisa al soporte.", "aviso")
        except FileNotFoundError as e:
            archivo = Path(getattr(e, "filename", "") or str(e)).name
            self.log(f"📛 No se encuentra el archivo o carpeta: {archivo}", "error")
            self.log("   👉 Revisa «Configurar rutas» para ver dónde busca el programa.", "aviso")
        except Exception as e:
            self.log(f"📛 Error inesperado: {type(e).__name__} — {e}", "error")
            self.log("   👉 Si se repite, copia el detalle técnico de abajo y pide ayuda.", "aviso")
            self.log(traceback.format_exc().strip(), "detalle")
            detail = f"{e}\n\nNo se ha modificado ningún archivo oficial."
            self.after(
                0,
                lambda detail=detail: messagebox.showerror(
                    "No se pudo completar", detail, parent=self,
                ),
            )
        finally:
            self._procesando = False
            self.after(0, self._habilitar_botones)
            self._estado("Listo")


    def _accion_cambiar_fechas(self):
        if self._procesando or not self._validar_expediente_activo():
            return
        MOD["expedient_ui"].open_case_period_dialog(self, self.id_expediente)

    def _accion_enlazar_periodo(self):
        """El período se deriva del expediente; si no puede enlazarse, se corrigen las fechas."""
        if self._procesando or not self._validar_expediente_activo():
            return
        database, service = MOD.get("gestor_bd"), MOD.get("expedient_service")
        connection = database.conectar(str(self.ruta_bd_expedientes))
        try:
            service.link_case_to_period(connection, self.id_expediente)
        except ValueError as error:
            self.log(f"No se pudo enlazar el período: {error}", "aviso")
            self.after(0, self._accion_cambiar_fechas)
            return
        finally:
            connection.close()
        self.log("Período del expediente enlazado.", "ok")
        self._refrescar_expediente()

    def _accion_crear_expediente(self):
        expedient_ui = MOD.get("expedient_ui")
        if not expedient_ui:
            self.log("No está disponible la interfaz de expedientes.", "error")
            return
        expedient_ui.open_create_case_dialog(self)

    def _validar_expediente_activo(self) -> bool:
        """Evita que los botones de resultado operen sin contexto auditable."""
        if not self.id_comunidad:
            self.log("Selecciona una comunidad antes de continuar.", "aviso")
            return False
        if not self.id_expediente:
            self.log("Crea o selecciona un expediente antes de continuar.", "aviso")
            return False
        if not MOD.get("case_workflow_actions"):
            self.log("No está disponible el flujo por expediente.", "error")
            return False
        return True

    def _accion_importar_modelo_inicial(self):
        """Pide sólo las fuentes de arranque; el trabajo ordinario será PDF."""
        if not self._validar_expediente_activo():
            return
        master = filedialog.askopenfilename(
            title="Selecciona el Excel maestro inicial",
            filetypes=[("Excel", "*.xlsx *.xls"), ("Todos", "*.*")],
        )
        if not master:
            return
        owners = readings = None
        include_companions = messagebox.askyesno(
            "Fuentes complementarias",
            "¿Quieres añadir ahora el listado de propietarios y las lecturas?\n\n"
            "Son opcionales. Si una comunidad no las usa, el sistema mostrará "
            "las incidencias necesarias sin inventar datos.",
            parent=self,
        )
        if include_companions:
            owners = filedialog.askopenfilename(
                title="Selecciona el listado de propietarios (opcional en otras comunidades)",
                filetypes=[("CSV", "*.csv"), ("Todos", "*.*")],
            )
            if not owners:
                return
            readings = filedialog.askopenfilename(
                title="Selecciona las lecturas complementarias",
                filetypes=[("Excel", "*.xlsx *.xls"), ("Todos", "*.*")],
            )
            if not readings:
                return
        self._en_hilo(
            lambda: self._importar_modelo_inicial_impl(master, owners, readings)
        )

    def _accion_generar_excel_expediente(self):
        if self._validar_expediente_activo():
            self._en_hilo(self._generar_excel_expediente_impl)

    def _accion_siguiente_paso(self):
        """Ejecuta la acción segura mostrada en el panel principal.

        El acceso lateral de reparto no debe saltarse la confirmación de
        fuentes ni intentar crear un Excel cuando el expediente aún no está
        preparado. La ruta se actualiza al refrescar el expediente.
        """
        command = getattr(self, "_workspace_command", None)
        if command is None:
            self.log("Selecciona un expediente para continuar con el reparto.", "aviso")
            return
        command()

    def _accion_calcular_reparto_expediente(self):
        if self._validar_expediente_activo():
            self._en_hilo(self._calcular_reparto_expediente_impl)

    def _accion_generar_cartas_expediente(self):
        if not self._validar_expediente_activo():
            return
        workflow = MOD["case_workflow_actions"]
        try:
            concepts = workflow.available_case_letter_concepts(
                self.ruta_bd_expedientes,
                id_case=self.id_expediente,
                active_community_id=self.id_comunidad,
                project_root=BASE_DIR,
            )
        except Exception as exc:
            self.log(f"No se pueden preparar las cartas: {exc}", "aviso")
            return
        if not concepts:
            self.log(
                "Aún no hay conceptos calculados y activos para las cartas. "
                "Genera el Excel oficial y calcula el reparto final.",
                "aviso",
            )
            return
        selected = self._dialogo_conceptos_cartas(concepts=concepts)
        if selected is not None:
            self._en_hilo(lambda: self._generar_cartas_expediente_impl(selected))

    def _accion_preparar_correo(self):
        if not self._validar_expediente_activo():
            return
        from mail_service import prepare_mail_run, generate_eml_drafts

        with closing(sqlite3.connect(str(self.ruta_bd_expedientes))) as connection:
            try:
                result = prepare_mail_run(connection, self.id_expediente, BASE_DIR / "salidas" / "correo")
                result = generate_eml_drafts(connection, result.id_mail_run)
            except (ValueError, sqlite3.Error, OSError) as error:
                messagebox.showerror("Correo de cartas", str(error), parent=self)
                return
        self.log(f"Borradores EML: {result.draft_count}; sin correo válido: "
                 f"{result.skipped_count}; duplicados: {result.duplicate_count}.", "info")
        self._ofrecer_abrir_carpeta(str(result.output_path))

    def _accion_enviar_correo(self):
        if not self._validar_expediente_activo():
            return
        from mail_service import exclude_mail_delivery, send_mail_run
        from mail_transport import SmtpSettings, SmtpTransport

        with closing(sqlite3.connect(str(self.ruta_bd_expedientes))) as connection:
            row = connection.execute(
                "SELECT id_mail_run,output_path FROM mail_runs WHERE id_case=? ORDER BY id_mail_run DESC LIMIT 1",
                (self.id_expediente,),
            ).fetchone()
            if row is None:
                messagebox.showwarning("Correo de cartas", "Prepara primero los borradores EML.", parent=self)
                return
            mail_id, folder = row
            candidates = connection.execute(
                "SELECT id_propietario,recipient FROM mail_deliveries WHERE id_mail_run=? AND status='draft' ORDER BY id_propietario",
                (mail_id,),
            ).fetchall()
        if not candidates:
            messagebox.showinfo("Correo de cartas", "No quedan borradores pendientes de envío.", parent=self)
            return
        choices = ", ".join(f"{owner}: {email}" for owner, email in candidates)
        excluded = simpledialog.askstring(
            "Revisar destinatarios",
            f"Destinatarios pendientes:\n{choices}\n\n"
            "IDs que quieres excluir, separados por comas (vacío para incluir todos):",
            parent=self,
        )
        if excluded is None:
            return
        if excluded.strip():
            try:
                selected = {int(item.strip()) for item in excluded.split(",")}
                if not selected.issubset({owner for owner, _email in candidates}):
                    raise ValueError("Algún propietario no pertenece a los borradores pendientes")
                reason = simpledialog.askstring("Motivo", "Motivo de la exclusión:", parent=self)
                if reason is None:
                    return
                with closing(sqlite3.connect(str(self.ruta_bd_expedientes))) as connection:
                    for owner in selected:
                        exclude_mail_delivery(connection, mail_id, owner, reason=reason)
            except ValueError as error:
                messagebox.showerror("Correo de cartas", str(error), parent=self)
                return
        with closing(sqlite3.connect(str(self.ruta_bd_expedientes))) as connection:
            counts = connection.execute(
                "SELECT status,COUNT(*) FROM mail_deliveries WHERE id_mail_run=? GROUP BY status", (mail_id,)
            ).fetchall()
        summary = ", ".join(f"{status}: {count}" for status, count in counts)
        if not messagebox.askyesno(
            "Confirmar envío SMTP",
            f"Revisa los borradores en {folder}.\n\n{summary}\n\n¿Confirmas el envío de este lote?",
            parent=self,
        ):
            return
        host = simpledialog.askstring("Servidor SMTP", "Servidor SMTP:", parent=self)
        port_text = simpledialog.askstring("Puerto SMTP", "Puerto con STARTTLS:", initialvalue="587", parent=self)
        username = simpledialog.askstring("Usuario SMTP", "Usuario:", parent=self)
        password = simpledialog.askstring("Contraseña SMTP", "Contraseña:", show="*", parent=self)
        if not all((host, port_text, username, password)):
            return
        try:
            port = int(port_text)
            settings = SmtpSettings(host, port, username, username)
            try:
                import keyring
                keyring.set_password("Regularizaciones SMTP", username, password)
                transport = SmtpTransport(settings)
            except Exception:
                transport = SmtpTransport(settings, session_password=password)
            with closing(sqlite3.connect(str(self.ruta_bd_expedientes))) as connection:
                sent, failed = send_mail_run(connection, mail_id, transport,
                                             confirmed_by=username)
            self.log(f"Correo: {sent} enviado(s), {failed} fallo(s).", "ok" if not failed else "aviso")
        except (ValueError, sqlite3.Error, OSError, ImportError) as error:
            messagebox.showerror("Correo de cartas", str(error), parent=self)

    def _accion_anadir_fuentes(self):
        expedient_ui = MOD.get("expedient_ui")
        ingestion = MOD.get("case_ingestion")
        database = MOD.get("gestor_bd")
        if not expedient_ui or not ingestion or not database:
            self.log("No está disponible la interfaz para añadir fuentes.", "error")
            return
        if not self.id_expediente:
            # No se reutiliza una comunidad abierta como destino implícito:
            # una carpeta de correo puede reunir distintos ejercicios y
            # comunidades. La bandeja clasifica antes de pedir un expediente.
            self._accion_bandeja_global()
            return
        if self.id_expediente:
            connection = database.conectar(str(self.ruta_bd_expedientes))
            try:
                ingestion.assert_case_belongs_to_community(
                    connection, self.id_expediente, self.id_comunidad
                )
            except LookupError as exc:
                self._limpiar_contexto_expediente()
                self._refrescar_expediente()
                self.log(str(exc), "aviso")
                return
            finally:
                connection.close()
        expedient_ui.open_add_sources_dialog(self, self.id_expediente)

    def _accion_bandeja_global(self):
        """Clasifica una carpeta mixta sin reutilizar el contexto activo."""
        ui = MOD.get("expedient_ui")
        if ui is None:
            self.log("No está disponible la bandeja global de fuentes.", "error")
            return
        ui.open_detect_communities_dialog(self)

    def _accion_panel_comunidades(self):
        """Expedientes de todas las comunidades con su estado y última salida."""
        ui = MOD.get("expedient_ui")
        if ui is None:
            self.log("No está disponible el panel de comunidades.", "error")
            return
        ui.open_cases_overview_dialog(self)

    def abrir_expediente(self, community_id: int, case_id: int) -> bool:
        """Activa exactamente esa comunidad y ese expediente (y su período).

        Seleccionar la comunidad por el desplegable abriría su primer
        expediente; aquí se pide el elegido para no cambiar de período.
        """
        if self._procesando:
            messagebox.showwarning("Espera", "Termina la operación actual antes de cambiar de expediente.")
            return False
        option = next((label for label, identifier in getattr(self, "_ids_comunidad", {}).items()
                       if identifier == community_id), None)
        if option is None:
            self._cargar_comunidades()
            option = next((label for label, identifier in getattr(self, "_ids_comunidad", {}).items()
                           if identifier == community_id), None)
        if option is None:
            self.log("La comunidad de ese expediente no está activa.", "aviso")
            return False
        self._limpiar_contexto_expediente()
        self.comunidad_actual.set(option)
        self.cb_comunidad.set(option)
        self.id_comunidad = community_id
        self._refrescar_lista_expedientes(select_case_id=case_id)
        found = self.id_expediente == case_id
        if not found:
            self.log("No se encontró el expediente elegido en esa comunidad.", "aviso")
        # La pantalla refleja siempre lo que quedó activo, aunque no sea el pedido.
        self._refrescar_expediente()
        return found

    def _accion_confirmar_fuentes(self):
        if self._procesando or not self._validar_expediente_activo():
            return
        ui = MOD.get("expedient_ui")
        if ui:
            ui.open_confirm_sources_dialog(self, self.id_expediente)

    def _accion_revalidar_perfil(self):
        if self._procesando or not self._validar_expediente_activo():
            return
        workflow = MOD.get("case_workflow_actions")
        database = MOD.get("gestor_bd")
        if not workflow or not database:
            self.log("No está disponible la revalidación del perfil Excel.", "error")
            return
        case_id, community_id = self.id_expediente, self.id_comunidad

        def work():
            connection = database.conectar(str(self.ruta_bd_expedientes))
            try:
                profile = workflow.revalidate_case_profile_registration(
                    connection,
                    id_case=case_id,
                    active_community_id=community_id,
                    project_root=BASE_DIR,
                )
            finally:
                connection.close()

            def completed():
                self._refrescar_lista_expedientes(select_case_id=case_id)
                self._refrescar_expediente()
                self.log(
                    f"Perfil {profile.key} revalidado. Genera de nuevo el Excel oficial.",
                    "ok",
                )
                messagebox.showinfo(
                    "Perfil revalidado",
                    "La configuración y la plantilla coinciden. "
                    "El expediente está listo para generar de nuevo el Excel oficial.",
                    parent=self,
                )

            self.after(0, completed)

        self._estado("Revalidando configuración del Excel", procesando=True)
        self._en_hilo(work)

    def _accion_reanalizar_fuentes(self):
        """Vuelve a leer las fuentes; si cambian, el expediente regresa a revisión."""
        if self._procesando or not self._validar_expediente_activo():
            return
        ingestion, database, ui, service = (
            MOD.get(name) for name in ("case_ingestion", "gestor_bd", "expedient_ui", "expedient_service")
        )
        if not all((ingestion, database, ui, service)):
            self.log("No está disponible el análisis de fuentes.", "error")
            return
        case_id, community_id = self.id_expediente, self.id_comunidad
        database_path = str(self.ruta_bd_expedientes)
        connection = database.conectar(database_path)
        try:
            status = service.get_case(connection, case_id).status
        finally:
            connection.close()
        advanced = status in {"ready_for_calculation", "calculated", "reconciled",
                              "deliveries_generated", "closed"}
        if advanced and not messagebox.askyesno(
            "Reevaluar fuentes",
            "Se volverán a leer todas las fuentes del expediente.\n\n"
            "Si cambia algún dato, el expediente vuelve a «En revisión» y tendrás que "
            "confirmar las fuentes y regenerar el Excel, el reparto y las cartas. "
            "Las versiones anteriores se conservan en el historial.\n\n¿Continuar?",
            parent=self,
        ):
            return

        def work():
            connection = database.conectar(database_path)
            try:
                ingestion.assert_case_belongs_to_community(connection, case_id, community_id)
                result = MOD["database_backup"].backup_database(database_path, reason="before_reanalysis")
                self.log(f"Copia de seguridad verificada: {result.backup_path.name}", "ok")
                for warning in result.cleanup_warnings:
                    self.log(warning, "aviso")
                outcome = ingestion.reevaluate_case_sources(connection, case_id)
            finally:
                connection.close()
            summary = ui.source_summary(result.document.document_kind for result in outcome.results)
            if outcome.reopened:
                message = (
                    f"{summary}\n\nHan cambiado datos de las fuentes: el expediente vuelve a "
                    "revisión. Confirma las fuentes y regenera el Excel."
                )
            elif outcome.changed:
                message = f"{summary}\n\nSe han actualizado los datos detectados."
            else:
                message = f"{summary}\n\nNo ha cambiado ningún dato."
            self.log(f"Fuentes reevaluadas: {summary}", "ok")

            def completed():
                self._refrescar_lista_expedientes(select_case_id=case_id)
                self._refrescar_expediente()
                if advanced and not outcome.reopened:
                    if messagebox.askyesno(
                        "Fuentes reevaluadas",
                        f"{message}\n\nEl Excel y el reparto siguen siendo válidos. "
                        "¿Quieres reabrir igualmente la revisión para corregir datos ya confirmados?",
                        parent=self,
                    ):
                        self._reabrir_revision(case_id)
                else:
                    messagebox.showinfo("Fuentes reevaluadas", message, parent=self)

            self.after(0, completed)

        self._estado("Reevaluando fuentes", procesando=True)
        self._en_hilo(work)

    def _reabrir_revision(self, case_id):
        database, service = MOD.get("gestor_bd"), MOD.get("expedient_service")
        connection = database.conectar(str(self.ruta_bd_expedientes))
        try:
            service.set_case_status(connection, case_id, "under_review")
        except ValueError as error:
            self.log(str(error), "aviso")
        finally:
            connection.close()
        self.log("Revisión reabierta: confirma las fuentes para volver a generar el Excel.", "ok")
        self._refrescar_expediente()

    def _accion_cuotas_cobradas(self):
        """Abre las cuotas del período del expediente, enlazándolo si aún no lo está."""
        ui, database, service = (MOD.get(name) for name in ("expedient_ui", "gestor_bd", "expedient_service"))
        if not ui:
            self.log("No está disponible la gestión de cuotas.", "error")
            return
        if self._procesando or not self._validar_expediente_activo():
            return
        connection = database.conectar(str(self.ruta_bd_expedientes))
        try:
            self.id_periodo = service.link_case_to_period(connection, self.id_expediente)
        except ValueError as error:
            self.log(f"No se pudo preparar el período del expediente: {error}", "aviso")
            return
        finally:
            connection.close()
        ui.open_service_fees_dialog(self, "ACS", period_id=self.id_periodo)

    def _periodo_enlazado(self):
        """Período del expediente activo, creándolo si aún no está enlazado."""
        if self._procesando or not self._validar_expediente_activo():
            return None
        database, service = MOD.get("gestor_bd"), MOD.get("expedient_service")
        connection = database.conectar(str(self.ruta_bd_expedientes))
        try:
            self.id_periodo = service.link_case_to_period(connection, self.id_expediente)
        except ValueError as error:
            self.log(f"No se pudo preparar el período del expediente: {error}", "aviso")
            return None
        finally:
            connection.close()
        return self.id_periodo

    def _accion_gastos_fijos(self):
        period_id = self._periodo_enlazado()
        if period_id and MOD.get("expedient_ui"):
            MOD["expedient_ui"].open_fixed_costs_dialog(self, period_id)

    def _accion_coherencia(self):
        period_id = self._periodo_enlazado()
        if period_id and MOD.get('expedient_ui'):
            MOD['expedient_ui'].open_coherence_dialog(self)

    def _accion_resolver_incidencias(self):
        review = MOD.get("document_review")
        ingestion = MOD.get("case_ingestion")
        database = MOD.get("gestor_bd")
        ui = MOD.get("expedient_ui")
        if not review or not ingestion or not database or not ui:
            self.log("No está disponible la revisión de incidencias.", "error")
            return
        if not self.id_expediente:
            self.log("Crea o selecciona un expediente antes de revisar incidencias.", "aviso")
            return
        connection = database.conectar(str(self.ruta_bd_expedientes))
        try:
            ingestion.assert_case_belongs_to_community(
                connection, self.id_expediente, self.id_comunidad
            )
            issues = review.list_open_issues(connection, self.id_expediente)
        except LookupError as exc:
            self._limpiar_contexto_expediente()
            self._refrescar_expediente()
            self.log(str(exc), "aviso")
            return
        finally:
            connection.close()
        if not issues:
            self._accion_confirmar_fuentes()
            return
        if any(issue.code in {"COUNTER_RESET", "READING_ZERO_REVIEW"} for issue in issues):
            ui.open_confirm_sources_dialog(self, self.id_expediente)
            return
        self._refrescar_bandeja_incidencias(issues)
        ui.open_issue_dialog(self, issues[0])

    def _accion_ver_historial_periodo(self):
        if not self.id_expediente:
            self.log("Selecciona un expediente para consultar su historial.", "aviso")
            return
        ui = MOD.get("expedient_ui")
        if ui is None:
            self.log("No está disponible el historial del expediente.", "error")
            return
        ui.open_case_history_dialog(self, self.id_expediente)

    def _refrescar_expediente(self):
        service = MOD.get("expedient_service")
        review = MOD.get("document_review")
        ingestion = MOD.get("case_ingestion")
        database = MOD.get("gestor_bd")
        if not all((service, review, ingestion, database)):
            self.log("No se pudo actualizar el estado del expediente: faltan módulos.", "error")
            return
        if not self.id_expediente:
            self._limpiar_contexto_expediente()
            return

        connection = database.conectar(str(self.ruta_bd_expedientes))
        try:
            case = ingestion.assert_case_belongs_to_community(
                connection, self.id_expediente, self.id_comunidad
            )
            document_count = ingestion.count_case_documents(
                connection, self.id_expediente
            )
            issues = review.list_open_issues(connection, self.id_expediente)
            review_summary = MOD.get("expedient_ui").review_summary(issues)
            open_count = review_summary.actionable_count
            technical_issue_count = review_summary.technical_count
            workflow = MOD.get("case_workflow_actions")
            template_row = connection.execute(
                """SELECT template_relative_path FROM excel_template_profiles
                   WHERE id_comunidad=? AND status='active'
                   ORDER BY id_template_profile DESC LIMIT 1""",
                (self.id_comunidad,),
            ).fetchone()
            has_registered_template = False
            profile_missing = template_row is None
            if template_row is not None:
                try:
                    template_path = (BASE_DIR / template_row["template_relative_path"]).resolve()
                    template_path.relative_to(BASE_DIR.resolve())
                    has_registered_template = template_path.is_file()
                except (TypeError, ValueError):
                    has_registered_template = False
            profile_issue = None
            if workflow:
                try:
                    profile = workflow.resolve_case_profile(
                        connection, id_case=self.id_expediente,
                        active_community_id=self.id_comunidad, project_root=BASE_DIR,
                    )
                    profile_label = profile.key
                except workflow.WorkflowBlockedError as error:
                    if profile_missing:
                        profile_label = "Se preparará automáticamente"
                    else:
                        profile_issue = str(error)
                        profile_label = f"Perfil pendiente: {error}"
            else:
                profile_label = "Perfil no disponible"
            readiness_report = None
            readiness_module = MOD.get("case_readiness")
            if readiness_module is not None:
                readiness_report = readiness_module.evaluate_case_readiness(
                    connection, self.id_expediente, BASE_DIR,
                )
        except LookupError as exc:
            self._limpiar_contexto_expediente()
            self._refrescar_lista_expedientes()
            self.log(str(exc), "aviso")
            return
        finally:
            connection.close()

        self.expediente_actual.set(case.name)
        self.id_periodo = case.period_id
        self._actualizar_banner(
            case,
            "período enlazado" if case.period_id else "el período se enlazará al confirmar las fuentes",
        )
        date_range = (
            f"{case.start_date.strftime('%d/%m/%Y')} – "
            f"{case.end_date.strftime('%d/%m/%Y')}"
        )
        status_labels = {
            "draft": "Borrador",
            "gathering_sources": "Recopilando fuentes",
            "under_review": "En revisión",
            "ready_for_calculation": "Listo para cálculo",
            "calculated": "Calculado",
            "reconciled": "Conciliado",
            "deliveries_generated": "Entregas generadas",
            "closed": "Cerrado",
        }
        status_label = status_labels.get(case.status, case.status)
        summary = (
            f"{case.name}\n{date_range}\n{document_count} fuente(s) · "
            f"Perfil: {profile_label}\n{open_count} decisión(es) pendiente(s)"
            f" · {technical_issue_count} comprobación(es) registrada(s)\n"
            f"Estado: {status_label}"
        )
        self._resumen_ejercicio(summary)
        self._refrescar_bandeja_incidencias(issues)
        workspace = MOD.get("expedient_ui").guided_workspace_state(
            has_case=True, document_count=document_count,
            open_issue_count=open_count, case_status=case.status,
            has_registered_template=has_registered_template,
            profile_issue=profile_issue,
            profile_missing=profile_missing,
            readiness_report=readiness_report,
        )
        self.after(0, lambda: self._actualizar_workspace(workspace))

        def update_metrics():
            metrics = getattr(self, "expediente_metricas", {})
            values = {
                "rango": date_range,
                "perfil": profile_label,
                "fuentes": str(document_count),
                "estado": status_label,
                "incidencias": f"{open_count} decisión(es)",
            }
            for key, value in values.items():
                if key in metrics:
                    metrics[key].configure(text=value)
            rail_color = (
                C["texto_sec"]
                if case.status == "draft"
                else C["primario"]
                if case.status in {"gathering_sources", "under_review"}
                else C["exito"]
            )
            self.carril_estado_expediente.configure(fg_color=rail_color)

        self.after(0, update_metrics)
        self.log(
            f"{case.name} · {date_range} · {document_count} fuente(s) · "
            f"estado {case.status} · {open_count} decisión(es) pendiente(s) · "
            f"{technical_issue_count} comprobación(es) registrada(s)",
            "info",
        )
        if open_count:
            self._actualizar_etapa("validacion")
            self.log(
                f"{open_count} decisión(es) por resolver antes de generar el Excel oficial "
                f"({technical_issue_count} comprobación(es) agrupada(s))",
                "aviso",
            )
        elif case.status == "ready_for_calculation":
            self._actualizar_etapa("calculo")
            self.log("Listo para generar el Excel oficial", "ok")
        elif case.status == "reconciled":
            self._actualizar_etapa("cartas")
            self.log("Reparto conciliado. Elige los conceptos activos y genera las cartas.", "ok")
        elif case.status == "deliveries_generated":
            self._actualizar_etapa("fin")
            self.log("Cartas generadas. Puedes abrir la carpeta de salida o cerrar el expediente.", "ok")
        elif document_count:
            self._actualizar_etapa("fuentes")
            self.log(
                f"Expediente actualizado · {document_count} fuente(s). "
                "Revisa las incidencias o completa sus fuentes.",
                "info",
            )

    def _actualizar_workspace(self, workspace):
        """Renders the safe next action without bypassing service-level gates."""
        action_map = {
            "crear_expediente": ("Crear expediente", self._accion_crear_expediente),
            "anadir_fuentes": ("Añadir fuentes", self._accion_anadir_fuentes),
            "importar_modelo": ("Importar modelo inicial", self._accion_importar_modelo_inicial),
            "resolver_incidencias": ("Resolver incidencias", self._accion_resolver_incidencias),
            "confirmar_fuentes": ("Confirmar fuentes", self._accion_confirmar_fuentes),
            "gestionar_periodos": ("Revisar fechas", self._accion_enlazar_periodo),
            "importar_propietarios": ("Importar propietarios", self._accion_anadir_fuentes),
            "revalidar_perfil": ("Revalidar perfil Excel", self._accion_revalidar_perfil),
            "generar_excel": ("Generar Excel oficial", self._accion_generar_excel_expediente),
            "revisar_gastos_fijos": ("Revisar gastos fijos", self._accion_gastos_fijos),
            "revisar_coherencia": ("Revisar coherencia", self._accion_coherencia),
            "calcular_reparto": ("Calcular reparto", self._accion_calcular_reparto_expediente),
            "generar_cartas": ("Generar cartas", self._accion_generar_cartas_expediente),
            "abrir_salidas": ("Abrir salidas", self._abrir_salidas),
            "preparar_correo": ("Preparar correos", self._accion_preparar_correo),
        }
        label, command = action_map[workspace.next_action]
        self._workspace_command = command
        self._workspace_action_label = label
        self.workspace_headline.configure(text=workspace.headline)
        self.workspace_detail.configure(text=workspace.detail)
        self.workspace_primary.configure(text=label, command=command)
        self._pintar_acciones_repetibles(workspace)
        for step in workspace.steps:
            row = self.workflow_step_rows.get(step.key)
            if row is not None:
                row.set_status(step.status)

    _REPEAT_LABELS = {
        "reevaluar_fuentes": "Reevaluar fuentes",
        "anadir_fuentes": "Añadir fuentes",
        "generar_excel": "Regenerar Excel",
        "calcular_reparto": "Recalcular reparto",
        "generar_cartas": "Repetir cartas",
        "abrir_salidas": "Abrir salidas",
        "enviar_correo": "Enviar correos",
    }

    def _pintar_acciones_repetibles(self, workspace):
        """Muestra sólo las etapas ya superadas que tiene sentido repetir."""
        frame = self.workspace_repeat_actions
        for child in frame.winfo_children():
            child.destroy()
        commands = {
            "reevaluar_fuentes": self._accion_reanalizar_fuentes,
            "anadir_fuentes": self._accion_anadir_fuentes,
            "generar_excel": self._accion_generar_excel_expediente,
            "calcular_reparto": self._accion_calcular_reparto_expediente,
            "generar_cartas": self._accion_generar_cartas_expediente,
            "abrir_salidas": self._abrir_salidas,
            "enviar_correo": self._accion_enviar_correo,
        }
        actions = [
            key for key in getattr(workspace, "repeat_actions", ())
            if key != workspace.next_action and key in commands
        ]
        if not actions:
            frame.pack_forget()
            return
        ctk.CTkLabel(
            frame, text="Repetir una etapa:", font=UIM.fuente(10, "bold"),
            text_color=C["texto_sec"],
        ).pack(side="left", padx=(0, 8))
        for key in actions:
            ctk.CTkButton(
                frame, text=self._REPEAT_LABELS[key], command=commands[key],
                height=32, corner_radius=8, font=UIM.fuente(10),
                **UIM.secondary_button_kwargs(),
            ).pack(side="left", padx=(0, 6))
        if not frame.winfo_manager():
            frame.pack(anchor="w", padx=24, pady=(0, 14), after=self.workspace_primary)

    def _refrescar_bandeja_incidencias(self, issues=()):
        issues = tuple(issues)
        self._issues_current = issues
        if self._issues_case_id != self.id_expediente:
            self._issues_case_id = self.id_expediente
            self._issues_page = 1

        def update_tray():
            tray = getattr(self, "bandeja_incidencias", None)
            if tray is None:
                return
            for child in tray.winfo_children():
                child.destroy()
            expedient_ui = MOD.get("expedient_ui")
            summary = expedient_ui.review_summary(issues)
            groups = summary.groups
            count = summary.actionable_count
            self.lbl_incidencias_bandeja.configure(
                text="Incidencias pendientes" + (f" · {count}" if count else ""))

            if not self.id_expediente:
                ctk.CTkLabel(
                    tray,
                    text="Elige un expediente para revisar sus datos pendientes.",
                    font=UIM.fuente(11),
                    text_color=C["texto_sec"],
                    wraplength=420,
                ).pack(anchor="w", padx=9, pady=12)
                return
            if not issues:
                ctk.CTkLabel(
                    tray,
                    text="Sin incidencias pendientes\nAñade fuentes o continúa con el cálculo.",
                    font=UIM.fuente(11, "bold"),
                    text_color=C["texto_sec"],
                    justify="left",
                    wraplength=420,
                ).pack(anchor="w", padx=9, pady=12)
                return

            visible_groups, current_page, total_pages = expedient_ui.issue_page(
                groups, self._issues_page,
            )
            self._issues_page = current_page
            controls = ctk.CTkFrame(tray, fg_color="transparent")
            controls.pack(fill="x", padx=4, pady=(2, 6))
            ctk.CTkLabel(
                controls,
                text=(
                    f"{summary.actionable_count} decisión(es) · "
                    f"{summary.technical_count} comprobación(es) · "
                    f"Página {current_page}/{total_pages}"
                ),
                font=UIM.fuente(10), text_color=C["texto_sec"],
            ).pack(side="left")
            if total_pages > 1:
                def change_page(target_page):
                    self._issues_page = target_page
                    self._refrescar_bandeja_incidencias(self._issues_current)

                ctk.CTkButton(
                    controls, text="Anterior", width=72, height=26,
                    corner_radius=7, command=lambda: change_page(current_page - 1),
                    **UIM.secondary_button_kwargs(),
                ).pack(side="right", padx=(5, 0))
                ctk.CTkButton(
                    controls, text="Siguiente", width=72, height=26,
                    corner_radius=7, command=lambda: change_page(current_page + 1),
                    **UIM.secondary_button_kwargs(),
                ).pack(side="right")

            for grouped in visible_groups:
                issue = grouped["representative"]
                reset_count = int(grouped["count"])
                is_reset_group = issue.code == "COUNTER_RESET" and reset_count > 1
                # Un único cero necesita el mismo diálogo que un grupo: el
                # genérico no escribe la lectura y parecía no guardar nada.
                is_zero_group = issue.code == "READING_ZERO_REVIEW"
                row = ctk.CTkFrame(
                    tray,
                    fg_color=C["panel"],
                    corner_radius=9,
                    border_width=1,
                    border_color=C["borde"],
                )
                row.pack(fill="x", padx=2, pady=4)
                marker = ctk.CTkFrame(
                    row, width=4, height=1, corner_radius=2, fg_color=C["primario"]
                )
                marker.pack(side="left", fill="y", padx=(7, 9), pady=7)
                marker.pack_propagate(False)
                detail = ctk.CTkFrame(row, fg_color="transparent")
                detail.pack(side="left", fill="both", expand=True, pady=7)
                ctk.CTkLabel(
                    detail,
                    text=(
                        f"{reset_count} reinicios de contador por revisar"
                        if is_reset_group else
                        f"{reset_count} lecturas a 0 por revisar" if is_zero_group else
                        expedient_ui.issue_title(issue)
                    ),
                    font=UIM.fuente(11, "bold"),
                    text_color=C["texto"],
                    anchor="w",
                ).pack(fill="x")
                ctk.CTkLabel(
                    detail,
                    text=(
                        f"{issue.archived_path.name} · Se aplica un único criterio temporal "
                        "a todas las viviendas de este informe; las lecturas originales "
                        "y la decisión se conservan en el histórico."
                        if is_reset_group else
                        (f"{issue.archived_path.name} · Aplica un único criterio: se conserva la lectura "
                         "canónica si existe; si no, el cero queda como lectura inicial auditada.")
                        if is_zero_group else
                        f"{expedient_ui.source_display_name(issue.archived_path)} · "
                        f"{expedient_ui.issue_message(issue)}"
                    ),
                    font=UIM.fuente(10),
                    text_color=C["texto_sec"],
                    anchor="w",
                    justify="left",
                    wraplength=560,
                ).pack(fill="x", pady=(2, 0))
                actions = ctk.CTkFrame(row, fg_color="transparent")
                actions.pack(side="right", padx=8, pady=7)
                if is_reset_group:
                    resolve_command = lambda current=issue, count=reset_count: expedient_ui.open_counter_reset_carry_forward_dialog(
                        self, current.id_case, current.id_document, count,
                        current.archived_path.name,
                    )
                elif is_zero_group:
                    resolve_command = lambda current=issue, count=reset_count: expedient_ui.open_initial_zero_confirmation_dialog(
                        self, current.id_case, current.id_document, count,
                        current.archived_path.name,
                    )
                else:
                    resolve_command = lambda current=issue: expedient_ui.open_issue_dialog(
                        self, current
                    )
                ctk.CTkButton(
                    actions,
                    text="Abrir archivo",
                    width=86,
                    height=28,
                    corner_radius=7,
                    fg_color="transparent",
                    border_width=1,
                    border_color=C["borde"],
                    text_color=C["primario"],
                    hover_color=C["acento_suave"],
                    command=lambda current=issue: expedient_ui.open_archived_file(
                        self, current
                    ),
                ).pack(pady=(0, 4))
                ctk.CTkButton(
                    actions,
                    text="Revisar grupo" if (is_reset_group or is_zero_group) else "Resolver",
                    width=86,
                    height=28,
                    corner_radius=7,
                    fg_color=C["primario"],
                    hover_color=C["primario_hover"],
                    command=resolve_command,
                ).pack()
                # Salida para las fuentes que no deben entrar en el expediente:
                # sin ella, una sola incidencia irresoluble lo bloquea entero.
                ctk.CTkButton(
                    actions,
                    text="Omitir",
                    width=86,
                    height=28,
                    corner_radius=7,
                    fg_color="transparent",
                    border_width=1,
                    border_color=C["borde"],
                    text_color=C["texto_sec"],
                    hover_color=C["acento_suave"],
                    command=lambda current=issue: expedient_ui.open_skip_source_dialog(
                        self, current
                    ),
                ).pack(pady=(4, 0))

        self.after(0, update_tray)

    def _limpiar_contexto_expediente(self):
        self.id_expediente = None
        self._issues_page = 1
        self._issues_case_id = None
        self._issues_current = ()
        self.expediente_actual.set("")
        self.expediente_seleccionado.set("")
        self._ids_expediente = {}
        if hasattr(self, "cb_expediente"):
            self.cb_expediente.configure(values=[])
        self.id_periodo = None
        self._resumen_ejercicio("Elige un expediente para continuar.")
        self._refrescar_bandeja_incidencias()
        self._actualizar_banner()

        def reset_case_state():
            metrics = getattr(self, "expediente_metricas", {})
            for key, value in {
                "rango": "—",
                "perfil": "—",
                "fuentes": "0",
                "estado": "—",
                "incidencias": "0 abiertas",
            }.items():
                if key in metrics:
                    metrics[key].configure(text=value)
            for dot in getattr(self, "etapas", {}).values():
                dot.configure(text="○", text_color=C["texto_sec"])
            if hasattr(self, "carril_estado_expediente"):
                self.carril_estado_expediente.configure(fg_color=C["borde"])

        self.after(0, reset_case_state)



    def _actualizar_etapa(self, activa):
        def _actualizar():
            aliases = {"validacion": "validar", "calculo": "reparto", "fin": "cartas"}
            activa_normalizada = aliases.get(activa, activa)
            colores = {"fuentes": C["primario"], "validar": C["primario"], "reparto": C["primario"], "cartas": C["exito"]}
            orden = ["fuentes", "validar", "reparto", "cartas"]
            for key, dot in getattr(self, "etapas", {}).items():
                if key in orden and activa_normalizada in orden and orden.index(key) <= orden.index(activa_normalizada):
                    dot.configure(text="●", text_color=colores.get(key, C["primario"]))
                else:
                    dot.configure(text="○", text_color=C["texto_sec"])
        self.after(0, _actualizar)







    # -----------------------------------------------------------------------
    # IMPLEMENTACIONES
    # -----------------------------------------------------------------------




    def _progreso_expediente(self, stage, details):
        """Traduce hitos técnicos a actividad que puede seguir el despacho."""
        if stage == 'coherence' or details.get('technical_stage') == 'coherence':
            for note in details.get('notes', ()):
                self.log(note, 'info')
            return
        if stage == "backup_database":
            self.log(f"Copia de seguridad verificada: {Path(details['path']).name}", "ok")
            for warning in details.get("warnings", ()):
                self.log(warning, "aviso")
            return
        messages = {
            "validate_case": ("validacion", "Comprobando el expediente seleccionado…"),
            "importar_modelo": ("fuentes", "Importando el modelo inicial archivado…"),
            "importar_complementarias": ("fuentes", "Incorporando propietarios y lecturas complementarias…"),
            "generar_excel": ("calculo", "Generando y comprobando el Excel oficial…"),
            "calcular_reparto": ("calculo", "Calculando el reparto final al céntimo…"),
            "generar_cartas": ("cartas", "Preparando las cartas por propietario…"),
            "reused": ("cartas", "Las cartas ya estaban generadas y auditadas."),
            "completed": ("fin", "Documentos generados y auditados."),
            "incomplete": ("cartas", "Algunas cartas necesitan revisión individual."),
        }
        key, text = messages.get(stage, ("calculo", "El expediente sigue procesándose…"))
        technical_messages = {
            "prepare_template": "Preparando la plantilla oficial…",
            "recalculate": "Recalculando las fórmulas del Excel…",
            "reconcile": "Comprobando que los importes cuadran…",
            "publish": "Publicando el Excel validado…",
            "validate_export": "Comprobando el Excel validado…",
            "generate_letter": "Generando una carta del lote…",
            "complete": "Reparto calculado y conciliado.",
        }
        technical = str(details.get("technical_stage") or "")
        if technical.startswith("write_"):
            text = "Incorporando datos validados al Excel oficial…"
        elif technical.startswith("calculate_"):
            text = "Distribuyendo un concepto y comprobando los céntimos…"
        elif technical in technical_messages:
            text = technical_messages[technical]
        self._actualizar_etapa(key)
        self._estado(text, procesando=True)
        self._resumen_ejercicio(text)
        self.log(f"  ▸ {text}", "info")

    def _refrescar_despues_de_accion(self, case_id):
        def refresh():
            self._refrescar_lista_expedientes(select_case_id=case_id)
            self._refrescar_expediente()
            self._avisar_cambio_expedientes()
        self.after(0, refresh)

    def _avisar_cambio_expedientes(self):
        """Paneles abiertos (p. ej. «Todas las comunidades») se actualizan solos."""
        for callback in tuple(getattr(self, "_observadores_expedientes", ())):
            try:
                callback()
            except Exception:
                self._observadores_expedientes.discard(callback)

    def _importar_modelo_inicial_impl(self, master, owners=None, readings=None):
        workflow = MOD["case_workflow_actions"]
        self._estado("Importando el modelo inicial…", procesando=True)
        self.log("━━━ IMPORTAR MODELO INICIAL ━━━", "titulo")
        self.log("  ℹ️ Este paso sirve sólo para arrancar desde un Excel histórico. "
                 "Las siguientes regularizaciones se nutrirán de PDF.", "info")
        try:
            result, companions = workflow.run_bootstrap_import(
                self.ruta_bd_expedientes,
                id_case=self.id_expediente,
                active_community_id=self.id_comunidad,
                project_root=BASE_DIR,
                master_path=master,
                owner_list_path=owners,
                readings_path=readings,
                progress=self._progreso_expediente,
            )
        except Exception:
            self._refrescar_despues_de_accion(self.id_expediente)
            raise
        self.log(f"  ✅ Modelo inicial incorporado ({result.imported_invoice_count} factura(s)).", "ok")
        if result.installed_template_path:
            self.log(
                "  ✅ Plantilla instalada y verificada: "
                f"{result.installed_template_path.name}", "ok"
            )
        if companions is not None:
            self.log(
                f"  ✅ Complementarias: {companions.imported_owner_count} propietario(s) y "
                f"{companions.imported_reading_count} lectura(s).", "ok"
            )
        if result.open_issue_count:
            self.log(
                f"  ⚠️ {result.open_issue_count} incidencia(s) pendiente(s). "
                "Ábrelas y confirma el dato solicitado antes de continuar.", "aviso"
            )
        else:
            self.log("  ✅ Fuentes importadas sin incidencias. Genera el Excel oficial.", "ok")
        self._refrescar_despues_de_accion(self.id_expediente)

    def _generar_excel_expediente_impl(self):
        workflow = MOD["case_workflow_actions"]
        self._estado("Generando Excel oficial…", procesando=True)
        self.log("━━━ GENERAR EXCEL OFICIAL ━━━", "titulo")
        try:
            result = workflow.run_generate_excel(
                self.ruta_bd_expedientes,
                id_case=self.id_expediente,
                active_community_id=self.id_comunidad,
                project_root=BASE_DIR,
                output_root=BASE_DIR,
                progress=self._progreso_expediente,
            )
        except Exception:
            # El preflight puede devolver el expediente a revisión y crear
            # incidencias concretas. Refrescar aquí evita que la pantalla siga
            # mostrando «Generar Excel» después de ese cambio de estado.
            self._refrescar_despues_de_accion(self.id_expediente)
            raise
        self.log(f"  ✅ Excel validado: {result.output_path.name}", "ok")
        if result.backup_path:
            self.log(f"  💾 Copia anterior: {result.backup_path.name}", "neutro")
        self._actualizar_etapa("calculo")
        self._refrescar_despues_de_accion(self.id_expediente)
        self.after(0, lambda: self._ofrecer_abrir_archivo(str(result.output_path), "Excel oficial"))

    def _calcular_reparto_expediente_impl(self):
        workflow = MOD["case_workflow_actions"]
        self._estado("Calculando reparto final…", procesando=True)
        self.log("━━━ CALCULAR REPARTO FINAL ━━━", "titulo")
        result = workflow.run_calculate_distribution(
            self.ruta_bd_expedientes,
            id_case=self.id_expediente,
            active_community_id=self.id_comunidad,
            project_root=BASE_DIR,
            progress=self._progreso_expediente,
        )
        self.log(
            f"  ✅ Reparto conciliado: {result.owner_result_count} resultado(s) "
            f"en {len(result.concept_totals_cents)} concepto(s).", "ok"
        )
        self._actualizar_etapa("cartas")
        self._refrescar_despues_de_accion(self.id_expediente)

    def _generar_cartas_expediente_impl(self, selected):
        workflow = MOD["case_workflow_actions"]
        self._estado("Generando cartas…", procesando=True)
        self.log("━━━ GENERAR CARTAS ━━━", "titulo")
        self.log("  ▸ Conceptos incluidos: " + ", ".join(selected), "info")
        result = workflow.run_generate_letters(
            self.ruta_bd_expedientes,
            id_case=self.id_expediente,
            active_community_id=self.id_comunidad,
            project_root=BASE_DIR,
            selected_concepts=tuple(selected),
            progress=self._progreso_expediente,
        )
        if result.failures:
            self.log(
                f"  ⚠️ {result.generated_count} carta(s) generada(s); "
                f"{len(result.failures)} requiere(n) revisión.", "aviso"
            )
            for failure in result.failures[:4]:
                self.log(f"     {failure}", "aviso")
        else:
            self.log(f"  ✅ {result.generated_count} carta(s) generada(s) y auditada(s).", "ok")
            self._actualizar_etapa("fin")
        self._refrescar_despues_de_accion(self.id_expediente)
        self.after(0, lambda: self._ofrecer_abrir_carpeta(str(result.output_path)))


    def _dialogo_conceptos_cartas(self, concepts=None):
        """Pide partidas activas del expediente y recuerda la preferencia comunitaria."""
        settings = MOD.get("letter_settings")
        if not settings or not MOD.get("gestor_bd"):
            return ()
        try:
            con = MOD["gestor_bd"].conectar(str(self.ruta_bd_expedientes))
            stored = settings.available_concepts(con)
            selected = set(settings.load_selected_concepts(con, self.id_comunidad))
            con.close()
        except Exception as exc:
            self.log(f"❌ No se pudo cargar la configuración de conceptos: {exc}", "error")
            return None
        if concepts is None:
            options = tuple((concept.key, concept.label) for concept in stored)
        else:
            options = tuple((str(key), str(label)) for key, label in concepts)
        allowed_keys = {key for key, _label in options}
        selected.intersection_update(allowed_keys)
        if not selected:
            selected = set(allowed_keys)
        if not options:
            self.log("⚠️  No hay conceptos activos disponibles para esta carta", "aviso")
            return None

        dialog = ctk.CTkToplevel(self)
        dialog.title("Conceptos de la carta")
        dialog.geometry("480x520")
        dialog.resizable(False, False)
        dialog.transient(self)
        dialog.grab_set()
        ctk.CTkLabel(dialog, text="¿Qué conceptos quieres incluir?",
                     font=UIM.fuente(17, "bold"), text_color=C["texto"]).pack(
                         anchor="w", padx=24, pady=(22, 4))
        ctk.CTkLabel(dialog, text="Solo se muestran los conceptos activos de este expediente. La selección se recuerda para esta comunidad.",
                     font=UIM.fuente(11), text_color=C["texto_sec"], wraplength=420,
                     justify="left").pack(anchor="w", padx=24, pady=(0, 14))
        variables = {}
        panel = ctk.CTkFrame(dialog, fg_color=C["acento_suave"], corner_radius=12)
        panel.pack(fill="both", expand=True, padx=20, pady=(0, 16))
        for key, label in options:
            variable = tk.BooleanVar(value=key in selected)
            variables[key] = variable
            ctk.CTkCheckBox(panel, text=label, variable=variable,
                            font=UIM.fuente(12), text_color=C["texto"]).pack(
                                anchor="w", padx=18, pady=8)
        result = {"value": None}
        def accept():
            chosen = tuple(key for key, variable in variables.items() if variable.get())
            if not chosen:
                messagebox.showwarning("Selección incompleta", "Selecciona al menos un concepto.", parent=dialog)
                return
            try:
                con = MOD["gestor_bd"].conectar(str(self.ruta_bd_expedientes))
                result["value"] = settings.save_selected_concepts(con, self.id_comunidad, chosen)
                con.close()
            except Exception as exc:
                messagebox.showerror("No se pudo guardar", str(exc), parent=dialog)
                return
            dialog.destroy()
        def cancel():
            dialog.destroy()
        buttons = ctk.CTkFrame(dialog, fg_color="transparent")
        buttons.pack(fill="x", padx=20, pady=(0, 18))
        ctk.CTkButton(buttons, text="Cancelar", command=cancel, width=110,
                      fg_color="transparent", border_width=1, border_color=C["borde"],
                      text_color=C["texto_sec"]).pack(side="right", padx=(8, 0))
        ctk.CTkButton(buttons, text="Continuar", command=accept, width=130,
                      fg_color=C["primario"], hover_color=C["primario_hover"]).pack(side="right")
        dialog.protocol("WM_DELETE_WINDOW", cancel)
        self.wait_window(dialog)
        if result["value"]:
            self.log("  ✓ Conceptos seleccionados: " + ", ".join(result["value"]), "info")
        return result["value"]





    # -----------------------------------------------------------------------
    # GESTIÓN DE EXCELS POR COMUNIDAD
    # -----------------------------------------------------------------------

    # -----------------------------------------------------------------------
    # UTILIDADES
    # -----------------------------------------------------------------------
    def _deshabilitar_botones(self):
        for btn in self.botones.values():
            btn.configure(state="disabled")
        for selector in (getattr(self, "cb_comunidad", None),
                         getattr(self, "cb_expediente", None)):
            if selector is not None:
                selector.configure(state="disabled")

    def _habilitar_botones(self):
        for btn in self.botones.values():
            btn.configure(state="normal")
        for selector in (getattr(self, "cb_comunidad", None),
                         getattr(self, "cb_expediente", None)):
            if selector is not None:
                selector.configure(state="normal")


    def _abrir_salidas(self):
        ruta = RUTA_CARTAS
        if sys.platform == "win32":
            os.startfile(str(ruta))
        else:
            os.system(f"xdg-open '{ruta}'")

    def _ofrecer_abrir_carpeta(self, carpeta: str):
        if messagebox.askyesno("Cartas generadas",
                                f"Cartas guardadas en:\n{carpeta}\n\n¿Abrir la carpeta?"):
            if sys.platform == "win32":
                os.startfile(carpeta)
            else:
                os.system(f"xdg-open '{carpeta}'")

    def _ofrecer_abrir_archivo(self, archivo: str, titulo: str):
        """Sólo ofrece abrir un resultado publicado correctamente."""
        if messagebox.askyesno(
            titulo,
            f"{titulo} guardado en:\n{archivo}\n\n¿Abrir ahora?",
            parent=self,
        ):
            if sys.platform == "win32":
                os.startfile(archivo)
            else:
                os.system(f"xdg-open '{archivo}'")

    # -----------------------------------------------------------------------
    # DIÁLOGOS
    # -----------------------------------------------------------------------
    def _preparar_dialogo(self, titulo: str, ancho: int, alto: int) -> ctk.CTkToplevel:
        """Crea un CTkToplevel centrado sobre la ventana principal, con fundido."""
        dialogo = ctk.CTkToplevel(self, fg_color=C["fondo"])
        dialogo.title(titulo)
        dialogo.resizable(False, False)
        dialogo.transient(self)
        self.update_idletasks()
        x = self.winfo_x() + (self.winfo_width()  - ancho) // 2
        y = self.winfo_y() + (self.winfo_height() - alto) // 2
        dialogo.geometry(f"{ancho}x{alto}+{x}+{y}")
        UIM.aparecer(dialogo)
        # grab_set diferido: CTkToplevel tarda unos ms en ser 'viewable'
        dialogo.after(150, lambda: self._grab_seguro(dialogo))
        return dialogo

    @staticmethod
    def _grab_seguro(dialogo):
        try:
            dialogo.grab_set()
        except tk.TclError:
            pass



    def _nueva_comunidad(self):
        """Ofrece el registro rápido existente o el alta guiada desde fuentes."""
        dialogo = self._preparar_dialogo("Añadir comunidad", 560, 450)
        panel = ctk.CTkFrame(
            dialogo,
            fg_color=C["panel"],
            corner_radius=16,
            border_width=1,
            border_color=C["borde"],
        )
        panel.pack(fill="both", expand=True, padx=18, pady=18)
        ctk.CTkLabel(
            panel,
            text="¿Cómo quieres añadirla?",
            font=UIM.fuente(20, "bold"),
            text_color=C["texto"],
        ).pack(anchor="w", padx=22, pady=(22, 2))
        ctk.CTkLabel(
            panel,
            text="Elige una incorporación básica o deja que las fuentes guíen la configuración.",
            font=UIM.fuente(11),
            text_color=C["texto_sec"],
            wraplength=470,
            justify="left",
        ).pack(anchor="w", padx=22, pady=(0, 16))

        def option(title, description, command, *, primary=False):
            row = ctk.CTkFrame(
                panel,
                fg_color=C["panel_2"],
                corner_radius=11,
                border_width=1,
                border_color=C["borde"],
            )
            row.pack(fill="x", padx=22, pady=5)
            copy = ctk.CTkFrame(row, fg_color="transparent")
            copy.pack(side="left", fill="both", expand=True, padx=14, pady=12)
            ctk.CTkLabel(
                copy, text=title, font=UIM.fuente(12, "bold"), text_color=C["texto"]
            ).pack(anchor="w")
            ctk.CTkLabel(
                copy, text=description, font=UIM.fuente(10), text_color=C["texto_sec"]
            ).pack(anchor="w", pady=(2, 0))

            def choose():
                dialogo.destroy()
                command()

            ctk.CTkButton(
                row,
                text="Elegir",
                command=choose,
                width=86,
                height=34,
                corner_radius=8,
                fg_color=C["primario"] if primary else C["acento_suave"],
                hover_color=C["primario_hover"] if primary else C["acento_suave_hover"],
                text_color=("#FFFFFF", "#FFFFFF") if primary else C["primario"],
            ).pack(side="right", padx=14)

        expedient_ui = MOD.get("expedient_ui")
        option(
            "Detectar desde carpeta",
            "Analiza una carpeta mixta y propone crear automáticamente las comunidades nuevas.",
            lambda: expedient_ui.open_detect_communities_dialog(self),
            primary=True,
        )
        option(
            "Alta masiva por subcarpetas",
            "Una subcarpeta por comunidad: crea las listas con propietarios, lecturas y expediente.",
            lambda: expedient_ui.open_batch_onboarding_dialog(self),
        )
        option(
            "Alta guiada desde fuentes",
            "Analiza propietarios, lecturas y facturas antes de crear el perfil.",
            lambda: expedient_ui.open_community_onboarding_dialog(self),
        )
        option(
            "Registro rápido",
            "Conserva el formulario básico para crear la comunidad directamente.",
            self._registro_rapido_comunidad,
        )

    def _registro_rapido_comunidad(self):
        """Diálogo para registrar una nueva comunidad."""
        dialogo = self._preparar_dialogo("Nueva Comunidad", 470, 320)

        campos = [
            ("Código de comunidad:", "codigo",  "Ej. 101"),
            ("Nombre completo:",     "nombre",  "CDAD. PROP. …"),
            ("CIF de la comunidad:", "cif",     "Ej. H12345674"),
            ("Nº de viviendas:",     "viviendas","Ej. 24"),
        ]
        entradas = {}
        for i, (label, clave, defecto) in enumerate(campos):
            ctk.CTkLabel(dialogo, text=label, font=UIM.fuente(12),
                         text_color=C["texto"]).grid(
                             row=i, column=0, padx=(24, 8), pady=10, sticky="w")
            # Ejemplos como placeholder: antes se insertaban como texto real y
            # podían guardarse por descuido como datos de la comunidad.
            e = ctk.CTkEntry(dialogo, font=UIM.fuente(12), width=230, height=32,
                             border_color=C["borde"], placeholder_text=defecto)
            e.grid(row=i, column=1, padx=(0, 24), pady=10)
            entradas[clave] = e

        def _crear():
            if not MOD.get("gestor_bd"):
                return
            codigo = entradas["codigo"].get().strip()
            nombre = entradas["nombre"].get().strip()
            cif = entradas["cif"].get().strip()
            viviendas = entradas["viviendas"].get().strip()
            problema = None
            if not codigo or not nombre:
                problema = "Indica el código y el nombre de la comunidad."
            elif viviendas and not viviendas.isdigit():
                problema = "El número de viviendas debe ser un número entero."
            elif cif:
                from provider_registry import valid_spanish_tax_id
                if not valid_spanish_tax_id(cif):
                    problema = f"El CIF «{cif}» no es válido."
            if problema:
                messagebox.showwarning("Nueva comunidad", problema, parent=dialogo)
                return
            try:
                con = MOD["gestor_bd"].conectar(str(RUTA_BD))
                id_com = MOD["gestor_bd"].obtener_o_crear_comunidad(
                    con,
                    entradas["codigo"].get().strip(),
                    entradas["nombre"].get().strip(),
                    cif=entradas["cif"].get().strip()
                )
                nviv = int(entradas["viviendas"].get() or 0)
                if nviv:
                    con.execute("UPDATE comunidades SET num_viviendas=? WHERE id_comunidad=?",
                                (nviv, id_com))
                    con.commit()
                con.close()
                cod = entradas["codigo"].get().strip()
                nom = entradas["nombre"].get().strip()
                self.log(f"✅ Comunidad '{cod}' registrada (id={id_com})", "ok")
                self._cargar_comunidades()
                dialogo.destroy()
                messagebox.showinfo(
                    "Comunidad guardada en la base de datos",
                    f"Se ha creado el registro directo:\n\n"
                    f"ID: {id_com}\nCódigo: {cod}\nNombre: {nom}\n"
                    f"CIF: {entradas['cif'].get().strip() or '—'}\n"
                    f"Viviendas: {nviv or '—'}\n\n"
                    "Ya aparece en el selector de comunidades.",
                )
            except Exception as e:
                messagebox.showerror("Error", str(e))

        ctk.CTkButton(dialogo, text="Guardar comunidad",
                      command=_crear,
                      height=38, corner_radius=10,
                      font=UIM.fuente(13, "bold"),
                      fg_color=C["exito"],
                      hover_color=C["exito_hover"],
                      border_width=1, border_color=C["exito"]).grid(
                          row=len(campos), column=0, columnspan=2, pady=18)

    def _accion_nueva_base_segura(self, settings_dialog=None):
        if self._procesando:
            messagebox.showwarning("Espera", "Termina la operación en curso antes de crear una nueva base.", parent=self)
            return
        reset, database = MOD.get("database_reset"), MOD.get("gestor_bd")
        if not reset or not database:
            messagebox.showerror("Nueva base segura", "No está disponible el reinicio seguro.", parent=self)
            return
        database_path = Path(self.ruta_bd_expedientes)
        if not messagebox.askyesno(
            "Nueva base segura",
            "Se guardará una copia de seguridad verificada de la base actual y se creará una base vacía.\n"
            "Las comunidades, expedientes y correcciones actuales quedarán en esa copia. "
            "Los archivos originales y las salidas se conservarán.\n\n"
            f"Base: {database_path}\n\n¿Crear la nueva base?",
            parent=settings_dialog or self, default="no", icon="warning",
        ):
            return

        office = MOD.get("office_settings")
        conserved_office = None
        if office and database_path.is_file():
            with closing(sqlite3.connect(database_path)) as source:
                conserved_office = office.load_office_settings(source)

        def work():
            try:
                def initialise(path):
                    # The replacement must have no live handles before Windows
                    # can publish it. sqlite3's context manager does not close.
                    connection = sqlite3.connect(path)
                    try:
                        with connection:
                            connection.execute("PRAGMA foreign_keys = ON")
                            connection.execute("PRAGMA journal_mode = DELETE")
                            for sql in (*database.TABLAS, *database.INDICES):
                                connection.execute(sql)
                            database.aplicar_migraciones(connection)
                        # Los datos del despacho son configuración, no datos de
                        # comunidades: la base nueva los conserva.
                        if office and conserved_office is not None and conserved_office.configured:
                            office.save_office_settings(connection, conserved_office)
                    finally:
                        connection.close()

                result = reset.reset_database(
                    database_path, backup_root=database_path.parent / "backups",
                    initialise=initialise,
                )
            except Exception as error:
                self.log(f"No se pudo crear la nueva base: {error}", "error")
                self.after(0, lambda detail=str(error): messagebox.showerror("Nueva base segura", detail, parent=self))
                return

            def completed():
                # Only discard live selections after a verified replacement exists.
                self.id_comunidad = self.id_periodo = None
                self.comunidad_actual.set("")
                self._ids_comunidad = {}
                self.cb_comunidad.configure(values=[])
                self._limpiar_contexto_expediente()
                self._cargar_comunidades()
                ui = MOD.get("expedient_ui")
                if ui:
                    self._actualizar_workspace(ui.guided_workspace_state(
                        has_case=False, document_count=0, open_issue_count=0, case_status="",
                    ))
                if settings_dialog is not None:
                    settings_dialog.destroy()
                self.log(f"Nueva base creada. Copia de seguridad: {result.backup_path}", "ok")
                messagebox.showinfo("Nueva base segura", f"La nueva base está lista.\n\nCopia de seguridad:\n{result.backup_path}", parent=self)

            self.after(0, completed)

        self._estado("Creando copia de seguridad y nueva base", procesando=True)
        self._en_hilo(work)

    def _datos_despacho(self):
        database, office = MOD.get("gestor_bd"), MOD.get("office_settings")
        if not database or not office or not RUTA_BD.exists():
            return None
        connection = database.conectar(str(self.ruta_bd_expedientes))
        try:
            return office.load_office_settings(connection)
        finally:
            connection.close()

    def _actualizar_titulo_despacho(self):
        settings = self._datos_despacho()
        name = settings.name if settings is not None and settings.configured else ""
        self.title(f"Regularizaciones · {name}" if name else "Regularizaciones")
        if hasattr(self, "lbl_despacho"):
            self.lbl_despacho.configure(
                text=name or "Flujo guiado de facturas, lecturas y cartas")

    def _comprobar_datos_despacho(self):
        settings = self._datos_despacho()
        if settings is not None and not settings.configured and MOD.get("expedient_ui"):
            MOD["expedient_ui"].open_office_settings_dialog(self, first_run=True)

    def _configurar_rutas(self):
        """Datos del despacho, ubicación de los datos y base nueva."""
        ventana = self._preparar_dialogo("Ajustes", 760, 520)
        ctk.CTkButton(
            ventana, text="Datos del despacho…",
            command=lambda: (ventana.destroy(), MOD["expedient_ui"].open_office_settings_dialog(self)),
            height=36, corner_radius=8, font=UIM.fuente(12, "bold"),
            fg_color=C["primario"], hover_color=C["primario_hover"],
        ).grid(row=0, column=0, columnspan=2, sticky="w", padx=20, pady=(4, 2))
        ctk.CTkLabel(
            ventana, text="Nombre, CIF, ciudad, firma, logo y mes de inicio del ejercicio.",
            font=UIM.fuente(11), text_color=C["texto_sec"],
        ).grid(row=1, column=0, columnspan=2, sticky="w", padx=20, pady=(0, 12))

        rutas = [
            ("Base de datos:",      str(self.ruta_bd_expedientes)),
            ("Fuentes archivadas:", str(self.ruta_archivo_expedientes)),
            ("Plantilla cartas:",   str(RUTA_PLANTILLA)),
            ("Proveedores propios:", str(RUTA_PROVEEDORES_DESPACHO)),
            ("Cartas generadas/:",  str(RUTA_CARTAS)),
        ]
        for i, (label, ruta) in enumerate(rutas):
            ctk.CTkLabel(ventana, text=label, font=UIM.fuente(11),
                         text_color=C["texto"], anchor="w", width=170).grid(
                row=i + 2, column=0, padx=(20, 4), pady=6, sticky="w")
            ctk.CTkLabel(ventana, text=ruta, font=UIM.fuente_mono(11),
                         text_color=C["primario"], anchor="w").grid(
                row=i + 2, column=1, padx=4, pady=6, sticky="w")

        ctk.CTkLabel(ventana,
                     text="Los datos del despacho se guardan en una carpeta escribible separada del programa.",
                     font=UIM.fuente(11),
                     text_color=C["texto_sec"]).grid(
            row=len(rutas) + 2, column=0, columnspan=2, pady=14, padx=20, sticky="w")
        ctk.CTkButton(
            ventana, text="Nueva base segura", command=lambda: self._accion_nueva_base_segura(ventana),
            height=36, corner_radius=8, font=UIM.fuente(11), **UIM.secondary_button_kwargs(),
        ).grid(row=len(rutas) + 3, column=0, columnspan=2, sticky="w", padx=20, pady=(8, 4))
        ctk.CTkLabel(
            ventana, text="Guarda una copia de seguridad verificada y empieza con una base vacía.",
            font=UIM.fuente(11), text_color=C["texto_sec"],
        ).grid(row=len(rutas) + 4, column=0, columnspan=2, sticky="w", padx=20, pady=4)


# ---------------------------------------------------------------------------
# PUNTO DE ENTRADA
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    if "--self-test" in sys.argv:
        APP_PATHS.prepare()
        from gestor_bd import conectar, crear_bd
        if not RUTA_BD.exists():
            crear_bd(str(RUTA_BD))
        with closing(conectar(str(RUTA_BD))) as connection:
            if connection.execute("PRAGMA integrity_check").fetchone()[0] != "ok":
                raise SystemExit("La base no supera la comprobación de integridad")
        print(f"Instalación lista: {RUTA_BD}")
    else:
        UIM.iniciar()
        app = AppGestionFincas()
        app.mainloop()
