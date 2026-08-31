"""
app.py
======
Interfaz gráfica principal del sistema de Gestión de Fincas.
Diseñada para usuarios no técnicos — un botón por acción, mensajes claros.

INTERFAZ: CustomTkinter (moderna, modo claro/oscuro, animaciones).

REQUISITOS:
    pip install customtkinter openpyxl python-docx pdfplumber
    (customtkinter se instala automáticamente si falta)

USO:
    python app.py
    (o doble clic en el .exe si está empaquetado con PyInstaller)

FLUJO DE UN USUARIO TÍPICO:
    1. Selecciona una comunidad del desplegable
    2. Selecciona el periodo (ej: 2024-2025)
    3. Pincha "Procesar Facturas" → lee los PDFs de la carpeta entrada/
    4. Pincha "Actualizar Excel" → vuelca datos en el Excel Maestro
    5. Pincha "Calcular Reparto" → calcula cuotas por vecino
    6. Pincha "Generar Cartas" → crea los .docx en salidas/cartas/
"""

import os
import re
import sys
import json
import shutil
import sqlite3
import threading
import traceback
from datetime import datetime, date
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext

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
# CUSTOMTKINTER (instalación automática si falta)
# ---------------------------------------------------------------------------
try:
    import customtkinter as ctk
except ImportError:
    import subprocess
    print("Instalando customtkinter (primera ejecución)…")
    codigo = subprocess.call(
        [sys.executable, "-m", "pip", "install", "customtkinter"])
    if codigo == 0:
        import customtkinter as ctk
    else:
        raise SystemExit(
            "No se pudo instalar customtkinter automáticamente.\n"
            "Ejecuta en una terminal:  pip install customtkinter"
        )

import ui_moderna as UIM
from ui_moderna import C

# ---------------------------------------------------------------------------
# RUTAS POR DEFECTO
# ---------------------------------------------------------------------------
BASE_DIR = Path(__file__).parent.parent  # sube un nivel desde core/

RUTA_BD         = BASE_DIR / "data" / "gestion.db"
RUTA_ENTRADA    = BASE_DIR / "entrada"
RUTA_PROCESADOS = BASE_DIR / "procesados"
RUTA_EXCELS     = BASE_DIR / "Excels_Maestros"
RUTA_PLANTILLA  = BASE_DIR / "plantillas" / "Plantilla_Cartas.docx"
RUTA_CARTAS     = BASE_DIR / "salidas" / "cartas"
RUTA_PROVEEDORES= BASE_DIR / "config" / "proveedores.json"


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
                   "expedient_ui", "case_workflow_actions"]:
        try:
            modulos[nombre] = __import__(nombre)
        except ImportError:
            modulos[nombre] = None
    return modulos


MOD = _importar_modulos()


# ---------------------------------------------------------------------------
# MOVIMIENTO SEGURO DE ARCHIVOS (evita WinError 32 si el archivo está abierto)
# ---------------------------------------------------------------------------
def mover_seguro(origen, destino, log=None, intentos: int = 3,
                 espera: float = 1.2) -> bool:
    """Mueve un archivo con reintentos; si sigue bloqueado, lo copia y avisa."""
    import time
    origen, destino = str(origen), str(destino)
    nombre = Path(origen).name
    for _ in range(intentos):
        try:
            shutil.move(origen, destino)
            return True
        except (PermissionError, OSError):
            time.sleep(espera)
    try:
        shutil.copy2(origen, destino)
        if log:
            log(f"  ⚠️  «{nombre}» está abierto en otro programa (¿Excel?).", "aviso")
            log("     Se copió a procesados/, pero el original sigue en entrada/.", "aviso")
            log("     👉 Ciérralo en Excel y bórralo de entrada/ (o vuelve a procesar).", "aviso")
    except Exception:
        if log:
            log(f"  ❌ No se pudo mover «{nombre}»: está abierto en otro programa.", "error")
            log("     👉 Cierra el archivo y pulsa de nuevo el botón.", "error")
    return False


# ---------------------------------------------------------------------------
# VENTANA PRINCIPAL
# ---------------------------------------------------------------------------
class AppGestionFincas(ctk.CTk):
    def __init__(self):
        super().__init__(fg_color=C["fondo"])
        self.title("Regularización de facturas")
        self.geometry("1180x760")
        self.minsize(1020, 680)
        self.resizable(True, True)

        # Estado
        self.comunidad_actual  = tk.StringVar(value="")
        self.periodo_actual     = tk.StringVar(value="")
        self.expediente_actual  = tk.StringVar(value="")
        self.expediente_seleccionado = tk.StringVar(value="")
        self.id_comunidad       = None
        self.id_periodo         = None
        self.id_expediente      = None
        self.ruta_bd_expedientes = RUTA_BD
        self.ruta_archivo_expedientes = BASE_DIR / "data" / "expedientes"
        self._procesando        = False

        self._crear_ui_v3()
        self._cargar_comunidades()
        self._verificar_estructura()
        self.log("", "bienvenida")
        self.log("  👋 Bienvenido. Selecciona comunidad y periodo,", "bienvenida")
        self.log("     deja los PDFs/Excels en entrada/ y pulsa PROCESAR TODO.", "bienvenida")
        self.log("  ✦ Flujo guiado activo · selección de fuentes y conceptos disponible", "bienvenida")
        UIM.aparecer(self)
        # El periodo se elige desde la tarjeta de contexto; no interrumpimos
        # el arranque con un diálogo heredado.

    # -----------------------------------------------------------------------
    # CONSTRUCCIÓN DE LA UI
    # -----------------------------------------------------------------------
    def _crear_ui_v3(self):
        """Panel principal del flujo de regularización, diseñado para operar por pasos."""
        header = ctk.CTkFrame(self, fg_color=C["panel"], corner_radius=0, height=78)
        header.pack(fill="x")
        header.pack_propagate(False)
        marca = ctk.CTkFrame(header, width=36, height=36, corner_radius=10, fg_color=C["primario"])
        marca.pack(side="left", padx=(28, 10), pady=20)
        marca.pack_propagate(False)
        ctk.CTkLabel(marca, text="R", font=UIM.fuente(18, "bold"), text_color="#FFFFFF").pack(expand=True)
        title = ctk.CTkFrame(header, fg_color="transparent")
        title.pack(side="left", pady=15)
        ctk.CTkLabel(title, text="Regularización", font=UIM.fuente(22, "bold"), text_color=C["texto"]).pack(anchor="w")
        ctk.CTkLabel(title, text="Facturas, consumos y cartas en un único flujo", font=UIM.fuente(11), text_color=C["texto_sec"]).pack(anchor="w", pady=(1, 0))
        self.interruptor = UIM.InterruptorTema(header)
        self.interruptor.pack(side="right", padx=26)
        self.linea = UIM.LineaGradiente(self, altura=2)
        self.linea.pack(fill="x")
        self.interruptor.al_cambiar(self.linea.refrescar)

        context = ctk.CTkFrame(self, fg_color=C["panel"], corner_radius=18, border_width=1, border_color=C["borde"])
        context.pack(fill="x", padx=24, pady=(18, 12))
        context.grid_columnconfigure(0, weight=3)
        context.grid_columnconfigure(1, weight=2)
        context.grid_columnconfigure(2, weight=0)
        ctk.CTkLabel(context, text="COMUNIDAD", font=UIM.fuente(10, "bold"), text_color=C["texto_sec"]).grid(row=0, column=0, sticky="w", padx=20, pady=(15, 3))
        ctk.CTkLabel(context, text="PERÍODO", font=UIM.fuente(10, "bold"), text_color=C["texto_sec"]).grid(row=0, column=1, sticky="w", padx=12, pady=(15, 3))
        self.cb_comunidad = UIM.ComboModerno(context, variable=self.comunidad_actual, values=[], width=460, height=40, command=lambda _v: self._on_comunidad_seleccionada())
        self.cb_comunidad.grid(row=1, column=0, sticky="ew", padx=(20, 12), pady=(0, 16))
        self.cb_periodo = UIM.ComboModerno(context, variable=self.periodo_actual, values=[], width=280, height=40, command=lambda _v: self._on_periodo_seleccionado())
        self.cb_periodo.grid(row=1, column=1, sticky="ew", padx=12, pady=(0, 16))
        ctk.CTkLabel(context, text="ELEGIR EXPEDIENTE", font=UIM.fuente(10, "bold"), text_color=C["texto_sec"]).grid(row=2, column=1, sticky="w", padx=12, pady=(0, 3))
        self.cb_expediente = UIM.ComboModerno(
            context,
            variable=self.expediente_seleccionado,
            values=[],
            width=280,
            height=40,
            command=lambda _v: self._on_expediente_seleccionado(),
        )
        self.cb_expediente.grid(row=3, column=1, sticky="ew", padx=12, pady=(0, 16))
        add = ctk.CTkFrame(context, fg_color="transparent")
        add.grid(row=0, column=2, rowspan=4, padx=(12, 18), pady=16)
        ctk.CTkButton(add, text="+ Comunidad", command=self._nueva_comunidad, height=32, corner_radius=9, fg_color="transparent", border_width=1, border_color=C["borde"], text_color=C["primario"], hover_color=C["acento_suave"]).pack(fill="x")
        ctk.CTkButton(add, text="Gestionar periodos", command=self._nuevo_periodo, height=32, corner_radius=9, fg_color=C["acento_suave"], text_color=C["primario"], hover_color=C["acento_suave_hover"]).pack(fill="x", pady=(7, 0))
        ctk.CTkButton(add, text="+ Expediente", command=self._accion_crear_expediente, height=32, corner_radius=9, fg_color=C["primario"], hover_color=C["primario_hover"]).pack(fill="x", pady=(7, 0))

        self.banner_periodo = ctk.CTkFrame(self, fg_color=C["banner_abierto"], corner_radius=14, border_width=1, border_color=C["borde"])
        self.banner_periodo.pack(fill="x", padx=24, pady=(0, 12))
        self.lbl_banner = ctk.CTkLabel(self.banner_periodo, text="Selecciona una comunidad y un ejercicio para preparar la regularización.", font=UIM.fuente(12), text_color=C["primario"], anchor="w")
        self.lbl_banner.pack(fill="x", padx=18, pady=(11, 5))
        self.etapas = {}
        tracker = ctk.CTkFrame(self.banner_periodo, fg_color="transparent")
        tracker.pack(fill="x", padx=18, pady=(1, 11))
        for index, (key, label) in enumerate((("fuentes", "01  Fuentes"), ("validacion", "02  Validar"), ("calculo", "03  Calcular"), ("cartas", "04  Cartas"), ("fin", "05  Terminado"))):
            cell = ctk.CTkFrame(tracker, fg_color="transparent")
            cell.pack(side="left", fill="x", expand=True)
            dot = ctk.CTkLabel(cell, text="○", font=UIM.fuente(16, "bold"), text_color=C["texto_sec"])
            dot.pack(side="left")
            ctk.CTkLabel(cell, text=label, font=UIM.fuente(10, "bold" if index == 0 else "normal"), text_color=C["texto_sec"]).pack(side="left", padx=4)
            self.etapas[key] = dot

        body = ctk.CTkFrame(self, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=24, pady=(0, 14))
        body.grid_columnconfigure(0, weight=1)
        body.grid_columnconfigure(1, weight=3)
        body.grid_columnconfigure(2, weight=2)
        body.grid_rowconfigure(0, weight=3)
        body.grid_rowconfigure(1, weight=2)

        nav = ctk.CTkFrame(body, fg_color=C["panel"], corner_radius=18, border_width=1, border_color=C["borde"])
        nav.grid(row=0, column=0, rowspan=2, sticky="nsew", padx=(0, 12))
        ctk.CTkLabel(nav, text="EL FLUJO", font=UIM.fuente(10, "bold"), text_color=C["texto_sec"]).pack(anchor="w", padx=18, pady=(19, 10))
        self.botones = {}
        acciones = (
            ("1", "Crear expediente", "Define el intervalo a revisar", self._accion_crear_expediente),
            ("2", "Importar modelo inicial", "Carga el Excel histórico una sola vez", self._accion_importar_modelo_inicial),
            ("3", "Añadir fuentes", "Archiva los PDF nuevos sin alterarlos", self._accion_anadir_fuentes),
            ("4", "Resolver incidencias", "Confirma solo los datos pendientes", self._accion_resolver_incidencias),
            ("5", "Generar Excel oficial", "Reconstruye el modelo desde datos validados", self._accion_generar_excel_expediente),
            ("6", "Calcular reparto final", "Cuadra cada concepto al céntimo", self._accion_calcular_reparto_expediente),
            ("7", "Generar cartas", "Prepara una carta auditada por propietario", self._accion_generar_cartas_expediente),
        )
        for number, name, description, command in acciones:
            row = ctk.CTkButton(nav, text=f"{number}   {name}\n     {description}", command=command, anchor="w", height=50, corner_radius=11, font=UIM.fuente(11, "bold"), fg_color=C["acento_suave"], hover_color=C["acento_suave_hover"], text_color=C["primario"])
            row.pack(fill="x", padx=12, pady=4)
            self.botones[name] = row
        ctk.CTkFrame(nav, fg_color=C["borde"], height=1).pack(fill="x", padx=16, pady=16)
        ctk.CTkLabel(nav, text="OTRAS ACCIONES", font=UIM.fuente(10, "bold"), text_color=C["texto_sec"]).pack(anchor="w", padx=18, pady=(0, 7))
        for name, command in (("Regularización guiada", self._accion_regularizacion_guiada), ("Procesar PDFs recibidos", self._accion_procesar_facturas), ("Abrir salidas", self._abrir_salidas), ("Configurar rutas", self._configurar_rutas)):
            button = ctk.CTkButton(nav, text=name, command=command, height=32, corner_radius=8, anchor="w", font=UIM.fuente(11), fg_color="transparent", hover_color=C["acento_suave"], text_color=C["texto_sec"])
            button.pack(fill="x", padx=12, pady=1)
            self.botones[name] = button

        hero = ctk.CTkFrame(body, fg_color=C["panel"], corner_radius=18, border_width=1, border_color=C["borde"])
        hero.grid(row=0, column=1, sticky="nsew", padx=(0, 12))
        ctk.CTkLabel(hero, text="Prepara una regularización", font=UIM.fuente(20, "bold"), text_color=C["texto"]).pack(anchor="w", padx=24, pady=(24, 4))
        ctk.CTkLabel(hero, text="Elige fechas, añade fuentes y resuelve solo los datos que falten.", font=UIM.fuente(12), text_color=C["texto_sec"], justify="left", wraplength=460).pack(anchor="w", padx=24)
        start = ctk.CTkButton(hero, text="Crear expediente", command=self._accion_crear_expediente, height=48, corner_radius=12, font=UIM.fuente(14, "bold"), fg_color=C["primario"], hover_color=C["primario_hover"])
        start.pack(anchor="w", padx=24, pady=(20, 18))
        self.botones["Crear expediente · principal"] = start
        issue_panel = ctk.CTkFrame(hero, fg_color=C["panel_2"], corner_radius=12)
        issue_panel.pack(fill="both", expand=True, padx=24, pady=(0, 24))
        issue_header = ctk.CTkFrame(issue_panel, fg_color="transparent")
        issue_header.pack(fill="x", padx=12, pady=(11, 3))
        ctk.CTkLabel(issue_header, text="INCIDENCIAS ABIERTAS", font=UIM.fuente(10, "bold"), text_color=C["texto_sec"]).pack(side="left")
        self.lbl_incidencias_bandeja = ctk.CTkLabel(issue_header, text="0", font=UIM.fuente(10, "bold"), text_color=C["primario"])
        self.lbl_incidencias_bandeja.pack(side="right")
        self.bandeja_incidencias = ctk.CTkScrollableFrame(
            issue_panel,
            fg_color="transparent",
            height=132,
        )
        self.bandeja_incidencias.pack(fill="both", expand=True, padx=7, pady=(0, 7))

        state = ctk.CTkFrame(body, fg_color=C["panel"], corner_radius=18, border_width=1, border_color=C["borde"])
        state.grid(row=0, column=2, sticky="nsew")
        ctk.CTkLabel(state, text="ESTADO DEL EXPEDIENTE", font=UIM.fuente(10, "bold"), text_color=C["texto_sec"]).pack(anchor="w", padx=18, pady=(20, 10))
        state_summary = ctk.CTkFrame(state, fg_color="transparent")
        state_summary.pack(fill="x", padx=18, pady=(0, 13))
        self.carril_estado_expediente = ctk.CTkFrame(
            state_summary, width=4, height=1, corner_radius=2, fg_color=C["borde"]
        )
        self.carril_estado_expediente.pack(side="left", fill="y", padx=(0, 9))
        self.carril_estado_expediente.pack_propagate(False)
        self.lbl_estado_resumen = ctk.CTkLabel(state_summary, text="Elige un expediente para continuar.", font=UIM.fuente(13, "bold"), text_color=C["texto"], justify="left", wraplength=225)
        self.lbl_estado_resumen.pack(side="left", anchor="w")
        self.expediente_metricas = {}
        for key, label, value in (("rango", "Fechas", "—"), ("perfil", "Perfil Excel", "—"), ("fuentes", "Fuentes", "0"), ("estado", "Estado", "—"), ("incidencias", "Incidencias", "0 abiertas")):
            line = ctk.CTkFrame(state, fg_color=C["panel_2"], corner_radius=9)
            line.pack(fill="x", padx=16, pady=4)
            ctk.CTkLabel(line, text=label, font=UIM.fuente(11), text_color=C["texto_sec"]).pack(side="left", padx=10, pady=9)
            metric = ctk.CTkLabel(line, text=value, font=UIM.fuente(11, "bold"), text_color=C["primario"])
            metric.pack(side="right", padx=10)
            self.expediente_metricas[key] = metric

        activity = ctk.CTkFrame(body, fg_color=C["panel"], corner_radius=18, border_width=1, border_color=C["borde"])
        activity.grid(row=1, column=1, columnspan=2, sticky="nsew")
        top = ctk.CTkFrame(activity, fg_color="transparent")
        top.pack(fill="x", padx=16, pady=(12, 6))
        ctk.CTkLabel(top, text="ACTIVIDAD", font=UIM.fuente(10, "bold"), text_color=C["texto_sec"]).pack(side="left")
        ctk.CTkButton(top, text="Limpiar", command=self._limpiar_log, width=68, height=25, corner_radius=7, fg_color="transparent", hover_color=C["acento_suave"], text_color=C["texto_sec"]).pack(side="right")
        log_wrap = ctk.CTkFrame(activity, fg_color="#102021", corner_radius=11)
        log_wrap.pack(fill="both", expand=True, padx=12, pady=(0, 12))
        mono = UIM.fuente_mono(10)
        self.log_area = tk.Text(log_wrap, font=mono, bg="#102021", fg="#D6E4E2", insertbackground="white", relief="flat", state="disabled", wrap="word", padx=12, pady=10, borderwidth=0, highlightthickness=0, cursor="arrow")
        scroll_log = ctk.CTkScrollbar(log_wrap, command=self.log_area.yview, fg_color="#102021", button_color="#2C4848", button_hover_color="#3F6665")
        self.log_area.configure(yscrollcommand=scroll_log.set)
        scroll_log.pack(side="right", fill="y", padx=(0, 3), pady=4)
        self.log_area.pack(side="left", fill="both", expand=True)
        for tag, color in (("hora", "#6D8988"), ("ok", "#7AD3B6"), ("error", "#FF9A9A"), ("aviso", "#F1C86A"), ("info", "#8BCAC6"), ("titulo", "#FFFFFF"), ("sep", "#29403F"), ("neutro", "#A2BCBA"), ("detalle", "#6D8988"), ("bienvenida", "#A2BCBA")):
            self.log_area.tag_config(tag, foreground=color)

        self.barra_estado = ctk.CTkFrame(self, fg_color=C["panel"], corner_radius=0, height=38, border_width=1, border_color=C["borde"])
        self.barra_estado.pack(fill="x", side="bottom")
        self.barra_estado.pack_propagate(False)
        self.punto_estado = UIM.PuntoEstado(self.barra_estado)
        self.punto_estado.pack(side="left", padx=(22, 6), pady=8)
        self.interruptor.al_cambiar(self.punto_estado.refrescar)
        self.lbl_estado = ctk.CTkLabel(self.barra_estado, text="Listo para empezar", font=UIM.fuente(11), text_color=C["texto"])
        self.lbl_estado.pack(side="left")
        self.lbl_bd = ctk.CTkLabel(self.barra_estado, text="Datos locales protegidos", font=UIM.fuente(10), text_color=C["texto_sec"])
        self.lbl_bd.pack(side="right", padx=20)
        self.progreso = ctk.CTkProgressBar(self.barra_estado, mode="indeterminate", width=140, height=6, corner_radius=3, progress_color=C["primario"], fg_color=C["acento_suave"])

    def _crear_ui(self):
        # ── CABECERA ────────────────────────────────────────────────────────
        cabecera = ctk.CTkFrame(self, fg_color=C["panel"], corner_radius=0,
                                height=64)
        cabecera.pack(fill="x")
        cabecera.pack_propagate(False)

        ctk.CTkLabel(cabecera, text="Regularización de facturas",
                     font=UIM.fuente(20, "bold"),
                     text_color=C["texto"]).pack(side="left",
                                                 padx=(24, 8), pady=14)
        ctk.CTkLabel(cabecera, text="●", font=UIM.fuente(9),
                     text_color=C["primario"]).pack(side="left", pady=14)
        ctk.CTkLabel(cabecera, text="Flujo guiado · edición 2 · cualquier despacho",
                     font=UIM.fuente(12),
                     text_color=C["texto_sec"]).pack(side="left",
                                                     padx=8, pady=14)

        # Interruptor claro / oscuro (a la derecha)
        self.interruptor = UIM.InterruptorTema(cabecera)
        self.interruptor.pack(side="right", padx=20)

        # Línea de acento degradada bajo la cabecera
        self.linea = UIM.LineaGradiente(self, altura=2)
        self.linea.pack(fill="x")
        self.interruptor.al_cambiar(self.linea.refrescar)

        # ── SELECTOR DE COMUNIDAD Y PERIODO ────────────────────────────────
        selector = ctk.CTkFrame(self, fg_color=C["panel"], corner_radius=16,
                                border_width=1, border_color=C["borde"])
        selector.pack(fill="x", padx=16, pady=(14, 0))

        ctk.CTkLabel(selector, text="🏢  COMUNIDAD", font=UIM.fuente(10, "bold"),
                     text_color=C["texto_sec"]).grid(
                         row=0, column=0, padx=(20, 8), pady=(14, 2), sticky="w")
        self.cb_comunidad = UIM.ComboModerno(
            selector, variable=self.comunidad_actual, values=[],
            width=400, height=38,
            command=lambda _v: self._on_comunidad_seleccionada())
        self.cb_comunidad.grid(row=1, column=0, padx=(18, 16),
                               pady=(2, 16), sticky="w")

        ctk.CTkLabel(selector, text="📅  PERIODO", font=UIM.fuente(10, "bold"),
                     text_color=C["texto_sec"]).grid(
                         row=0, column=1, padx=(8, 8), pady=(14, 2), sticky="w")
        self.cb_periodo = UIM.ComboModerno(
            selector, variable=self.periodo_actual, values=[],
            width=170, height=38,
            command=lambda _v: self._on_periodo_seleccionado())
        self.cb_periodo.grid(row=1, column=1, padx=(6, 8),
                             pady=(2, 16), sticky="w")

        ctk.CTkButton(selector, text="＋ Nuevo periodo",
                      command=self._nuevo_periodo,
                      width=140, height=34, corner_radius=8,
                      font=UIM.fuente(12, "bold"),
                      fg_color=C["acento_suave"],
                      hover_color=C["acento_suave_hover"],
                      text_color=C["primario"]).grid(
                          row=1, column=2, padx=8, pady=(4, 14), sticky="w")

        # ── BANNER DE PERIODO ACTIVO ────────────────────────────────────────
        self.banner_periodo = ctk.CTkFrame(
            self, fg_color=C["banner_abierto"], corner_radius=12,
            border_width=1, border_color=C["borde"])
        self.banner_periodo.pack(fill="x", padx=16, pady=(10, 0))
        self.lbl_banner = ctk.CTkLabel(
            self.banner_periodo,
            text="Selecciona una comunidad y un periodo para empezar.",
            font=UIM.fuente(12), text_color=C["primario"],
            anchor="w", justify="left")
        self.lbl_banner.pack(fill="x", padx=14, pady=7)
        self.etapas = {}
        etapas = (("fuentes", "Fuentes"), ("validacion", "Validación"),
                  ("calculo", "Cálculo"), ("cartas", "Cartas"), ("fin", "Listo"))
        tracker = ctk.CTkFrame(self.banner_periodo, fg_color="transparent")
        tracker.pack(fill="x", padx=14, pady=(0, 9))
        for key, label in etapas:
            item = ctk.CTkFrame(tracker, fg_color="transparent")
            item.pack(side="left", expand=True, fill="x")
            dot = ctk.CTkLabel(item, text="○", font=UIM.fuente(16, "bold"), text_color=C["texto_sec"])
            dot.pack(side="left")
            ctk.CTkLabel(item, text=label, font=UIM.fuente(10), text_color=C["texto_sec"]).pack(side="left", padx=3)
            self.etapas[key] = dot

        # ── CUERPO PRINCIPAL ────────────────────────────────────────────────
        cuerpo = ctk.CTkFrame(self, fg_color="transparent")
        cuerpo.pack(fill="both", expand=True, padx=16, pady=12)
        cuerpo.columnconfigure(0, weight=1)
        cuerpo.columnconfigure(1, weight=2)
        cuerpo.rowconfigure(0, weight=1)

        # ── PANEL IZQUIERDO: ACCIONES ───────────────────────────────────────
        panel_acc = ctk.CTkFrame(cuerpo, fg_color=C["panel"], corner_radius=16,
                                 border_width=1, border_color=C["borde"])
        panel_acc.grid(row=0, column=0, sticky="nsew", padx=(0, 10))

        self.botones = {}

        # ── Acción principal ──
        ctk.CTkLabel(panel_acc, text="ACCIÓN PRINCIPAL",
                     font=UIM.fuente(10, "bold"),
                     text_color=C["texto_sec"]).pack(
                         anchor="w", padx=18, pady=(18, 6))

        btn_todo = ctk.CTkButton(
            panel_acc, text="▶   PROCESAR TODO",
            command=self._accion_todo_en_uno,
            height=56, corner_radius=14,
            font=UIM.fuente(15, "bold"),
            fg_color=C["primario"], hover_color=C["primario_hover"])
        btn_todo.pack(fill="x", padx=16, pady=(0, 4))
        self.botones["🔄  TODO EN UNO"] = btn_todo

        ctk.CTkLabel(panel_acc,
                     text="Ingesta entrada/ → BD → reparto → Excel → cartas",
                     font=UIM.fuente(10),
                     text_color=C["texto_sec"]).pack(
                         anchor="w", padx=18, pady=(0, 10))

        ctk.CTkFrame(panel_acc, fg_color=C["borde"], height=1).pack(
            fill="x", padx=16, pady=(2, 10))

        # ── Paso a paso ──
        ctk.CTkLabel(panel_acc, text="PASO A PASO",
                     font=UIM.fuente(10, "bold"),
                     text_color=C["texto_sec"]).pack(
                         anchor="w", padx=18, pady=(0, 6))

        pasos = [
            ("🧾  Regularización guiada", self._accion_regularizacion_guiada),
            ("📥  Procesar facturas", self._accion_procesar_facturas),
            ("📊  Regenerar Excel",   self._accion_actualizar_excel),
            ("🔢  Calcular reparto",  self._accion_calcular_reparto),
            ("✉️  Generar cartas",    self._accion_generar_cartas),
        ]
        claves = ["🧾  Regularización guiada", "📥  Procesar Facturas", "📊  Actualizar Excel",
                  "🔢  Calcular Reparto", "✉️  Generar Cartas"]
        for (texto, cmd), clave in zip(pasos, claves):
            btn = ctk.CTkButton(
                panel_acc, text=texto, command=cmd,
                height=42, corner_radius=10, anchor="w",
                font=UIM.fuente(13, "bold"),
                fg_color=C["acento_suave"],
                hover_color=C["acento_suave_hover"],
                text_color=C["primario"])
            btn.pack(fill="x", padx=16, pady=3)
            self.botones[clave] = btn

        ctk.CTkFrame(panel_acc, fg_color=C["borde"], height=1).pack(
            fill="x", padx=16, pady=10)

        # ── Utilidades ──
        utils = [
            ("📂  Abrir carpeta entrada",        self._abrir_entrada),
            ("📂  Abrir carpeta salidas",        self._abrir_salidas),
            ("⚙️  Configurar rutas",             self._configurar_rutas),
            ("➕  Nueva comunidad",              self._nueva_comunidad),
            ("🔄  Sincronizar comunidades",      self._accion_sincronizar_comunidades),
            ("📥  Importar desde Excel Maestro", self._accion_importar_excel),
        ]
        for texto, cmd in utils:
            ctk.CTkButton(
                panel_acc, text=texto, command=cmd,
                height=32, corner_radius=8, anchor="w",
                font=UIM.fuente(12),
                fg_color="transparent",
                hover_color=C["acento_suave"],
                text_color=C["texto_sec"]).pack(fill="x", padx=10, pady=1)

        # ── PANEL DERECHO: LOG Y ESTADO ─────────────────────────────────────
        panel_log = ctk.CTkFrame(cuerpo, fg_color=C["panel"], corner_radius=16,
                                 border_width=1, border_color=C["borde"])
        panel_log.grid(row=0, column=1, sticky="nsew")

        # Cabecera del log
        log_header = ctk.CTkFrame(panel_log, fg_color="transparent")
        log_header.pack(fill="x")
        ctk.CTkLabel(log_header, text="ACTIVIDAD",
                     font=UIM.fuente(10, "bold"),
                     text_color=C["texto_sec"]).pack(side="left",
                                                     padx=16, pady=8)
        ctk.CTkButton(log_header, text="Limpiar",
                      command=self._limpiar_log,
                      width=70, height=26, corner_radius=8,
                      font=UIM.fuente(11),
                      fg_color="transparent",
                      hover_color=C["acento_suave"],
                      text_color=C["texto_sec"]).pack(side="right",
                                                      padx=10, pady=6)

        # Área de log (estilo terminal moderno — oscura en ambos temas)
        marco_log = ctk.CTkFrame(panel_log, fg_color="#0D1117",
                                 corner_radius=10)
        marco_log.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        mono = UIM.fuente_mono(11)
        self.log_area = tk.Text(
            marco_log,
            font=mono,
            bg="#0D1117", fg="#C9D1D9",
            insertbackground="white",
            relief="flat",
            state="disabled",
            wrap="word",
            padx=14, pady=12,
            borderwidth=0, highlightthickness=0,
            cursor="arrow",
            spacing1=1, spacing3=1,
        )
        scroll_log = ctk.CTkScrollbar(marco_log, command=self.log_area.yview,
                                      fg_color="#0D1117",
                                      button_color="#30363D",
                                      button_hover_color="#484F58")
        self.log_area.configure(yscrollcommand=scroll_log.set)
        scroll_log.pack(side="right", fill="y", padx=(0, 2), pady=4)
        self.log_area.pack(side="left", fill="both", expand=True)

        # Tags de estilo para el log
        self.log_area.tag_config("hora",    foreground="#484F58",
                                  font=(mono[0], 9))
        self.log_area.tag_config("ok",      foreground="#4AC26B")
        self.log_area.tag_config("error",   foreground="#FF8182",
                                  background="#2A1215",
                                  lmargin1=6, lmargin2=26,
                                  spacing1=2, spacing3=2)
        self.log_area.tag_config("aviso",   foreground="#F0B72F",
                                  lmargin1=6, lmargin2=26)
        self.log_area.tag_config("info",    foreground="#6CB6FF")
        self.log_area.tag_config("titulo",  foreground="#E6EDF3",
                                  font=(mono[0], 12, "bold"),
                                  spacing1=10, spacing3=2)
        self.log_area.tag_config("sep",     foreground="#21262D",
                                  spacing3=4)
        self.log_area.tag_config("neutro",  foreground="#8B949E")
        self.log_area.tag_config("detalle", foreground="#57606A",
                                  font=(mono[0], 9),
                                  lmargin1=26, lmargin2=26)
        self.log_area.tag_config("bienvenida", foreground="#8B949E",
                                  font=(mono[0], 10, "italic"))

        # ── BARRA DE ESTADO ─────────────────────────────────────────────────
        self.barra_estado = ctk.CTkFrame(self, fg_color=C["panel"],
                                         corner_radius=0, height=36,
                                         border_width=1,
                                         border_color=C["borde"])
        self.barra_estado.pack(fill="x", side="bottom")
        self.barra_estado.pack_propagate(False)

        self.punto_estado = UIM.PuntoEstado(self.barra_estado)
        self.punto_estado.pack(side="left", padx=(14, 4), pady=7)
        self.interruptor.al_cambiar(self.punto_estado.refrescar)

        self.lbl_estado = ctk.CTkLabel(self.barra_estado, text="Listo",
                                       font=UIM.fuente(12),
                                       text_color=C["texto"])
        self.lbl_estado.pack(side="left", padx=(0, 12))

        self.lbl_bd = ctk.CTkLabel(self.barra_estado,
                                   text=f"BD: {RUTA_BD}",
                                   font=UIM.fuente(10),
                                   text_color=C["texto_sec"])
        self.lbl_bd.pack(side="right", padx=14)

        # Barra de progreso (oculta por defecto)
        self.progreso = ctk.CTkProgressBar(
            self.barra_estado, mode="indeterminate",
            width=130, height=6, corner_radius=3,
            progress_color=C["primario"], fg_color=C["acento_suave"])

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

    def _limpiar_log(self):
        self.log_area.configure(state="normal")
        self.log_area.delete("1.0", "end")
        self.log_area.configure(state="disabled")

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
        for carpeta in [RUTA_ENTRADA, RUTA_PROCESADOS,
                        RUTA_EXCELS, RUTA_CARTAS,
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
        sel = self.comunidad_actual.get()
        self.id_comunidad = self._ids_comunidad.get(sel)
        if not self.id_comunidad or not MOD.get("gestor_bd"):
            return
        try:
            con = MOD["gestor_bd"].conectar(str(RUTA_BD))
            rows = con.execute(
                "SELECT id_periodo, nombre FROM periodos WHERE id_comunidad=? ORDER BY fecha_inicio DESC",
                (self.id_comunidad,)
            ).fetchall()
            con.close()
            opciones = [r["nombre"] for r in rows]
            self.cb_periodo.configure(values=opciones)
            self._ids_periodo = {r["nombre"]: r["id_periodo"] for r in rows}
            if opciones:
                self.cb_periodo.set(opciones[0])
                self._on_periodo_seleccionado()
            else:
                self._actualizar_banner()
        except Exception:
            pass

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

    def _on_periodo_seleccionado(self, event=None):
        nombre = self.periodo_actual.get()
        self.id_periodo = self._ids_periodo.get(nombre) if hasattr(self, "_ids_periodo") else None
        self._actualizar_banner()

    def _actualizar_banner(self):
        """Actualiza el banner con la info del periodo activo y su rango de fechas."""
        if not self.id_periodo or not MOD.get("gestor_bd") or not RUTA_BD.exists():
            if hasattr(self, "lbl_banner"):
                self.lbl_banner.configure(
                    text="Selecciona una comunidad y un periodo para empezar.",
                    text_color=C["primario"])
                self.banner_periodo.configure(fg_color=C["banner_abierto"])
            return
        try:
            con = MOD["gestor_bd"].conectar(str(RUTA_BD))
            r = con.execute(
                "SELECT nombre, fecha_inicio, fecha_fin, estado FROM periodos "
                "WHERE id_periodo=?", (self.id_periodo,)
            ).fetchone()
            # Contar facturas en este periodo
            n_facturas = con.execute(
                "SELECT COUNT(*) FROM facturas WHERE id_periodo=?",
                (self.id_periodo,)
            ).fetchone()[0]
            con.close()

            if r:
                fi = r["fecha_inicio"] or "?"
                ff = r["fecha_fin"] or "abierto"
                estado = r["estado"] or "abierto"
                try:
                    dias = (date.fromisoformat(ff[:10]) -
                            date.fromisoformat(fi[:10])).days
                    dur = f"{dias} días"
                except Exception:
                    dur = ""
                abierto = estado == "abierto"
                icono   = "📅" if abierto else "🔒"
                texto = (
                    f"{icono}  Periodo activo: {r['nombre']}  ·  "
                    f"{fi} → {ff}  ({dur})  ·  "
                    f"{n_facturas} factura(s) registrada(s)   "
                    f"— Las facturas procesadas se asignarán al periodo que corresponda por fecha."
                )
                self.lbl_banner.configure(
                    text=texto,
                    text_color=C["primario"] if abierto else C["texto_sec"])
                self.banner_periodo.configure(
                    fg_color=C["banner_abierto"] if abierto else C["banner_cerrado"])
        except Exception:
            pass

    # -----------------------------------------------------------------------
    # ACCIONES PRINCIPALES — se ejecutan en hilo secundario
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
        finally:
            self._procesando = False
            self.after(0, self._habilitar_botones)
            self._estado("Listo")

    def _accion_procesar_facturas(self):
        self._en_hilo(self._procesar_facturas_impl)

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

    def _accion_anadir_fuentes(self):
        expedient_ui = MOD.get("expedient_ui")
        ingestion = MOD.get("case_ingestion")
        database = MOD.get("gestor_bd")
        if not expedient_ui or not ingestion or not database:
            self.log("No está disponible la interfaz para añadir fuentes.", "error")
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

    def _accion_resolver_incidencias(self):
        review = MOD.get("document_review")
        ingestion = MOD.get("case_ingestion")
        database = MOD.get("gestor_bd")
        if not review or not ingestion or not database:
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
            self.log("No hay incidencias pendientes en este expediente.", "ok")
            return
        self._refrescar_bandeja_incidencias(issues)
        self.log(
            "Elige Resolver en la incidencia concreta de la bandeja.",
            "info",
        )

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
            open_count = len(issues)
            workflow = MOD.get("case_workflow_actions")
            if workflow:
                try:
                    profile = workflow.resolve_case_profile(
                        connection, id_case=self.id_expediente,
                        active_community_id=self.id_comunidad, project_root=BASE_DIR,
                    )
                    profile_label = profile.key
                except workflow.WorkflowBlockedError as error:
                    profile_label = f"Perfil pendiente: {error}"
            else:
                profile_label = "Perfil no disponible"
        except LookupError as exc:
            self._limpiar_contexto_expediente()
            self._refrescar_lista_expedientes()
            self.log(str(exc), "aviso")
            return
        finally:
            connection.close()

        self.expediente_actual.set(case.name)
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
            f"Perfil: {profile_label}\n{open_count} incidencia(s) abierta(s)\n"
            f"Estado: {status_label}"
        )
        self._resumen_ejercicio(summary)
        self._refrescar_bandeja_incidencias(issues)

        def update_metrics():
            metrics = getattr(self, "expediente_metricas", {})
            values = {
                "rango": date_range,
                "perfil": profile_label,
                "fuentes": str(document_count),
                "estado": status_label,
                "incidencias": f"{open_count} abiertas",
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
            f"estado {case.status} · {open_count} incidencia(s) abierta(s)",
            "info",
        )
        if open_count:
            self._actualizar_etapa("validacion")
            self.log(f"{open_count} incidencia(s) por resolver antes de generar el Excel oficial", "aviso")
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

    def _refrescar_bandeja_incidencias(self, issues=()):
        issues = tuple(issues)

        def update_tray():
            tray = getattr(self, "bandeja_incidencias", None)
            if tray is None:
                return
            for child in tray.winfo_children():
                child.destroy()
            self.lbl_incidencias_bandeja.configure(text=str(len(issues)))

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

            expedient_ui = MOD.get("expedient_ui")
            for issue in issues:
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
                    text=issue.field_name,
                    font=UIM.fuente(11, "bold"),
                    text_color=C["texto"],
                    anchor="w",
                ).pack(fill="x")
                ctk.CTkLabel(
                    detail,
                    text=f"{issue.archived_path.name} · {issue.message}",
                    font=UIM.fuente(10),
                    text_color=C["texto_sec"],
                    anchor="w",
                    justify="left",
                    wraplength=260,
                ).pack(fill="x", pady=(2, 0))
                actions = ctk.CTkFrame(row, fg_color="transparent")
                actions.pack(side="right", padx=8, pady=7)
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
                    text="Resolver",
                    width=86,
                    height=28,
                    corner_radius=7,
                    fg_color=C["primario"],
                    hover_color=C["primario_hover"],
                    command=lambda current=issue: expedient_ui.open_issue_dialog(
                        self, current
                    ),
                ).pack()

        self.after(0, update_tray)

    def _limpiar_contexto_expediente(self):
        self.id_expediente = None
        self.expediente_actual.set("")
        self.expediente_seleccionado.set("")
        self._ids_expediente = {}
        if hasattr(self, "cb_expediente"):
            self.cb_expediente.configure(values=[])
        self._resumen_ejercicio("Elige un expediente para continuar.")
        self._refrescar_bandeja_incidencias()

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

    def _accion_regularizacion_guiada(self):
        if not self._validar_seleccion():
            return
        referencia = filedialog.askopenfilename(
            title="Selecciona el Excel de referencia económica",
            filetypes=[("Excel", "*.xlsx *.xls"), ("Todos", "*.*")],
        )
        if not referencia:
            return
        propietarios = filedialog.askopenfilename(
            title="Selecciona el listado de propietarios",
            filetypes=[("CSV", "*.csv"), ("Todos", "*.*")],
        )
        if not propietarios:
            return
        lecturas = filedialog.askopenfilename(
            title="Selecciona el Excel de lecturas de contadores",
            filetypes=[("Excel", "*.xls *.xlsx"), ("Todos", "*.*")],
        )
        if not lecturas:
            return
        self._en_hilo(lambda: self._regularizacion_guiada_impl(referencia, propietarios, lecturas))

    def _regularizacion_guiada_impl(self, referencia, propietarios, lecturas):
        flujo = MOD.get("regularization_flow")
        gbd = MOD.get("gestor_bd")
        if not flujo or not gbd:
            self.log("❌ No está disponible el flujo de regularización guiada", "error")
            return
        self._estado("Importando fuentes…", procesando=True)
        self.log("━━━ REGULARIZACIÓN GUIADA ━━━", "titulo")
        def progress(stage, details):
            labels = {
                "reading_owners": ("fuentes", "Leyendo propietarios…"),
                "reading_reference": ("fuentes", "Leyendo Excel de referencia…"),
                "validating_reference": ("validacion", "Validando importes y periodo…"),
                "reading_individual_readings": ("fuentes", "Leyendo contadores…"),
                "individual_readings_imported": ("validacion", "Lecturas validadas"),
                "period_calculated": ("calculo", "Resultados calculados"),
                "regularization_completed": ("fin", "Regularización lista"),
            }
            if stage in labels:
                key, text = labels[stage]
                self._actualizar_etapa(key)
                self._estado(text, procesando=True)
                self._resumen_ejercicio(text)
                self.log(f"  ▸ {text}", "info")
                if stage == "individual_readings_imported" and details.get("carried_forward"):
                    self.log(f"  ⚠️ {details['carried_forward']} contador(es) con lectura anterior arrastrada", "aviso")
        con = gbd.conectar(str(RUTA_BD))
        try:
            resultado = flujo.run_regularization(
                con, self.id_comunidad, referencia, propietarios, lecturas, progress=progress
            )
        finally:
            con.close()
        self.id_periodo = resultado["period_id"]
        try:
            self.after(0, self._cargar_periodos)
        except Exception:
            pass
        self._actualizar_etapa("cartas")
        self._resumen_ejercicio(
            f"Cálculo cuadrado para {resultado['reading_summary'].participating_properties} propiedades. Elige los conceptos para crear las cartas."
        )
        self.log(f"  ✅ {resultado['results']} resultados calculados; conciliación cuadrada", "ok")
        self.log("  ℹ️ Ahora puedes generar el Excel y las cartas desde los botones del flujo.", "info")
        self.after(0, self._accion_generar_cartas)

    def _actualizar_etapa(self, activa):
        def _actualizar():
            colores = {"fuentes": C["primario"], "validacion": C["primario"], "calculo": C["primario"], "cartas": C["primario"], "fin": C["exito"]}
            orden = ["fuentes", "validacion", "calculo", "cartas", "fin"]
            for key, dot in getattr(self, "etapas", {}).items():
                if orden.index(key) <= orden.index(activa):
                    dot.configure(text="●", text_color=colores.get(key, C["primario"]))
                else:
                    dot.configure(text="○", text_color=C["texto_sec"])
        self.after(0, _actualizar)

    def _accion_actualizar_excel(self):
        self._en_hilo(self._actualizar_excel_impl)

    def _accion_calcular_reparto(self):
        self._en_hilo(self._calcular_reparto_impl)

    def _accion_generar_cartas(self):
        if not self._validar_seleccion():
            return
        selected = self._dialogo_conceptos_cartas()
        if selected is None:
            return
        self._en_hilo(lambda: self._generar_cartas_impl(selected))

    def _accion_todo_en_uno(self):
        self._en_hilo(self._todo_en_uno_impl)

    def _accion_sincronizar_comunidades(self):
        self._en_hilo(self._sincronizar_comunidades_impl)

    def _accion_importar_excel(self):
        self._en_hilo(self._importar_excel_impl)

    # -----------------------------------------------------------------------
    # IMPLEMENTACIONES
    # -----------------------------------------------------------------------
    def _validar_seleccion(self) -> bool:
        if not self.id_comunidad:
            self.log("⚠️  Selecciona una comunidad primero", "aviso")
            return False
        if not self.id_periodo:
            self.log("⚠️  Selecciona un periodo primero", "aviso")
            return False
        return True

    def _detectar_o_crear_periodo_para_factura(self, con, gbd, datos: dict,
                                                id_comunidad: int = None):
        """
        Dado los datos de una factura (con fecha_inicio / fecha_fin o fecha_factura),
        busca en la BD el periodo de la comunidad que solapa con esas fechas.
        Si no existe, crea uno automáticamente con nombre 'YYYY-YYYY' según el año fiscal.
        Devuelve id_periodo o None si no se puede determinar.
        """
        id_com = id_comunidad or self.id_comunidad
        fecha_str = (datos.get("fecha_inicio") or
                     datos.get("fecha_factura") or
                     datos.get("fecha_fin"))
        if not fecha_str:
            return None

        try:
            fecha = date.fromisoformat(fecha_str[:10])
        except (ValueError, TypeError):
            return None

        # Buscar periodos existentes de esta comunidad que solapen con la fecha
        rows = con.execute(
            "SELECT id_periodo, nombre, fecha_inicio, fecha_fin FROM periodos "
            "WHERE id_comunidad=? AND fecha_inicio IS NOT NULL AND fecha_fin IS NOT NULL "
            "ORDER BY fecha_inicio DESC",
            (id_com,)
        ).fetchall()

        for r in rows:
            try:
                fi = date.fromisoformat(r["fecha_inicio"][:10])
                ff = date.fromisoformat(r["fecha_fin"][:10])
                if fi <= fecha <= ff:
                    return r["id_periodo"]
            except (ValueError, TypeError):
                continue

        # No hay periodo que cubra esta fecha → crear uno automático
        # Año fiscal: sep YYYY → ago YYYY+1  (o ene→dic si es período simple)
        if fecha.month >= 9:
            anio_ini, anio_fin = fecha.year, fecha.year + 1
            f_ini = date(anio_ini, 9, 1)
            f_fin = date(anio_fin, 8, 31)
        else:
            anio_ini, anio_fin = fecha.year - 1, fecha.year
            f_ini = date(anio_ini, 9, 1)
            f_fin = date(anio_fin, 8, 31)

        nombre_auto = f"{anio_ini}-{anio_fin}"
        self.log(
            f"     📅 Periodo '{nombre_auto}' creado automáticamente para factura de {fecha_str[:10]}",
            "aviso"
        )
        try:
            id_per = gbd.obtener_o_crear_periodo(
                con, id_com, nombre_auto, str(f_ini)
            )
            con.execute(
                "UPDATE periodos SET fecha_fin=?, estado='abierto' WHERE id_periodo=?",
                (str(f_fin), id_per)
            )
            con.commit()
            # Refrescar la lista de periodos en la UI
            self.after(0, self._on_comunidad_seleccionada)
            return id_per
        except Exception as e:
            self.log(f"     ⚠️  No se pudo crear periodo automático: {e}", "error")
            return None

    def _procesar_facturas_impl(self):
        if not self._validar_seleccion():
            return
        if not MOD.get("lector_pdf") or not MOD.get("gestor_bd"):
            self.log("❌ Módulos lector_pdf o gestor_bd no disponibles", "error")
            return

        self._estado("Procesando facturas…", procesando=True)
        self.log("━━━ PROCESAR FACTURAS ━━━", "titulo")
        self.log(
            "  ℹ️  Las facturas se asignan al periodo que corresponda por su fecha\n"
            "     (independientemente del periodo seleccionado en la UI).",
            "neutro"
        )

        archivos = sorted([
            f for f in os.listdir(RUTA_ENTRADA)
            if f.lower().endswith(".pdf")
        ])

        if not archivos:
            self.log(f"📂 Carpeta entrada/ vacía: {RUTA_ENTRADA}", "aviso")
            return

        self.log(f"📂 {len(archivos)} archivo(s) encontrado(s)", "info")

        gbd  = MOD["gestor_bd"]
        lpdf = MOD["lector_pdf"]
        con  = gbd.conectar(str(RUTA_BD))

        ok_count = err_count = dup_count = 0

        # Obtener el código de la comunidad activa para usarlo como fallback
        codigo_comunidad_activa = self.comunidad_actual.get().split(" — ")[0].strip()

        for nombre in archivos:
            ruta = RUTA_ENTRADA / nombre

            if gbd.archivo_ya_procesado(con, nombre):
                self.log(f"  ⏭  {nombre} — ya procesado", "neutro")
                dup_count += 1
                continue

            # --- Detectar código de comunidad desde el nombre del archivo ---
            # Intentamos extraer un código numérico del nombre (ej: "644_Endesa.pdf",
            # "Endesa-644.pdf", "644.pdf"). Si no hay número, usamos la comunidad activa.
            m_codigo = re.search(r'(?<![A-Za-z])(\d{3,6})(?![A-Za-z0-9])', nombre)
            codigo = m_codigo.group(1) if m_codigo else codigo_comunidad_activa

            # Resolver el id_comunidad real a partir del código detectado
            id_com_factura = self.id_comunidad  # fallback: comunidad seleccionada en UI
            if m_codigo and codigo != codigo_comunidad_activa:
                row_com = con.execute(
                    "SELECT id_comunidad, nombre FROM comunidades WHERE codigo=?",
                    (codigo,)
                ).fetchone()
                if row_com:
                    id_com_factura = row_com["id_comunidad"]
                    self.log(
                        f"     🏘️  Comunidad detectada del nombre: {codigo} — {row_com['nombre']}",
                        "info"
                    )
                else:
                    self.log(
                        f"     ⚠️  Código '{codigo}' no registrado en BD → usando comunidad activa",
                        "aviso"
                    )

            resultado = lpdf.procesar_archivo(
                str(ruta), codigo,
                con_bd=con,
                ruta_proveedores=str(RUTA_PROVEEDORES)
            )

            if resultado["ok"] and resultado["tipo"] == "FACTURA":
                datos = resultado["datos"]
                datos["id_comunidad"] = id_com_factura

                # ── Detectar periodo automáticamente por la fecha de la factura ──
                # La factura puede ser de cualquier año; la insertamos en el periodo
                # correcto independientemente del periodo seleccionado en la UI.
                id_periodo_factura = self._detectar_o_crear_periodo_para_factura(
                    con, gbd, datos, id_comunidad=id_com_factura
                )
                id_per_fallback = (self.id_periodo if id_com_factura == self.id_comunidad
                                   else None)
                datos["id_periodo"] = id_periodo_factura or id_per_fallback

                id_fac = gbd.insertar_factura(con, datos)
                if id_fac:
                    gbd.marcar_archivo_procesado(con, nombre,
                                                  resultado["hash_md5"], id_fac)
                    dest = RUTA_PROCESADOS / nombre
                    mover_seguro(ruta, dest, self.log)
                    tipo = datos.get("tipo_suministro", "?")
                    imp  = datos.get("importe_total", 0)
                    self.log(f"  ✅ {nombre[:45]:<45} {tipo:<15} {imp:>9.2f} €", "ok")
                    ok_count += 1
                else:
                    self.log(f"  ⚠️  {nombre} — duplicado en BD", "aviso")
                    dup_count += 1

            elif resultado["ok"] and resultado["tipo"] == "LECTURA_METRIGEST":
                nvecs    = len(resultado["datos"].get("vecinos", []))
                tipo_lec = resultado["datos"].get("tipo", "?")
                self.log(f"  📊 {nombre[:45]:<45} LECTURA {tipo_lec} — {nvecs} vecinos", "info")

                # Importar lecturas Metrigest automáticamente a la BD
                if MOD.get("importar_lecturas_metrigest"):
                    res_lec = MOD["importar_lecturas_metrigest"].importar_lecturas_pdf(
                        ruta_pdf=str(ruta),
                        ruta_bd=str(RUTA_BD),
                        mover_procesado=True,
                        ruta_proveedores=str(RUTA_PROVEEDORES),
                        verbose=False,
                    )
                    if res_lec.get("ok"):
                        ins = res_lec.get("lecturas_insertadas", 0)
                        dup = res_lec.get("lecturas_duplicadas", 0)
                        nf  = res_lec.get("vecinos_no_encontrados", [])
                        self.log(
                            f"     → {ins} lecturas insertadas, {dup} duplicadas"
                            + (f", {len(nf)} vecinos no encontrados" if nf else ""),
                            "ok" if ins > 0 else "aviso"
                        )
                        if nf:
                            self.log(f"     → Sin match: {', '.join(str(v) for v in nf[:8])}", "aviso")
                    else:
                        self.log(f"     ❌ {res_lec.get('error','?')}", "error")
                        # Aunque falle la inserción de lecturas, marcar como procesado
                        gbd.marcar_archivo_procesado(con, nombre,
                                                      resultado["hash_md5"],
                                                      resultado="ok",
                                                      notas=f"LECTURA {tipo_lec} (sin vecinos)")
                else:
                    # Módulo no disponible: solo marcar
                    gbd.marcar_archivo_procesado(con, nombre,
                                                  resultado["hash_md5"],
                                                  resultado="ok",
                                                  notas=f"LECTURA {tipo_lec}")
                    self.log("     ⚠️  importar_lecturas_metrigest no disponible", "aviso")
                ok_count += 1
            else:
                motivo = resultado.get('motivo', '?')
                detalle = resultado.get('detalle', '')
                # Mensajes de error más descriptivos
                if motivo == "CUPS_NO_COINCIDE":
                    cups_pdf = resultado.get('cups_pdf', '?')
                    cups_esp = resultado.get('cups_esperado', '?')
                    self.log(
                        f"  ❌ {nombre[:40]:<40} CUPS incorrecto\n"
                        f"       PDF tiene: {cups_pdf}\n"
                        f"       Esperado:  {cups_esp}",
                        "error"
                    )
                elif motivo == "COMUNIDAD_CONTRADICTORIA":
                    self.log(
                        f"  ❌ {nombre[:40]:<40} {motivo}\n"
                        f"       {detalle}",
                        "error"
                    )
                elif motivo == "SIN_TEXTO":
                    self.log(
                        f"  ❌ {nombre[:40]:<40} Sin texto extraíble\n"
                        f"       (PDF escaneado — necesita OCR o Tesseract instalado)",
                        "error"
                    )
                elif motivo == "PROVEEDOR_NO_IDENTIFICADO":
                    self.log(
                        f"  ❌ {nombre[:40]:<40} Proveedor no reconocido\n"
                        f"       Revisa proveedores.json o renombra el archivo",
                        "error"
                    )
                else:
                    self.log(f"  ❌ {nombre[:45]:<45} {motivo}: {detalle[:60]}", "error")
                err_count += 1

        con.close()
        self.log(f"\n  ✅ Procesados: {ok_count}  |  ⚠️  Duplicados: {dup_count}  |  ❌ Errores: {err_count}", "info")

    def _actualizar_excel_impl(self):
        """
        Regenera el Excel Maestro COMPLETO desde la BD (todos los periodos).
        La BD es la fuente de verdad; el Excel es un informe de salida.
        El libro anterior queda en Excels_Maestros/backups/.
        """
        if not self._validar_seleccion():
            return
        try:
            import excel_generator
        except ImportError as e:
            self.log(f"❌ No se pudo cargar excel_generator.py: {e}", "error")
            return

        self._estado("Regenerando Excel…", procesando=True)
        self.log("━━━ REGENERAR EXCEL MAESTRO ━━━", "titulo")

        codigo = self.comunidad_actual.get().split(" — ")[0].strip()
        self.log(f"📊 Comunidad {codigo}: reconstruyendo libro completo desde la BD…", "info")

        resultado = excel_generator.regenerar_excel_comunidad(
            codigo,
            ruta_bd=str(RUTA_BD),
            ruta_excels=str(RUTA_EXCELS),
            log=self.log,
        )

        if resultado["ok"]:
            self.log(f"  ✅ {Path(resultado['archivo']).name} regenerado "
                     f"({len(resultado['periodos_volcados'])} periodos, "
                     f"{resultado['filas_escritas']} filas)", "ok")
            if resultado.get("backup"):
                self.log(f"  💾 Anterior guardado en backups/{Path(resultado['backup']).name}", "neutro")
        else:
            for err in resultado.get("errores", []):
                self.log(f"  ❌ {err}", "error")

    def _progreso_expediente(self, stage, details):
        """Traduce hitos técnicos a actividad que puede seguir el despacho."""
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
            self._actualizar_banner()
        self.after(0, refresh)

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
        result = workflow.run_generate_excel(
            self.ruta_bd_expedientes,
            id_case=self.id_expediente,
            active_community_id=self.id_comunidad,
            project_root=BASE_DIR,
            output_root=BASE_DIR,
            progress=self._progreso_expediente,
        )
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

    def _calcular_reparto_impl(self):
        if not self._validar_seleccion():
            return
        if not MOD.get("motor_reparto") or not MOD.get("gestor_bd"):
            self.log("❌ Módulo motor_reparto no disponible", "error")
            return

        self._estado("Calculando reparto…", procesando=True)
        self.log("━━━ CALCULAR REPARTO ━━━", "titulo")

        con = MOD["gestor_bd"].conectar(str(RUTA_BD))
        resultado = MOD["motor_reparto"].calcular_reparto(
            con=con,
            id_comunidad=self.id_comunidad,
            id_periodo=self.id_periodo,
            sobrescribir=True
        )
        con.close()

        if resultado.get("ok"):
            rc = resultado["resumen_costes"]
            rr = resultado["resumen_repartos"]
            self.log(f"  ✅ Periodo: {resultado['periodo']}", "ok")
            self.log(f"  👥 Vecinos procesados: {resultado['vecinos_procesados']}", "info")
            self.log(f"  💶 Coste total ACS:   {rc['coste_total_acs']:>10.2f} €", "info")
            self.log(f"  💶 Coste total CALEF: {rc['coste_total_cal']:>10.2f} €", "info")
            self.log(f"  📊 Media por vecino ACS:   {rr['media_acs_vecino']:>8.2f} €", "neutro")
            self.log(f"  📊 Media por vecino CALEF: {rr['media_cal_vecino']:>8.2f} €", "neutro")
            self.log(f"  ✅ Registros guardados: {resultado['registros_guardados']}", "ok")
            for aviso in resultado.get("avisos", [])[:5]:
                self.log(f"  ⚠️  {aviso}", "aviso")
        else:
            self.log(f"  ❌ {resultado.get('error', 'Error desconocido')}", "error")

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

    def _generar_cartas_impl(self, selected_concepts=None):
        if not self._validar_seleccion():
            return
        if not MOD.get("carta_writer"):
            self.log("❌ Módulo carta_writer no disponible", "error")
            return
        if not RUTA_PLANTILLA.exists():
            self.log(f"❌ No se encuentra la plantilla: {RUTA_PLANTILLA}", "error")
            return

        self._estado("Generando cartas…", procesando=True)
        self.log("━━━ GENERAR CARTAS ━━━", "titulo")
        if selected_concepts:
            self.log(f"  ▸ Conceptos incluidos: {', '.join(selected_concepts)}", "info")

        nombre_periodo = self.periodo_actual.get()
        carpeta_salida = RUTA_CARTAS / nombre_periodo
        carpeta_salida.mkdir(parents=True, exist_ok=True)

        resultado = MOD["carta_writer"].generar_todas_las_cartas(
            ruta_bd=str(RUTA_BD),
            id_comunidad=self.id_comunidad,
            nombre_periodo=nombre_periodo,
            ruta_plantilla=str(RUTA_PLANTILLA),
            carpeta_salida=str(carpeta_salida),
            selected_concepts=selected_concepts,
        )

        if resultado.get("ok"):
            self.log(f"  ✅ {resultado['cartas_generadas']} cartas generadas", "ok")
            self.log(f"  📁 Carpeta: {carpeta_salida}", "info")
        else:
            self.log(f"  ❌ {resultado.get('error','')}", "error")
            for err in resultado.get("errores", [])[:10]:
                self.log(f"     {err}", "error")

        # Ofrecer abrir la carpeta
        if resultado.get("cartas_generadas", 0) > 0:
            self.after(0, lambda: self._ofrecer_abrir_carpeta(str(carpeta_salida)))

    def _sincronizar_comunidades_impl(self):
        """
        Lee LISTADO_COMUNIDADES.xlsx (o LISTADO_COMUNIDADES.xlsx en la raíz del proyecto)
        e inserta en la BD todas las comunidades que aún no estén registradas.
        No toca comunidades ya existentes — solo añade las nuevas.
        """
        self._estado("Sincronizando comunidades…", procesando=True)
        self.log("━━━ SINCRONIZAR COMUNIDADES ━━━", "titulo")

        if not MOD.get("gestor_bd"):
            self.log("❌ Módulo gestor_bd no disponible", "error")
            return

        # Buscar el archivo LISTADO en varias ubicaciones típicas
        candidatos = [
            BASE_DIR / "LISTADO_COMUNIDADES.xlsx",
            BASE_DIR / "data_fuente" / "LISTADO_COMUNIDADES.xlsx",
            BASE_DIR / "config" / "LISTADO_COMUNIDADES.xlsx",
            Path(__file__).parent / "LISTADO_COMUNIDADES.xlsx",
        ]
        ruta_listado = next((p for p in candidatos if p.exists()), None)

        if ruta_listado is None:
            self.log("❌ Falta el archivo LISTADO_COMUNIDADES.xlsx", "error")
            self.log("   👉 Copia tu listado de comunidades (Excel) en la carpeta del programa:", "aviso")
            self.log(f"      {BASE_DIR}", "aviso")
            self.log("      y renómbralo a LISTADO_COMUNIDADES.xlsx. Después vuelve a pulsar este botón.", "aviso")
            return

        self.log(f"📄 Leyendo: {ruta_listado.name}", "info")

        try:
            import openpyxl
            wb = openpyxl.load_workbook(str(ruta_listado), read_only=True, data_only=True)
            ws = wb.active

            # Detectar cabecera (fila 1)
            cabecera = [str(c.value).strip().lower() if c.value else "" for c in next(ws.iter_rows(min_row=1, max_row=1))]

            # Mapear columnas por nombre o por posición por defecto
            col_codigo = next((i for i, h in enumerate(cabecera) if "num" in h or "codigo" in h or "empresa" in h), 0)
            col_nombre = next((i for i, h in enumerate(cabecera) if "raz" in h or "nombre" in h or "social" in h), 1)
            col_cif    = next((i for i, h in enumerate(cabecera) if "nif" in h or "cif" in h), 2)

            gbd = MOD["gestor_bd"]
            con = gbd.conectar(str(RUTA_BD))

            nuevas = 0
            ya_existentes = 0
            errores = 0

            for fila in ws.iter_rows(min_row=2, values_only=True):
                if not fila or fila[col_codigo] is None:
                    continue

                codigo_raw = str(fila[col_codigo]).strip()
                nombre_raw = str(fila[col_nombre]).strip() if fila[col_nombre] else f"COMUNIDAD {codigo_raw}"
                cif_raw    = str(fila[col_cif]).strip()    if len(fila) > col_cif and fila[col_cif] else None

                if not codigo_raw or codigo_raw.lower() in ("none", ""):
                    continue

                # Comprobar si ya existe
                existe = con.execute(
                    "SELECT id_comunidad FROM comunidades WHERE codigo=?", (codigo_raw,)
                ).fetchone()

                if existe:
                    ya_existentes += 1
                    # Actualizar CIF si faltaba
                    if cif_raw:
                        con.execute(
                            "UPDATE comunidades SET cif=? WHERE codigo=? AND (cif IS NULL OR cif='')",
                            (cif_raw, codigo_raw)
                        )
                    # Asegurar que tiene Excel aunque ya existiera en la BD
                    self._crear_excel_si_no_existe(codigo_raw, nombre_raw)
                    continue

                try:
                    gbd.obtener_o_crear_comunidad(con, codigo_raw, nombre_raw, cif=cif_raw)
                    self.log(f"  ➕ {codigo_raw} — {nombre_raw[:50]}", "ok")
                    nuevas += 1
                    # Crear Excel maestro si no existe
                    self._crear_excel_si_no_existe(codigo_raw, nombre_raw)
                except Exception as e:
                    self.log(f"  ❌ {codigo_raw}: {e}", "error")
                    errores += 1

            con.commit()
            con.close()
            wb.close()

            self.log(
                f"\n  ✅ Nuevas: {nuevas}  |  Ya existían: {ya_existentes}  |  Errores: {errores}",
                "ok" if errores == 0 else "aviso"
            )
            if nuevas > 0:
                self.after(0, self._cargar_comunidades)

        except Exception as e:
            self.log(f"❌ Error leyendo LISTADO: {e}", "error")

    def _importar_excel_impl(self):
        """
        Lee las facturas del periodo activo desde el Excel Maestro
        e inserta las que falten en la BD. Útil para periodos históricos
        cuya información ya estaba en el Excel pero no fue importada por PDF.
        """
        if not self._validar_seleccion():
            return
        if not MOD.get("importar_excel_maestro"):
            self.log("❌ Módulo importar_excel_maestro no disponible", "error")
            return

        self._estado("Importando desde Excel Maestro…", procesando=True)
        self.log("━━━ IMPORTAR DESDE EXCEL MAESTRO ━━━", "titulo")

        # Buscar el Excel de esta comunidad
        codigo = self.comunidad_actual.get().split(" — ")[0]
        nombre_periodo = self.periodo_actual.get()
        posibles = list(RUTA_EXCELS.glob(f"*{codigo}*.xlsx"))
        posibles = [p for p in posibles if "bak" not in p.name.lower()
                    and "~" not in p.name
                    and "PLANTILLA" not in p.name.upper()]

        if not posibles:
            self.log(f"❌ La comunidad {codigo} no tiene Excel Maestro todavía", "error")
            self.log("   👉 Este botón importa datos DESDE un Excel ya existente.", "aviso")
            self.log("      Si lo que quieres es crearlo, pulsa «Regenerar Excel»", "aviso")
            self.log("      (lo genera desde la base de datos) y no hace falta importar nada.", "aviso")
            return

        ruta_excel = str(posibles[0])
        self.log(f"📊 Leyendo: {Path(ruta_excel).name}", "info")
        self.log(f"📅 Periodo: {nombre_periodo}", "info")

        try:
            resultado = MOD["importar_excel_maestro"].importar_excel_a_bd(
                ruta_excel    = ruta_excel,
                ruta_bd       = str(RUTA_BD),
                id_comunidad  = self.id_comunidad,
                id_periodo    = self.id_periodo,
                nombre_año    = nombre_periodo,
                verbose       = False,
            )
        except Exception as e:
            self.log(f"❌ Error inesperado: {e}", "error")
            return

        if resultado.get("ok") or resultado.get("facturas_importadas", 0) > 0:
            fi = resultado.get("facturas_importadas", 0)
            fd = resultado.get("duplicadas", 0)
            self.log(f"  ✅ {fi} factura(s) importadas, {fd} ya existían", "ok")
            for hoja, det in resultado.get("detalle_por_hoja", {}).items():
                n = det.get("importadas", 0)
                d = det.get("duplicadas", 0)
                e = det.get("errores", 0)
                if n or d or e:
                    self.log(f"     {hoja:<20} {n:>3} nuevas, {d:>3} dup, {e:>3} err",
                             "ok" if e == 0 else "aviso")
            for err in resultado.get("errores", [])[:5]:
                self.log(f"  ⚠️  {err}", "aviso")
            if fi > 0:
                self.log(
                    "\n  ℹ️  Facturas cargadas. Ejecuta ahora:\n"
                    "      1. Calcular Reparto  →  para calcular cuotas\n"
                    "      2. Generar Cartas    →  para crear los documentos",
                    "info"
                )
                self.after(0, self._actualizar_banner)
        else:
            err = resultado.get("error", "Sin detalles")
            self.log(f"  ❌ {err}", "error")
            for e in resultado.get("errores", [])[:5]:
                self.log(f"     {e}", "aviso")

    def _todo_en_uno_impl(self):
        """
        Ejecuta el flujo completo con el pipeline headless (el mismo código
        que el modo desatendido por línea de comandos):
          ingesta entrada/ → BD → reparto → Excel regenerado → cartas.
        """
        try:
            import pipeline
        except ImportError as e:
            self.log(f"❌ No se pudo cargar pipeline.py: {e}", "error")
            return

        self._estado("Procesando todo…", procesando=True)
        codigo = None
        try:
            codigo = self.comunidad_actual.get().split(" — ")[0].strip() or None
        except Exception:
            pass

        resultado = pipeline.procesar_todo(
            codigo_comunidad=codigo,
            hacer_reparto=True,
            hacer_cartas=True,
            log_callback=self.log,
        )

        if resultado.get("ok"):
            self.log("\n✅ Proceso completo finalizado", "ok")
        else:
            self.log("\n⚠️ Proceso finalizado con incidencias — revisa el log", "aviso")

    # -----------------------------------------------------------------------
    # GESTIÓN DE EXCELS POR COMUNIDAD
    # -----------------------------------------------------------------------
    def _crear_excel_si_no_existe(self, codigo: str, nombre: str = "") -> "Path | None":
        """
        Asegura que Comunidad_{codigo}.xlsx existe en Excels_Maestros/.
        Orden de búsqueda/creación:
          1. Si ya existe → devuelve su ruta
          2. Si existe Comunidad_PLANTILLA.xlsx → lo copia
          3. Si no → usa excel_writer.crear_plantilla_excel() para generarlo limpio
        """
        nombre_archivo = f"Comunidad_{codigo}.xlsx"
        ruta_destino   = RUTA_EXCELS / nombre_archivo

        if ruta_destino.exists():
            return ruta_destino

        # Buscar plantilla explícita (blank, sin datos)
        candidatos_plantilla = [
            RUTA_EXCELS / "Comunidad_PLANTILLA.xlsx",
            BASE_DIR / "plantillas" / "Comunidad_PLANTILLA.xlsx",
        ]
        plantilla = next((p for p in candidatos_plantilla if p.exists()), None)
        if plantilla:
            shutil.copy2(str(plantilla), str(ruta_destino))
            self.log(f"  📊 Excel {nombre_archivo} creado desde Comunidad_PLANTILLA.xlsx", "ok")
            return ruta_destino

        # Generar Excel desde cero con la estructura correcta
        if MOD.get("excel_writer") and hasattr(MOD["excel_writer"], "crear_plantilla_excel"):
            try:
                MOD["excel_writer"].crear_plantilla_excel(
                    str(ruta_destino), codigo, nombre
                )
                self.log(f"  📊 Excel {nombre_archivo} generado (plantilla nueva)", "ok")
                return ruta_destino
            except Exception as e:
                self.log(f"  ❌ No se pudo generar {nombre_archivo}: {e}", "error")
                return None

        self.log(f"  ❌ No se pudo crear {nombre_archivo}: módulo excel_writer no disponible", "error")
        return None

    # -----------------------------------------------------------------------
    # UTILIDADES
    # -----------------------------------------------------------------------
    def _deshabilitar_botones(self):
        for btn in self.botones.values():
            btn.configure(state="disabled")
        for selector in (getattr(self, "cb_comunidad", None),
                         getattr(self, "cb_periodo", None),
                         getattr(self, "cb_expediente", None)):
            if selector is not None:
                selector.configure(state="disabled")

    def _habilitar_botones(self):
        for btn in self.botones.values():
            btn.configure(state="normal")
        for selector in (getattr(self, "cb_comunidad", None),
                         getattr(self, "cb_periodo", None),
                         getattr(self, "cb_expediente", None)):
            if selector is not None:
                selector.configure(state="normal")

    def _abrir_entrada(self):
        os.startfile(str(RUTA_ENTRADA)) if sys.platform == "win32" else \
            os.system(f"xdg-open '{RUTA_ENTRADA}'")

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

    def _dialogo_seleccionar_periodo_inicio(self):
        """
        Diálogo que aparece al arrancar el programa.
        Permite elegir rango de fechas exacto (día/mes/año) para el periodo de trabajo.
        Si ya hay un periodo seleccionado para la comunidad activa, lo precarga.
        """
        if not MOD.get("gestor_bd") or not RUTA_BD.exists():
            return  # Sin BD no podemos hacer nada útil

        dialogo = self._preparar_dialogo("Seleccionar Periodo de Trabajo", 580, 520)

        MESES = ["Ene","Feb","Mar","Abr","May","Jun",
                 "Jul","Ago","Sep","Oct","Nov","Dic"]

        # ── Cabecera ─────────────────────────────────────────────────────────
        ctk.CTkLabel(dialogo, text="SELECCIONAR PERIODO DE TRABAJO",
                     font=UIM.fuente(13, "bold"), fg_color=C["primario"],
                     text_color="#FFFFFF", corner_radius=0, height=42
                     ).pack(fill="x")

        ctk.CTkLabel(dialogo,
                     text="Define el rango exacto del periodo que quieres gestionar.\n"
                          "Las facturas y lecturas se filtrarán por estos días exactos.\n"
                          "El consumo se calculará proporcionalmente al día.",
                     font=UIM.fuente(11), text_color=C["texto_sec"],
                     justify="center").pack(pady=(12, 4))

        # ── Comunidad ────────────────────────────────────────────────────────
        frm_com = ctk.CTkFrame(dialogo, fg_color="transparent")
        frm_com.pack(fill="x", padx=24, pady=4)
        ctk.CTkLabel(frm_com, text="Comunidad:", font=UIM.fuente(12),
                     text_color=C["texto"], width=140,
                     anchor="w").pack(side="left")

        var_comunidad = tk.StringVar(value=self.comunidad_actual.get())
        opciones_com  = list(getattr(self, "_ids_comunidad", {}).keys()) or [""]
        cb_com = UIM.ComboModerno(frm_com, variable=var_comunidad,
                                  values=opciones_com, width=340, height=34)
        cb_com.pack(side="left")

        # ── Periodos existentes ───────────────────────────────────────────────
        frm_per = ctk.CTkFrame(dialogo, fg_color="transparent")
        frm_per.pack(fill="x", padx=24, pady=4)
        ctk.CTkLabel(frm_per, text="Periodo existente:", font=UIM.fuente(12),
                     text_color=C["texto"], width=140,
                     anchor="w").pack(side="left")
        var_per_existente = tk.StringVar(value="— Nuevo periodo —")
        cb_per_existente  = UIM.ComboModerno(frm_per, variable=var_per_existente,
                                             values=["— Nuevo periodo —"],
                                             width=260, height=34)
        cb_per_existente.pack(side="left", padx=(0, 8))
        ctk.CTkLabel(frm_per, text="(o define uno nuevo abajo)",
                     font=UIM.fuente(10),
                     text_color=C["texto_sec"]).pack(side="left")

        # Actualizar lista de periodos al cambiar comunidad
        _periodos_cache: dict[str, dict] = {}

        def _actualizar_periodos(*_):
            sel_com = var_comunidad.get()
            id_com  = getattr(self, "_ids_comunidad", {}).get(sel_com)
            if not id_com:
                return
            try:
                con = MOD["gestor_bd"].conectar(str(RUTA_BD))
                rows = con.execute(
                    "SELECT id_periodo, nombre, fecha_inicio, fecha_fin FROM periodos "
                    "WHERE id_comunidad=? ORDER BY fecha_inicio DESC",
                    (id_com,)
                ).fetchall()
                con.close()
                _periodos_cache.clear()
                opts = ["— Nuevo periodo —"]
                for r in rows:
                    fi = r["fecha_inicio"] or ""
                    ff = r["fecha_fin"]    or "abierto"
                    etiq = f"{r['nombre']}  ({fi} → {ff})"
                    opts.append(etiq)
                    _periodos_cache[etiq] = {
                        "id":    r["id_periodo"],
                        "nombre": r["nombre"],
                        "inicio": r["fecha_inicio"] or "",
                        "fin":    r["fecha_fin"]    or "",
                    }
                cb_per_existente.configure(values=opts)
                var_per_existente.set(opts[1] if len(opts) > 1 else opts[0])
                _precargar_desde_existente()
            except Exception:
                pass

        var_comunidad.trace_add("write", _actualizar_periodos)

        # ── Separador ────────────────────────────────────────────────────────
        ctk.CTkFrame(dialogo, fg_color=C["borde"], height=1).pack(
            fill="x", padx=16, pady=8)

        def _crear_selector_fecha(parent, label_text: str,
                                   dia_def: int, mes_def: int, anio_def: int):
            frm = ctk.CTkFrame(parent, fg_color="transparent")
            frm.pack(fill="x", padx=24, pady=5)
            ctk.CTkLabel(frm, text=label_text, font=UIM.fuente(11, "bold"),
                         text_color=C["primario"], width=190,
                         anchor="w").pack(side="left")
            var_dia  = tk.StringVar(value=str(dia_def).zfill(2))
            var_mes  = tk.StringVar(value=MESES[mes_def - 1])
            var_anio = tk.StringVar(value=str(anio_def))
            frm_sel  = ctk.CTkFrame(frm, fg_color="transparent")
            frm_sel.pack(side="left")
            estilo = dict(state="readonly", height=30, font=UIM.fuente(12),
                          dropdown_font=UIM.fuente(12),
                          border_color=C["borde"], button_color=C["primario"],
                          button_hover_color=C["primario_hover"])
            anio_act = datetime.now().year
            ctk.CTkComboBox(frm_sel, variable=var_dia,
                            values=[str(d).zfill(2) for d in range(1, 32)],
                            width=66, **estilo).pack(side="left", padx=(0, 4))
            ctk.CTkComboBox(frm_sel, variable=var_mes,
                            values=MESES, width=74,
                            **estilo).pack(side="left", padx=(0, 4))
            ctk.CTkComboBox(frm_sel, variable=var_anio,
                            values=[str(a) for a in range(2015, anio_act + 3)],
                            width=82, **estilo).pack(side="left")
            return var_dia, var_mes, var_anio

        def _vars_a_fecha(vd, vm, va) -> str:
            mes_num = MESES.index(vm.get()) + 1
            return f"{va.get()}-{mes_num:02d}-{vd.get()}"

        now       = datetime.now()
        anio_ini  = now.year - 1 if now.month < 9 else now.year

        ctk.CTkLabel(dialogo, text="Fecha de INICIO del periodo:",
                     font=UIM.fuente(11), text_color=C["texto_sec"],
                     anchor="w").pack(anchor="w", padx=24)
        vd_ini, vm_ini, va_ini = _crear_selector_fecha(
            dialogo, "  Día  /  Mes  /  Año:", 1, 9, anio_ini)

        ctk.CTkFrame(dialogo, fg_color=C["borde"], height=1).pack(
            fill="x", padx=16, pady=4)
        ctk.CTkLabel(dialogo, text="Fecha de FIN del periodo:",
                     font=UIM.fuente(11), text_color=C["texto_sec"],
                     anchor="w").pack(anchor="w", padx=24)
        vd_fin, vm_fin, va_fin = _crear_selector_fecha(
            dialogo, "  Día  /  Mes  /  Año:", 31, 8, anio_ini + 1)

        # ── Nombre y resumen ─────────────────────────────────────────────────
        frm_nom = ctk.CTkFrame(dialogo, fg_color="transparent")
        frm_nom.pack(fill="x", padx=24, pady=(8, 0))
        ctk.CTkLabel(frm_nom, text="Nombre:", font=UIM.fuente(12),
                     text_color=C["texto"], width=140,
                     anchor="w").pack(side="left")
        var_nombre = tk.StringVar(value=f"{anio_ini}-{anio_ini+1}")
        ctk.CTkEntry(frm_nom, textvariable=var_nombre,
                     font=UIM.fuente(12), width=150, height=32,
                     border_color=C["borde"]).pack(side="left", padx=(0, 12))

        lbl_dur = ctk.CTkLabel(frm_nom, text="", font=UIM.fuente(11),
                               text_color=C["primario"])
        lbl_dur.pack(side="left")

        def _actualizar_dur(*_):
            try:
                from datetime import date as _d
                ini = _d.fromisoformat(_vars_a_fecha(vd_ini, vm_ini, va_ini))
                fin = _d.fromisoformat(_vars_a_fecha(vd_fin, vm_fin, va_fin))
                dias = (fin - ini).days
                meses = round(dias / 30.44, 1)
                lbl_dur.configure(
                    text=f"→ {dias} días ({meses} meses)",
                    text_color=C["exito"] if dias > 0 else C["alerta"]
                )
            except Exception:
                lbl_dur.configure(text="")

        for v in (vd_ini, vm_ini, va_ini, vd_fin, vm_fin, va_fin):
            v.trace_add("write", _actualizar_dur)
        _actualizar_dur()

        def _precargar_desde_existente(*_):
            """Rellena los selectores con las fechas del periodo seleccionado."""
            etiq = var_per_existente.get()
            info = _periodos_cache.get(etiq)
            if not info:
                return
            var_nombre.set(info["nombre"])
            for fecha_str, vd, vm, va in [
                (info["inicio"], vd_ini, vm_ini, va_ini),
                (info["fin"],    vd_fin, vm_fin, va_fin),
            ]:
                if not fecha_str or len(fecha_str) < 10:
                    continue
                try:
                    partes = fecha_str[:10].split("-")
                    va.set(partes[0])
                    vm.set(MESES[int(partes[1]) - 1])
                    vd.set(partes[2].zfill(2))
                except Exception:
                    pass

        var_per_existente.trace_add("write", _precargar_desde_existente)

        # Cargar periodos de la comunidad actual al abrir
        _actualizar_periodos()

        # ── Botones ───────────────────────────────────────────────────────────
        frm_bots = ctk.CTkFrame(dialogo, fg_color="transparent")
        frm_bots.pack(pady=14)

        def _aplicar():
            sel_com = var_comunidad.get()
            id_com  = getattr(self, "_ids_comunidad", {}).get(sel_com)
            if not id_com:
                messagebox.showwarning("Sin comunidad", "Selecciona una comunidad.")
                return
            nombre = var_nombre.get().strip()
            if not nombre:
                messagebox.showwarning("Sin nombre", "Escribe un nombre para el periodo.")
                return
            try:
                from datetime import date as _d
                f_ini = _vars_a_fecha(vd_ini, vm_ini, va_ini)
                f_fin = _vars_a_fecha(vd_fin, vm_fin, va_fin)
                if _d.fromisoformat(f_fin) <= _d.fromisoformat(f_ini):
                    messagebox.showerror("Fechas inválidas",
                                         "La fecha de fin debe ser posterior al inicio.")
                    return

                con = MOD["gestor_bd"].conectar(str(RUTA_BD))
                # Cambiar comunidad si hace falta
                if sel_com != self.comunidad_actual.get():
                    self.comunidad_actual.set(sel_com)
                    self.id_comunidad = id_com
                    self._on_comunidad_seleccionada()

                id_per = MOD["gestor_bd"].obtener_o_crear_periodo(
                    con, id_com, nombre, f_ini)
                con.execute(
                    "UPDATE periodos SET fecha_fin=?, fecha_inicio=?, estado='abierto' "
                    "WHERE id_periodo=?",
                    (f_fin, f_ini, id_per)
                )
                con.commit()
                con.close()

                dias = (_d.fromisoformat(f_fin) - _d.fromisoformat(f_ini)).days
                self.log(
                    f"📅 Periodo activo: {nombre}  |  {f_ini} → {f_fin}  ({dias} días)",
                    "titulo"
                )
                self._on_comunidad_seleccionada()
                self.cb_periodo.set(nombre)
                self._on_periodo_seleccionado()
                dialogo.destroy()
            except Exception as e:
                messagebox.showerror("Error", str(e))

        ctk.CTkButton(frm_bots, text="✅  Trabajar con este periodo",
                      command=_aplicar,
                      height=38, corner_radius=10,
                      font=UIM.fuente(13, "bold"),
                      fg_color=C["primario"],
                      hover_color=C["primario_hover"]).pack(side="left", padx=8)

        ctk.CTkButton(frm_bots, text="Omitir",
                      command=dialogo.destroy,
                      width=80, height=38, corner_radius=10,
                      font=UIM.fuente(12),
                      fg_color="transparent",
                      hover_color=C["acento_suave"],
                      text_color=C["texto_sec"]).pack(side="left")

    def _nuevo_periodo(self):
        """
        Diálogo para crear o seleccionar un periodo con fechas exactas (día/mes/año).
        Soporta periodos de cualquier duración — 8 meses, 10 meses, año completo, etc.
        """
        dialogo = self._preparar_dialogo("Definir Periodo de Regularización", 520, 440)

        # ── Cabecera ─────────────────────────────────────────────────────────
        ctk.CTkLabel(dialogo, text="NUEVO PERIODO",
                     font=UIM.fuente(13, "bold"), fg_color=C["primario"],
                     text_color="#FFFFFF", height=38).pack(fill="x")

        ctk.CTkLabel(dialogo,
                     text="Define el rango exacto. Las facturas y lecturas se filtrarán\n"
                          "por estos días, calculando consumo proporcional al día.",
                     font=UIM.fuente(10), text_color=C["texto_sec"],
                     justify="center").pack(pady=(12, 6))

        # ── Nombre ───────────────────────────────────────────────────────────
        frm_nombre = ctk.CTkFrame(dialogo, fg_color="transparent")
        frm_nombre.pack(fill="x", padx=24, pady=4)
        ctk.CTkLabel(frm_nombre, text="Nombre del periodo:", font=UIM.fuente(12),
                     text_color=C["texto"], width=170,
                     anchor="w").pack(side="left")
        ent_nombre = ctk.CTkEntry(frm_nombre, font=UIM.fuente(12),
                                  width=150, height=32,
                                  border_color=C["borde"])
        ent_nombre.insert(0, "2024-2025")
        ent_nombre.pack(side="left")

        # ── Función auxiliar para crear un selector de fecha ─────────────────
        MESES = ["Ene","Feb","Mar","Abr","May","Jun",
                 "Jul","Ago","Sep","Oct","Nov","Dic"]

        def _crear_selector_fecha(parent, label_text: str, dia_def: int,
                                   mes_def: int, anio_def: int):
            """Crea una fila con selector de día, mes y año. Devuelve variables."""
            frm = ctk.CTkFrame(parent, fg_color="transparent")
            frm.pack(fill="x", padx=24, pady=6)

            ctk.CTkLabel(frm, text=label_text, font=UIM.fuente(12),
                         text_color=C["texto"], width=170,
                         anchor="w").pack(side="left")

            var_dia  = tk.StringVar(value=str(dia_def).zfill(2))
            var_mes  = tk.StringVar(value=MESES[mes_def - 1])
            var_anio = tk.StringVar(value=str(anio_def))

            frm_sel = ctk.CTkFrame(frm, fg_color="transparent")
            frm_sel.pack(side="left")

            estilo = dict(state="readonly", height=30, font=UIM.fuente(12),
                          dropdown_font=UIM.fuente(12),
                          border_color=C["borde"], button_color=C["primario"],
                          button_hover_color=C["primario_hover"])

            # Día
            dias = [str(d).zfill(2) for d in range(1, 32)]
            ctk.CTkComboBox(frm_sel, variable=var_dia, values=dias,
                            width=66, **estilo).pack(side="left", padx=(0, 4))

            # Mes
            ctk.CTkComboBox(frm_sel, variable=var_mes, values=MESES,
                            width=74, **estilo).pack(side="left", padx=(0, 4))

            # Año
            anio_actual = datetime.now().year
            anios = [str(a) for a in range(2015, anio_actual + 3)]
            ctk.CTkComboBox(frm_sel, variable=var_anio, values=anios,
                            width=82, **estilo).pack(side="left")

            return var_dia, var_mes, var_anio

        def _vars_a_fecha(var_dia, var_mes, var_anio) -> str:
            mes_num = MESES.index(var_mes.get()) + 1
            return f"{var_anio.get()}-{mes_num:02d}-{var_dia.get()}"

        now = datetime.now()
        anio_ini = now.year - 1 if now.month < 9 else now.year

        # ── Separador ────────────────────────────────────────────────────────
        ctk.CTkFrame(dialogo, fg_color=C["borde"], height=1).pack(
            fill="x", padx=16, pady=4)
        ctk.CTkLabel(dialogo, text="Fecha de INICIO (lectura inicial del periodo):",
                     font=UIM.fuente(11, "bold"), text_color=C["primario"],
                     anchor="w").pack(anchor="w", padx=24)

        vd_ini, vm_ini, va_ini = _crear_selector_fecha(
            dialogo, "  Día / Mes / Año:", 1, 9, anio_ini)

        ctk.CTkFrame(dialogo, fg_color=C["borde"], height=1).pack(
            fill="x", padx=16, pady=4)
        ctk.CTkLabel(dialogo, text="Fecha de FIN (lectura final del periodo):",
                     font=UIM.fuente(11, "bold"), text_color=C["primario"],
                     anchor="w").pack(anchor="w", padx=24)

        vd_fin, vm_fin, va_fin = _crear_selector_fecha(
            dialogo, "  Día / Mes / Año:", 31, 8, anio_ini + 1)

        # ── Etiqueta de resumen ───────────────────────────────────────────────
        lbl_resumen = ctk.CTkLabel(dialogo, text="", font=UIM.fuente(11),
                                   text_color=C["primario"])
        lbl_resumen.pack(pady=(4, 0))

        def _actualizar_resumen(*_):
            try:
                f_ini = _vars_a_fecha(vd_ini, vm_ini, va_ini)
                f_fin = _vars_a_fecha(vd_fin, vm_fin, va_fin)
                from datetime import date as _d
                ini = _d.fromisoformat(f_ini)
                fin = _d.fromisoformat(f_fin)
                dias = (fin - ini).days
                meses = round(dias / 30.44, 1)
                lbl_resumen.configure(
                    text=f"Duración: {dias} días  ≈  {meses} meses",
                    text_color=C["exito"] if dias > 0 else C["alerta"]
                )
            except Exception:
                lbl_resumen.configure(text="")

        for v in (vd_ini, vm_ini, va_ini, vd_fin, vm_fin, va_fin):
            v.trace_add("write", _actualizar_resumen)
        _actualizar_resumen()

        # ── Botón Crear ───────────────────────────────────────────────────────
        def _crear():
            nombre = ent_nombre.get().strip()
            if not nombre:
                messagebox.showwarning("Faltan datos", "Escribe un nombre para el periodo.")
                return
            if not self.id_comunidad:
                messagebox.showwarning("Sin comunidad", "Selecciona una comunidad primero.")
                return
            try:
                f_ini = _vars_a_fecha(vd_ini, vm_ini, va_ini)
                f_fin = _vars_a_fecha(vd_fin, vm_fin, va_fin)
                from datetime import date as _d
                if _d.fromisoformat(f_fin) <= _d.fromisoformat(f_ini):
                    messagebox.showerror("Fechas inválidas",
                                         "La fecha de fin debe ser posterior al inicio.")
                    return
                con = MOD["gestor_bd"].conectar(str(RUTA_BD))
                id_per = MOD["gestor_bd"].obtener_o_crear_periodo(
                    con, self.id_comunidad, nombre, f_ini
                )
                con.execute(
                    "UPDATE periodos SET fecha_fin=?, estado='abierto' WHERE id_periodo=?",
                    (f_fin, id_per)
                )
                con.commit()
                con.close()
                dias = (_d.fromisoformat(f_fin) - _d.fromisoformat(f_ini)).days
                self.log(
                    f"✅ Periodo '{nombre}' creado: {f_ini} → {f_fin} ({dias} días)",
                    "ok"
                )
                self._on_comunidad_seleccionada()
                self.cb_periodo.set(nombre)
                self._on_periodo_seleccionado()
                dialogo.destroy()
            except Exception as e:
                messagebox.showerror("Error", str(e))

        ctk.CTkButton(dialogo, text="✅  Crear Periodo",
                      command=_crear,
                      height=38, corner_radius=10,
                      font=UIM.fuente(13, "bold"),
                      fg_color=C["primario"],
                      hover_color=C["primario_hover"]).pack(pady=14)

    def _nueva_comunidad(self):
        """Diálogo para registrar una nueva comunidad."""
        dialogo = self._preparar_dialogo("Nueva Comunidad", 470, 320)

        campos = [
            ("Código (ej: 644):",   "codigo",  "644"),
            ("Nombre completo:",     "nombre",  "CDAD. PROP. …"),
            ("CIF de la comunidad:", "cif",     "H99258139"),
            ("Nº de viviendas:",     "viviendas","120"),
        ]
        entradas = {}
        for i, (label, clave, defecto) in enumerate(campos):
            ctk.CTkLabel(dialogo, text=label, font=UIM.fuente(12),
                         text_color=C["texto"]).grid(
                             row=i, column=0, padx=(24, 8), pady=10, sticky="w")
            e = ctk.CTkEntry(dialogo, font=UIM.fuente(12), width=230, height=32,
                             border_color=C["borde"])
            e.insert(0, defecto)
            e.grid(row=i, column=1, padx=(0, 24), pady=10)
            entradas[clave] = e

        def _crear():
            if not MOD.get("gestor_bd"):
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
                self._crear_excel_si_no_existe(cod, nom)
                self._cargar_comunidades()
                dialogo.destroy()
            except Exception as e:
                messagebox.showerror("Error", str(e))

        ctk.CTkButton(dialogo, text="Registrar Comunidad",
                      command=_crear,
                      height=38, corner_radius=10,
                      font=UIM.fuente(13, "bold"),
                      fg_color=C["primario"],
                      hover_color=C["primario_hover"]).grid(
                          row=len(campos), column=0, columnspan=2, pady=18)

    def _configurar_rutas(self):
        """Muestra las rutas actuales y permite cambiarlas."""
        ventana = self._preparar_dialogo("Configuración de Rutas", 660, 350)

        rutas = [
            ("Base de datos:",      str(RUTA_BD)),
            ("Carpeta entrada/:",   str(RUTA_ENTRADA)),
            ("Carpeta procesados/:",str(RUTA_PROCESADOS)),
            ("Excels Maestros/:",   str(RUTA_EXCELS)),
            ("Plantilla cartas:",   str(RUTA_PLANTILLA)),
            ("Cartas generadas/:",  str(RUTA_CARTAS)),
        ]
        for i, (label, ruta) in enumerate(rutas):
            ctk.CTkLabel(ventana, text=label, font=UIM.fuente(11),
                         text_color=C["texto"], anchor="w", width=170).grid(
                row=i, column=0, padx=(20, 4), pady=6, sticky="w")
            ctk.CTkLabel(ventana, text=ruta, font=UIM.fuente_mono(11),
                         text_color=C["primario"], anchor="w").grid(
                row=i, column=1, padx=4, pady=6, sticky="w")

        ctk.CTkLabel(ventana,
                     text="Para cambiar las rutas, edita las constantes al inicio de app.py",
                     font=UIM.fuente(11),
                     text_color=C["texto_sec"]).grid(
            row=len(rutas), column=0, columnspan=2, pady=14, padx=20)


# ---------------------------------------------------------------------------
# PUNTO DE ENTRADA
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    UIM.iniciar()
    app = AppGestionFincas()
    app.mainloop()
