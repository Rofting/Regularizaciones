"""
app.py
======
Interfaz gráfica principal del sistema de Gestión de Fincas.
Diseñada para usuarios no técnicos — un botón por acción, mensajes claros.

REQUISITOS:
    pip install openpyxl python-docx pdfplumber

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
from tkinter import ttk, filedialog, messagebox, scrolledtext

# ---------------------------------------------------------------------------
# RUTAS POR DEFECTO
# ---------------------------------------------------------------------------
# Calculadas relativas al directorio donde está app.py
BASE_DIR = Path(__file__).parent.parent  # sube un nivel desde core/

RUTA_BD         = BASE_DIR / "data" / "gestion.db"
RUTA_ENTRADA    = BASE_DIR / "entrada"
RUTA_PROCESADOS = BASE_DIR / "procesados"
RUTA_EXCELS     = BASE_DIR / "Excels_Maestros"
RUTA_PLANTILLA  = BASE_DIR / "plantillas" / "Plantilla_Cartas.docx"
RUTA_CARTAS     = BASE_DIR / "salidas" / "cartas"
RUTA_PROVEEDORES= BASE_DIR / "config" / "proveedores.json"

# Tema moderno (paleta definida en ui_moderna.py)
try:
    import ui_moderna as UIM
except ImportError:
    UIM = None

if UIM:
    COLOR_FONDO      = UIM.TEMA["fondo"]
    COLOR_PANEL      = UIM.TEMA["panel"]
    COLOR_PRIMARIO   = UIM.TEMA["primario"]
    COLOR_ACENTO     = UIM.TEMA["acento"]
    COLOR_EXITO      = UIM.TEMA["exito"]
    COLOR_ALERTA     = UIM.TEMA["alerta"]
    COLOR_NEUTRO     = UIM.TEMA["texto_sec"]
    COLOR_BORDE      = UIM.TEMA["borde"]
else:  # respaldo por si falta ui_moderna.py
    COLOR_FONDO      = "#F5F5F0"
    COLOR_PANEL      = "#FFFFFF"
    COLOR_PRIMARIO   = "#1A3A5C"
    COLOR_ACENTO     = "#2E86C1"
    COLOR_EXITO      = "#1E8449"
    COLOR_ALERTA     = "#C0392B"
    COLOR_NEUTRO     = "#7F8C8D"
    COLOR_BORDE      = "#DCE1E7"

COLOR_TEXTO      = "#111827"
FUENTE_TITULO    = ("Segoe UI", 14, "bold")
FUENTE_SECCION   = ("Segoe UI", 10, "bold")
FUENTE_NORMAL    = ("Segoe UI", 10)
FUENTE_PEQUEÑA   = ("Segoe UI", 9)
FUENTE_MONO      = ("Consolas", 9)


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
                   "importar_excel_maestro"]:
        try:
            modulos[nombre] = __import__(nombre)
        except ImportError as e:
            modulos[nombre] = None
    return modulos


MOD = _importar_modulos()


# ---------------------------------------------------------------------------
# VENTANA PRINCIPAL
# ---------------------------------------------------------------------------
class AppGestionFincas(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Gestión de Fincas — Meditrade")
        self.geometry("900x680")
        self.minsize(800, 600)
        self.configure(bg=COLOR_FONDO)
        self.resizable(True, True)

        # Estado
        self.comunidad_actual  = tk.StringVar(value="")
        self.periodo_actual     = tk.StringVar(value="")
        self.id_comunidad       = None
        self.id_periodo         = None
        self._procesando        = False

        self._crear_ui()
        self._cargar_comunidades()
        self._verificar_estructura()
        # Mostrar selector de periodo al arrancar (tras un breve retraso
        # para que la ventana principal ya esté renderizada)
        self.after(300, self._dialogo_seleccionar_periodo_inicio)

    # -----------------------------------------------------------------------
    # CONSTRUCCIÓN DE LA UI
    # -----------------------------------------------------------------------
    def _crear_ui(self):
        if UIM:
            UIM.aplicar_estilo_ttk(self)

        # ── CABECERA ────────────────────────────────────────────────────────
        cabecera = tk.Frame(self, bg=COLOR_PANEL, height=62)
        cabecera.pack(fill="x")
        cabecera.pack_propagate(False)

        tk.Label(cabecera, text="Gestión de Fincas",
                 font=("Segoe UI", 16, "bold"),
                 bg=COLOR_PANEL, fg=COLOR_TEXTO).pack(side="left", padx=(24, 8), pady=14)
        tk.Label(cabecera, text="●", font=("Segoe UI", 8),
                 bg=COLOR_PANEL, fg=COLOR_ACENTO).pack(side="left", pady=14)
        tk.Label(cabecera, text="Meditrade",
                 font=FUENTE_PEQUEÑA,
                 bg=COLOR_PANEL, fg=COLOR_NEUTRO).pack(side="left", padx=8, pady=14)

        # Línea de acento bajo la cabecera (degradado indigo → transparente)
        linea = tk.Canvas(self, height=2, bg=COLOR_FONDO, highlightthickness=0)
        linea.pack(fill="x")
        def _pintar_linea(event=None):
            linea.delete("all")
            w = max(linea.winfo_width(), 2)
            tramos = 60
            for i in range(tramos):
                c = UIM.mezclar(COLOR_PRIMARIO, COLOR_FONDO, i / tramos) if UIM else COLOR_PRIMARIO
                linea.create_rectangle(w * i / tramos, 0, w * (i + 1) / tramos, 2,
                                       fill=c, outline=c)
        linea.bind("<Configure>", _pintar_linea)

        # ── SELECTOR DE COMUNIDAD Y PERIODO ────────────────────────────────
        selector = tk.Frame(self, bg=COLOR_PANEL,
                            highlightbackground=COLOR_BORDE, highlightthickness=1)
        selector.pack(fill="x", padx=16, pady=(14, 0))

        estilo_cb = {"style": "Moderno.TCombobox"} if UIM else {}

        tk.Label(selector, text="COMUNIDAD", font=("Segoe UI", 8, "bold"),
                 bg=COLOR_PANEL, fg=COLOR_NEUTRO).grid(
                     row=0, column=0, padx=(18, 8), pady=(10, 0), sticky="w")
        self.cb_comunidad = ttk.Combobox(selector, textvariable=self.comunidad_actual,
                                          state="readonly", width=42,
                                          font=FUENTE_NORMAL, **estilo_cb)
        self.cb_comunidad.grid(row=1, column=0, padx=(16, 16), pady=(2, 12), sticky="w")
        self.cb_comunidad.bind("<<ComboboxSelected>>", self._on_comunidad_seleccionada)

        tk.Label(selector, text="PERIODO", font=("Segoe UI", 8, "bold"),
                 bg=COLOR_PANEL, fg=COLOR_NEUTRO).grid(
                     row=0, column=1, padx=(8, 8), pady=(10, 0), sticky="w")
        self.cb_periodo = ttk.Combobox(selector, textvariable=self.periodo_actual,
                                        state="readonly", width=16,
                                        font=FUENTE_NORMAL, **estilo_cb)
        self.cb_periodo.grid(row=1, column=1, padx=(6, 8), pady=(2, 12), sticky="w")
        self.cb_periodo.bind("<<ComboboxSelected>>", self._on_periodo_seleccionado)

        if UIM:
            btn_np = UIM.BotonSuave(selector, texto="＋ Nuevo periodo",
                                    comando=self._nuevo_periodo, alto=32, width=130)
            btn_np.grid(row=1, column=2, padx=8, pady=(2, 12), sticky="w")
        else:
            tk.Button(selector, text="+ Nuevo periodo",
                      font=FUENTE_PEQUEÑA, bg=COLOR_PANEL, fg=COLOR_ACENTO,
                      relief="flat", cursor="hand2",
                      command=self._nuevo_periodo).grid(row=1, column=2, padx=8)

        # ── BANNER DE PERIODO ACTIVO ────────────────────────────────────────
        self.banner_periodo = tk.Frame(self, bg="#EEF2FF",
                                        highlightbackground="#C7D2FE",
                                        highlightthickness=1)
        self.banner_periodo.pack(fill="x", padx=16, pady=(8, 0))
        self.lbl_banner = tk.Label(
            self.banner_periodo,
            text="Selecciona una comunidad y un periodo para empezar.",
            font=("Segoe UI", 9), bg="#EEF2FF", fg=COLOR_PRIMARIO,
            anchor="w", padx=14, pady=6
        )
        self.lbl_banner.pack(fill="x")

        # ── CUERPO PRINCIPAL ────────────────────────────────────────────────
        cuerpo = tk.Frame(self, bg=COLOR_FONDO)
        cuerpo.pack(fill="both", expand=True, padx=16, pady=12)
        cuerpo.columnconfigure(0, weight=1)
        cuerpo.columnconfigure(1, weight=2)
        cuerpo.rowconfigure(0, weight=1)

        # ── PANEL IZQUIERDO: ACCIONES ───────────────────────────────────────
        panel_acc = tk.Frame(cuerpo, bg=COLOR_PANEL,
                             relief="flat",
                             highlightbackground=COLOR_BORDE,
                             highlightthickness=1)
        panel_acc.grid(row=0, column=0, sticky="nsew", padx=(0, 8))

        self.botones = {}

        if UIM:
            # ── Acción principal ──
            tk.Label(panel_acc, text="ACCIÓN PRINCIPAL",
                     font=("Segoe UI", 8, "bold"),
                     bg=COLOR_PANEL, fg=COLOR_NEUTRO).pack(
                         anchor="w", padx=16, pady=(16, 6))

            btn_todo = UIM.BotonModerno(
                panel_acc, texto="▶   PROCESAR TODO",
                comando=self._accion_todo_en_uno,
                color=COLOR_PRIMARIO, alto=52, radio=12,
                fuente=("Segoe UI", 12, "bold"))
            btn_todo.pack(fill="x", padx=14, pady=(0, 2))
            self.botones["🔄  TODO EN UNO"] = btn_todo

            tk.Label(panel_acc,
                     text="Ingesta entrada/ → BD → reparto → Excel → cartas",
                     font=("Segoe UI", 8), bg=COLOR_PANEL,
                     fg=COLOR_NEUTRO).pack(anchor="w", padx=16, pady=(0, 10))

            tk.Frame(panel_acc, bg=COLOR_BORDE, height=1).pack(
                fill="x", padx=14, pady=(2, 10))

            # ── Paso a paso ──
            tk.Label(panel_acc, text="PASO A PASO",
                     font=("Segoe UI", 8, "bold"),
                     bg=COLOR_PANEL, fg=COLOR_NEUTRO).pack(
                         anchor="w", padx=16, pady=(0, 6))

            pasos = [
                ("📥  Procesar facturas", self._accion_procesar_facturas),
                ("📊  Regenerar Excel",   self._accion_actualizar_excel),
                ("🔢  Calcular reparto",  self._accion_calcular_reparto),
                ("✉️  Generar cartas",    self._accion_generar_cartas),
            ]
            claves = ["📥  Procesar Facturas", "📊  Actualizar Excel",
                      "🔢  Calcular Reparto", "✉️  Generar Cartas"]
            for (texto, cmd), clave in zip(pasos, claves):
                btn = UIM.BotonSuave(panel_acc, texto=texto, comando=cmd, alto=36)
                btn.pack(fill="x", padx=14, pady=3)
                self.botones[clave] = btn

            tk.Frame(panel_acc, bg=COLOR_BORDE, height=1).pack(
                fill="x", padx=14, pady=10)

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
                UIM.BotonFantasma(panel_acc, texto=texto,
                                  comando=cmd).pack(fill="x", padx=6)
        else:
            # Respaldo clásico si falta ui_moderna.py
            acciones = [
                ("📥  Procesar Facturas", self._accion_procesar_facturas, COLOR_ACENTO),
                ("📊  Actualizar Excel",  self._accion_actualizar_excel,  COLOR_ACENTO),
                ("🔢  Calcular Reparto",  self._accion_calcular_reparto,  COLOR_ACENTO),
                ("✉️  Generar Cartas",    self._accion_generar_cartas,    COLOR_EXITO),
                ("🔄  TODO EN UNO",       self._accion_todo_en_uno,       COLOR_PRIMARIO),
            ]
            for texto, cmd, color in acciones:
                btn = tk.Button(panel_acc, text=texto,
                                font=("Segoe UI", 10, "bold"),
                                bg=color, fg="white", relief="flat",
                                cursor="hand2", pady=10, command=cmd)
                btn.pack(fill="x", padx=12, pady=6)
                self.botones[texto] = btn

        # ── PANEL DERECHO: LOG Y ESTADO ─────────────────────────────────────
        panel_log = tk.Frame(cuerpo, bg=COLOR_PANEL,
                             relief="flat",
                             highlightbackground=COLOR_BORDE,
                             highlightthickness=1)
        panel_log.grid(row=0, column=1, sticky="nsew")

        # Cabecera del log
        log_header = tk.Frame(panel_log, bg=COLOR_PANEL)
        log_header.pack(fill="x")
        tk.Label(log_header, text="ACTIVIDAD",
                 font=("Segoe UI", 8, "bold"), bg=COLOR_PANEL,
                 fg=COLOR_NEUTRO).pack(side="left", padx=14, pady=8)
        limpiar = tk.Label(log_header, text="Limpiar",
                           font=("Segoe UI", 9), bg=COLOR_PANEL,
                           fg=COLOR_NEUTRO, cursor="hand2")
        limpiar.pack(side="right", padx=12, pady=8)
        limpiar.bind("<Button-1>", lambda e: self._limpiar_log())
        limpiar.bind("<Enter>", lambda e: limpiar.config(fg=COLOR_ACENTO))
        limpiar.bind("<Leave>", lambda e: limpiar.config(fg=COLOR_NEUTRO))

        # Área de log
        self.log_area = scrolledtext.ScrolledText(
            panel_log,
            font=FUENTE_MONO,
            bg="#0D1117", fg="#C9D1D9",
            insertbackground="white",
            relief="flat",
            state="disabled",
            wrap="word",
            padx=8, pady=8
        )
        self.log_area.pack(fill="both", expand=True)

        # Tags de color para el log
        self.log_area.tag_config("ok",      foreground="#3FB950")
        self.log_area.tag_config("error",   foreground="#F85149")
        self.log_area.tag_config("aviso",   foreground="#E3B341")
        self.log_area.tag_config("info",    foreground="#58A6FF")
        self.log_area.tag_config("titulo",  foreground="#FFFFFF",
                                  font=("Consolas", 9, "bold"))
        self.log_area.tag_config("neutro",  foreground="#8B949E")

        # ── BARRA DE ESTADO ─────────────────────────────────────────────────
        self.barra_estado = tk.Frame(self, bg=COLOR_PANEL, height=30,
                                     highlightbackground=COLOR_BORDE,
                                     highlightthickness=1)
        self.barra_estado.pack(fill="x", side="bottom")
        self.barra_estado.pack_propagate(False)

        if UIM:
            self.punto_estado = UIM.PuntoEstado(self.barra_estado, fondo=COLOR_PANEL)
            self.punto_estado.pack(side="left", padx=(12, 4), pady=5)
        else:
            self.punto_estado = None

        self.lbl_estado = tk.Label(self.barra_estado, text="Listo",
                                    font=FUENTE_PEQUEÑA,
                                    bg=COLOR_PANEL, fg=COLOR_TEXTO)
        self.lbl_estado.pack(side="left", padx=(0, 12))

        self.lbl_bd = tk.Label(self.barra_estado,
                                text=f"BD: {RUTA_BD}",
                                font=("Segoe UI", 8),
                                bg=COLOR_PANEL, fg=COLOR_NEUTRO)
        self.lbl_bd.pack(side="right", padx=12)

        # Barra de progreso (oculta por defecto)
        estilo_pb = {"style": "Moderno.Horizontal.TProgressbar"} if UIM else {}
        self.progreso = ttk.Progressbar(self.barra_estado, mode="indeterminate",
                                         length=120, **estilo_pb)

    # -----------------------------------------------------------------------
    # LOG
    # -----------------------------------------------------------------------
    def log(self, mensaje: str, tipo: str = "neutro"):
        """Añade una línea al log en el hilo principal."""
        def _escribir():
            self.log_area.configure(state="normal")
            hora = datetime.now().strftime("%H:%M:%S")
            self.log_area.insert("end", f"[{hora}] ", "neutro")
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
            self.lbl_estado.config(text=texto)
            if self.punto_estado is not None:
                self.punto_estado.procesando(procesando)
            if procesando:
                self.progreso.pack(side="left", padx=8)
                self.progreso.start(12)
            else:
                self.progreso.stop()
                self.progreso.pack_forget()
        self.after(0, _act)

    # -----------------------------------------------------------------------
    # CARGA DE DATOS
    # -----------------------------------------------------------------------
    def _verificar_estructura(self):
        """Crea las carpetas necesarias si no existen."""
        for carpeta in [RUTA_ENTRADA, RUTA_PROCESADOS,
                        RUTA_EXCELS, RUTA_CARTAS,
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
            self.cb_comunidad["values"] = opciones
            self._ids_comunidad = {f"{r['codigo']} — {r['nombre']}": r["id_comunidad"]
                                    for r in rows}

            if opciones:
                self.cb_comunidad.set(opciones[0])
                self._on_comunidad_seleccionada()
        except Exception:
            pass

    def _on_comunidad_seleccionada(self, event=None):
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
            self.cb_periodo["values"] = opciones
            self._ids_periodo = {r["nombre"]: r["id_periodo"] for r in rows}
            if opciones:
                self.cb_periodo.set(opciones[0])
                self._on_periodo_seleccionado()
            else:
                self._actualizar_banner()
        except Exception:
            pass

    def _on_periodo_seleccionado(self, event=None):
        nombre = self.periodo_actual.get()
        self.id_periodo = self._ids_periodo.get(nombre) if hasattr(self, "_ids_periodo") else None
        self._actualizar_banner()

    def _actualizar_banner(self):
        """Actualiza el banner con la info del periodo activo y su rango de fechas."""
        if not self.id_periodo or not MOD.get("gestor_bd") or not RUTA_BD.exists():
            if hasattr(self, "lbl_banner"):
                self.lbl_banner.config(
                    text="Selecciona una comunidad y un periodo para empezar.",
                    fg=COLOR_PRIMARIO, bg="#EEF2FF"
                )
                self.banner_periodo.config(bg="#EEF2FF",
                                            highlightbackground="#C7D2FE")
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
                color_bg  = "#EEF2FF" if estado == "abierto" else "#F9FAFB"
                color_brd = "#C7D2FE" if estado == "abierto" else "#E5E7EB"
                icono     = "📅" if estado == "abierto" else "🔒"
                texto = (
                    f"{icono}  Periodo activo: {r['nombre']}  ·  "
                    f"{fi} → {ff}  ({dur})  ·  "
                    f"{n_facturas} factura(s) registrada(s)   "
                    f"— Las facturas procesadas se asignarán al periodo que corresponda por fecha."
                )
                self.lbl_banner.config(text=texto, bg=color_bg, fg=COLOR_PRIMARIO)
                self.banner_periodo.config(bg=color_bg,
                                            highlightbackground=color_brd)
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
        except Exception as e:
            self.log(f"❌ Error inesperado: {e}", "error")
            self.log(traceback.format_exc(), "error")
        finally:
            self._procesando = False
            self.after(0, self._habilitar_botones)
            self._estado("Listo")

    def _accion_procesar_facturas(self):
        self._en_hilo(self._procesar_facturas_impl)

    def _accion_actualizar_excel(self):
        self._en_hilo(self._actualizar_excel_impl)

    def _accion_calcular_reparto(self):
        self._en_hilo(self._calcular_reparto_impl)

    def _accion_generar_cartas(self):
        self._en_hilo(self._generar_cartas_impl)

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
                    shutil.move(str(ruta), str(dest))
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

    def _generar_cartas_impl(self):
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

        nombre_periodo = self.periodo_actual.get()
        carpeta_salida = RUTA_CARTAS / nombre_periodo
        carpeta_salida.mkdir(parents=True, exist_ok=True)

        resultado = MOD["carta_writer"].generar_todas_las_cartas(
            ruta_bd=str(RUTA_BD),
            id_comunidad=self.id_comunidad,
            nombre_periodo=nombre_periodo,
            ruta_plantilla=str(RUTA_PLANTILLA),
            carpeta_salida=str(carpeta_salida)
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
            self.log("❌ No se encontró LISTADO_COMUNIDADES.xlsx", "error")
            self.log(f"   Colócalo en: {BASE_DIR}/LISTADO_COMUNIDADES.xlsx", "aviso")
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
            self.log(f"❌ No se encontró Excel para comunidad {codigo}", "error")
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
            if hasattr(btn, "set_estado"):
                btn.set_estado("disabled")
            else:
                btn.configure(state="disabled", bg=COLOR_NEUTRO)

    def _habilitar_botones(self):
        colores = {
            "🔄  TODO EN UNO":  COLOR_PRIMARIO,
            "✉️  Generar Cartas": COLOR_EXITO,
        }
        for texto, btn in self.botones.items():
            if hasattr(btn, "set_estado"):
                btn.set_estado("normal")
            else:
                btn.configure(state="normal", bg=colores.get(texto, COLOR_ACENTO))

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

    def _dialogo_seleccionar_periodo_inicio(self):
        """
        Diálogo que aparece al arrancar el programa.
        Permite elegir rango de fechas exacto (día/mes/año) para el periodo de trabajo.
        Si ya hay un periodo seleccionado para la comunidad activa, lo precarga.
        """
        if not MOD.get("gestor_bd") or not RUTA_BD.exists():
            return  # Sin BD no podemos hacer nada útil

        dialogo = tk.Toplevel(self)
        dialogo.title("Seleccionar Periodo de Trabajo")
        dialogo.geometry("540x440")
        dialogo.configure(bg=COLOR_FONDO)
        dialogo.resizable(False, False)
        dialogo.transient(self)
        dialogo.grab_set()
        # Centrar sobre la ventana principal
        self.update_idletasks()
        x = self.winfo_x() + (self.winfo_width()  - 540) // 2
        y = self.winfo_y() + (self.winfo_height() - 440) // 2
        dialogo.geometry(f"540x440+{x}+{y}")

        MESES = ["Ene","Feb","Mar","Abr","May","Jun",
                 "Jul","Ago","Sep","Oct","Nov","Dic"]

        # ── Cabecera ─────────────────────────────────────────────────────────
        tk.Label(dialogo, text="SELECCIONAR PERIODO DE TRABAJO",
                 font=FUENTE_SECCION, bg=COLOR_PRIMARIO, fg="white"
                 ).pack(fill="x", ipady=10)

        tk.Label(dialogo,
                 text="Define el rango exacto del periodo que quieres gestionar.\n"
                      "Las facturas y lecturas se filtrarán por estos días exactos.\n"
                      "El consumo se calculará proporcionalmente al día.",
                 font=("Segoe UI", 9), bg=COLOR_FONDO, fg=COLOR_NEUTRO,
                 justify="center").pack(pady=(10, 2))

        # ── Comunidad ────────────────────────────────────────────────────────
        frm_com = tk.Frame(dialogo, bg=COLOR_FONDO)
        frm_com.pack(fill="x", padx=24, pady=4)
        tk.Label(frm_com, text="Comunidad:", font=FUENTE_NORMAL,
                 bg=COLOR_FONDO, width=18, anchor="w").pack(side="left")

        var_comunidad = tk.StringVar(value=self.comunidad_actual.get())
        opciones_com  = list(getattr(self, "_ids_comunidad", {}).keys())
        cb_com = ttk.Combobox(frm_com, textvariable=var_comunidad,
                               values=opciones_com, state="readonly",
                               width=34, font=FUENTE_NORMAL)
        cb_com.pack(side="left")

        # ── Periodos existentes ───────────────────────────────────────────────
        frm_per = tk.Frame(dialogo, bg=COLOR_FONDO)
        frm_per.pack(fill="x", padx=24, pady=4)
        tk.Label(frm_per, text="Periodo existente:", font=FUENTE_NORMAL,
                 bg=COLOR_FONDO, width=18, anchor="w").pack(side="left")
        var_per_existente = tk.StringVar(value="— Nuevo periodo —")
        cb_per_existente  = ttk.Combobox(frm_per, textvariable=var_per_existente,
                                          state="readonly", width=22, font=FUENTE_NORMAL)
        cb_per_existente.pack(side="left", padx=(0, 8))
        tk.Label(frm_per, text="(o define uno nuevo abajo)",
                 font=("Segoe UI", 8), bg=COLOR_FONDO, fg=COLOR_NEUTRO).pack(side="left")

        # Actualizar lista de periodos al cambiar comunidad
        _periodos_cache: dict[str, int] = {}

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
                cb_per_existente["values"] = opts
                var_per_existente.set(opts[1] if len(opts) > 1 else opts[0])
                _precargar_desde_existente()
            except Exception:
                pass

        var_comunidad.trace_add("write", _actualizar_periodos)

        # ── Separador ────────────────────────────────────────────────────────
        tk.Frame(dialogo, bg=COLOR_BORDE, height=1).pack(fill="x", padx=16, pady=6)

        def _crear_selector_fecha(parent, label_text: str,
                                   dia_def: int, mes_def: int, anio_def: int):
            frm = tk.Frame(parent, bg=COLOR_FONDO)
            frm.pack(fill="x", padx=24, pady=5)
            tk.Label(frm, text=label_text, font=("Segoe UI", 9, "bold"),
                     bg=COLOR_FONDO, fg=COLOR_PRIMARIO, width=28, anchor="w").pack(side="left")
            var_dia  = tk.StringVar(value=str(dia_def).zfill(2))
            var_mes  = tk.StringVar(value=MESES[mes_def - 1])
            var_anio = tk.StringVar(value=str(anio_def))
            frm_sel  = tk.Frame(frm, bg=COLOR_FONDO)
            frm_sel.pack(side="left")
            ttk.Combobox(frm_sel, textvariable=var_dia,
                         values=[str(d).zfill(2) for d in range(1, 32)],
                         width=4, state="readonly",
                         font=FUENTE_NORMAL).pack(side="left", padx=(0, 4))
            ttk.Combobox(frm_sel, textvariable=var_mes,
                         values=MESES, width=5, state="readonly",
                         font=FUENTE_NORMAL).pack(side="left", padx=(0, 4))
            anio_act = datetime.now().year
            ttk.Combobox(frm_sel, textvariable=var_anio,
                         values=[str(a) for a in range(2015, anio_act + 3)],
                         width=6, state="readonly",
                         font=FUENTE_NORMAL).pack(side="left")
            return var_dia, var_mes, var_anio

        def _vars_a_fecha(vd, vm, va) -> str:
            mes_num = MESES.index(vm.get()) + 1
            return f"{va.get()}-{mes_num:02d}-{vd.get()}"

        now       = datetime.now()
        anio_ini  = now.year - 1 if now.month < 9 else now.year

        tk.Label(dialogo, text="Fecha de INICIO del periodo:",
                 font=("Segoe UI", 9), bg=COLOR_FONDO, fg=COLOR_NEUTRO
                 ).pack(anchor="w", padx=24)
        vd_ini, vm_ini, va_ini = _crear_selector_fecha(
            dialogo, "  Día  /  Mes  /  Año:", 1, 9, anio_ini)

        tk.Frame(dialogo, bg=COLOR_BORDE, height=1).pack(fill="x", padx=16, pady=2)
        tk.Label(dialogo, text="Fecha de FIN del periodo:",
                 font=("Segoe UI", 9), bg=COLOR_FONDO, fg=COLOR_NEUTRO
                 ).pack(anchor="w", padx=24)
        vd_fin, vm_fin, va_fin = _crear_selector_fecha(
            dialogo, "  Día  /  Mes  /  Año:", 31, 8, anio_ini + 1)

        # ── Nombre y resumen ─────────────────────────────────────────────────
        frm_nom = tk.Frame(dialogo, bg=COLOR_FONDO)
        frm_nom.pack(fill="x", padx=24, pady=(6, 0))
        tk.Label(frm_nom, text="Nombre:", font=FUENTE_NORMAL,
                 bg=COLOR_FONDO, width=18, anchor="w").pack(side="left")
        var_nombre = tk.StringVar(value=f"{anio_ini}-{anio_ini+1}")
        ent_nombre = tk.Entry(frm_nom, textvariable=var_nombre,
                               font=FUENTE_NORMAL, width=18)
        ent_nombre.pack(side="left", padx=(0, 12))

        lbl_dur = tk.Label(frm_nom, text="", font=("Segoe UI", 8, "italic"),
                            bg=COLOR_FONDO, fg=COLOR_ACENTO)
        lbl_dur.pack(side="left")

        def _actualizar_dur(*_):
            try:
                from datetime import date as _d
                ini = _d.fromisoformat(_vars_a_fecha(vd_ini, vm_ini, va_ini))
                fin = _d.fromisoformat(_vars_a_fecha(vd_fin, vm_fin, va_fin))
                dias = (fin - ini).days
                meses = round(dias / 30.44, 1)
                lbl_dur.config(
                    text=f"→ {dias} días ({meses} meses)",
                    fg=COLOR_EXITO if dias > 0 else COLOR_ALERTA
                )
            except Exception:
                lbl_dur.config(text="")

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
        frm_bots = tk.Frame(dialogo, bg=COLOR_FONDO)
        frm_bots.pack(pady=10)

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

        tk.Button(frm_bots, text="✅  Trabajar con este periodo",
                  font=FUENTE_SECCION,
                  bg=COLOR_PRIMARIO, fg="white",
                  relief="flat", cursor="hand2", padx=12, pady=6,
                  command=_aplicar).pack(side="left", padx=8)

        tk.Button(frm_bots, text="Omitir",
                  font=FUENTE_PEQUEÑA,
                  bg=COLOR_PANEL, fg=COLOR_NEUTRO,
                  relief="flat", cursor="hand2",
                  command=dialogo.destroy).pack(side="left")

    def _nuevo_periodo(self):
        """
        Diálogo para crear o seleccionar un periodo con fechas exactas (día/mes/año).
        Soporta periodos de cualquier duración — 8 meses, 10 meses, año completo, etc.
        """
        dialogo = tk.Toplevel(self)
        dialogo.title("Definir Periodo de Regularización")
        dialogo.geometry("480x360")
        dialogo.configure(bg=COLOR_FONDO)
        dialogo.resizable(False, False)
        dialogo.transient(self)
        dialogo.grab_set()

        # ── Cabecera ─────────────────────────────────────────────────────────
        tk.Label(dialogo, text="NUEVO PERIODO",
                 font=FUENTE_SECCION, bg=COLOR_PRIMARIO, fg="white"
                 ).pack(fill="x", ipady=8)

        tk.Label(dialogo,
                 text="Define el rango exacto. Las facturas y lecturas se filtrarán\n"
                      "por estos días, calculando consumo proporcional al día.",
                 font=("Segoe UI", 8), bg=COLOR_FONDO, fg=COLOR_NEUTRO,
                 justify="center").pack(pady=(10, 4))

        # ── Nombre ───────────────────────────────────────────────────────────
        frm_nombre = tk.Frame(dialogo, bg=COLOR_FONDO)
        frm_nombre.pack(fill="x", padx=24, pady=4)
        tk.Label(frm_nombre, text="Nombre del periodo:", font=FUENTE_NORMAL,
                 bg=COLOR_FONDO, width=22, anchor="w").pack(side="left")
        ent_nombre = tk.Entry(frm_nombre, font=FUENTE_NORMAL, width=18)
        ent_nombre.insert(0, "2024-2025")
        ent_nombre.pack(side="left")

        # ── Función auxiliar para crear un selector de fecha ─────────────────
        MESES = ["Ene","Feb","Mar","Abr","May","Jun",
                 "Jul","Ago","Sep","Oct","Nov","Dic"]

        def _crear_selector_fecha(parent, label_text: str, dia_def: int,
                                   mes_def: int, anio_def: int):
            """Crea una fila con selector de día, mes y año. Devuelve variables."""
            frm = tk.Frame(parent, bg=COLOR_FONDO)
            frm.pack(fill="x", padx=24, pady=6)

            tk.Label(frm, text=label_text, font=FUENTE_NORMAL,
                     bg=COLOR_FONDO, width=22, anchor="w").pack(side="left")

            var_dia  = tk.StringVar(value=str(dia_def).zfill(2))
            var_mes  = tk.StringVar(value=MESES[mes_def - 1])
            var_anio = tk.StringVar(value=str(anio_def))

            frm_sel = tk.Frame(frm, bg=COLOR_FONDO)
            frm_sel.pack(side="left")

            # Día
            dias = [str(d).zfill(2) for d in range(1, 32)]
            cb_dia = ttk.Combobox(frm_sel, textvariable=var_dia,
                                   values=dias, width=4, state="readonly",
                                   font=FUENTE_NORMAL)
            cb_dia.pack(side="left", padx=(0, 4))

            # Mes
            cb_mes = ttk.Combobox(frm_sel, textvariable=var_mes,
                                   values=MESES, width=5, state="readonly",
                                   font=FUENTE_NORMAL)
            cb_mes.pack(side="left", padx=(0, 4))

            # Año
            anio_actual = datetime.now().year
            anios = [str(a) for a in range(2015, anio_actual + 3)]
            cb_anio = ttk.Combobox(frm_sel, textvariable=var_anio,
                                    values=anios, width=6, state="readonly",
                                    font=FUENTE_NORMAL)
            cb_anio.pack(side="left")

            return var_dia, var_mes, var_anio

        def _vars_a_fecha(var_dia, var_mes, var_anio) -> str:
            mes_num = MESES.index(var_mes.get()) + 1
            return f"{var_anio.get()}-{mes_num:02d}-{var_dia.get()}"

        now = datetime.now()
        anio_ini = now.year - 1 if now.month < 9 else now.year

        # ── Separador ────────────────────────────────────────────────────────
        tk.Frame(dialogo, bg=COLOR_BORDE, height=1).pack(fill="x", padx=16, pady=4)
        tk.Label(dialogo, text="Fecha de INICIO (lectura inicial del periodo):",
                 font=("Segoe UI", 9, "bold"), bg=COLOR_FONDO, fg=COLOR_PRIMARIO
                 ).pack(anchor="w", padx=24)

        vd_ini, vm_ini, va_ini = _crear_selector_fecha(
            dialogo, "  Día / Mes / Año:", 1, 9, anio_ini)

        tk.Frame(dialogo, bg=COLOR_BORDE, height=1).pack(fill="x", padx=16, pady=4)
        tk.Label(dialogo, text="Fecha de FIN (lectura final del periodo):",
                 font=("Segoe UI", 9, "bold"), bg=COLOR_FONDO, fg=COLOR_PRIMARIO
                 ).pack(anchor="w", padx=24)

        vd_fin, vm_fin, va_fin = _crear_selector_fecha(
            dialogo, "  Día / Mes / Año:", 31, 8, anio_ini + 1)

        # ── Etiqueta de resumen ───────────────────────────────────────────────
        lbl_resumen = tk.Label(dialogo, text="", font=("Segoe UI", 9, "italic"),
                                bg=COLOR_FONDO, fg=COLOR_ACENTO)
        lbl_resumen.pack(pady=(2, 0))

        def _actualizar_resumen(*_):
            try:
                f_ini = _vars_a_fecha(vd_ini, vm_ini, va_ini)
                f_fin = _vars_a_fecha(vd_fin, vm_fin, va_fin)
                from datetime import date as _d
                ini = _d.fromisoformat(f_ini)
                fin = _d.fromisoformat(f_fin)
                dias = (fin - ini).days
                meses = round(dias / 30.44, 1)
                lbl_resumen.config(
                    text=f"Duración: {dias} días  ≈  {meses} meses",
                    fg=COLOR_EXITO if dias > 0 else COLOR_ALERTA
                )
            except Exception:
                lbl_resumen.config(text="", fg=COLOR_NEUTRO)

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

        tk.Button(dialogo, text="✅  Crear Periodo",
                  font=FUENTE_SECCION,
                  bg=COLOR_PRIMARIO, fg="white",
                  relief="flat", cursor="hand2",
                  command=_crear).pack(pady=12)

    def _nueva_comunidad(self):
        """Diálogo para registrar una nueva comunidad."""
        dialogo = tk.Toplevel(self)
        dialogo.title("Nueva Comunidad")
        dialogo.geometry("420x280")
        dialogo.configure(bg=COLOR_FONDO)
        dialogo.transient(self)
        dialogo.grab_set()

        campos = [
            ("Código (ej: 644):",   "codigo",  "644"),
            ("Nombre completo:",     "nombre",  "CDAD. PROP. …"),
            ("CIF de la comunidad:", "cif",     "H99258139"),
            ("Nº de viviendas:",     "viviendas","120"),
        ]
        entradas = {}
        for i, (label, clave, defecto) in enumerate(campos):
            tk.Label(dialogo, text=label, font=FUENTE_NORMAL,
                     bg=COLOR_FONDO).grid(row=i, column=0, padx=20, pady=6, sticky="w")
            e = tk.Entry(dialogo, font=FUENTE_NORMAL, width=28)
            e.insert(0, defecto)
            e.grid(row=i, column=1, padx=8, pady=6)
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

        tk.Button(dialogo, text="Registrar Comunidad",
                  font=FUENTE_SECCION,
                  bg=COLOR_PRIMARIO, fg="white",
                  relief="flat", cursor="hand2",
                  command=_crear).grid(row=len(campos), column=0,
                                       columnspan=2, pady=16)

    def _configurar_rutas(self):
        """Muestra las rutas actuales y permite cambiarlas."""
        ventana = tk.Toplevel(self)
        ventana.title("Configuración de Rutas")
        ventana.geometry("600x320")
        ventana.configure(bg=COLOR_FONDO)
        ventana.transient(self)
        ventana.grab_set()

        rutas = [
            ("Base de datos:",      str(RUTA_BD)),
            ("Carpeta entrada/:",   str(RUTA_ENTRADA)),
            ("Carpeta procesados/:",str(RUTA_PROCESADOS)),
            ("Excels Maestros/:",   str(RUTA_EXCELS)),
            ("Plantilla cartas:",   str(RUTA_PLANTILLA)),
            ("Cartas generadas/:",  str(RUTA_CARTAS)),
        ]
        for i, (label, ruta) in enumerate(rutas):
            tk.Label(ventana, text=label, font=FUENTE_PEQUEÑA,
                     bg=COLOR_FONDO, anchor="w", width=22).grid(
                row=i, column=0, padx=16, pady=4, sticky="w")
            tk.Label(ventana, text=ruta, font=FUENTE_MONO,
                     bg=COLOR_FONDO, fg=COLOR_ACENTO, anchor="w").grid(
                row=i, column=1, padx=4, pady=4, sticky="w")

        tk.Label(ventana,
                 text="Para cambiar las rutas, edita las constantes al inicio de app.py",
                 font=FUENTE_PEQUEÑA, bg=COLOR_FONDO, fg=COLOR_NEUTRO).grid(
            row=len(rutas), column=0, columnspan=2, pady=12, padx=16)


# ---------------------------------------------------------------------------
# PUNTO DE ENTRADA
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    app = AppGestionFincas()
    app.mainloop()
