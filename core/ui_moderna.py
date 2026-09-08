"""
ui_moderna.py
=============
Capa visual basada en CustomTkinter para la app de Gestión de Fincas.

REQUISITO:  pip install customtkinter
            (app.py intenta instalarlo automáticamente si falta)

Incluye:
  - C:        paleta de colores con pares (claro, oscuro) — CustomTkinter
              cambia todos los widgets automáticamente al alternar el modo
  - fuente(): fuentes modernas (Segoe UI Variable en Win11, Segoe UI si no)
  - InterruptorTema: conmutador claro/oscuro con persistencia en disco
  - PuntoEstado:     indicador circular con pulso animado al procesar
  - LineaGradiente:  línea de acento degradada bajo la cabecera
  - aparecer():      animación de fundido de entrada para ventanas
"""

import json
import tkinter as tk
import tkinter.font as tkfont
from pathlib import Path

import customtkinter as ctk

# ---------------------------------------------------------------------------
# PALETA — cada clave: (modo claro, modo oscuro)
# ---------------------------------------------------------------------------
C = {
    "fondo":              ("#F3F6F8", "#0E1718"),
    "panel":              ("#FFFFFF", "#171A23"),
    "panel_2":            ("#F8FAFB", "#172224"),
    "borde":              ("#D9E3E6", "#2A3A3C"),
    "primario":           ("#0F766E", "#48B8AD"),
    "primario_hover":     ("#0B615B", "#69D1C5"),
    "acento_suave":       ("#E5F4F1", "#173B39"),
    "acento_suave_hover": ("#D3ECE8", "#20514D"),
    "texto":              ("#0F172A", "#E7EAF3"),
    "texto_sec":          ("#64748B", "#8B93A7"),
    "exito":              ("#059669", "#34D399"),
    "exito_hover":        ("#047857", "#5EEAD4"),
    "alerta":             ("#DC2626", "#F87171"),
    "aviso":              ("#D97706", "#FBBF24"),
    "banner_abierto":     ("#E5F4F1", "#173B39"),
    "banner_cerrado":     ("#EEF2F3", "#182426"),
    "log_fondo":          ("#0D1117", "#0D1117"),
}

# Compatibilidad con código antiguo que use TEMA["clave"]
TEMA = {k: v[0] for k, v in C.items()}

RUTA_PREFS = Path(__file__).parent.parent / "config" / "ui_prefs.json"


# ---------------------------------------------------------------------------
# MODO CLARO / OSCURO
# ---------------------------------------------------------------------------
def modo_actual() -> str:
    """'light' o 'dark'."""
    return "dark" if ctk.get_appearance_mode().lower() == "dark" else "light"


def color(par) -> str:
    """Devuelve el color del par (claro, oscuro) según el modo activo.
    Útil para widgets tk clásicos (Canvas, Text) que no entienden tuplas."""
    if isinstance(par, str):
        par = C[par]
    return par[1] if modo_actual() == "dark" else par[0]


def cargar_preferencia() -> str:
    try:
        return json.loads(RUTA_PREFS.read_text(encoding="utf-8")).get("tema", "light")
    except Exception:
        return "light"


def guardar_preferencia(tema: str):
    try:
        RUTA_PREFS.parent.mkdir(parents=True, exist_ok=True)
        RUTA_PREFS.write_text(json.dumps({"tema": tema}), encoding="utf-8")
    except Exception:
        pass


def iniciar():
    """Configura CustomTkinter antes de crear la ventana principal."""
    ctk.set_appearance_mode(cargar_preferencia())
    ctk.set_default_color_theme("green")
    ctk.set_widget_scaling(1.0)


# ---------------------------------------------------------------------------
# FUENTES
# ---------------------------------------------------------------------------
_familia_cache = None


def familia() -> str:
    """Elige la mejor fuente disponible del sistema."""
    global _familia_cache
    if _familia_cache is None:
        disponibles = set(tkfont.families())
        for candidata in ("Aptos Display", "Inter", "Trebuchet MS",
                          "Segoe UI Variable Display", "Segoe UI Variable", "Segoe UI"):
            if candidata in disponibles:
                _familia_cache = candidata
                break
        else:
            _familia_cache = "Segoe UI"
    return _familia_cache


def fuente(tam: int = 13, peso: str = "normal") -> ctk.CTkFont:
    return ctk.CTkFont(family=familia(), size=tam, weight=peso)


def fuente_mono(tam: int = 12) -> tuple:
    disponibles = set(tkfont.families())
    fam = "Cascadia Code" if "Cascadia Code" in disponibles else "Consolas"
    return (fam, tam)


# ---------------------------------------------------------------------------
# ANIMACIÓN DE ENTRADA (fundido)
# ---------------------------------------------------------------------------
def aparecer(ventana, duracion_ms: int = 220):
    """Fundido de entrada (alpha 0 → 1). Silencioso si el SO no lo soporta."""
    try:
        ventana.attributes("-alpha", 0.0)
        pasos = max(int(duracion_ms / 16), 1)

        def _paso(i=0):
            try:
                ventana.attributes("-alpha", min(1.0, (i + 1) / pasos))
                if i + 1 < pasos:
                    ventana.after(16, _paso, i + 1)
            except tk.TclError:
                pass

        ventana.after(10, _paso)
    except tk.TclError:
        pass


# ---------------------------------------------------------------------------
# UTILIDADES DE COLOR (compatibilidad)
# ---------------------------------------------------------------------------
def _hex_a_rgb(c):
    c = c.lstrip("#")
    return tuple(int(c[i:i + 2], 16) for i in (0, 2, 4))


def _rgb_a_hex(rgb):
    return "#%02x%02x%02x" % tuple(max(0, min(255, int(v))) for v in rgb)


def mezclar(c1: str, c2: str, t: float) -> str:
    a, b = _hex_a_rgb(c1), _hex_a_rgb(c2)
    return _rgb_a_hex(tuple(a[i] + (b[i] - a[i]) * t for i in range(3)))


def oscurecer(c: str, t: float = 0.15) -> str:
    return mezclar(c, "#000000", t)


# ---------------------------------------------------------------------------
# INTERRUPTOR DE TEMA  ☀️ / 🌙
# ---------------------------------------------------------------------------
class InterruptorTema(ctk.CTkFrame):
    """Conmutador claro/oscuro. Guarda la preferencia y notifica a los
    widgets tk clásicos registrados con .al_cambiar(callback)."""

    def __init__(self, master, **kw):
        kw.setdefault("fg_color", "transparent")
        super().__init__(master, **kw)
        self._callbacks = []

        self._icono = ctk.CTkLabel(self, text="☀️", font=fuente(14),
                                   text_color=C["texto_sec"])
        self._icono.pack(side="left", padx=(0, 6))

        self._sw = ctk.CTkSwitch(
            self, text="", width=42, command=self._alternar,
            progress_color=C["primario"], button_color=("#FFFFFF", "#E7EAF3"),
            fg_color=("#CBD2E0", "#3A4160")
        )
        self._sw.pack(side="left")
        if modo_actual() == "dark":
            self._sw.select()
            self._icono.configure(text="🌙")

    def al_cambiar(self, callback):
        self._callbacks.append(callback)

    def _alternar(self):
        nuevo = "dark" if self._sw.get() else "light"
        ctk.set_appearance_mode(nuevo)
        guardar_preferencia(nuevo)
        self._icono.configure(text="🌙" if nuevo == "dark" else "☀️")
        for cb in self._callbacks:
            try:
                cb(nuevo)
            except Exception:
                pass


# ---------------------------------------------------------------------------
# PUNTO DE ESTADO (pulso animado)
# ---------------------------------------------------------------------------
class PuntoEstado(tk.Canvas):
    """Círculo de estado: verde fijo (listo) o pulso de acento (procesando).
    Widget tk clásico → se re-pinta al cambiar el tema vía refrescar()."""

    DIAM = 10

    def __init__(self, master, **kw):
        super().__init__(master, width=18, height=18,
                         bg=color("panel"), highlightthickness=0, bd=0, **kw)
        self._pulsando = False
        self._fase = 0.0
        self._anim_id = None
        self._dibujar(color("exito"))

    def refrescar(self, _modo=None):
        self.configure(bg=color("panel"))
        if not self._pulsando:
            self._dibujar(color("exito"))

    def _dibujar(self, c):
        self.delete("all")
        m = (18 - self.DIAM) / 2
        self.create_oval(m, m, m + self.DIAM, m + self.DIAM, fill=c, outline=c)

    def _pulso(self):
        self._anim_id = None
        if not self._pulsando:
            self._dibujar(color("exito"))
            return
        self._fase = (self._fase + 0.12) % 2.0
        t = self._fase if self._fase <= 1.0 else 2.0 - self._fase
        self._dibujar(mezclar(color("acento_suave"), color("primario"), t))
        self._anim_id = self.after(50, self._pulso)

    def procesando(self, activo: bool):
        self._pulsando = activo
        if activo and self._anim_id is None:
            self._pulso()
        elif not activo:
            if self._anim_id:
                self.after_cancel(self._anim_id)
                self._anim_id = None
            self._dibujar(color("exito"))


# ---------------------------------------------------------------------------
# LÍNEA DEGRADADA (acento bajo la cabecera)
# ---------------------------------------------------------------------------
class LineaGradiente(tk.Canvas):
    """Línea de 2 px con degradado primario → fondo. Se repinta al
    redimensionar y al cambiar de tema (vía refrescar())."""

    def __init__(self, master, altura: int = 2, **kw):
        super().__init__(master, height=altura, bg=color("fondo"),
                         highlightthickness=0, bd=0, **kw)
        self._altura = altura
        self.bind("<Configure>", lambda e: self._pintar())

    def refrescar(self, _modo=None):
        self.configure(bg=color("fondo"))
        self._pintar()

    def _pintar(self):
        self.delete("all")
        w = max(self.winfo_width(), 2)
        c1, c2 = color("primario"), color("fondo")
        tramos = 80
        for i in range(tramos):
            c = mezclar(c1, c2, i / tramos)
            self.create_rectangle(w * i / tramos, 0,
                                  w * (i + 1) / tramos, self._altura,
                                  fill=c, outline=c)


# ---------------------------------------------------------------------------
# COMBO MODERNO — desplegable con buscador y lista scrollable
# ---------------------------------------------------------------------------
class ComboModerno(ctk.CTkFrame):
    """
    Sustituto moderno de CTkComboBox:
      - Campo redondeado con flecha
      - Panel flotante con esquinas redondeadas
      - Buscador integrado cuando hay más de 7 opciones
      - Resalta la opción seleccionada
    API compatible: set(), get(), configure(values=[...]), variable=, command=
    """

    def __init__(self, master, variable=None, values=None, command=None,
                 width=340, height=36, **kw):
        super().__init__(master, fg_color=C["panel"], corner_radius=9,
                         border_width=1, border_color=C["borde"],
                         width=width, height=height, **kw)
        self.pack_propagate(False)
        self.grid_propagate(False)
        self._variable = variable if variable is not None else tk.StringVar()
        self._values   = list(values or [])
        self._command  = command
        self._panel    = None
        self._bind_id  = None
        self._grab_prev = None

        self._lbl = ctk.CTkLabel(self, textvariable=self._variable,
                                 font=fuente(13), text_color=C["texto"],
                                 anchor="w")
        self._lbl.pack(side="left", fill="both", expand=True, padx=(12, 4))
        self._flecha = ctk.CTkLabel(self, text="▾", font=fuente(14),
                                    text_color=C["texto_sec"], width=22)
        self._flecha.pack(side="right", padx=(0, 10))
        for w in (self, self._lbl, self._flecha):
            w.bind("<Button-1>", self._alternar)
        try:
            self.configure(cursor="hand2")
        except Exception:
            pass

    # ── API compatible con CTkComboBox ──────────────────────────────────
    def set(self, valor: str):
        self._variable.set(valor)

    def get(self) -> str:
        return self._variable.get()

    def configure(self, require_redraw=False, **kw):
        if "values" in kw:
            self._values = list(kw.pop("values") or [])
        kw.pop("state", None)   # siempre "readonly" de facto
        if kw:
            super().configure(require_redraw=require_redraw, **kw)

    # ── apertura / cierre del panel ──────────────────────────────────────
    def _alternar(self, _=None):
        if self._panel and self._panel.winfo_exists():
            self._cerrar()
        else:
            self._abrir()

    def _abrir(self):
        self._cerrar()
        self.update_idletasks()
        x = self.winfo_rootx()
        y = self.winfo_rooty() + self.winfo_height() + 4
        w = max(self.winfo_width(), 230)

        con_buscador = len(self._values) > 7
        alto_lista = min(9, max(2, len(self._values))) * 34 + 10
        alto_total = alto_lista + (54 if con_buscador else 14)

        # Si hay un grab activo (diálogo modal), soltarlo mientras esté abierto
        raiz = self.winfo_toplevel()
        try:
            actual = raiz.grab_current()
            if actual is not None:
                self._grab_prev = actual
                actual.grab_release()
        except Exception:
            self._grab_prev = None

        panel = ctk.CTkToplevel(self)
        panel.overrideredirect(True)
        panel.attributes("-topmost", True)
        panel.geometry(f"{w}x{alto_total}+{x}+{y}")

        marco = ctk.CTkFrame(panel, fg_color=C["panel"], corner_radius=10,
                             border_width=1, border_color=C["borde"])
        marco.pack(fill="both", expand=True)

        filtro = tk.StringVar()
        if con_buscador:
            ent = ctk.CTkEntry(marco, textvariable=filtro,
                               placeholder_text="Buscar…",
                               height=32, font=fuente(12),
                               border_color=C["borde"])
            ent.pack(fill="x", padx=8, pady=(8, 4))
            ent.bind("<Escape>", lambda e: self._cerrar())
            panel.after(140, lambda: ent.focus_set() if ent.winfo_exists() else None)

        lista = ctk.CTkScrollableFrame(marco, fg_color="transparent",
                                       height=alto_lista)
        lista.pack(fill="both", expand=True, padx=4, pady=(0, 6))

        def _rellenar(*_):
            for hijo in lista.winfo_children():
                hijo.destroy()
            txt = filtro.get().lower().strip()
            visibles = ([v for v in self._values if txt in v.lower()]
                        if txt else self._values)
            actual = self._variable.get()
            for v in visibles[:400]:
                es_actual = (v == actual)
                ctk.CTkButton(
                    lista, text=v, anchor="w", height=30, corner_radius=6,
                    font=fuente(12),
                    fg_color=C["acento_suave"] if es_actual else "transparent",
                    hover_color=C["acento_suave_hover"] if es_actual
                                else C["acento_suave"],
                    text_color=C["primario"] if es_actual else C["texto"],
                    command=lambda val=v: self._elegir(val)
                ).pack(fill="x", padx=2, pady=1)
            if not visibles:
                ctk.CTkLabel(lista, text="Sin resultados",
                             font=fuente(12), text_color=C["texto_sec"]
                             ).pack(pady=12)

        filtro.trace_add("write", _rellenar)
        _rellenar()

        panel.bind("<Escape>", lambda e: self._cerrar())
        self._panel = panel
        aparecer(panel, 130)

        # Cerrar al hacer clic fuera (en la ventana principal / diálogo)
        self._bind_id = raiz.bind("<ButtonPress-1>", self._click_fuera, add="+")

    def _click_fuera(self, evento):
        if not (self._panel and self._panel.winfo_exists()):
            return
        w = evento.widget
        while w is not None:
            if w is self:
                return  # clic en el propio combo → lo gestiona _alternar
            w = getattr(w, "master", None)
        self._cerrar()

    def _elegir(self, valor: str):
        self._variable.set(valor)
        self._cerrar()
        if self._command:
            try:
                self._command(valor)
            except Exception:
                pass

    def _cerrar(self):
        if self._bind_id:
            try:
                self.winfo_toplevel().unbind("<ButtonPress-1>", self._bind_id)
            except Exception:
                pass
            self._bind_id = None
        if self._panel and self._panel.winfo_exists():
            self._panel.destroy()
        self._panel = None
        if self._grab_prev is not None:
            try:
                if self._grab_prev.winfo_exists():
                    self._grab_prev.grab_set()
            except Exception:
                pass
            self._grab_prev = None
