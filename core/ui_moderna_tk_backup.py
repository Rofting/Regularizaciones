"""
ui_moderna.py
=============
Tema visual y widgets modernos para la app de Gestión de Fincas.
Solo tkinter estándar — sin dependencias nuevas.

Incluye:
  - TEMA: paleta de colores minimalista (indigo + grises suaves)
  - BotonModerno: botón redondeado con animación de hover y pulsación
  - BotonSuave:   botón secundario "soft" (fondo claro, tinta de acento)
  - BotonFantasma: botón de texto con hover sutil (utilidades)
  - PuntoEstado:  indicador circular con animación de pulso al procesar
  - aplicar_estilo_ttk: estiliza Combobox/Progressbar acorde al tema
"""

import tkinter as tk
from tkinter import ttk

# ---------------------------------------------------------------------------
# PALETA
# ---------------------------------------------------------------------------
TEMA = {
    "fondo":        "#F4F5F7",
    "panel":        "#FFFFFF",
    "borde":        "#E5E7EB",
    "primario":     "#4F46E5",   # indigo
    "primario_osc": "#4338CA",
    "acento":       "#6366F1",
    "acento_suave": "#EEF2FF",   # fondo de botones secundarios
    "acento_borde": "#C7D2FE",
    "exito":        "#10B981",
    "exito_osc":    "#059669",
    "alerta":       "#EF4444",
    "aviso":        "#F59E0B",
    "texto":        "#111827",
    "texto_sec":    "#6B7280",
    "deshabilitado":"#D1D5DB",
    "log_fondo":    "#0D1117",
}

FUENTE_UI       = ("Segoe UI", 10)
FUENTE_UI_BOLD  = ("Segoe UI", 10, "bold")
FUENTE_TITULO   = ("Segoe UI", 15, "bold")
FUENTE_MINI     = ("Segoe UI", 8)
FUENTE_SECCION  = ("Segoe UI", 8, "bold")


# ---------------------------------------------------------------------------
# UTILIDADES DE COLOR
# ---------------------------------------------------------------------------
def _hex_a_rgb(c):
    c = c.lstrip("#")
    return tuple(int(c[i:i+2], 16) for i in (0, 2, 4))


def _rgb_a_hex(rgb):
    return "#%02x%02x%02x" % tuple(max(0, min(255, int(v))) for v in rgb)


def mezclar(c1: str, c2: str, t: float) -> str:
    """Interpola entre dos colores hex. t=0 → c1, t=1 → c2."""
    a, b = _hex_a_rgb(c1), _hex_a_rgb(c2)
    return _rgb_a_hex(tuple(a[i] + (b[i] - a[i]) * t for i in range(3)))


def oscurecer(c: str, t: float = 0.15) -> str:
    return mezclar(c, "#000000", t)


# ---------------------------------------------------------------------------
# BOTÓN MODERNO (redondeado, animado)
# ---------------------------------------------------------------------------
class BotonModerno(tk.Canvas):
    """
    Botón con esquinas redondeadas dibujado en Canvas.
    Animación suave de color al pasar el ratón y efecto de pulsación.
    """

    def __init__(self, master, texto="", comando=None,
                 color=None, color_texto="white", color_hover=None,
                 radio=10, alto=42, fuente=FUENTE_UI_BOLD,
                 fondo=None, **kw):
        fondo = fondo or TEMA["panel"]
        super().__init__(master, height=alto, bg=fondo,
                         highlightthickness=0, bd=0, **kw)
        self._texto       = texto
        self._comando     = comando
        self._color_base  = color or TEMA["primario"]
        self._color_hover = color_hover or oscurecer(self._color_base, 0.18)
        self._color_texto = color_texto
        self._color_actual = self._color_base
        self._radio  = radio
        self._alto   = alto
        self._fuente = fuente
        self._estado = "normal"
        self._anim_id = None
        self._t_actual = 0.0        # 0 = base, 1 = hover
        self._t_destino = 0.0
        self._presionado = False

        self.configure(cursor="hand2")
        self.bind("<Configure>",       lambda e: self._dibujar())
        self.bind("<Enter>",           self._al_entrar)
        self.bind("<Leave>",           self._al_salir)
        self.bind("<ButtonPress-1>",   self._al_pulsar)
        self.bind("<ButtonRelease-1>", self._al_soltar)

    # ── dibujo ──────────────────────────────────────────────────────────
    def _dibujar(self):
        self.delete("all")
        w = max(self.winfo_width(), 2)
        h = self._alto
        margen_y = 2 if self._presionado else 0
        color = (TEMA["deshabilitado"] if self._estado == "disabled"
                 else self._color_actual)
        self.create_polygon(
            self._puntos_redondeados(1, 1 + margen_y, w - 1, h - 1),
            smooth=True, fill=color, outline=color)
        fg = "white" if self._estado == "disabled" else self._color_texto
        self.create_text(w // 2, h // 2 + margen_y, text=self._texto,
                         font=self._fuente, fill=fg)

    def _puntos_redondeados(self, x1, y1, x2, y2):
        r = min(self._radio, (y2 - y1) // 2)
        return [
            x1+r, y1,  x2-r, y1,  x2, y1,  x2, y1+r,
            x2, y2-r,  x2, y2,  x2-r, y2,  x1+r, y2,
            x1, y2,  x1, y2-r,  x1, y1+r,  x1, y1,
        ]

    # ── animación ───────────────────────────────────────────────────────
    def _animar(self):
        self._anim_id = None
        paso = 0.22
        if abs(self._t_destino - self._t_actual) < 0.02:
            self._t_actual = self._t_destino
        else:
            self._t_actual += paso if self._t_destino > self._t_actual else -paso
            self._anim_id = self.after(16, self._animar)
        self._color_actual = mezclar(self._color_base, self._color_hover,
                                     max(0.0, min(1.0, self._t_actual)))
        self._dibujar()

    def _ir_a(self, destino: float):
        self._t_destino = destino
        if self._anim_id is None:
            self._animar()

    # ── eventos ─────────────────────────────────────────────────────────
    def _al_entrar(self, _=None):
        if self._estado == "normal":
            self._ir_a(1.0)

    def _al_salir(self, _=None):
        self._presionado = False
        if self._estado == "normal":
            self._ir_a(0.0)

    def _al_pulsar(self, _=None):
        if self._estado == "normal":
            self._presionado = True
            self._dibujar()

    def _al_soltar(self, evento=None):
        if self._estado != "normal" or not self._presionado:
            return
        self._presionado = False
        self._dibujar()
        if self._comando and 0 <= evento.x <= self.winfo_width() \
                and 0 <= evento.y <= self.winfo_height():
            self._comando()

    # ── API pública ─────────────────────────────────────────────────────
    def set_estado(self, estado: str):
        """'normal' o 'disabled'."""
        self._estado = estado
        self.configure(cursor="hand2" if estado == "normal" else "arrow")
        if estado == "disabled" and self._anim_id:
            self.after_cancel(self._anim_id)
            self._anim_id = None
            self._t_actual = self._t_destino = 0.0
            self._color_actual = self._color_base
        self._dibujar()

    def set_texto(self, texto: str):
        self._texto = texto
        self._dibujar()


class BotonSuave(BotonModerno):
    """Botón secundario: fondo suave de acento, texto de acento."""

    def __init__(self, master, texto="", comando=None, alto=36, **kw):
        super().__init__(
            master, texto=texto, comando=comando,
            color=TEMA["acento_suave"],
            color_hover=TEMA["acento_borde"],
            color_texto=TEMA["primario"],
            alto=alto, radio=9,
            fuente=("Segoe UI", 9, "bold"),
            **kw)

    def _dibujar(self):
        # Igual que el padre pero el texto siempre en color de acento
        self.delete("all")
        w = max(self.winfo_width(), 2)
        h = self._alto
        margen_y = 2 if self._presionado else 0
        if self._estado == "disabled":
            relleno, fg = "#F3F4F6", TEMA["deshabilitado"]
        else:
            relleno, fg = self._color_actual, self._color_texto
        self.create_polygon(
            self._puntos_redondeados(1, 1 + margen_y, w - 1, h - 1),
            smooth=True, fill=relleno, outline=relleno)
        self.create_text(w // 2, h // 2 + margen_y, text=self._texto,
                         font=self._fuente, fill=fg)


class BotonFantasma(tk.Label):
    """Botón de texto plano con hover sutil (para utilidades)."""

    def __init__(self, master, texto="", comando=None,
                 fondo=None, **kw):
        fondo = fondo or TEMA["panel"]
        super().__init__(master, text=texto, font=("Segoe UI", 9),
                         bg=fondo, fg=TEMA["texto_sec"],
                         anchor="w", padx=12, pady=6,
                         cursor="hand2", **kw)
        self._fondo = fondo
        self._comando = comando
        self.bind("<Enter>", lambda e: self.config(
            bg=TEMA["acento_suave"], fg=TEMA["primario"]))
        self.bind("<Leave>", lambda e: self.config(
            bg=self._fondo, fg=TEMA["texto_sec"]))
        self.bind("<Button-1>", lambda e: comando() if comando else None)


# ---------------------------------------------------------------------------
# PUNTO DE ESTADO (con pulso animado)
# ---------------------------------------------------------------------------
class PuntoEstado(tk.Canvas):
    """Círculo de estado: verde fijo (listo) o pulso indigo (procesando)."""

    DIAM = 10

    def __init__(self, master, fondo=None, **kw):
        fondo = fondo or TEMA["panel"]
        super().__init__(master, width=18, height=18, bg=fondo,
                         highlightthickness=0, bd=0, **kw)
        self._pulsando = False
        self._fase = 0.0
        self._anim_id = None
        self._dibujar(TEMA["exito"])

    def _dibujar(self, color):
        self.delete("all")
        m = (18 - self.DIAM) / 2
        self.create_oval(m, m, m + self.DIAM, m + self.DIAM,
                         fill=color, outline=color)

    def _pulso(self):
        self._anim_id = None
        if not self._pulsando:
            self._dibujar(TEMA["exito"])
            return
        self._fase = (self._fase + 0.12) % 2.0
        t = self._fase if self._fase <= 1.0 else 2.0 - self._fase
        self._dibujar(mezclar(TEMA["acento_suave"], TEMA["primario"], t))
        self._anim_id = self.after(50, self._pulso)

    def procesando(self, activo: bool):
        self._pulsando = activo
        if activo and self._anim_id is None:
            self._pulso()
        elif not activo:
            if self._anim_id:
                self.after_cancel(self._anim_id)
                self._anim_id = None
            self._dibujar(TEMA["exito"])


# ---------------------------------------------------------------------------
# ESTILO TTK (Combobox, Progressbar)
# ---------------------------------------------------------------------------
def aplicar_estilo_ttk(root):
    estilo = ttk.Style(root)
    try:
        estilo.theme_use("clam")
    except tk.TclError:
        pass
    estilo.configure(
        "Moderno.TCombobox",
        fieldbackground=TEMA["panel"],
        background=TEMA["panel"],
        bordercolor=TEMA["borde"],
        arrowcolor=TEMA["texto_sec"],
        lightcolor=TEMA["panel"],
        darkcolor=TEMA["panel"],
        relief="flat",
        padding=6,
    )
    estilo.map("Moderno.TCombobox",
               bordercolor=[("focus", TEMA["acento"])],
               lightcolor=[("focus", TEMA["panel"])])
    estilo.configure(
        "Moderno.Horizontal.TProgressbar",
        troughcolor=TEMA["acento_suave"],
        background=TEMA["primario"],
        bordercolor=TEMA["acento_suave"],
        lightcolor=TEMA["primario"],
        darkcolor=TEMA["primario"],
        thickness=4,
    )
    return estilo
