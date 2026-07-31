"""
carta_writer.py
===============
Genera una carta Word (.docx) por vecino usando Plantilla_Cartas.docx
y los datos de la tabla 'repartos' de la BD SQLite.

DISEÑO 2026 — moderno, visual y completo:
  - Cabecera corporativa índigo con nombre de comunidad y periodo
  - Caja de RESULTADO destacada al principio (lo primero que ve el cliente)
  - Bloque de datos del propietario completo (vivienda, coeficiente, fechas)
  - Tablas de suministro con lecturas, consumo, precio medio y diferencia
  - Gráfica doble: barras cobrado vs real + donut de distribución del coste
  - Gráfica de consumo del vecino frente a la media de la comunidad
  - Bloque "¿Cómo se ha calculado?" en lenguaje claro
  - Pie con instrucciones según haya que pagar o devolver

USO:
    python carta_writer.py --bd gestion.db --comunidad 644 --periodo 2024-2025
                           --plantilla plantillas/Plantilla_Cartas.docx
                           --salida salidas/cartas/
"""

import os
import re
import sqlite3
import argparse
import tempfile
from datetime import datetime
from pathlib import Path
from typing import Optional

from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

try:
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    import matplotlib.patches as mpatches
    import numpy as np
    _HAS_MPL = True
except ImportError:
    _HAS_MPL = False


# ---------------------------------------------------------------------------
# PALETA MODERNA (índigo corporativo + semáforo suave)
# ---------------------------------------------------------------------------

C_PRIMARIO       = "#312E81"   # índigo profundo (cabeceras)
C_PRIMARIO_MED   = "#4F46E5"   # índigo vivo (acentos, barras "real")
C_PRIMARIO_SUAVE = "#EEF2FF"   # fondo suave de cabeceras de columna
C_ACENTO_2       = "#0D9488"   # teal (ACS en el donut)
C_TEXTO_HEADER   = "#E0E7FF"
C_GRIS_ETQ       = "#6B7280"
C_DARK           = "#111827"
C_ZEBRA          = "#F8FAFC"
C_VERDE          = "#047857"
C_ROJO           = "#B91C1C"
C_FONDO_PAGAR    = "#FEE2E2"
C_FONDO_DEVOLVER = "#D1FAE5"
C_FONDO_CERO     = "#EEF2FF"
C_BARRA_COB      = "#C7D2FE"   # índigo claro (cobrado)
C_BARRA_REAL     = "#4F46E5"   # índigo (real)
C_BARRA_MEDIA    = "#CBD5E1"   # gris (media comunidad)
C_FONDO_GRAF     = "#FFFFFF"
FUENTE_CARTA     = "Segoe UI"

MESES_ES = {
    1: "enero",    2: "febrero",  3: "marzo",    4: "abril",
    5: "mayo",     6: "junio",    7: "julio",    8: "agosto",
    9: "septiembre", 10: "octubre", 11: "noviembre", 12: "diciembre",
}


# ---------------------------------------------------------------------------
# LECTURA DE DATOS DESDE LA BD
# ---------------------------------------------------------------------------

def _obtener_repartos_vecino(con: sqlite3.Connection,
                              id_propietario: int,
                              id_periodo: int) -> dict:
    """Devuelve todos los datos necesarios para la carta de un vecino."""
    cur = con.cursor()

    vecino = cur.execute("""
        SELECT p.nombre_propietario, p.codigo_vivienda, p.coeficiente,
               per.nombre AS periodo_nombre, per.fecha_inicio, per.fecha_fin
        FROM propietarios p
        JOIN periodos per ON per.id_periodo = ?
        WHERE p.id_propietario = ?
    """, (id_periodo, id_propietario)).fetchone()

    if not vecino:
        return None
    v = dict(vecino)

    repartos = cur.execute("""
        SELECT tipo_suministro, lectura_inicial, lectura_final, consumo_real,
               importe_cobrado, importe_real, diferencia, notas
        FROM repartos
        WHERE id_propietario = ? AND id_periodo = ?
        ORDER BY tipo_suministro
    """, (id_propietario, id_periodo)).fetchall()

    datos = {
        "vecino": {
            "nombre":      v["nombre_propietario"],
            "vivienda":    v["codigo_vivienda"],
            "coeficiente": v["coeficiente"],
        },
        "periodo": {
            "nombre":       v["periodo_nombre"],
            "fecha_inicio": v["fecha_inicio"],
            "fecha_fin":    v["fecha_fin"],
        },
        "calef":  None,
        "acs":    None,
        "extras": [],
        "total":  {"cobrado": 0.0, "real": 0.0, "diferencia": 0.0},
    }

    for r in repartos:
        rd = dict(r)
        if rd["tipo_suministro"] == "CALEFACCION":
            datos["calef"] = rd
        elif rd["tipo_suministro"] == "ACS":
            datos["acs"] = rd
        datos["total"]["cobrado"]    += rd["importe_cobrado"]
        datos["total"]["real"]       += rd["importe_real"]
        datos["total"]["diferencia"] += rd["diferencia"]

    for k in datos["total"]:
        datos["total"][k] = round(datos["total"][k], 2)

    return datos


def obtener_todos_los_vecinos_con_repartos(con: sqlite3.Connection,
                                            id_comunidad: int,
                                            id_periodo: int) -> list:
    """Devuelve la lista completa de vecinos con sus repartos calculados."""
    cur = con.cursor()
    vecinos = cur.execute("""
        SELECT DISTINCT p.id_propietario
        FROM repartos r
        JOIN propietarios p ON p.id_propietario = r.id_propietario
        WHERE p.id_comunidad = ? AND r.id_periodo = ?
        ORDER BY p.codigo_vivienda
    """, (id_comunidad, id_periodo)).fetchall()

    resultado = []
    for (id_prop,) in vecinos:
        datos = _obtener_repartos_vecino(con, id_prop, id_periodo)
        if datos:
            resultado.append(datos)
    return resultado


def _calcular_medias(todos: list) -> dict:
    """Media de consumo e importe real por suministro sobre todos los vecinos."""
    medias = {"calef_consumo": 0.0, "acs_consumo": 0.0,
              "calef_importe": 0.0, "acs_importe": 0.0,
              "n_vecinos": len(todos)}
    if not todos:
        return medias
    n = len(todos)
    for d in todos:
        if d.get("calef"):
            medias["calef_consumo"] += d["calef"].get("consumo_real") or 0
            medias["calef_importe"] += d["calef"].get("importe_real") or 0
        if d.get("acs"):
            medias["acs_consumo"] += d["acs"].get("consumo_real") or 0
            medias["acs_importe"] += d["acs"].get("importe_real") or 0
    for k in ("calef_consumo", "acs_consumo", "calef_importe", "acs_importe"):
        medias[k] = round(medias[k] / n, 2)
    return medias


# ---------------------------------------------------------------------------
# FORMATEADO DE NÚMEROS Y FECHAS
# ---------------------------------------------------------------------------

def _fmt_euros(valor: float) -> str:
    if valor is None:
        return "0,00 €"
    return f"{valor:,.2f} €".replace(",", "X").replace(".", ",").replace("X", ".")


def _fmt_consumo(valor: float, unidad: str) -> str:
    if unidad in ("m3", "m\u00b3"):
        unidad = "m\u00b3"
    if valor is None or valor == 0:
        return f"0 {unidad}"
    if unidad == "kWh":
        return f"{valor:,.0f} kWh".replace(",", ".")
    txt = f"{valor:.1f}".rstrip("0").rstrip(".").replace(".", ",")
    return f"{txt} {unidad}"


def _fmt_fecha_es(iso: str) -> str:
    """'2024-09-01' → '1 de septiembre de 2024'"""
    try:
        d = datetime.fromisoformat(iso[:10])
        return f"{d.day} de {MESES_ES[d.month]} de {d.year}"
    except Exception:
        return iso or "—"


# ---------------------------------------------------------------------------
# HELPERS DE ESTILO DOCX
# ---------------------------------------------------------------------------

def _rgb(hex_str: str) -> RGBColor:
    h = hex_str.lstrip("#")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def _set_color_celda(celda, hex_color: str) -> None:
    tc = celda._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color.lstrip("#"))
    tcPr.append(shd)


def _quitar_margen_parrafo(parrafo, despues: int = 6) -> None:
    pf = parrafo.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after  = Pt(despues)


def _set_ancho_col(tabla, col_idx: int, ancho_cm: float) -> None:
    dxa = int(ancho_cm * 567)
    for celda in tabla.columns[col_idx].cells:
        tc = celda._tc
        tcPr = tc.get_or_add_tcPr()
        for old in tcPr.findall(qn("w:tcW")):
            tcPr.remove(old)
        tcW = OxmlElement("w:tcW")
        tcW.set(qn("w:w"),    str(dxa))
        tcW.set(qn("w:type"), "dxa")
        tcPr.append(tcW)


def _sin_borde_tabla(tabla) -> None:
    tbl   = tabla._tbl
    tblPr = tbl.find(qn("w:tblPr"))
    if tblPr is None:
        tblPr = OxmlElement("w:tblPr")
        tbl.insert(0, tblPr)
    for old in tblPr.findall(qn("w:tblBorders")):
        tblPr.remove(old)
    tblBorders = OxmlElement("w:tblBorders")
    for lado in ("top", "left", "bottom", "right", "insideH", "insideV"):
        borde = OxmlElement(f"w:{lado}")
        borde.set(qn("w:val"),   "none")
        borde.set(qn("w:sz"),    "0")
        borde.set(qn("w:space"), "0")
        borde.set(qn("w:color"), "auto")
        tblBorders.append(borde)
    tblPr.append(tblBorders)


def _escribir_celda(celda, texto: str,
                    negrita: bool = False,
                    tamaño: float = 10,
                    color_hex: str = None,
                    alineacion=WD_ALIGN_PARAGRAPH.LEFT,
                    despues: int = 6) -> None:
    for p in celda.paragraphs:
        p._element.getparent().remove(p._element)
    p = celda.add_paragraph()
    _quitar_margen_parrafo(p, despues)
    p.alignment = alineacion
    run = p.add_run(texto)
    run.bold      = negrita
    run.font.size = Pt(tamaño)
    run.font.name = FUENTE_CARTA
    if color_hex:
        run.font.color.rgb = _rgb(color_hex)


def _parrafo(doc, texto: str, tam: float = 9, color: str = C_DARK,
             negrita: bool = False, alin=WD_ALIGN_PARAGRAPH.JUSTIFY,
             despues: int = 6):
    p = doc.add_paragraph()
    _quitar_margen_parrafo(p, despues)
    p.alignment = alin
    run = p.add_run(texto)
    run.font.size = Pt(tam)
    run.font.name = FUENTE_CARTA
    run.bold = negrita
    run.font.color.rgb = _rgb(color)
    return p


# ---------------------------------------------------------------------------
# BLOQUES DE DISEÑO
# ---------------------------------------------------------------------------

def _bloque_cabecera(doc: Document, nombre_comunidad: str,
                     nombre_periodo: str) -> None:
    """Cabecera índigo: título + comunidad + periodo."""
    tabla = doc.add_table(rows=2, cols=1)
    _sin_borde_tabla(tabla)
    _set_ancho_col(tabla, 0, 16.0)

    c1 = tabla.rows[0].cells[0]
    _set_color_celda(c1, C_PRIMARIO)
    _escribir_celda(c1, "REGULARIZACIÓN DE CONSUMOS  ·  CALEFACCIÓN Y ACS",
                    negrita=True, tamaño=14, color_hex="#FFFFFF",
                    alineacion=WD_ALIGN_PARAGRAPH.CENTER, despues=2)

    c2 = tabla.rows[1].cells[0]
    _set_color_celda(c2, C_PRIMARIO_MED)
    _escribir_celda(c2,
                    f"{nombre_comunidad.upper()}   ·   PERIODO {nombre_periodo}   ·   Meditrade Administración",
                    negrita=False, tamaño=9, color_hex=C_TEXTO_HEADER,
                    alineacion=WD_ALIGN_PARAGRAPH.CENTER, despues=2)

    p = doc.add_paragraph()
    _quitar_margen_parrafo(p, 4)


def _caja_resultado(doc: Document, diferencia_total: float) -> None:
    """
    Caja de RESULTADO al principio: lo primero que ve el cliente.
    Explica en una frase qué significa el número.
    """
    if diferencia_total > 0.005:
        fondo, color_v = C_FONDO_PAGAR, C_ROJO
        titulo  = "RESULTADO: IMPORTE A PAGAR"
        importe = f"+ {_fmt_euros(abs(diferencia_total))}"
        explica = ("Su consumo real ha sido superior a lo cobrado a cuenta durante el año. "
                   "La diferencia se cargará en su próximo recibo de comunidad.")
    elif diferencia_total < -0.005:
        fondo, color_v = C_FONDO_DEVOLVER, C_VERDE
        titulo  = "RESULTADO: IMPORTE A SU FAVOR"
        importe = f"− {_fmt_euros(abs(diferencia_total))}"
        explica = ("Ha pagado a cuenta más de lo que realmente ha consumido. "
                   "La diferencia se abonará en su próximo recibo de comunidad.")
    else:
        fondo, color_v = C_FONDO_CERO, C_PRIMARIO
        titulo  = "RESULTADO: SIN DIFERENCIA"
        importe = _fmt_euros(0)
        explica = "Lo cobrado a cuenta coincide con su consumo real. No hay regularización."

    tabla = doc.add_table(rows=2, cols=2)
    _sin_borde_tabla(tabla)
    _set_ancho_col(tabla, 0, 10.0)
    _set_ancho_col(tabla, 1, 6.0)

    c_tit = tabla.rows[0].cells[0]
    _set_color_celda(c_tit, fondo)
    _escribir_celda(c_tit, titulo, negrita=True, tamaño=12,
                    color_hex=C_DARK, despues=0)

    c_imp = tabla.rows[0].cells[1]
    _set_color_celda(c_imp, fondo)
    _escribir_celda(c_imp, importe, negrita=True, tamaño=16,
                    color_hex=color_v,
                    alineacion=WD_ALIGN_PARAGRAPH.RIGHT, despues=0)

    c_exp = tabla.rows[1].cells[0].merge(tabla.rows[1].cells[1])
    _set_color_celda(c_exp, fondo)
    _escribir_celda(c_exp, explica, negrita=False, tamaño=9,
                    color_hex=C_GRIS_ETQ, despues=2)

    p = doc.add_paragraph()
    _quitar_margen_parrafo(p, 4)


def _bloque_info(doc: Document, fecha_str: str, vivienda: str,
                 propietario: str, coeficiente, periodo: str,
                 fecha_ini: str, fecha_fin: str, n_vecinos: int) -> None:
    """Ficha del propietario en dos columnas de etiqueta|valor."""
    coef_txt = (f"{coeficiente:.4f}".rstrip("0").rstrip(".").replace(".", ",") + " %"
                if isinstance(coeficiente, (int, float)) and coeficiente else "—")
    pares = [
        ("Propietario/a", propietario,             "Fecha",   fecha_str),
        ("Vivienda",      vivienda,                "Periodo", periodo),
        ("Coeficiente",   coef_txt,                "Desde",   _fmt_fecha_es(fecha_ini)),
        ("Vecinos en el reparto", str(n_vecinos or "—"), "Hasta", _fmt_fecha_es(fecha_fin)),
    ]

    tabla = doc.add_table(rows=len(pares), cols=4)
    _sin_borde_tabla(tabla)
    for idx, ancho in enumerate([3.2, 5.8, 2.4, 4.6]):
        _set_ancho_col(tabla, idx, ancho)

    for i, (etq1, val1, etq2, val2) in enumerate(pares):
        fila = tabla.rows[i]
        if i % 2 == 1:
            for c in fila.cells:
                _set_color_celda(c, C_ZEBRA)
        _escribir_celda(fila.cells[0], etq1, tamaño=8.5, color_hex=C_GRIS_ETQ, despues=2)
        _escribir_celda(fila.cells[1], str(val1), negrita=True, tamaño=9,
                        color_hex=C_DARK, despues=2)
        _escribir_celda(fila.cells[2], etq2, tamaño=8.5, color_hex=C_GRIS_ETQ, despues=2)
        _escribir_celda(fila.cells[3], str(val2), negrita=True, tamaño=9,
                        color_hex=C_DARK, despues=2)

    p = doc.add_paragraph()
    _quitar_margen_parrafo(p, 4)


def _seccion_titulo(doc: Document, icono_texto: str) -> None:
    tabla = doc.add_table(rows=1, cols=1)
    _sin_borde_tabla(tabla)
    _set_ancho_col(tabla, 0, 16.0)
    c = tabla.rows[0].cells[0]
    _set_color_celda(c, C_PRIMARIO)
    _escribir_celda(c, icono_texto, negrita=True, tamaño=10,
                    color_hex="#FFFFFF", despues=2)


def _tabla_suministro(doc: Document, titulo: str,
                      lectura_ini, lectura_fin, consumo, unidad: str,
                      importe_cobrado: float, importe_real: float,
                      diferencia: float) -> None:
    """
    Bloque de un suministro:
      lecturas, consumo, precio medio, cobrado vs real y diferencia.
    """
    _seccion_titulo(doc, titulo)

    precio_medio = (importe_real / consumo) if consumo else None
    precio_txt   = (f"{precio_medio:.4f} €/{unidad}".replace(".", ",")
                    if precio_medio else "—")

    if diferencia < -0.005:
        dif_txt, dif_color = f"− {_fmt_euros(abs(diferencia))}", C_VERDE
    elif diferencia > 0.005:
        dif_txt, dif_color = f"+ {_fmt_euros(diferencia)}", C_ROJO
    else:
        dif_txt, dif_color = "0,00 €", C_DARK

    filas = [
        ("Lectura inicial del contador", _fmt_consumo(lectura_ini, unidad), None),
        ("Lectura final del contador",   _fmt_consumo(lectura_fin, unidad), None),
        ("Consumo del periodo",          _fmt_consumo(consumo, unidad),     None),
        ("Precio medio aplicado",        precio_txt,                        None),
        ("Cobrado a cuenta (recibos)",   _fmt_euros(importe_cobrado),       None),
        ("Coste real de su consumo",     _fmt_euros(importe_real),          None),
        ("DIFERENCIA",                   dif_txt,                           dif_color),
    ]

    tabla = doc.add_table(rows=len(filas), cols=2)
    _sin_borde_tabla(tabla)
    _set_ancho_col(tabla, 0, 10.5)
    _set_ancho_col(tabla, 1, 5.5)

    for i, (concepto, valor, color_valor) in enumerate(filas):
        fila = tabla.rows[i]
        es_dif = concepto == "DIFERENCIA"
        if es_dif:
            for c in fila.cells:
                _set_color_celda(c, C_PRIMARIO_SUAVE)
        elif i % 2 == 1:
            for c in fila.cells:
                _set_color_celda(c, C_ZEBRA)
        _escribir_celda(fila.cells[0], concepto,
                        negrita=es_dif, tamaño=9,
                        color_hex=C_PRIMARIO if es_dif else C_DARK, despues=2)
        _escribir_celda(fila.cells[1], valor,
                        negrita=es_dif or "Coste real" in concepto, tamaño=9.5,
                        color_hex=color_valor or C_DARK,
                        alineacion=WD_ALIGN_PARAGRAPH.RIGHT, despues=2)

    p = doc.add_paragraph()
    _quitar_margen_parrafo(p, 4)


# ---------------------------------------------------------------------------
# GRÁFICAS (matplotlib)
# ---------------------------------------------------------------------------

def _estilo_ejes(ax):
    for lado in ("top", "right", "left"):
        ax.spines[lado].set_visible(False)
    ax.spines["bottom"].set_color("#E5E7EB")
    ax.tick_params(colors=C_GRIS_ETQ, labelsize=7.5)
    ax.set_facecolor(C_FONDO_GRAF)


def _grafica_resumen(calef_cob, calef_real, acs_cob, acs_real) -> Optional[str]:
    """Figura doble: barras cobrado vs real + donut de distribución del coste."""
    if not _HAS_MPL:
        return None
    try:
        fig, (ax1, ax2) = plt.subplots(
            1, 2, figsize=(16 / 2.54, 5.6 / 2.54), dpi=150,
            gridspec_kw={"width_ratios": [1.55, 1.0]})
        fig.patch.set_facecolor(C_FONDO_GRAF)

        # ── Barras agrupadas ──
        grupos   = ["Calefacción", "ACS"]
        cobrados = [calef_cob or 0, acs_cob or 0]
        reales   = [calef_real or 0, acs_real or 0]
        x = np.arange(len(grupos))
        ancho = 0.36
        b1 = ax1.bar(x - ancho / 2, cobrados, ancho, color=C_BARRA_COB,
                     label="Cobrado a cuenta", zorder=3)
        b2 = ax1.bar(x + ancho / 2, reales, ancho, color=C_BARRA_REAL,
                     label="Coste real", zorder=3)
        tope = max(cobrados + reales) or 1
        for barras in (b1, b2):
            for bar in barras:
                v = bar.get_height()
                if v > 0:
                    ax1.text(bar.get_x() + bar.get_width() / 2, v + tope * 0.03,
                             _fmt_euros(v).replace(" €", "€"),
                             ha="center", va="bottom", fontsize=6.8, color=C_DARK)
        ax1.set_xticks(x)
        ax1.set_xticklabels(grupos, fontsize=8.5, color=C_DARK)
        ax1.set_yticks([])
        ax1.set_ylim(0, tope * 1.22)
        _estilo_ejes(ax1)
        ax1.legend(fontsize=6.8, frameon=False, loc="upper center",
                   bbox_to_anchor=(0.5, 1.16), ncol=2)
        ax1.set_title("Cobrado vs coste real", fontsize=8, color=C_GRIS_ETQ, pad=16)
        ax1.grid(axis="y", color="#F1F5F9", zorder=0)

        # ── Donut distribución del coste real ──
        vals = [max(calef_real or 0, 0), max(acs_real or 0, 0)]
        if sum(vals) > 0:
            colores = [C_PRIMARIO_MED, C_ACENTO_2]
            wedges, _txt, autos = ax2.pie(
                vals, colors=colores, startangle=90,
                autopct=lambda p: f"{p:.0f}%" if p > 4 else "",
                pctdistance=0.75,
                wedgeprops=dict(width=0.42, edgecolor="white"))
            for a in autos:
                a.set_fontsize(7)
                a.set_color("white")
                a.set_fontweight("bold")
            ax2.text(0, 0.06, _fmt_euros(sum(vals)).replace(" €", " €"),
                     ha="center", va="center", fontsize=8.2,
                     color=C_DARK, fontweight="bold")
            ax2.text(0, -0.22, "coste real total", ha="center", va="center",
                     fontsize=6, color=C_GRIS_ETQ)
            ax2.legend(wedges, ["Calefacción", "ACS"], fontsize=6.8,
                       frameon=False, loc="upper center",
                       bbox_to_anchor=(0.5, 1.18), ncol=2)
            ax2.set_title("Distribución de su coste", fontsize=8,
                          color=C_GRIS_ETQ, pad=16)
        else:
            ax2.axis("off")

        plt.tight_layout(pad=0.6)
        tmp = tempfile.mktemp(suffix=".png")
        fig.savefig(tmp, dpi=150, bbox_inches="tight", facecolor=C_FONDO_GRAF)
        plt.close(fig)
        return tmp
    except Exception:
        return None


def _grafica_vs_media(calef_consumo, acs_consumo, medias: dict) -> Optional[str]:
    """Barras horizontales: consumo del vecino frente a la media de la comunidad."""
    if not _HAS_MPL or not medias:
        return None
    try:
        tiene_calef = (calef_consumo or 0) > 0 or (medias.get("calef_consumo") or 0) > 0
        tiene_acs   = (acs_consumo or 0) > 0 or (medias.get("acs_consumo") or 0) > 0
        if not tiene_calef and not tiene_acs:
            return None

        n_planos = int(tiene_calef) + int(tiene_acs)
        fig, ejes = plt.subplots(1, n_planos,
                                 figsize=(16 / 2.54, 4.2 / 2.54), dpi=150)
        fig.patch.set_facecolor(C_FONDO_GRAF)
        if n_planos == 1:
            ejes = [ejes]

        def _plano(ax, titulo, propio, media, unidad):
            etiquetas = ["Usted", "Media\ncomunidad"]
            valores   = [propio or 0, media or 0]
            colores   = [C_BARRA_REAL, C_BARRA_MEDIA]
            barras = ax.barh(etiquetas, valores, 0.52, color=colores, zorder=3)
            tope = max(valores) or 1
            for bar, v in zip(barras, valores):
                ax.text(v + tope * 0.02, bar.get_y() + bar.get_height() / 2,
                        _fmt_consumo(v, unidad),
                        va="center", ha="left", fontsize=7, color=C_DARK)
            ax.set_xlim(0, tope * 1.3)
            ax.set_xticks([])
            ax.invert_yaxis()
            _estilo_ejes(ax)
            ax.spines["bottom"].set_visible(False)
            ax.set_title(titulo, fontsize=8, color=C_GRIS_ETQ, pad=6)
            ax.grid(False)

        idx = 0
        if tiene_calef:
            _plano(ejes[idx], "Consumo de calefacción",
                   calef_consumo, medias.get("calef_consumo"), "kWh")
            idx += 1
        if tiene_acs:
            _plano(ejes[idx], "Consumo de ACS",
                   acs_consumo, medias.get("acs_consumo"), "m3")

        plt.tight_layout(pad=0.6)
        tmp = tempfile.mktemp(suffix=".png")
        fig.savefig(tmp, dpi=150, bbox_inches="tight", facecolor=C_FONDO_GRAF)
        plt.close(fig)
        return tmp
    except Exception:
        return None


def _insertar_imagen(doc: Document, ruta_png: str, ancho_cm: float = 16.0):
    if ruta_png and os.path.exists(ruta_png):
        try:
            p = doc.add_paragraph()
            _quitar_margen_parrafo(p, 4)
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.add_run().add_picture(ruta_png, width=Cm(ancho_cm))
        finally:
            try:
                os.remove(ruta_png)
            except OSError:
                pass
        return True
    return False


def _bloque_comparativa_fallback(doc, calef_cob, calef_real, acs_cob, acs_real):
    """Tabla resumen si matplotlib no está disponible."""
    tabla = doc.add_table(rows=3, cols=4)
    _sin_borde_tabla(tabla)
    for idx in range(4):
        _set_ancho_col(tabla, idx, 4.0)
    cabeceras = ["", "Cobrado", "Real", "Diferencia"]
    for j, txt in enumerate(cabeceras):
        c = tabla.rows[0].cells[j]
        _set_color_celda(c, C_PRIMARIO_SUAVE)
        _escribir_celda(c, txt, negrita=True, tamaño=9, color_hex=C_PRIMARIO,
                        alineacion=WD_ALIGN_PARAGRAPH.RIGHT if j else WD_ALIGN_PARAGRAPH.LEFT,
                        despues=2)
    for i, (nom, cob, real) in enumerate(
            [("Calefacción", calef_cob, calef_real), ("ACS", acs_cob, acs_real)]):
        dif = (real or 0) - (cob or 0)
        fila = tabla.rows[i + 1]
        color = C_VERDE if dif < -0.005 else (C_ROJO if dif > 0.005 else C_DARK)
        signo = "− " if dif < -0.005 else ("+ " if dif > 0.005 else "")
        _escribir_celda(fila.cells[0], nom, tamaño=9, color_hex=C_DARK, despues=2)
        _escribir_celda(fila.cells[1], _fmt_euros(cob), tamaño=9, color_hex=C_DARK,
                        alineacion=WD_ALIGN_PARAGRAPH.RIGHT, despues=2)
        _escribir_celda(fila.cells[2], _fmt_euros(real), tamaño=9, color_hex=C_DARK,
                        alineacion=WD_ALIGN_PARAGRAPH.RIGHT, despues=2)
        _escribir_celda(fila.cells[3], signo + _fmt_euros(abs(dif)), tamaño=9,
                        color_hex=color, alineacion=WD_ALIGN_PARAGRAPH.RIGHT, despues=2)


def _meses_periodo(fecha_ini: str, fecha_fin: str) -> int:
    """Número de meses (redondeado) entre dos fechas ISO."""
    try:
        d1 = datetime.fromisoformat(fecha_ini[:10])
        d2 = datetime.fromisoformat(fecha_fin[:10])
        return max(1, round((d2 - d1).days / 30.44))
    except Exception:
        return 12


def _tabla_cuota_fija(doc: Document, cuota_fija: dict) -> None:
    """
    Sección CUOTA FIJA: lo cobrado mensualmente frente al coste fijo real.
    cuota_fija = {"CALEFACCION": {...}, "ACS": {...}} con
                 precio_cobrado, precio_real, meses, cobrado, real, diferencia
    """
    if not cuota_fija:
        return
    _seccion_titulo(doc, "CUOTA FIJA")

    nombres = {"CALEFACCION": "Cuota fija calefacción",
               "ACS":         "Cuota fija ACS"}
    tabla = doc.add_table(rows=len(cuota_fija) + 2, cols=4)
    _sin_borde_tabla(tabla)
    for idx, ancho in enumerate([7.0, 3.0, 3.0, 3.0]):
        _set_ancho_col(tabla, idx, ancho)

    # Cabecera
    for j, txt in enumerate(["Concepto", "Cobrado", "Real", "Diferencia"]):
        c = tabla.rows[0].cells[j]
        _set_color_celda(c, C_PRIMARIO_SUAVE)
        _escribir_celda(c, txt, negrita=True, tamaño=9, color_hex=C_PRIMARIO,
                        alineacion=WD_ALIGN_PARAGRAPH.RIGHT if j else WD_ALIGN_PARAGRAPH.LEFT,
                        despues=2)

    total_cob = total_real = total_dif = 0.0
    for i, (tipo, cf) in enumerate(sorted(cuota_fija.items())):
        fila = tabla.rows[i + 1]
        if i % 2 == 1:
            for c in fila.cells:
                _set_color_celda(c, C_ZEBRA)
        precio_txt = f"{cf['precio_cobrado']:.2f}".replace(".", ",")
        concepto = (f"{nombres.get(tipo, tipo)}  "
                    f"({precio_txt} €/mes × {cf['meses']} meses)")
        dif = cf["diferencia"]
        color = C_VERDE if dif < -0.005 else (C_ROJO if dif > 0.005 else C_DARK)
        signo = "− " if dif < -0.005 else ("+ " if dif > 0.005 else "")
        _escribir_celda(fila.cells[0], concepto, tamaño=9, color_hex=C_DARK, despues=2)
        _escribir_celda(fila.cells[1], _fmt_euros(cf["cobrado"]), tamaño=9,
                        color_hex=C_DARK, alineacion=WD_ALIGN_PARAGRAPH.RIGHT, despues=2)
        _escribir_celda(fila.cells[2], _fmt_euros(cf["real"]), tamaño=9,
                        color_hex=C_DARK, alineacion=WD_ALIGN_PARAGRAPH.RIGHT, despues=2)
        _escribir_celda(fila.cells[3], signo + _fmt_euros(abs(dif)), tamaño=9,
                        color_hex=color, alineacion=WD_ALIGN_PARAGRAPH.RIGHT, despues=2)
        total_cob += cf["cobrado"]; total_real += cf["real"]; total_dif += dif

    # Fila total
    fila = tabla.rows[len(cuota_fija) + 1]
    for c in fila.cells:
        _set_color_celda(c, C_PRIMARIO_SUAVE)
    color = C_VERDE if total_dif < -0.005 else (C_ROJO if total_dif > 0.005 else C_DARK)
    signo = "− " if total_dif < -0.005 else ("+ " if total_dif > 0.005 else "")
    _escribir_celda(fila.cells[0], "TOTAL CUOTA FIJA", negrita=True, tamaño=9,
                    color_hex=C_PRIMARIO, despues=2)
    _escribir_celda(fila.cells[1], _fmt_euros(total_cob), negrita=True, tamaño=9,
                    color_hex=C_DARK, alineacion=WD_ALIGN_PARAGRAPH.RIGHT, despues=2)
    _escribir_celda(fila.cells[2], _fmt_euros(total_real), negrita=True, tamaño=9,
                    color_hex=C_DARK, alineacion=WD_ALIGN_PARAGRAPH.RIGHT, despues=2)
    _escribir_celda(fila.cells[3], signo + _fmt_euros(abs(total_dif)), negrita=True,
                    tamaño=9, color_hex=color,
                    alineacion=WD_ALIGN_PARAGRAPH.RIGHT, despues=2)

    p = doc.add_paragraph()
    _quitar_margen_parrafo(p, 4)


def _tabla_totales(doc: Document, dif_fija: float, dif_calef: float,
                   dif_acs: float, total_general: float) -> None:
    """Resumen de totales por concepto + total general."""
    _seccion_titulo(doc, "RESUMEN DE LA REGULARIZACIÓN")
    conceptos = [("Cuota fija", dif_fija), ("Calefacción", dif_calef),
                 ("ACS", dif_acs)]
    conceptos = [(n, v) for n, v in conceptos if abs(v) > 0.004 or n != "Cuota fija"]

    tabla = doc.add_table(rows=len(conceptos) + 1, cols=2)
    _sin_borde_tabla(tabla)
    _set_ancho_col(tabla, 0, 10.5)
    _set_ancho_col(tabla, 1, 5.5)

    for i, (nombre, val) in enumerate(conceptos):
        fila = tabla.rows[i]
        if i % 2 == 1:
            for c in fila.cells:
                _set_color_celda(c, C_ZEBRA)
        color = C_VERDE if val < -0.005 else (C_ROJO if val > 0.005 else C_DARK)
        signo = "− " if val < -0.005 else ("+ " if val > 0.005 else "")
        _escribir_celda(fila.cells[0], f"Diferencia {nombre.lower()}",
                        tamaño=9, color_hex=C_DARK, despues=2)
        _escribir_celda(fila.cells[1], signo + _fmt_euros(abs(val)), tamaño=9,
                        color_hex=color, alineacion=WD_ALIGN_PARAGRAPH.RIGHT, despues=2)

    fila = tabla.rows[len(conceptos)]
    fondo = (C_FONDO_PAGAR if total_general > 0.005
             else C_FONDO_DEVOLVER if total_general < -0.005 else C_FONDO_CERO)
    color = C_ROJO if total_general > 0.005 else (C_VERDE if total_general < -0.005 else C_PRIMARIO)
    signo = "+ " if total_general > 0.005 else ("− " if total_general < -0.005 else "")
    for c in fila.cells:
        _set_color_celda(c, fondo)
    _escribir_celda(fila.cells[0], "IMPORTE TOTAL DE LA REGULARIZACIÓN",
                    negrita=True, tamaño=11, color_hex=C_DARK, despues=2)
    _escribir_celda(fila.cells[1], signo + _fmt_euros(abs(total_general)),
                    negrita=True, tamaño=12, color_hex=color,
                    alineacion=WD_ALIGN_PARAGRAPH.RIGHT, despues=2)

    p = doc.add_paragraph()
    _quitar_margen_parrafo(p, 4)


# ---------------------------------------------------------------------------
# GENERACIÓN DE UNA CARTA
# ---------------------------------------------------------------------------

def generar_carta(datos: dict, ruta_plantilla: str, ruta_salida: str) -> str:
    """Genera una carta Word moderna para un vecino."""
    import shutil
    shutil.copy2(ruta_plantilla, ruta_salida)
    doc = Document(ruta_salida)

    # Limpiar el cuerpo preservando el header con logo de Meditrade
    for par in list(doc.paragraphs):
        par._element.getparent().remove(par._element)
    for tabla in list(doc.tables):
        tabla._element.getparent().remove(tabla._element)

    vecino  = datos["vecino"]
    periodo = datos["periodo"]
    calef   = datos.get("calef") or {}
    acs     = datos.get("acs")   or {}
    total   = datos["total"]
    medias  = datos.get("medias") or {}
    cuota_fija = datos.get("cuota_fija") or {}
    dif_fija   = round(sum(cf["diferencia"] for cf in cuota_fija.values()), 2)
    total_general = round(total["diferencia"] + dif_fija, 2)
    nombre_comunidad = datos.get("comunidad", "COMUNIDAD DE PROPIETARIOS")

    nombre_periodo = periodo["nombre"]
    hoy       = datetime.now()
    fecha_str = f"Zaragoza, {hoy.day} de {MESES_ES[hoy.month]} de {hoy.year}"

    # 1. CABECERA
    _bloque_cabecera(doc, nombre_comunidad, nombre_periodo)

    # 2. RESULTADO DESTACADO (lo más importante, primero)
    _caja_resultado(doc, total_general)

    # 3. FICHA DEL PROPIETARIO
    _bloque_info(doc,
                 fecha_str   = fecha_str,
                 vivienda    = vecino["vivienda"],
                 propietario = vecino["nombre"],
                 coeficiente = vecino.get("coeficiente"),
                 periodo     = nombre_periodo,
                 fecha_ini   = periodo.get("fecha_inicio") or "",
                 fecha_fin   = periodo.get("fecha_fin") or "",
                 n_vecinos   = medias.get("n_vecinos", 0))

    # 4. INTRO BREVE
    _parrafo(doc,
             "En cumplimiento del acuerdo adoptado en la última Junta General, le remitimos "
             f"el detalle de la regularización de Calefacción y Agua Caliente Sanitaria (ACS) del periodo {nombre_periodo}: "
             "lecturas de su contador, consumo, coste real e importes cobrados a cuenta.",
             tam=9, despues=8)

    # 5. DETALLE POR SUMINISTRO
    if calef:
        _tabla_suministro(doc, "CALEFACCIÓN",
                          calef.get("lectura_inicial", 0), calef.get("lectura_final", 0),
                          calef.get("consumo_real", 0), "kWh",
                          calef.get("importe_cobrado", 0.0), calef.get("importe_real", 0.0),
                          calef.get("diferencia", 0.0))
    if acs:
        _tabla_suministro(doc, "AGUA CALIENTE SANITARIA (ACS)",
                          acs.get("lectura_inicial", 0), acs.get("lectura_final", 0),
                          acs.get("consumo_real", 0), "m3",
                          acs.get("importe_cobrado", 0.0), acs.get("importe_real", 0.0),
                          acs.get("diferencia", 0.0))

    # 5b. CUOTA FIJA (si la comunidad la regulariza)
    _tabla_cuota_fija(doc, cuota_fija)

    # 5c. RESUMEN DE TOTALES
    _tabla_totales(doc,
                   dif_fija  = dif_fija,
                   dif_calef = calef.get("diferencia", 0.0) if calef else 0.0,
                   dif_acs   = acs.get("diferencia", 0.0) if acs else 0.0,
                   total_general = total_general)

    # 6. GRÁFICAS
    calef_cob, calef_real = calef.get("importe_cobrado", 0.0), calef.get("importe_real", 0.0)
    acs_cob,   acs_real   = acs.get("importe_cobrado", 0.0),   acs.get("importe_real", 0.0)

    _seccion_titulo(doc, "SU REGULARIZACIÓN EN CIFRAS")
    png1 = _grafica_resumen(calef_cob, calef_real, acs_cob, acs_real)
    if not _insertar_imagen(doc, png1, 16.0):
        _bloque_comparativa_fallback(doc, calef_cob, calef_real, acs_cob, acs_real)

    png2 = _grafica_vs_media(calef.get("consumo_real"), acs.get("consumo_real"), medias)
    if png2:
        _insertar_imagen(doc, png2, 15.0)

    # 7. CIERRE Y FIRMA
    _parrafo(doc,
             "Quedamos a su disposición para cualquier aclaración. Puede verificar los datos con sus recibos "
             "y comunicarnos cualquier discrepancia en un plazo de 30 días. Reciba un cordial saludo.",
             tam=9, despues=10)
    _parrafo(doc, "La Administración", tam=10, negrita=True,
             alin=WD_ALIGN_PARAGRAPH.RIGHT, despues=2)
    _parrafo(doc, "Meditrade Administración de Fincas", tam=8,
             color=C_GRIS_ETQ, alin=WD_ALIGN_PARAGRAPH.RIGHT, despues=0)

    doc.save(ruta_salida)
    return ruta_salida


# ---------------------------------------------------------------------------
# GENERACIÓN EN LOTE — todas las cartas del periodo
# ---------------------------------------------------------------------------

def generar_todas_las_cartas(ruta_bd: str,
                              id_comunidad: int,
                              nombre_periodo: str,
                              ruta_plantilla: str,
                              carpeta_salida: str) -> dict:
    """Genera una carta por cada vecino con repartos calculados para el periodo."""
    os.makedirs(carpeta_salida, exist_ok=True)

    con = sqlite3.connect(ruta_bd)
    con.row_factory = sqlite3.Row

    periodo_row = con.execute(
        "SELECT id_periodo FROM periodos WHERE id_comunidad=? AND nombre=?",
        (id_comunidad, nombre_periodo)
    ).fetchone()

    if not periodo_row:
        con.close()
        return {"ok": False, "error": f"Periodo '{nombre_periodo}' no encontrado"}

    id_periodo = periodo_row["id_periodo"]

    comunidad = con.execute(
        "SELECT nombre FROM comunidades WHERE id_comunidad=?",
        (id_comunidad,)
    ).fetchone()
    nombre_comunidad = comunidad["nombre"] if comunidad else ""

    # Cuota fija: precios mensuales cobrado/real desde config_suministro
    per_info = con.execute(
        "SELECT fecha_inicio, fecha_fin FROM periodos WHERE id_periodo=?",
        (id_periodo,)).fetchone()
    meses = _meses_periodo(per_info["fecha_inicio"] or "",
                           per_info["fecha_fin"] or "") if per_info else 12
    cuota_fija = {}
    for cfg in con.execute(
            "SELECT tipo_suministro, precio_fijo_facturado, precio_fijo_real "
            "FROM config_suministro WHERE id_comunidad=? AND id_periodo=? "
            "AND tipo_suministro IN ('ACS','CALEFACCION')",
            (id_comunidad, id_periodo)).fetchall():
        p_cob  = cfg["precio_fijo_facturado"] or 0.0
        p_real = cfg["precio_fijo_real"] or 0.0
        if p_cob or p_real:
            cobrado = round(p_cob * meses, 2)
            real    = round(p_real * meses, 2)
            cuota_fija[cfg["tipo_suministro"]] = {
                "precio_cobrado": p_cob, "precio_real": p_real,
                "meses": meses, "cobrado": cobrado, "real": real,
                "diferencia": round(real - cobrado, 2),
            }

    todos = obtener_todos_los_vecinos_con_repartos(con, id_comunidad, id_periodo)
    con.close()

    if not todos:
        return {"ok": False, "error": "No hay repartos calculados para este periodo"}

    medias = _calcular_medias(todos)

    cartas_ok = 0
    errores   = []

    for datos in todos:
        datos["comunidad"]  = nombre_comunidad
        datos["medias"]     = medias
        datos["cuota_fija"] = cuota_fija

        vivienda = re.sub(r"[^\w\s-]", "", datos["vecino"]["vivienda"]).strip()
        vivienda = re.sub(r"\s+", "_", vivienda).upper()
        nombre   = re.sub(r"[^\w\s-]", "", datos["vecino"]["nombre"]).strip()
        nombre   = re.sub(r"\s+", "_", nombre).upper()[:25]

        nombre_archivo = f"CARTA_{vivienda}_{nombre}.docx"
        ruta_salida    = os.path.join(carpeta_salida, nombre_archivo)

        try:
            generar_carta(datos, ruta_plantilla, ruta_salida)
            cartas_ok += 1
            print(f"  OK  {datos['vecino']['vivienda']:<22} {datos['vecino']['nombre'][:30]}")
        except Exception as e:
            errores.append(f"{datos['vecino']['vivienda']}: {e}")
            print(f"  ERR {datos['vecino']['vivienda']}: {e}")

    return {
        "ok":               len(errores) == 0,
        "cartas_generadas": cartas_ok,
        "errores":          errores,
        "carpeta_salida":   carpeta_salida,
        "periodo":          nombre_periodo,
    }


# ---------------------------------------------------------------------------
# EJECUCIÓN DIRECTA
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Genera cartas Word por vecino")
    parser.add_argument("--bd",        required=True,           help="Ruta a gestion.db")
    parser.add_argument("--comunidad", required=True, type=int, help="ID comunidad")
    parser.add_argument("--periodo",   required=True,           help="Nombre periodo ej: 2024-2025")
    parser.add_argument("--plantilla", required=True,           help="Ruta a Plantilla_Cartas.docx")
    parser.add_argument("--salida",    default="./cartas_generadas/",
                        help="Carpeta de salida (default: ./cartas_generadas/)")
    args = parser.parse_args()

    print(f"\n=== GENERANDO CARTAS  —  {args.periodo} ===\n")
    resultado = generar_todas_las_cartas(
        ruta_bd          = args.bd,
        id_comunidad     = args.comunidad,
        nombre_periodo   = args.periodo,
        ruta_plantilla   = args.plantilla,
        carpeta_salida   = args.salida,
    )

    print(f"\n{'=' * 52}")
    print(f"Resultado : {'OK' if resultado['ok'] else 'CON ERRORES'}")
    print(f"Cartas    : {resultado.get('cartas_generadas', 0)}")
    print(f"Carpeta   : {resultado.get('carpeta_salida')}")
    if resultado.get("errores"):
        print(f"Errores   : {len(resultado['errores'])}")
        for e in resultado["errores"][:5]:
            print(f"  - {e}")
