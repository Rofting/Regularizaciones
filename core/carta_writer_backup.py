"""
carta_writer.py
===============
Genera una carta Word (.docx) por vecino usando Plantilla_Cartas.docx
y los datos de la tabla 'repartos' de la BD SQLite.

Diseño moderno y minimalista con:
  - Cabecera corporativa azul marino
  - Bloques de información limpios
  - Tablas de suministro con colores semafórico en diferencias
  - Gráfica comparativa cobrado vs real (matplotlib, con fallback elegante)
  - Caja de importe total con color según resultado

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
from docx.shared import Pt, RGBColor, Cm, Twips
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

try:
    from docx.enum.table import WD_TABLE_ALIGNMENT
    _HAS_TABLE_ALIGN = True
except ImportError:
    _HAS_TABLE_ALIGN = False

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
# CONSTANTES DE COLOR Y DISEÑO
# ---------------------------------------------------------------------------

C_AZUL_MARINO    = "#1A3A5C"
C_AZUL_HEADER2   = "#1E3A5C"
C_TEXTO_HEADER   = "#D6ECF7"
C_GRIS_ETQ       = "#7F8C8D"
C_DARK           = "#2C3E50"
C_FONDO_CABCOL   = "#EBF5FB"
C_VERDE          = "#1E8449"
C_ROJO           = "#C0392B"
C_FONDO_PAGAR    = "#FADBD8"
C_FONDO_DEVOLVER = "#D5F5E3"
C_FONDO_CERO     = "#EBF5FB"
C_BARRA_COB      = "#85C1E9"
C_BARRA_REAL     = "#1A3A5C"
C_FONDO_GRAF     = "#F8F9FA"

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
    """
    Devuelve todos los datos necesarios para la carta de un vecino.
    Estructura:
    {
        'vecino':   {nombre, vivienda, coeficiente},
        'periodo':  {nombre, fecha_inicio, fecha_fin},
        'calef':    {lec_ini, lec_fin, consumo, importe_cobrado, importe_real, diferencia},
        'acs':      {lec_ini, lec_fin, consumo, importe_cobrado, importe_real, diferencia},
        'extras':   [],
        'total':    {cobrado, real, diferencia}
    }
    """
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


# ---------------------------------------------------------------------------
# FORMATEADO DE NÚMEROS
# ---------------------------------------------------------------------------

def _fmt_euros(valor: float) -> str:
    """Formatea un número como '1.234,56 €'"""
    if valor is None:
        return "0,00 €"
    return f"{valor:,.2f} €".replace(",", "X").replace(".", ",").replace("X", ".")


def _fmt_consumo(valor: float, unidad: str) -> str:
    """Formatea consumo con unidad: '24.500 kWh' o '43 m3'"""
    if valor is None or valor == 0:
        return f"0 {unidad}"
    if unidad == "kWh":
        return f"{valor:,.0f} kWh".replace(",", ".")
    return f"{valor:.1f} m3".rstrip("0").rstrip(".")


# ---------------------------------------------------------------------------
# HELPERS DE ESTILO DOCX
# ---------------------------------------------------------------------------

def _rgb(hex_str: str) -> RGBColor:
    """Convierte '#RRGGBB' a RGBColor."""
    h = hex_str.lstrip("#")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def _set_color_celda(celda, hex_color: str) -> None:
    """Aplica color de fondo a una celda."""
    tc = celda._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color.lstrip("#"))
    tcPr.append(shd)


def _quitar_margen_parrafo(parrafo) -> None:
    """Pone space_before=0 y space_after=6pt (60 twips) al párrafo."""
    pf = parrafo.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after  = Pt(6)


def _set_ancho_col(tabla, col_idx: int, ancho_cm: float) -> None:
    """Fija el ancho de todas las celdas de una columna."""
    dxa = int(ancho_cm * 567)  # 1 cm = 567 DXA
    for celda in tabla.columns[col_idx].cells:
        tc = celda._tc
        tcPr = tc.get_or_add_tcPr()
        # Eliminar w:tcW previo si existe
        for old in tcPr.findall(qn("w:tcW")):
            tcPr.remove(old)
        tcW = OxmlElement("w:tcW")
        tcW.set(qn("w:w"),    str(dxa))
        tcW.set(qn("w:type"), "dxa")
        tcPr.append(tcW)


def _sin_borde_tabla(tabla) -> None:
    """Quita todos los bordes visibles de una tabla."""
    tbl   = tabla._tbl
    tblPr = tbl.find(qn("w:tblPr"))
    if tblPr is None:
        tblPr = OxmlElement("w:tblPr")
        tbl.insert(0, tblPr)
    # Eliminar tblBorders previo
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
                    alineacion=WD_ALIGN_PARAGRAPH.LEFT) -> None:
    """Escribe texto en una celda con estilo completo."""
    for p in celda.paragraphs:
        p._element.getparent().remove(p._element)
    p = celda.add_paragraph()
    _quitar_margen_parrafo(p)
    p.alignment = alineacion
    run = p.add_run(texto)
    run.bold      = negrita
    run.font.size = Pt(tamaño)
    if color_hex:
        run.font.color.rgb = _rgb(color_hex)


def _quitar_margen_celda(celda) -> None:
    """Pone márgenes internos de celda a 0 (excepto left=0.1cm)."""
    tc    = celda._tc
    tcPr  = tc.get_or_add_tcPr()
    tcMar = OxmlElement("w:tcMar")
    for lado, val in [("top", "0"), ("bottom", "0"),
                      ("left", "56"), ("right", "56")]:
        m = OxmlElement(f"w:{lado}")
        m.set(qn("w:w"),    val)
        m.set(qn("w:type"), "dxa")
        tcMar.append(m)
    tcPr.append(tcMar)


# ---------------------------------------------------------------------------
# BLOQUES DE DISEÑO
# ---------------------------------------------------------------------------

def _bloque_cabecera(doc: Document, nombre_comunidad: str) -> None:
    """
    Cabecera corporativa: tabla 1 col, 2 filas.
    Fila 1 (azul marino): "REGULARIZACIÓN DE CUOTAS · Meditrade Administración"
    Fila 2 (azul oscuro): nombre comunidad en gris claro.
    """
    tabla = doc.add_table(rows=2, cols=1)
    _sin_borde_tabla(tabla)
    _set_ancho_col(tabla, 0, 16.0)

    # Fila 1 — título principal
    c1 = tabla.rows[0].cells[0]
    _set_color_celda(c1, C_AZUL_MARINO)
    _escribir_celda(c1, "REGULARIZACIÓN DE CUOTAS  ·  Meditrade Administración",
                    negrita=True, tamaño=14,
                    color_hex="#FFFFFF",
                    alineacion=WD_ALIGN_PARAGRAPH.CENTER)

    # Fila 2 — nombre comunidad
    c2 = tabla.rows[1].cells[0]
    _set_color_celda(c2, C_AZUL_HEADER2)
    _escribir_celda(c2, nombre_comunidad.upper(),
                    negrita=False, tamaño=10,
                    color_hex=C_TEXTO_HEADER,
                    alineacion=WD_ALIGN_PARAGRAPH.CENTER)

    p = doc.add_paragraph()
    _quitar_margen_parrafo(p)


def _bloque_info(doc: Document,
                 fecha_str: str,
                 vivienda: str,
                 propietario: str,
                 periodo: str) -> None:
    """
    Pequeña tabla 2 cols: etiqueta | valor.
    Etiquetas en gris, valores en oscuro.
    """
    filas_data = [
        ("Fecha",       fecha_str),
        ("Propiedad",   vivienda),
        ("Propietario", propietario),
        ("Período",     periodo),
    ]

    tabla = doc.add_table(rows=len(filas_data), cols=2)
    _sin_borde_tabla(tabla)
    _set_ancho_col(tabla, 0, 3.0)
    _set_ancho_col(tabla, 1, 13.0)

    for i, (etq, val) in enumerate(filas_data):
        fila = tabla.rows[i]
        _escribir_celda(fila.cells[0], etq,
                        negrita=False, tamaño=9,
                        color_hex=C_GRIS_ETQ,
                        alineacion=WD_ALIGN_PARAGRAPH.LEFT)
        _escribir_celda(fila.cells[1], val,
                        negrita=True, tamaño=9,
                        color_hex=C_DARK,
                        alineacion=WD_ALIGN_PARAGRAPH.LEFT)

    p = doc.add_paragraph()
    _quitar_margen_parrafo(p)


def _seccion_titulo(doc: Document, icono_texto: str) -> None:
    """Barra de sección: tabla 1 celda con fondo azul marino y texto blanco bold."""
    tabla = doc.add_table(rows=1, cols=1)
    _sin_borde_tabla(tabla)
    _set_ancho_col(tabla, 0, 16.0)
    c = tabla.rows[0].cells[0]
    _set_color_celda(c, C_AZUL_MARINO)
    _escribir_celda(c, icono_texto,
                    negrita=True, tamaño=10,
                    color_hex="#FFFFFF",
                    alineacion=WD_ALIGN_PARAGRAPH.LEFT)


def _tabla_suministro(doc: Document,
                      titulo: str,
                      lectura_ini,
                      lectura_fin,
                      consumo,
                      unidad: str,
                      importe_cobrado: float,
                      importe_real: float,
                      diferencia: float) -> None:
    """
    Bloque completo para un suministro (CALEFACCIÓN o ACS):
      - Barra de título azul
      - Tabla 4 cols: Concepto | Cobrado | Real | Diferencia
      - Fila IMPORTE con diferencia coloreada
    """
    _seccion_titulo(doc, titulo)

    ANCHOS = [6.0, 3.5, 3.5, 3.5]
    CABECERAS = ["Concepto", "Cobrado", "Real", "Diferencia"]

    # Datos de filas (texto en cada columna)
    lec_ini_txt = _fmt_consumo(lectura_ini, unidad)
    lec_fin_txt = _fmt_consumo(lectura_fin, unidad)
    consumo_txt = _fmt_consumo(consumo,     unidad)

    filas = [
        ("Lectura inicial",  lec_ini_txt, lec_ini_txt, "—"),
        ("Lectura final",    lec_fin_txt, lec_fin_txt, "—"),
        ("Consumo",          consumo_txt, consumo_txt, "—"),
        ("IMPORTE",
         _fmt_euros(importe_cobrado),
         _fmt_euros(importe_real),
         None),  # diferencia se tratará aparte
    ]

    # Signo y color de la diferencia
    if diferencia < -0.005:
        dif_txt   = f"− {_fmt_euros(abs(diferencia))}"
        dif_color = C_VERDE
    elif diferencia > 0.005:
        dif_txt   = f"+ {_fmt_euros(diferencia)}"
        dif_color = C_ROJO
    else:
        dif_txt   = "0,00 €"
        dif_color = C_DARK

    # Tabla
    n_filas = len(filas) + 1  # +1 cabecera
    tabla = doc.add_table(rows=n_filas, cols=4)
    _sin_borde_tabla(tabla)
    for col_idx, ancho in enumerate(ANCHOS):
        _set_ancho_col(tabla, col_idx, ancho)

    # Fila cabecera
    fila_cab = tabla.rows[0]
    for col_idx, cab_txt in enumerate(CABECERAS):
        c = fila_cab.cells[col_idx]
        _set_color_celda(c, C_FONDO_CABCOL)
        alin = WD_ALIGN_PARAGRAPH.RIGHT if col_idx > 0 else WD_ALIGN_PARAGRAPH.LEFT
        _escribir_celda(c, cab_txt,
                        negrita=True, tamaño=9,
                        color_hex=C_AZUL_MARINO,
                        alineacion=alin)

    # Filas de datos
    for i, (concepto, col_cob, col_real, col_dif) in enumerate(filas):
        fila_obj = tabla.rows[i + 1]
        es_importe = (concepto == "IMPORTE")

        # Concepto
        c0 = fila_obj.cells[0]
        _escribir_celda(c0, concepto,
                        negrita=es_importe, tamaño=9,
                        color_hex=C_DARK,
                        alineacion=WD_ALIGN_PARAGRAPH.LEFT)

        # Cobrado
        c1 = fila_obj.cells[1]
        _escribir_celda(c1, col_cob,
                        negrita=es_importe, tamaño=9,
                        color_hex=C_DARK,
                        alineacion=WD_ALIGN_PARAGRAPH.RIGHT)

        # Real
        c2 = fila_obj.cells[2]
        _escribir_celda(c2, col_real,
                        negrita=es_importe, tamaño=9,
                        color_hex=C_DARK,
                        alineacion=WD_ALIGN_PARAGRAPH.RIGHT)

        # Diferencia
        c3 = fila_obj.cells[3]
        if es_importe:
            _escribir_celda(c3, dif_txt,
                            negrita=True, tamaño=9,
                            color_hex=dif_color,
                            alineacion=WD_ALIGN_PARAGRAPH.RIGHT)
        else:
            _escribir_celda(c3, "—",
                            negrita=False, tamaño=9,
                            color_hex=C_GRIS_ETQ,
                            alineacion=WD_ALIGN_PARAGRAPH.CENTER)

    p = doc.add_paragraph()
    _quitar_margen_parrafo(p)


def _grafica_comparativa(calef_cob: float, calef_real: float,
                          acs_cob: float,   acs_real: float) -> Optional[str]:
    """
    Genera una gráfica de barras horizontales agrupadas (cobrado vs real).
    Devuelve la ruta del PNG temporal o None si matplotlib no está disponible
    o si ocurre cualquier error.
    """
    if not _HAS_MPL:
        return None

    try:
        fig, ax = plt.subplots(figsize=(15 / 2.54, 5 / 2.54), dpi=110)
        fig.patch.set_facecolor(C_FONDO_GRAF)
        ax.set_facecolor(C_FONDO_GRAF)

        grupos   = ["CALEFACCIÓN", "ACS"]
        cobrados = [calef_cob, acs_cob]
        reales   = [calef_real, acs_real]

        y      = np.arange(len(grupos))
        altura = 0.32

        barras_cob  = ax.barh(y + altura / 2, cobrados, altura,
                               color=C_BARRA_COB, label="Cobrado")
        barras_real = ax.barh(y - altura / 2, reales,   altura,
                               color=C_BARRA_REAL, label="Real")

        # Labels en cada barra
        for bar, val in zip(barras_cob, cobrados):
            if val and val > 0:
                ax.text(bar.get_width() + max(cobrados + reales) * 0.01,
                        bar.get_y() + bar.get_height() / 2,
                        _fmt_euros(val),
                        va="center", ha="left", fontsize=7, color=C_DARK)

        for bar, val in zip(barras_real, reales):
            if val and val > 0:
                ax.text(bar.get_width() + max(cobrados + reales) * 0.01,
                        bar.get_y() + bar.get_height() / 2,
                        _fmt_euros(val),
                        va="center", ha="left", fontsize=7, color=C_DARK)

        ax.set_yticks(y)
        ax.set_yticklabels(grupos, fontsize=8, color=C_DARK)
        ax.set_xlabel("Importe (€)", fontsize=7, color=C_GRIS_ETQ)
        ax.tick_params(axis="x", labelsize=7, colors=C_GRIS_ETQ)
        ax.tick_params(axis="y", colors=C_DARK)

        # Quitar bordes superior y derecho
        ax.spines["top"].set_visible(False)
        ax.spines["right"].set_visible(False)
        ax.spines["left"].set_color("#CCCCCC")
        ax.spines["bottom"].set_color("#CCCCCC")

        legend = ax.legend(fontsize=7, framealpha=0,
                           handles=[
                               mpatches.Patch(color=C_BARRA_COB,  label="Cobrado"),
                               mpatches.Patch(color=C_BARRA_REAL, label="Real"),
                           ])
        for text in legend.get_texts():
            text.set_color(C_DARK)

        plt.tight_layout(pad=0.5)

        tmp = tempfile.mktemp(suffix=".png")
        fig.savefig(tmp, dpi=110, bbox_inches="tight",
                    facecolor=C_FONDO_GRAF)
        plt.close(fig)
        return tmp

    except Exception:
        return None


def _bloque_comparativa(doc: Document,
                         calef_cob: float, calef_real: float,
                         acs_cob: float,   acs_real: float) -> None:
    """
    Inserta la sección COMPARATIVA.
    Usa gráfica matplotlib si está disponible, en caso contrario una tabla texto.
    """
    _seccion_titulo(doc, "COMPARATIVA  COBRADO vs REAL")

    tmp_png = _grafica_comparativa(calef_cob, calef_real, acs_cob, acs_real)

    if tmp_png and os.path.exists(tmp_png):
        try:
            p = doc.add_paragraph()
            _quitar_margen_parrafo(p)
            run = p.add_run()
            run.add_picture(tmp_png, width=Cm(15))
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        finally:
            try:
                os.remove(tmp_png)
            except OSError:
                pass
    else:
        # Fallback elegante: tabla resumen
        ANCHOS_FB = [4.0, 4.0, 4.0, 4.0]
        tabla = doc.add_table(rows=3, cols=4)
        _sin_borde_tabla(tabla)
        for idx, a in enumerate(ANCHOS_FB):
            _set_ancho_col(tabla, idx, a)

        cabeceras_fb = ["", "Cobrado", "Real", "Diferencia"]
        for j, txt in enumerate(cabeceras_fb):
            c = tabla.rows[0].cells[j]
            _set_color_celda(c, C_FONDO_CABCOL)
            alin = WD_ALIGN_PARAGRAPH.RIGHT if j > 0 else WD_ALIGN_PARAGRAPH.LEFT
            _escribir_celda(c, txt,
                            negrita=True, tamaño=9,
                            color_hex=C_AZUL_MARINO,
                            alineacion=alin)

        filas_fb = [
            ("Calefacción", calef_cob, calef_real, calef_real - calef_cob),
            ("ACS",         acs_cob,   acs_real,   acs_real  - acs_cob),
        ]
        for i, (nom, cob, real, dif) in enumerate(filas_fb):
            fila_obj = tabla.rows[i + 1]
            _escribir_celda(fila_obj.cells[0], nom,
                            negrita=False, tamaño=9,
                            color_hex=C_DARK,
                            alineacion=WD_ALIGN_PARAGRAPH.LEFT)
            _escribir_celda(fila_obj.cells[1], _fmt_euros(cob),
                            negrita=False, tamaño=9,
                            color_hex=C_DARK,
                            alineacion=WD_ALIGN_PARAGRAPH.RIGHT)
            _escribir_celda(fila_obj.cells[2], _fmt_euros(real),
                            negrita=False, tamaño=9,
                            color_hex=C_DARK,
                            alineacion=WD_ALIGN_PARAGRAPH.RIGHT)
            dif_color = C_VERDE if dif < -0.005 else (C_ROJO if dif > 0.005 else C_DARK)
            signo = "− " if dif < -0.005 else ("+ " if dif > 0.005 else "")
            _escribir_celda(fila_obj.cells[3],
                            signo + _fmt_euros(abs(dif)),
                            negrita=False, tamaño=9,
                            color_hex=dif_color,
                            alineacion=WD_ALIGN_PARAGRAPH.RIGHT)

    p = doc.add_paragraph()
    _quitar_margen_parrafo(p)


def _caja_total(doc: Document, diferencia_total: float) -> None:
    """
    Caja IMPORTE TOTAL: tabla 1 fila 2 celdas.
    Color de fondo y texto según si hay que pagar, devolver o es cero.
    """
    if diferencia_total > 0.005:
        fondo   = C_FONDO_PAGAR
        color_v = C_ROJO
        signo   = "+"
    elif diferencia_total < -0.005:
        fondo   = C_FONDO_DEVOLVER
        color_v = C_VERDE
        signo   = "−"
    else:
        fondo   = C_FONDO_CERO
        color_v = C_AZUL_MARINO
        signo   = ""

    tabla = doc.add_table(rows=1, cols=2)
    _sin_borde_tabla(tabla)
    _set_ancho_col(tabla, 0, 9.0)
    _set_ancho_col(tabla, 1, 7.0)

    fila = tabla.rows[0]

    # Col 1: etiqueta
    c1 = fila.cells[0]
    _set_color_celda(c1, fondo)
    _escribir_celda(c1, "IMPORTE TOTAL",
                    negrita=True, tamaño=12,
                    color_hex=C_DARK,
                    alineacion=WD_ALIGN_PARAGRAPH.LEFT)

    # Col 2: importe con signo
    importe_str = f"{signo} {_fmt_euros(abs(diferencia_total))}" if signo else _fmt_euros(0)
    c2 = fila.cells[1]
    _set_color_celda(c2, fondo)
    _escribir_celda(c2, importe_str,
                    negrita=True, tamaño=13,
                    color_hex=color_v,
                    alineacion=WD_ALIGN_PARAGRAPH.RIGHT)

    p = doc.add_paragraph()
    _quitar_margen_parrafo(p)


# ---------------------------------------------------------------------------
# GENERACIÓN DE UNA CARTA
# ---------------------------------------------------------------------------

def generar_carta(datos: dict,
                  ruta_plantilla: str,
                  ruta_salida: str) -> str:
    """
    Genera una carta Word para un vecino.

    Args:
        datos:          Dict con estructura de _obtener_repartos_vecino()
        ruta_plantilla: Ruta a Plantilla_Cartas.docx
        ruta_salida:    Ruta del archivo .docx a generar

    Returns:
        Ruta del archivo generado
    """
    import shutil
    shutil.copy2(ruta_plantilla, ruta_salida)
    doc = Document(ruta_salida)

    # Limpiar el cuerpo preservando el header con logo de Meditrade
    for par in list(doc.paragraphs):
        p_elem = par._element
        p_elem.getparent().remove(p_elem)

    for tabla in list(doc.tables):
        tabla._element.getparent().remove(tabla._element)

    # --- Datos de trabajo ---
    vecino          = datos["vecino"]
    periodo         = datos["periodo"]
    calef           = datos.get("calef") or {}
    acs             = datos.get("acs")   or {}
    total           = datos["total"]
    nombre_comunidad = datos.get("comunidad", "COMUNIDAD DE PROPIETARIOS")

    nombre_periodo  = periodo["nombre"]
    hoy             = datetime.now()
    fecha_str       = f"Zaragoza, {hoy.day} de {MESES_ES[hoy.month]} de {hoy.year}"

    # ---- 1. CABECERA CORPORATIVA ----
    _bloque_cabecera(doc, nombre_comunidad)

    # ---- 2. BLOQUE INFO ----
    _bloque_info(
        doc,
        fecha_str    = fecha_str,
        vivienda     = vecino["vivienda"],
        propietario  = vecino["nombre"],
        periodo      = nombre_periodo,
    )

    # ---- 3. TEXTO INTRO ----
    p_intro = doc.add_paragraph()
    _quitar_margen_parrafo(p_intro)
    p_intro.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    run_intro = p_intro.add_run(
        "En cumplimiento del acuerdo adoptado en la última Junta General Ordinaria, "
        "le comunicamos el resultado de la regularización anual de los consumos de "
        "Calefacción y Agua Caliente Sanitaria (ACS) correspondientes al período "
        f"{nombre_periodo}. A continuación encontrará el detalle de lecturas, "
        "consumos e importes cobrados frente a los reales."
    )
    run_intro.font.size = Pt(9)
    run_intro.font.color.rgb = _rgb(C_DARK)

    p_sep = doc.add_paragraph()
    _quitar_margen_parrafo(p_sep)

    # ---- 4. SECCIÓN CALEFACCIÓN ----
    if calef:
        _tabla_suministro(
            doc,
            titulo          = "CALEFACCIÓN",
            lectura_ini     = calef.get("lectura_inicial", 0),
            lectura_fin     = calef.get("lectura_final",   0),
            consumo         = calef.get("consumo_real",    0),
            unidad          = "kWh",
            importe_cobrado = calef.get("importe_cobrado", 0.0),
            importe_real    = calef.get("importe_real",    0.0),
            diferencia      = calef.get("diferencia",      0.0),
        )

    # ---- 5. SECCIÓN ACS ----
    if acs:
        _tabla_suministro(
            doc,
            titulo          = "AGUA CALIENTE SANITARIA (ACS)",
            lectura_ini     = acs.get("lectura_inicial", 0),
            lectura_fin     = acs.get("lectura_final",   0),
            consumo         = acs.get("consumo_real",    0),
            unidad          = "m3",
            importe_cobrado = acs.get("importe_cobrado", 0.0),
            importe_real    = acs.get("importe_real",    0.0),
            diferencia      = acs.get("diferencia",      0.0),
        )

    # ---- 6. GRÁFICA COMPARATIVA ----
    calef_cob  = calef.get("importe_cobrado", 0.0) if calef else 0.0
    calef_real = calef.get("importe_real",    0.0) if calef else 0.0
    acs_cob    = acs.get("importe_cobrado",   0.0) if acs   else 0.0
    acs_real   = acs.get("importe_real",      0.0) if acs   else 0.0

    _bloque_comparativa(doc, calef_cob, calef_real, acs_cob, acs_real)

    # ---- 7. CAJA TOTAL ----
    _caja_total(doc, total["diferencia"])

    # ---- 8. CIERRE ----
    p_cierre = doc.add_paragraph()
    _quitar_margen_parrafo(p_cierre)
    p_cierre.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    run_cierre = p_cierre.add_run(
        "Quedamos a su disposición para cualquier aclaración o consulta. "
        "Le rogamos verifique los datos con la documentación de sus recibos "
        "y nos comunique cualquier discrepancia. Reciba un cordial saludo."
    )
    run_cierre.font.size = Pt(9)
    run_cierre.font.color.rgb = _rgb(C_DARK)

    doc.add_paragraph()

    p_firma = doc.add_paragraph()
    _quitar_margen_parrafo(p_firma)
    p_firma.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run_firma = p_firma.add_run("La Administración")
    run_firma.bold      = True
    run_firma.font.size = Pt(10)
    run_firma.font.color.rgb = _rgb(C_DARK)

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
    """
    Genera una carta por cada vecino con repartos calculados para el periodo.

    Returns:
        {ok, cartas_generadas, errores, carpeta_salida, periodo}
    """
    os.makedirs(carpeta_salida, exist_ok=True)

    con = sqlite3.connect(ruta_bd)
    con.row_factory = sqlite3.Row

    # Obtener id_periodo
    periodo_row = con.execute(
        "SELECT id_periodo FROM periodos WHERE id_comunidad=? AND nombre=?",
        (id_comunidad, nombre_periodo)
    ).fetchone()

    if not periodo_row:
        con.close()
        return {"ok": False, "error": f"Periodo '{nombre_periodo}' no encontrado"}

    id_periodo = periodo_row["id_periodo"]

    # Obtener nombre de la comunidad
    comunidad = con.execute(
        "SELECT nombre FROM comunidades WHERE id_comunidad=?",
        (id_comunidad,)
    ).fetchone()
    nombre_comunidad = comunidad["nombre"] if comunidad else ""

    todos = obtener_todos_los_vecinos_con_repartos(con, id_comunidad, id_periodo)
    con.close()

    if not todos:
        return {"ok": False, "error": "No hay repartos calculados para este periodo"}

    cartas_ok = 0
    errores   = []

    for datos in todos:
        datos["comunidad"] = nombre_comunidad

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
        "ok":              len(errores) == 0,
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
