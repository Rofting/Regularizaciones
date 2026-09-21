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

try:
    from consumption_charts import render_consumption_charts
except ImportError:
    render_consumption_charts = None

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
except ImportError as _e:
    _HAS_MPL = False
    print(f"⚠️  AVISO: no se generarán gráficas — falta una dependencia: {_e}")
    print("   Instala con:  pip install matplotlib numpy")


# ---------------------------------------------------------------------------
# PALETA MODERNA (índigo corporativo + semáforo suave)
# ---------------------------------------------------------------------------

# Paleta corporativa configurable: azul marino + rojo de acento
C_PRIMARIO       = "#1B365D"   # azul marino corporativo (cabeceras)
C_PRIMARIO_MED   = "#24507A"   # azul medio (banda secundaria)
C_PRIMARIO_SUAVE = "#EAF1F8"   # azul muy claro (fondos de cabecera de columna)
C_ROJO_CORP      = "#C8102E"   # rojo corporativo del logo
C_TEXTO_HEADER   = "#D6E4F0"
C_GRIS_ETQ       = "#6B7280"
C_DARK           = "#111827"
C_ZEBRA          = "#F7FAFC"
C_VERDE          = "#047857"
C_ROJO           = "#C8102E"   # rojo corporativo también para "a pagar"
C_FONDO_PAGAR    = "#FBE9EB"
C_FONDO_DEVOLVER = "#D1FAE5"
C_FONDO_CERO     = "#EAF1F8"
C_BARRA_COB      = "#A9C6E8"   # azul claro (cobrado)
C_BARRA_REAL     = "#1B365D"   # azul marino (real)
C_ACENTO_2       = "#C8102E"   # rojo (ACS en el donut)
C_BARRA_MEDIA    = "#CBD5E1"   # gris (media comunidad)
C_FONDO_GRAF     = "#FFFFFF"
FUENTE_CARTA     = "Trebuchet MS"

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
                              id_periodo: int,
                              selected_concepts: tuple[str, ...] | None = None) -> dict:
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

    if not repartos:
        # Formato nuevo de la fase 1: resultados trazables en céntimos.
        nuevos = cur.execute(
            """
            SELECT r.concept_key, r.consumption, r.billed_cents,
                   r.actual_cents, r.difference_cents
            FROM owner_concept_results r
            WHERE r.id_propietario=? AND r.id_periodo=?
            ORDER BY r.concept_key
            """,
            (id_propietario, id_periodo),
        ).fetchall()
        agrupados = {}
        for row in nuevos:
            if selected_concepts and row[0] not in selected_concepts:
                continue
            tipo = "ACS" if str(row[0]).startswith("acs_") else "CALEFACCION"
            item = agrupados.setdefault(tipo, {
                "tipo_suministro": tipo, "lectura_inicial": None,
                "lectura_final": None, "consumo_real": None,
                "importe_cobrado": 0.0, "importe_real": 0.0,
                "diferencia": 0.0, "notas": None,
            })
            item["importe_cobrado"] += (row[2] or 0) / 100
            item["importe_real"] += (row[3] or 0) / 100
            item["diferencia"] += (row[4] or 0) / 100
            if row[1] is not None:
                item["consumo_real"] = (item["consumo_real"] or 0.0) + float(row[1])
        reading_rows = cur.execute(
            """
            SELECT fecha_lectura, valor_acumulado, estado, notas
            FROM lecturas_vecino
            WHERE id_propietario=? AND id_periodo=? AND tipo='ACS'
            ORDER BY fecha_lectura
            """, (id_propietario, id_periodo)
        ).fetchall()
        if reading_rows and "ACS" in agrupados:
            agrupados["ACS"]["lectura_inicial"] = reading_rows[0][1]
            agrupados["ACS"]["lectura_final"] = reading_rows[-1][1]
            if any(row[2] == "estimado" for row in reading_rows):
                agrupados["ACS"]["notas"] = next((row[3] for row in reading_rows if row[3]), None)
        repartos = tuple(agrupados.values())

    datos = {
        "id_propietario": id_propietario,
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
                                            id_periodo: int,
                                            selected_concepts: tuple[str, ...] | None = None) -> list:
    """Devuelve la lista completa de vecinos con sus repartos calculados."""
    cur = con.cursor()
    vecinos = cur.execute("""
        SELECT DISTINCT p.id_propietario
        FROM repartos r
        JOIN propietarios p ON p.id_propietario = r.id_propietario
        WHERE p.id_comunidad = ? AND r.id_periodo = ?
        ORDER BY p.codigo_vivienda
    """, (id_comunidad, id_periodo)).fetchall()
    if not vecinos:
        vecinos = cur.execute(
            """
            SELECT DISTINCT p.id_propietario
            FROM owner_concept_results r
            JOIN propietarios p ON p.id_propietario=r.id_propietario
            WHERE p.id_comunidad=? AND r.id_periodo=?
            ORDER BY p.codigo_vivienda
            """, (id_comunidad, id_periodo)
        ).fetchall()

    resultado = []
    for (id_prop,) in vecinos:
        datos = _obtener_repartos_vecino(con, id_prop, id_periodo, selected_concepts)
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
                     nombre_periodo: str, oficina: str = "Administración de fincas") -> None:
    """Cabecera índigo: título + comunidad + periodo."""
    tabla = doc.add_table(rows=2, cols=1)
    _sin_borde_tabla(tabla)
    _set_ancho_col(tabla, 0, 16.0)

    c1 = tabla.rows[0].cells[0]
    _set_color_celda(c1, C_PRIMARIO)
    _escribir_celda(c1, "REGULARIZACIÓN DE CONSUMOS",
                    negrita=True, tamaño=12, color_hex="#FFFFFF",
                    alineacion=WD_ALIGN_PARAGRAPH.CENTER, despues=1)

    c2 = tabla.rows[1].cells[0]
    _set_color_celda(c2, C_PRIMARIO_MED)
    _escribir_celda(c2,
                    f"{nombre_comunidad.upper()}   ·   PERIODO {nombre_periodo}   ·   {oficina}",
                    negrita=False, tamaño=8, color_hex=C_TEXTO_HEADER,
                    alineacion=WD_ALIGN_PARAGRAPH.CENTER, despues=1)

    p = doc.add_paragraph()
    _quitar_margen_parrafo(p, 3)


def _tabla_conceptos_activos(doc: Document, conceptos: list[dict]) -> None:
    """Detalle compacto de conceptos canónicos sin inventar un servicio.

    Los conceptos llegan ya filtrados por el perfil del expediente, por lo que
    una comunidad sin calefacción nunca muestra una columna ni un rótulo de
    calefacción. Fijo, variable, créditos y extras permanecen como filas
    independientes para que el propietario pueda revisarlos.
    """
    _seccion_titulo(doc, "DETALLE DE LA REGULARIZACIÓN")
    tabla = doc.add_table(rows=len(conceptos) + 2, cols=4)
    _sin_borde_tabla(tabla)
    for index, width in enumerate((7.2, 2.9, 2.9, 3.0)):
        _set_ancho_col(tabla, index, width)
    for index, title in enumerate(("Concepto", "Cobrado", "Coste real", "Diferencia")):
        cell = tabla.rows[0].cells[index]
        _set_color_celda(cell, C_PRIMARIO_SUAVE)
        _escribir_celda(cell, title, negrita=True, tamaño=8.5, color_hex=C_PRIMARIO,
                        alineacion=WD_ALIGN_PARAGRAPH.LEFT if index == 0 else WD_ALIGN_PARAGRAPH.RIGHT,
                        despues=1)
    total_billed = total_actual = total_difference = 0.0
    for row_index, concept in enumerate(conceptos, start=1):
        row = tabla.rows[row_index]
        if row_index % 2 == 0:
            for cell in row.cells:
                _set_color_celda(cell, C_ZEBRA)
        billed = float(concept.get("importe_cobrado", 0.0) or 0.0)
        actual = float(concept.get("importe_real", 0.0) or 0.0)
        difference = float(concept.get("diferencia", actual - billed) or 0.0)
        colour = C_VERDE if difference < -0.005 else (C_ROJO if difference > 0.005 else C_DARK)
        sign = "− " if difference < -0.005 else ("+ " if difference > 0.005 else "")
        _escribir_celda(row.cells[0], str(concept["label"]), tamaño=8.5, color_hex=C_DARK, despues=1)
        _escribir_celda(row.cells[1], _fmt_euros(billed), tamaño=8.5, color_hex=C_DARK,
                        alineacion=WD_ALIGN_PARAGRAPH.RIGHT, despues=1)
        _escribir_celda(row.cells[2], _fmt_euros(actual), tamaño=8.5, color_hex=C_DARK,
                        alineacion=WD_ALIGN_PARAGRAPH.RIGHT, despues=1)
        _escribir_celda(row.cells[3], sign + _fmt_euros(abs(difference)), tamaño=8.5,
                        color_hex=colour, alineacion=WD_ALIGN_PARAGRAPH.RIGHT, despues=1)
        total_billed += billed
        total_actual += actual
        total_difference += difference
    row = tabla.rows[-1]
    for cell in row.cells:
        _set_color_celda(cell, C_PRIMARIO_SUAVE)
    colour = C_VERDE if total_difference < -0.005 else (C_ROJO if total_difference > 0.005 else C_PRIMARIO)
    sign = "− " if total_difference < -0.005 else ("+ " if total_difference > 0.005 else "")
    _escribir_celda(row.cells[0], "TOTAL", negrita=True, tamaño=9, color_hex=C_PRIMARIO, despues=1)
    _escribir_celda(row.cells[1], _fmt_euros(total_billed), negrita=True, tamaño=9, color_hex=C_DARK,
                    alineacion=WD_ALIGN_PARAGRAPH.RIGHT, despues=1)
    _escribir_celda(row.cells[2], _fmt_euros(total_actual), negrita=True, tamaño=9, color_hex=C_DARK,
                    alineacion=WD_ALIGN_PARAGRAPH.RIGHT, despues=1)
    _escribir_celda(row.cells[3], sign + _fmt_euros(abs(total_difference)), negrita=True,
                    tamaño=9, color_hex=colour, alineacion=WD_ALIGN_PARAGRAPH.RIGHT, despues=1)
    p = doc.add_paragraph()
    _quitar_margen_parrafo(p, 3)


def _limpiar_identidad_plantilla(doc: Document, identity: dict) -> None:
    """Elimina marca heredada y aplica identidad portable en cabecera y pie."""
    for section in doc.sections:
        for story in (section.header, section.footer):
            for paragraph in list(story.paragraphs):
                paragraph._element.getparent().remove(paragraph._element)
            for table in list(story.tables):
                table._element.getparent().remove(table._element)
        header = section.header.add_paragraph()
        header.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        run = header.add_run(identity.get("office_name", "Administración de fincas"))
        run.font.name = FUENTE_CARTA
        run.font.size = Pt(7)
        run.font.color.rgb = _rgb(C_GRIS_ETQ)
        logo_path = identity.get("logo_path")
        if logo_path and os.path.isfile(str(logo_path)):
            try:
                header.alignment = WD_ALIGN_PARAGRAPH.LEFT
                header.add_run().add_picture(str(logo_path), width=Cm(1.0))
            except Exception:
                pass
        footer = section.footer.add_paragraph()
        footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = footer.add_run(identity.get("footer", "Atención de la comunidad"))
        run.font.name = FUENTE_CARTA
        run.font.size = Pt(7)
        run.font.color.rgb = _rgb(C_GRIS_ETQ)


def _caja_resultado(doc: Document, diferencia_total: float) -> None:
    """
    Caja de RESULTADO al principio: lo primero que ve el cliente.
    Explica en una frase qué significa el número.
    """
    if diferencia_total > 0.005:
        fondo, color_v = C_FONDO_PAGAR, C_ROJO
        titulo  = "RESULTADO: IMPORTE A PAGAR"
        importe = f"+ {_fmt_euros(abs(diferencia_total))}"
        explica = ("El importe que le corresponde por los gastos reales de la comunidad "
                   "ha sido superior a lo cobrado a cuenta durante el año. "
                   "La diferencia se cargará en su próximo recibo de comunidad.")
    elif diferencia_total < -0.005:
        fondo, color_v = C_FONDO_DEVOLVER, C_VERDE
        titulo  = "RESULTADO: IMPORTE A SU FAVOR"
        importe = f"− {_fmt_euros(abs(diferencia_total))}"
        explica = ("El importe que le corresponde por los gastos reales de la comunidad "
                   "ha sido inferior a lo cobrado a cuenta durante el año. "
                   "La diferencia se abonará en su próximo recibo de comunidad.")
    else:
        fondo, color_v = C_FONDO_CERO, C_PRIMARIO
        titulo  = "RESULTADO: SIN DIFERENCIA"
        importe = _fmt_euros(0)
        explica = ("El importe que le corresponde por los gastos reales de la comunidad "
                   "coincide con lo cobrado a cuenta durante el año. No hay regularización.")

    tabla = doc.add_table(rows=2, cols=2)
    _sin_borde_tabla(tabla)
    _set_ancho_col(tabla, 0, 10.0)
    _set_ancho_col(tabla, 1, 6.0)

    c_tit = tabla.rows[0].cells[0]
    _set_color_celda(c_tit, fondo)
    _escribir_celda(c_tit, titulo, negrita=True, tamaño=11,
                    color_hex=C_DARK, despues=0)

    c_imp = tabla.rows[0].cells[1]
    _set_color_celda(c_imp, fondo)
    _escribir_celda(c_imp, importe, negrita=True, tamaño=14,
                    color_hex=color_v,
                    alineacion=WD_ALIGN_PARAGRAPH.RIGHT, despues=0)

    c_exp = tabla.rows[1].cells[0].merge(tabla.rows[1].cells[1])
    _set_color_celda(c_exp, fondo)
    _escribir_celda(c_exp, explica, negrita=False, tamaño=8,
                    color_hex=C_GRIS_ETQ, despues=1)

    p = doc.add_paragraph()
    _quitar_margen_parrafo(p, 3)


def _bloque_info(doc: Document, fecha_str: str, vivienda: str,
                 propietario: str, participacion_registral,
                 coeficiente_aplicado, periodo: str,
                 fecha_ini: str, fecha_fin: str, n_vecinos: int) -> None:
    """Ficha del propietario en dos columnas de etiqueta|valor."""
    registral_txt = (
        f"{participacion_registral:.4f}".rstrip("0").rstrip(".").replace(".", ",") + " %"
        if isinstance(participacion_registral, (int, float)) else "—"
    )
    aplicado_txt = (
        f"{coeficiente_aplicado * 100:.4f}".rstrip("0").rstrip(".").replace(".", ",") + " %"
        if isinstance(coeficiente_aplicado, (int, float)) else "—"
    )
    pares = [
        ("Propietario/a", propietario,             "Fecha",   fecha_str),
        ("Vivienda",      vivienda,                "Periodo", periodo),
        ("Participación registral", registral_txt, "Coeficiente aplicado", aplicado_txt),
        ("Desde", _fmt_fecha_es(fecha_ini), "Hasta", _fmt_fecha_es(fecha_fin)),
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
        _escribir_celda(fila.cells[0], etq1, tamaño=8, color_hex=C_GRIS_ETQ, despues=1)
        _escribir_celda(fila.cells[1], str(val1), negrita=True, tamaño=8.5,
                        color_hex=C_DARK, despues=1)
        _escribir_celda(fila.cells[2], etq2, tamaño=8, color_hex=C_GRIS_ETQ, despues=1)
        _escribir_celda(fila.cells[3], str(val2), negrita=True, tamaño=8.5,
                        color_hex=C_DARK, despues=1)

    p = doc.add_paragraph()
    _quitar_margen_parrafo(p, 3)


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
    _quitar_margen_parrafo(p, 3)


def _tabla_suministros_combinada(doc: Document, calef: dict, acs: dict) -> None:
    """
    Una sola tabla con dos columnas de datos: CALEFACCIÓN y ACS.
    Ahorra media página frente a dos tablas separadas.
    """
    _seccion_titulo(doc, "DETALLE DE CONSUMOS E IMPORTES")

    def _v(d, clave, defecto=0.0):
        return (d or {}).get(clave, defecto) or defecto

    def _precio(d, unidad):
        cons = _v(d, "consumo_real")
        if not cons:
            return "—"
        return f"{_v(d, 'importe_real') / cons:.4f} €/{unidad}".replace(".", ",")

    def _dif_txt(d):
        dif = _v(d, "diferencia")
        if dif < -0.005:
            return f"− {_fmt_euros(abs(dif))}", C_VERDE
        if dif > 0.005:
            return f"+ {_fmt_euros(dif)}", C_ROJO
        return "0,00 €", C_DARK

    filas = [
        ("Lectura inicial",
         _fmt_consumo(_v(calef, "lectura_inicial"), "kWh") if calef else "—",
         _fmt_consumo(_v(acs, "lectura_inicial"), "m³") if acs else "—", None, None),
        ("Lectura final",
         _fmt_consumo(_v(calef, "lectura_final"), "kWh") if calef else "—",
         _fmt_consumo(_v(acs, "lectura_final"), "m³") if acs else "—", None, None),
        ("Consumo del periodo",
         _fmt_consumo(_v(calef, "consumo_real"), "kWh") if calef else "—",
         _fmt_consumo(_v(acs, "consumo_real"), "m³") if acs else "—", None, None),
        ("Precio medio aplicado",
         _precio(calef, "kWh") if calef else "—",
         _precio(acs, "m³") if acs else "—", None, None),
        ("Cobrado a cuenta",
         _fmt_euros(_v(calef, "importe_cobrado")) if calef else "—",
         _fmt_euros(_v(acs, "importe_cobrado")) if acs else "—", None, None),
        ("Coste real",
         _fmt_euros(_v(calef, "importe_real")) if calef else "—",
         _fmt_euros(_v(acs, "importe_real")) if acs else "—", None, None),
    ]

    tabla = doc.add_table(rows=len(filas) + 2, cols=3)
    _sin_borde_tabla(tabla)
    for idx, ancho in enumerate([6.4, 4.8, 4.8]):
        _set_ancho_col(tabla, idx, ancho)

    # Cabecera
    for j, txt in enumerate(["Concepto", "CALEFACCIÓN", "ACS"]):
        c = tabla.rows[0].cells[j]
        _set_color_celda(c, C_PRIMARIO_SUAVE)
        _escribir_celda(c, txt, negrita=True, tamaño=8.5, color_hex=C_PRIMARIO,
                        alineacion=WD_ALIGN_PARAGRAPH.RIGHT if j else WD_ALIGN_PARAGRAPH.LEFT,
                        despues=1)

    for i, (concepto, v_cal, v_acs, _x, _y) in enumerate(filas):
        fila = tabla.rows[i + 1]
        if i % 2 == 1:
            for c in fila.cells:
                _set_color_celda(c, C_ZEBRA)
        neg = concepto == "Coste real"
        _escribir_celda(fila.cells[0], concepto, tamaño=8.5, negrita=neg,
                        color_hex=C_DARK, despues=1)
        _escribir_celda(fila.cells[1], v_cal, tamaño=8.5, negrita=neg,
                        color_hex=C_DARK, alineacion=WD_ALIGN_PARAGRAPH.RIGHT, despues=1)
        _escribir_celda(fila.cells[2], v_acs, tamaño=8.5, negrita=neg,
                        color_hex=C_DARK, alineacion=WD_ALIGN_PARAGRAPH.RIGHT, despues=1)

    # Fila DIFERENCIA
    fila = tabla.rows[len(filas) + 1]
    for c in fila.cells:
        _set_color_celda(c, C_PRIMARIO_SUAVE)
    t_cal, col_cal = _dif_txt(calef) if calef else ("—", C_DARK)
    t_acs, col_acs = _dif_txt(acs) if acs else ("—", C_DARK)
    _escribir_celda(fila.cells[0], "DIFERENCIA", negrita=True, tamaño=8.5,
                    color_hex=C_PRIMARIO, despues=1)
    _escribir_celda(fila.cells[1], t_cal, negrita=True, tamaño=9,
                    color_hex=col_cal, alineacion=WD_ALIGN_PARAGRAPH.RIGHT, despues=1)
    _escribir_celda(fila.cells[2], t_acs, negrita=True, tamaño=9,
                    color_hex=col_acs, alineacion=WD_ALIGN_PARAGRAPH.RIGHT, despues=1)

    p = doc.add_paragraph()
    _quitar_margen_parrafo(p, 3)


# ---------------------------------------------------------------------------
# GRÁFICAS (matplotlib)
# ---------------------------------------------------------------------------

def _estilo_ejes(ax):
    for lado in ("top", "right", "left"):
        ax.spines[lado].set_visible(False)
    ax.spines["bottom"].set_color("#E5E7EB")
    ax.tick_params(colors=C_GRIS_ETQ, labelsize=7.5)
    ax.set_facecolor(C_FONDO_GRAF)


def _grafica_unica(calef: dict, acs: dict, medias: dict) -> Optional[str]:
    """
    Una sola tira de 3 paneles (compacta, para carta de 1 página):
      1. Barras cobrado vs coste real por suministro
      2. Donut de distribución del coste real
      3. Su consumo respecto a la media de la comunidad (en %)
    """
    if not _HAS_MPL:
        print("⚠️  Gráfica omitida: matplotlib/numpy no disponibles "
              "(pip install matplotlib numpy)")
        return None
    try:
        calef_cob  = (calef or {}).get("importe_cobrado", 0.0) or 0.0
        calef_real = (calef or {}).get("importe_real",    0.0) or 0.0
        acs_cob    = (acs   or {}).get("importe_cobrado", 0.0) or 0.0
        acs_real   = (acs   or {}).get("importe_real",    0.0) or 0.0
        c_cons     = (calef or {}).get("consumo_real",    0.0) or 0.0
        a_cons     = (acs   or {}).get("consumo_real",    0.0) or 0.0
        m_cal      = (medias or {}).get("calef_consumo",  0.0) or 0.0
        m_acs      = (medias or {}).get("acs_consumo",    0.0) or 0.0

        tiene_media = (m_cal > 0 and c_cons >= 0) or (m_acs > 0 and a_cons >= 0)
        n_paneles = 3 if tiene_media else 2

        fig, ejes = plt.subplots(
            1, n_paneles, figsize=(16.4 / 2.54, 4.4 / 2.54), dpi=150,
            gridspec_kw={"width_ratios": [1.35, 0.85, 1.15][:n_paneles]})
        fig.patch.set_facecolor(C_FONDO_GRAF)
        ax1, ax2 = ejes[0], ejes[1]

        # ── Panel 1: barras cobrado vs real ──
        grupos   = ["Calefacción", "ACS"]
        cobrados = [calef_cob, acs_cob]
        reales   = [calef_real, acs_real]
        x = np.arange(len(grupos))
        ancho = 0.36
        b1 = ax1.bar(x - ancho / 2, cobrados, ancho, color=C_BARRA_COB,
                     label="Cobrado", zorder=3)
        b2 = ax1.bar(x + ancho / 2, reales, ancho, color=C_BARRA_REAL,
                     label="Real", zorder=3)
        tope = max(cobrados + reales) or 1
        for barras in (b1, b2):
            for bar in barras:
                v = bar.get_height()
                if v > 0:
                    ax1.text(bar.get_x() + bar.get_width() / 2, v + tope * 0.03,
                             _fmt_euros(v).replace(",00 €", "€").replace(" €", "€"),
                             ha="center", va="bottom", fontsize=6, color=C_DARK)
        ax1.set_xticks(x)
        ax1.set_xticklabels(grupos, fontsize=7.5, color=C_DARK)
        ax1.set_yticks([])
        ax1.set_ylim(0, tope * 1.24)
        _estilo_ejes(ax1)
        ax1.legend(fontsize=6, frameon=False, loc="upper center",
                   bbox_to_anchor=(0.5, 1.22), ncol=2)
        ax1.set_title("Cobrado vs real", fontsize=7.5, color=C_GRIS_ETQ, pad=16)
        ax1.grid(axis="y", color="#F1F5F9", zorder=0)

        # ── Panel 2: donut ──
        vals = [max(calef_real, 0), max(acs_real, 0)]
        if sum(vals) > 0:
            wedges, _t, autos = ax2.pie(
                vals, colors=[C_BARRA_REAL, C_ACENTO_2], startangle=90,
                autopct=lambda p: f"{p:.0f}%" if p > 5 else "",
                pctdistance=0.72,
                wedgeprops=dict(width=0.44, edgecolor="white"))
            for a in autos:
                a.set_fontsize(6); a.set_color("white"); a.set_fontweight("bold")
            ax2.text(0, 0.02, _fmt_euros(sum(vals)), ha="center", va="center",
                     fontsize=7, color=C_DARK, fontweight="bold")
            ax2.legend(wedges, ["Calef.", "ACS"], fontsize=6, frameon=False,
                       loc="upper center", bbox_to_anchor=(0.5, 1.24), ncol=2)
            ax2.set_title("Su coste real", fontsize=7.5, color=C_GRIS_ETQ, pad=16)
        else:
            ax2.axis("off")

        # ── Panel 3: consumo vs media (en % de la media) ──
        if tiene_media:
            ax3 = ejes[2]
            cats, pcts, etiq = [], [], []
            if m_cal > 0:
                cats.append("Calefacción"); pcts.append(c_cons / m_cal * 100)
                etiq.append(_fmt_consumo(c_cons, "kWh"))
            if m_acs > 0:
                cats.append("ACS"); pcts.append(a_cons / m_acs * 100)
                etiq.append(_fmt_consumo(a_cons, "m³"))
            colores = [C_BARRA_REAL if p <= 100 else C_ACENTO_2 for p in pcts]
            barras = ax3.barh(cats, pcts, 0.5, color=colores, zorder=3)
            tope3 = max(pcts + [100]) * 1.35
            for bar, p, e in zip(barras, pcts, etiq):
                ax3.text(bar.get_width() + tope3 * 0.02,
                         bar.get_y() + bar.get_height() / 2,
                         f"{e} ({p:.0f}%)", va="center", ha="left",
                         fontsize=6, color=C_DARK)
            ax3.axvline(100, color=C_GRIS_ETQ, lw=0.8, ls="--", zorder=4)
            import matplotlib.transforms as mtrans
            _tr = mtrans.blended_transform_factory(ax3.transData, ax3.transAxes)
            ax3.text(100, -0.14, "media comunidad", transform=_tr, ha="center",
                     fontsize=5.5, color=C_GRIS_ETQ, clip_on=False)
            ax3.set_xlim(0, tope3)
            ax3.set_xticks([])
            ax3.invert_yaxis()
            ax3.tick_params(axis="y", labelsize=7, colors=C_DARK)
            _estilo_ejes(ax3)
            ax3.spines["bottom"].set_visible(False)
            ax3.set_title("Su consumo vs media", fontsize=7.5,
                          color=C_GRIS_ETQ, pad=8)

        plt.tight_layout(pad=0.4, w_pad=1.2)
        tmp = tempfile.mktemp(suffix=".png")
        fig.savefig(tmp, dpi=150, bbox_inches="tight", facecolor=C_FONDO_GRAF)
        plt.close(fig)
        return tmp
    except Exception:
        import traceback
        print("⚠️  Error generando la gráfica (se usará tabla de respaldo):")
        traceback.print_exc()
        return None


def _insertar_imagen(doc: Document, ruta_png: str, ancho_cm: float = 16.0):
    if ruta_png and os.path.exists(ruta_png):
        try:
            p = doc.add_paragraph()
            _quitar_margen_parrafo(p, 3)
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.add_run().add_picture(str(ruta_png), width=Cm(ancho_cm))
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
        _escribir_celda(fila.cells[0], nom, tamaño=8.5, color_hex=C_DARK, despues=1)
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
        concepto = f"{nombres.get(tipo, tipo)} ({cf['meses']} × {precio_txt} €)"
        dif = cf["diferencia"]
        color = C_VERDE if dif < -0.005 else (C_ROJO if dif > 0.005 else C_DARK)
        signo = "− " if dif < -0.005 else ("+ " if dif > 0.005 else "")
        _escribir_celda(fila.cells[0], concepto, tamaño=8.5, color_hex=C_DARK, despues=1)
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
    _quitar_margen_parrafo(p, 3)


def _tabla_totales(doc: Document, dif_fija: float, dif_calef: float,
                   dif_acs: float, total_general: float) -> None:
    """Resumen de totales por concepto + total general."""
    _seccion_titulo(doc, "RESUMEN DE LA REGULARIZACIÓN")
    conceptos = [("cuota fija", dif_fija), ("calefacción", dif_calef),
                 ("ACS", dif_acs)]
    conceptos = [(n, v) for n, v in conceptos if abs(v) > 0.004 or n != "cuota fija"]

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
        _escribir_celda(fila.cells[0], f"Diferencia {nombre}",
                        tamaño=8.5, color_hex=C_DARK, despues=1)
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
    _quitar_margen_parrafo(p, 3)


# ---------------------------------------------------------------------------
# GENERACIÓN DE UNA CARTA
# ---------------------------------------------------------------------------

def generar_carta(datos: dict, ruta_plantilla: str, ruta_salida: str) -> str:
    """Genera una carta Word moderna para un vecino."""
    import shutil
    shutil.copy2(ruta_plantilla, ruta_salida)
    doc = Document(ruta_salida)

    # El contenido se reconstruye; la identidad heredada no debe filtrarse a
    # otro despacho aunque la plantilla proceda de una instalación anterior.
    for par in list(doc.paragraphs):
        par._element.getparent().remove(par._element)
    for tabla in list(doc.tables):
        tabla._element.getparent().remove(tabla._element)
    identity = datos.get("identity") or {}
    _limpiar_identidad_plantilla(doc, identity)

    vecino  = datos["vecino"]
    periodo = datos["periodo"]
    calef   = datos.get("calef") or {}
    acs     = datos.get("acs")   or {}
    total   = datos["total"]
    medias  = datos.get("medias") or {}
    cuota_fija = datos.get("cuota_fija") or {}
    conceptos_activos = list(datos.get("conceptos") or ())
    dif_fija   = round(sum(cf["diferencia"] for cf in cuota_fija.values()), 2)
    total_general = round(total["diferencia"] + dif_fija, 2)
    if conceptos_activos:
        total_general = round(sum(float(item.get("diferencia", 0.0) or 0.0)
                                  for item in conceptos_activos), 2)
    nombre_comunidad = datos.get("comunidad", "COMUNIDAD DE PROPIETARIOS")

    nombre_periodo = periodo["nombre"]
    hoy       = datetime.now()
    city = str(identity.get("city", "")).strip()
    fecha_text = f"{hoy.day} de {MESES_ES[hoy.month]} de {hoy.year}"
    fecha_str = f"{city}, {fecha_text}" if city else fecha_text

    # 1. CABECERA
    _bloque_cabecera(
        doc, nombre_comunidad, nombre_periodo,
        identity.get("office_name", "Administración de fincas"),
    )

    # 2. RESULTADO DESTACADO (lo más importante, primero)
    _caja_resultado(doc, total_general)

    # 3. FICHA DEL PROPIETARIO
    _bloque_info(doc,
                 fecha_str   = fecha_str,
                 vivienda    = vecino["vivienda"],
                 propietario = vecino["nombre"],
                 participacion_registral = vecino.get("participacion_registral", vecino.get("coeficiente")),
                 coeficiente_aplicado = vecino.get("coeficiente_aplicado"),
                 periodo     = nombre_periodo,
                 fecha_ini   = periodo.get("fecha_inicio") or "",
                 fecha_fin   = periodo.get("fecha_fin") or "",
                 n_vecinos   = medias.get("n_vecinos", 0))

    # 4. INTRO BREVE (una línea)
    intro = (
        "Conforme al acuerdo de la última Junta General, le remitimos la regularización "
        f"del periodo {nombre_periodo} con los conceptos detallados a continuación."
        if conceptos_activos else
        "Conforme al acuerdo de la última Junta General, le remitimos la regularización "
        f"de Calefacción y ACS del periodo {nombre_periodo} según las lecturas de su contador individual."
    )
    _parrafo(doc, intro, tam=8.5, despues=6)

    # 5. DETALLE COMBINADO (una sola tabla: calefacción + ACS)
    if conceptos_activos:
        _tabla_conceptos_activos(doc, conceptos_activos)
    else:
        _tabla_suministros_combinada(doc, calef, acs)
        # 5b. CUOTA FIJA (si la comunidad la regulariza)
        _tabla_cuota_fija(doc, cuota_fija)
        # 5c. RESUMEN DE TOTALES
        _tabla_totales(doc,
                       dif_fija  = dif_fija,
                       dif_calef = calef.get("diferencia", 0.0) if calef else 0.0,
                       dif_acs   = acs.get("diferencia", 0.0) if acs else 0.0,
                       total_general = total_general)

    # 6. COMPARATIVA DE CONSUMO (franjas de vecinos + histórico propio)
    grafica = datos.get("consumo_grafica") or {}
    png = None
    if render_consumption_charts and grafica:
        png = render_consumption_charts(
            tempfile.mktemp(suffix=".png"),
            owner_consumption=grafica.get("owner_consumption", 0),
            neighbor_consumptions=grafica.get("neighbor_consumptions", []),
            history=grafica.get("history", []),
            unit=grafica.get("unit", "m³"),
        )
    if not png and not conceptos_activos:
        png = _grafica_unica(calef, acs, medias)
    if png:
        _insertar_imagen(doc, png, 16.2)
    elif not conceptos_activos:
        calef_cob, calef_real = calef.get("importe_cobrado", 0.0), calef.get("importe_real", 0.0)
        acs_cob,   acs_real   = acs.get("importe_cobrado", 0.0),   acs.get("importe_real", 0.0)
        _bloque_comparativa_fallback(doc, calef_cob, calef_real, acs_cob, acs_real)

    # 7. CIERRE Y FIRMA (compacto)
    _parrafo(doc,
             "Puede verificar los datos con sus recibos y comunicarnos cualquier discrepancia "
             "en un plazo de 30 días. Quedamos a su disposición. Reciba un cordial saludo.",
             tam=8.5, despues=4)
    _parrafo(doc, identity.get("signature", "La Administración"),
             tam=8.5, negrita=True, alin=WD_ALIGN_PARAGRAPH.RIGHT, despues=0)

    doc.save(ruta_salida)
    return ruta_salida


# ---------------------------------------------------------------------------
# GENERACIÓN EN LOTE — todas las cartas del periodo
# ---------------------------------------------------------------------------

def generar_todas_las_cartas(ruta_bd: str,
                              id_comunidad: int,
                              nombre_periodo: str,
                              ruta_plantilla: str,
                              carpeta_salida: str,
                              selected_concepts: tuple[str, ...] | None = None) -> dict:
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

    todos = obtener_todos_los_vecinos_con_repartos(con, id_comunidad, id_periodo, selected_concepts)
    consumo_vecinos = [
        float((datos.get("acs") or {}).get("consumo_real") or 0.0)
        for datos in todos
    ]
    # Conserva ejercicios anteriores cuando existen lecturas acumuladas.
    historicos_por_vecino = {}
    for datos in todos:
        owner_id = datos.get("id_propietario")
        historicos = con.execute(
            """
            SELECT per.nombre, ini.valor_acumulado, fin.valor_acumulado
            FROM periodos per
            JOIN lecturas_vecino ini ON ini.id_periodo=per.id_periodo
                AND ini.id_propietario=? AND ini.tipo='ACS' AND ini.fecha_lectura=per.fecha_inicio
            JOIN lecturas_vecino fin ON fin.id_periodo=per.id_periodo
                AND fin.id_propietario=? AND fin.tipo='ACS' AND fin.fecha_lectura=per.fecha_fin
            WHERE per.id_comunidad=? AND per.id_periodo<>?
            ORDER BY per.fecha_inicio
            """,
            (owner_id, owner_id, id_comunidad, id_periodo),
        ).fetchall()
        historicos_por_vecino[owner_id] = [
            (row[0], max(0.0, float(row[2]) - float(row[1]))) for row in historicos
        ]
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
        datos["consumo_grafica"] = {
            "owner_consumption": float((datos.get("acs") or {}).get("consumo_real") or 0.0),
            "neighbor_consumptions": consumo_vecinos,
            "history": historicos_por_vecino.get(datos.get("id_propietario"), []),
            "unit": "m³",
        }

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
