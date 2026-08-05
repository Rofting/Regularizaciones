"""
excel_writer.py
===============
Vuelca datos de la BD SQLite en el Excel Maestro (Comunidad_XXX.xlsx),
añadiendo un nuevo bloque de año en cada pestaña o actualizando uno existente.

PESTAÑAS QUE ACTUALIZA:
    GAS           → facturas de gas (Baser, Galp, etc.)
    ELECTRICIDAD  → facturas de luz (Endesa, Energía XXI, etc.)
    AGUA          → facturas de agua
    OTROS GASTOS  → mantenimiento + reparaciones amortizadas
    LECTURAS ACS M3      → resumen de lecturas Metrigest ACS
    LECTURAS CALEF KWH   → resumen de lecturas Metrigest Calef

DISEÑO:
    - Nunca borra celdas con fórmulas existentes — solo inserta filas nuevas
    - Busca el bloque del año (ej: '2025-2026') en la col A de cada pestaña
    - Si no existe, crea el bloque al final respetando el patrón de años anteriores
    - Si existe, añade las filas nuevas dentro del bloque existente (antes de la Suma)
    - Las fórmulas de Suma, Días, precios medios se escriben en formato Excel
    - Preserva el formato de número de las celdas contiguas copiando estilos
"""

import os
import shutil
import sqlite3
from datetime import datetime, date
from typing import Optional

import openpyxl
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, numbers
from openpyxl.utils import get_column_letter


class CupsMismatchError(Exception):
    """
    El CUPS del PDF no coincide con el CUPS registrado en la celda C4 del Excel.
    Detiene el volcado para evitar contaminar la hoja de una comunidad incorrecta.
    """
    def __init__(self, cups_pdf: str, cups_excel: str, hoja: str = ""):
        self.cups_pdf   = cups_pdf
        self.cups_excel = cups_excel
        self.hoja       = hoja
        super().__init__(
            f"CUPS no coincide en hoja '{hoja}': "
            f"PDF='{cups_pdf}' | Excel C4='{cups_excel}'"
        )


# ---------------------------------------------------------------------------
# MAPA DE COLUMNAS — exactamente como están en el Excel Maestro
# ---------------------------------------------------------------------------

# GAS: cabecera en fila 8, datos desde fila 11
GAS_COLS = {
    "fecha_factura":   "B",   # Fecha fra.
    "dias":            "C",   # Días (=F-D)
    "fecha_inicio":    "D",   # Fecha inicial
    "lec_ini":         "E",   # Lectura inicial
    "fecha_fin":       "F",   # Fecha final
    "lec_fin":         "G",   # Lectura final
    "consumo_m3":      "H",   # Consumo m3
    "consumo_kwh":     "I",   # Consumo kWh
    "termino_fijo":    "J",   # Término fijo €
    "termino_var":     "K",   # Término variable €
    "importe_total":   "L",   # Importe TOTAL
    "fecha_pago":      "M",   # Fecha pago
    "nota":            "N",   # Nota (proveedor)
}
GAS_FILA_CABECERA = 8
GAS_FILA_TITULO   = 2   # "FACTURAS SUMINISTRO GAS NATURAL"

# ELECTRICIDAD: cabecera en fila 8, datos desde fila 11
ELEC_COLS = {
    "fecha_factura":   "B",   # Fecha fra.
    "fecha_inicio":    "C",   # Fecha inicial
    "fecha_fin":       "D",   # Fecha final
    "consumo_kwh":     "E",   # Consumo kWh
    "termino_fijo":    "F",   # Término fijo €
    "termino_var":     "G",   # Término variable €
    "importe_total":   "H",   # Importe TOTAL
    "fecha_pago":      "I",   # Fecha pago
    "nota":            "J",   # Nota (proveedor)
}
ELEC_FILA_CABECERA = 8

# AGUA: cabecera en fila variable, datos debajo
AGUA_COLS = {
    "fecha_factura":   "B",
    "periodo":         "C",
    "fecha_inicio":    "D",
    "lec_ini":         "E",
    "fecha_fin":       "F",
    "lec_fin":         "G",
    "consumo_m3":      "H",
    "total":           "T",   # TOTAL IVA incl.
    "c_variable":      "V",   # C. Variable
    "c_fija":          "W",   # C. Fija
}
AGUA_FILA_CABECERA = 12  # estimada, se detecta dinámicamente

# LECTURAS ACS/CALEF: cabecera en fila 6
LECT_COLS = {
    "fecha":           "B",
    "fecha_ini":       "C",
    "lec_ini_val":     "D",
    "fecha_fin":       "E",
    "lec_fin_val":     "F",
    "consumo":         "G",
    "cuota_var":       "H",
    "cuota_fija":      "I",
    "total":           "J",
    "precio_unit":     "L",
}

# OTROS GASTOS: columnas de resumen por ejercicio
OTROS_FILA_CABECERA_RESUMEN = 5   # fila con "€/MES", "2020-2021"...


# ---------------------------------------------------------------------------
# UTILIDADES
# ---------------------------------------------------------------------------

def _col_num(letra: str) -> int:
    """Convierte letra de columna a número (A=1, B=2...)."""
    from openpyxl.utils import column_index_from_string
    return column_index_from_string(letra)


def _fecha_excel(valor) -> Optional[date]:
    """Convierte string ISO, datetime o date a date Python."""
    if valor is None:
        return None
    if isinstance(valor, (datetime, date)):
        return valor if isinstance(valor, date) else valor.date()
    try:
        return datetime.strptime(str(valor)[:10], "%Y-%m-%d").date()
    except ValueError:
        return None


def _copiar_estilo(celda_origen, celda_destino):
    """Copia el formato numérico y fuente de una celda a otra."""
    if celda_origen.has_style:
        celda_destino.font       = Font(
            name=celda_origen.font.name,
            size=celda_origen.font.size,
            bold=celda_origen.font.bold
        )
        celda_destino.number_format = celda_origen.number_format
        celda_destino.alignment     = Alignment(
            horizontal=celda_origen.alignment.horizontal
        )


def _encontrar_bloque_año(ws, nombre_año: str) -> Optional[int]:
    """
    Busca la fila donde está el bloque de un año (ej: '2025-2026') en col A.
    Devuelve el número de fila o None si no existe.
    """
    for row in ws.iter_rows(min_col=1, max_col=1, values_only=False):
        cell = row[0]
        if cell.value and str(cell.value).strip() == nombre_año:
            return cell.row
    return None


def _fin_datos_hoja(ws, max_col: int = 14) -> int:
    """Última fila con algún valor en las columnas A..N (para colocar bloques
    nuevos sin pisar los anteriores, aunque falte la fila Suma)."""
    ultima = 0
    for fila in ws.iter_rows(min_col=1, max_col=max_col):
        for c in fila:
            if c.value is not None:
                ultima = max(ultima, c.row)
    return ultima


def _meses_entre(fecha_ini, fecha_fin) -> int:
    """Meses (redondeados) entre dos fechas ISO/datetime. Mínimo 1, defecto 3."""
    from datetime import date as _d, datetime as _dt
    try:
        def _p(x):
            if isinstance(x, _dt):
                return x.date()
            if isinstance(x, _d):
                return x
            return _d.fromisoformat(str(x)[:10])
        dias = (_p(fecha_fin) - _p(fecha_ini)).days
        return max(1, round(dias / 30.44))
    except Exception:
        return 3


def _encontrar_fila_suma(ws, fila_inicio_bloque: int) -> Optional[int]:
    """
    Busca la fila 'Suma ….' dentro de un bloque (en col B), a partir de fila_inicio.
    """
    for i in range(fila_inicio_bloque + 1, fila_inicio_bloque + 60):
        cell = ws.cell(row=i, column=2)  # col B
        if cell.value and "Suma" in str(cell.value):
            return i
    return None


def _siguiente_fila_libre_bloque(ws, fila_suma: int) -> int:
    """Devuelve la fila justo antes de la fila Suma."""
    return fila_suma  # insertamos antes de la fila Suma


def _ultimo_bloque_año(ws) -> tuple[int, str]:
    """
    Devuelve (fila, nombre_año) del último bloque de año en col A.
    """
    ultimo = (0, "")
    for row in ws.iter_rows(min_col=1, max_col=1, values_only=False):
        cell = row[0]
        if cell.value and str(cell.value).strip().count("-") == 1:
            partes = str(cell.value).strip().split("-")
            if len(partes) == 2 and partes[0].isdigit() and partes[1].isdigit():
                ultimo = (cell.row, str(cell.value).strip())
    return ultimo


# ---------------------------------------------------------------------------
# ESCRITURA EN PESTAÑA GAS
# ---------------------------------------------------------------------------

def _escribir_fila_gas(ws, fila: int, factura: dict, es_primera_del_bloque: bool):
    """Escribe una fila de factura de gas en la posición indicada."""
    f = factura

    # Col B: fecha factura
    ws[f"B{fila}"] = _fecha_excel(f.get("fecha_factura"))
    ws[f"B{fila}"].number_format = "DD/MM/YYYY"

    # Col C: días (fórmula =F-D)
    ws[f"C{fila}"] = f"=F{fila}-D{fila}"

    # Col D: fecha inicio
    ws[f"D{fila}"] = _fecha_excel(f.get("fecha_inicio"))
    ws[f"D{fila}"].number_format = "DD/MM/YYYY"

    # Col E: lectura inicial (vacía si no hay)
    if f.get("lec_ini"):
        ws[f"E{fila}"] = f.get("lec_ini")

    # Col F: fecha fin
    ws[f"F{fila}"] = _fecha_excel(f.get("fecha_fin"))
    ws[f"F{fila}"].number_format = "DD/MM/YYYY"

    # Col G: lectura final
    if f.get("lec_fin"):
        ws[f"G{fila}"] = f.get("lec_fin")

    # Col H: consumo m3
    if f.get("consumo_m3"):
        ws[f"H{fila}"] = f.get("consumo_m3")

    # Col I: consumo kWh
    if f.get("consumo_kwh"):
        ws[f"I{fila}"] = f.get("consumo_kwh")

    # Col J: término fijo
    if f.get("termino_fijo"):
        ws[f"J{fila}"] = f.get("termino_fijo")
        ws[f"J{fila}"].number_format = "#,##0.00"

    # Col K: término variable
    if f.get("termino_variable"):
        ws[f"K{fila}"] = f.get("termino_variable")
        ws[f"K{fila}"].number_format = "#,##0.00"

    # Col L: importe total — si J y K existen, fórmula; si no, valor directo
    if f.get("termino_fijo") and f.get("termino_variable"):
        ws[f"L{fila}"] = f"=SUM(J{fila}:K{fila})"
    else:
        ws[f"L{fila}"] = f.get("importe_total", 0)
    ws[f"L{fila}"].number_format = "#,##0.00"

    # Col N: nota (proveedor abreviado)
    ws[f"N{fila}"] = _abreviar_proveedor(f.get("proveedor", ""))

    # Col O: precio m3 (solo si existen término variable Y consumo — evita #DIV/0!)
    if f.get("consumo_m3") and f.get("termino_variable"):
        ws[f"O{fila}"] = f"=K{fila}/H{fila}"

    # Col Q: precio kWh (idem)
    if f.get("consumo_kwh") and f.get("termino_variable"):
        ws[f"Q{fila}"] = f"=K{fila}/I{fila}"

    # Cols T/U: consumo invierno / verano (según el mes de la fecha inicial).
    # Referencia la columna que realmente tenga dato: H (m3) si existe, si no
    # I (kWh) — algunos proveedores/periodos solo informan uno de los dos
    # (ej. Baser en kWh), y antes esto se quedaba en blanco si no había m3.
    col_consumo = "H" if f.get("consumo_m3") else ("I" if f.get("consumo_kwh") else None)
    if col_consumo:
        try:
            mes = int(str(f.get("fecha_inicio") or f.get("fecha_factura"))[5:7])
        except (TypeError, ValueError):
            mes = 0
        if mes in (6, 7, 8, 9):
            ws[f"U{fila}"] = f"={col_consumo}{fila}"       # verano
        elif mes:
            ws[f"T{fila}"] = f"={col_consumo}{fila}"       # invierno


def _escribir_fila_suma_gas(ws, fila_suma: int, fila_ini: int, fila_fin: int):
    """Escribe la fila 'Suma ….' de un bloque GAS."""
    ws[f"B{fila_suma}"] = "Suma …."
    ws[f"C{fila_suma}"] = f"=SUM(C{fila_ini}:C{fila_fin})"
    ws[f"H{fila_suma}"] = f"=SUM(H{fila_ini}:H{fila_fin})"
    ws[f"I{fila_suma}"] = f"=SUM(I{fila_ini}:I{fila_fin})"
    ws[f"J{fila_suma}"] = f"=SUM(J{fila_ini}:J{fila_fin})"
    ws[f"K{fila_suma}"] = f"=SUM(K{fila_ini}:K{fila_fin})"
    ws[f"L{fila_suma}"] = f"=SUM(L{fila_ini}:L{fila_fin})"
    # Precio medio m3 y kWh (IFERROR: algunos periodos solo tienen datos de
    # uno de los dos, dejando el otro en 0 y provocando #DIV/0!)
    ws[f"P{fila_suma}"] = f"=IFERROR(K{fila_suma}/H{fila_suma},0)"
    ws[f"R{fila_suma}"] = f"=IFERROR(K{fila_suma}/I{fila_suma},0)"
    ws[f"P{fila_suma}"].number_format = "#,##0.0000"
    ws[f"R{fila_suma}"].number_format = "#,##0.0000"
    for col in ["H", "I", "J", "K", "L"]:
        ws[f"{col}{fila_suma}"].number_format = "#,##0.00"


# ---------------------------------------------------------------------------
# ESCRITURA EN PESTAÑA ELECTRICIDAD
# ---------------------------------------------------------------------------

def _escribir_fila_elec(ws, fila: int, factura: dict):
    """Escribe una fila de factura eléctrica."""
    f = factura

    ws[f"B{fila}"] = _fecha_excel(f.get("fecha_factura"))
    ws[f"B{fila}"].number_format = "DD/MM/YYYY"

    ws[f"C{fila}"] = _fecha_excel(f.get("fecha_inicio"))
    ws[f"C{fila}"].number_format = "DD/MM/YYYY"

    ws[f"D{fila}"] = _fecha_excel(f.get("fecha_fin"))
    ws[f"D{fila}"].number_format = "DD/MM/YYYY"

    if f.get("consumo_total"):
        ws[f"E{fila}"] = f.get("consumo_total")

    if f.get("termino_fijo"):
        ws[f"F{fila}"] = f.get("termino_fijo")
        ws[f"F{fila}"].number_format = "#,##0.00"

    if f.get("termino_variable"):
        ws[f"G{fila}"] = f.get("termino_variable")
        ws[f"G{fila}"].number_format = "#,##0.00"

    if f.get("termino_fijo") and f.get("termino_variable"):
        ws[f"H{fila}"] = f"=F{fila}+G{fila}"
    else:
        ws[f"H{fila}"] = f.get("importe_total", 0)
    ws[f"H{fila}"].number_format = "#,##0.00"

    ws[f"J{fila}"] = _abreviar_proveedor(f.get("proveedor", ""))

    if f.get("consumo_total"):
        ws[f"L{fila}"] = f"=G{fila}/E{fila}"
        ws[f"L{fila}"].number_format = "#,##0.0000"

    # Cols N/O: consumo invierno / verano (según el mes de la fecha inicial),
    # igual que T/U en GAS — lo usa el reparto estacional ACS/CALEFACCION.
    if f.get("consumo_total"):
        try:
            mes = int(str(f.get("fecha_inicio") or f.get("fecha_factura"))[5:7])
        except (TypeError, ValueError):
            mes = 0
        if mes in (6, 7, 8, 9):
            ws[f"O{fila}"] = f"=E{fila}"       # verano
        elif mes:
            ws[f"N{fila}"] = f"=E{fila}"       # invierno


def _escribir_fila_suma_elec(ws, fila_suma: int, fila_ini: int, fila_fin: int):
    ws[f"B{fila_suma}"] = "Suma …."
    ws[f"D{fila_suma}"] = f"=D{fila_fin}-C{fila_ini}"
    ws[f"E{fila_suma}"] = f"=SUM(E{fila_ini}:E{fila_fin})"
    ws[f"F{fila_suma}"] = f"=SUM(F{fila_ini}:F{fila_fin})"
    ws[f"G{fila_suma}"] = f"=SUM(G{fila_ini}:G{fila_fin})"
    ws[f"H{fila_suma}"] = f"=SUM(H{fila_ini}:H{fila_fin})"
    ws[f"M{fila_suma}"] = f"=G{fila_suma}/E{fila_suma}"
    ws[f"M{fila_suma}"].number_format = "#,##0.0000"
    for col in ["E", "F", "G", "H"]:
        ws[f"{col}{fila_suma}"].number_format = "#,##0.00"


# ---------------------------------------------------------------------------
# REPARTO ESTACIONAL ACS/CALEFACCION (GAS y ELECTRICIDAD)
# ---------------------------------------------------------------------------
#
# Cada factura ya se clasifica como consumo de invierno o de verano al
# escribirla (ver _escribir_fila_gas/_escribir_fila_elec, columnas T/U en GAS
# y N/O en ELECTRICIDAD). Lo que faltaba era, por cada bloque de año:
#   1. Sumar esas columnas en la fila Suma.
#   2. Calcular, en la fila JUSTO DEBAJO de la Suma, la fraccion invierno/
#      verano de ESE MISMO bloque (T_ratio = T/H, U_ratio = U/H) — es la
#      misma "fila ayudante" que usa el Excel de referencia.
#   3. Con esa fraccion, repartir el gasto del bloque entre ACS y
#      CALEFACCION: el verano se asume 100% ACS (no hay calefaccion
#      encendida), y el invierno se pondera con esa misma fraccion de verano
#      (formula tal cual la usa el Excel de referencia en GAS, replicada
#      igual aqui: V=(U+T*U_ratio)/H, W=T*(1-U_ratio)/H).
#   4. El %fijo/%variable del contrato (X/Y en GAS, R/S en ELECTRICIDAD) solo
#      se calcula de verdad en el periodo MAS RECIENTE (el unico con
#      facturas desglosadas); todos los periodos anteriores heredan ese
#      mismo % en cascada hacia atras — igual que en el Excel de referencia.
#      En ELECTRICIDAD, además, el propio Excel de referencia SOLO calcula
#      el reparto ACS/CALEF (P/Q) en el periodo más reciente y lo hereda
#      hacia atrás igual que el %fijo/variable (a diferencia de GAS, que sí
#      calcula V/W por su cuenta en cada periodo).
#
# Como el periodo N+1 no existe todavia cuando se escribe el periodo N (los
# periodos se vuelcan en orden cronologico), esto se recalcula ENTERO cada
# vez que se anade un periodo nuevo (ver llamada al final de
# actualizar_excel_maestro) — barato y siempre deja el libro consistente.

def _listar_bloques_año(ws) -> list[tuple[str, int, int]]:
    """
    Devuelve [(nombre_año, fila_bloque, fila_suma), ...] ordenados
    cronologicamente (los bloques ya se insertan en ese orden en col A).
    """
    bloques = []
    for row in ws.iter_rows(min_col=1, max_col=1):
        cell = row[0]
        v = cell.value
        if isinstance(v, str) and v.strip().count("-") == 1:
            partes = v.strip().split("-")
            if len(partes) == 2 and partes[0].isdigit() and partes[1].isdigit():
                fila_suma = _encontrar_fila_suma(ws, cell.row)
                if fila_suma:
                    bloques.append((v.strip(), cell.row, fila_suma))
    bloques.sort(key=lambda b: b[1])
    return bloques


def recalcular_reparto_estacional_gas(wb):
    """Rellena T/U/V/W/X/Y de la fila Suma (y su fila ayudante) de cada bloque de GAS."""
    if "GAS" not in wb.sheetnames:
        return
    ws = wb["GAS"]
    bloques = _listar_bloques_año(ws)
    if not bloques:
        return

    for (_, fila_bloque, fila_suma) in bloques:
        ws[f"T{fila_suma}"] = f"=SUM(T{fila_bloque + 1}:T{fila_suma - 1})"
        ws[f"U{fila_suma}"] = f"=SUM(U{fila_bloque + 1}:U{fila_suma - 1})"

        # Denominador H (m3) o I (kWh): algunos bloques mezclan facturas de
        # las dos unidades (cambio de proveedor a mitad de ejercicio: Baser
        # solo informa kWh). Un bloque asi puede dejar H con un residuo
        # pequeño de las pocas facturas en m3 e I con el grueso en kWh (o
        # al reves) — usar "el que sea cero" como antes falla porque ninguno
        # de los dos es realmente cero. Se usa el MAYOR de los dos: el mas
        # pequeño es residuo de la unidad minoritaria de ese bloque.
        d = f"IF(H{fila_suma}>I{fila_suma},H{fila_suma},I{fila_suma})"

        # Fila ayudante (justo debajo de la Suma, en el hueco antes del
        # siguiente bloque): fraccion invierno/verano de ESTE bloque.
        fila_ayuda = fila_suma + 1
        ws[f"T{fila_ayuda}"] = f"=IFERROR(T{fila_suma}/({d}),0)"
        ws[f"U{fila_ayuda}"] = f"=IFERROR(U{fila_suma}/({d}),0)"

        ws[f"V{fila_suma}"] = f"=IFERROR((U{fila_suma}+T{fila_suma}*U{fila_ayuda})/({d}),0)"
        ws[f"W{fila_suma}"] = f"=IFERROR(T{fila_suma}*(1-U{fila_ayuda})/({d}),0)"

    # %fijo/%variable: solo el periodo mas reciente lo calcula de verdad
    # (unico con termino fijo/variable desglosado); el resto hereda en cascada.
    _, _, fila_ultimo = bloques[-1]
    ws[f"X{fila_ultimo}"] = f"=IFERROR(J{fila_ultimo}/L{fila_ultimo},0)"
    ws[f"Y{fila_ultimo}"] = f"=1-X{fila_ultimo}"

    for i in range(len(bloques) - 2, -1, -1):
        _, _, fila = bloques[i]
        _, _, fila_sig = bloques[i + 1]
        ws[f"X{fila}"] = f"=X{fila_sig}"
        ws[f"Y{fila}"] = f"=Y{fila_sig}"


def recalcular_reparto_estacional_elec(wb):
    """
    Rellena N/O/P/Q/R/S de la fila Suma de cada bloque de ELECTRICIDAD.
    A diferencia de GAS, el Excel de referencia solo calcula el reparto
    ACS/CALEF (P/Q) y %fijo/variable (R/S) en el periodo MAS RECIENTE, y el
    resto de periodos simplemente heredan ese mismo valor (no cada uno
    calcula el suyo con su propia estacionalidad, como sí hace GAS).
    """
    if "ELECTRICIDAD" not in wb.sheetnames:
        return
    ws = wb["ELECTRICIDAD"]
    bloques = _listar_bloques_año(ws)
    if not bloques:
        return

    for (_, fila_bloque, fila_suma) in bloques:
        ws[f"N{fila_suma}"] = f"=SUM(N{fila_bloque + 1}:N{fila_suma - 1})"
        ws[f"O{fila_suma}"] = f"=SUM(O{fila_bloque + 1}:O{fila_suma - 1})"

    _, _, fila_ultimo = bloques[-1]
    fila_ayuda = fila_ultimo + 1
    ws[f"N{fila_ayuda}"] = f"=IFERROR(N{fila_ultimo}/E{fila_ultimo},0)"
    ws[f"O{fila_ayuda}"] = f"=IFERROR(O{fila_ultimo}/E{fila_ultimo},0)"

    ws[f"R{fila_ultimo}"] = f"=IFERROR(F{fila_ultimo}/H{fila_ultimo},0)"
    ws[f"S{fila_ultimo}"] = f"=IFERROR(G{fila_ultimo}/H{fila_ultimo},0)"
    ws[f"P{fila_ultimo}"] = f"=IFERROR((O{fila_ultimo}+N{fila_ultimo}*O{fila_ayuda})/E{fila_ultimo},0)"
    ws[f"Q{fila_ultimo}"] = f"=1-P{fila_ultimo}"

    for i in range(len(bloques) - 2, -1, -1):
        _, _, fila = bloques[i]
        ws[f"P{fila}"] = f"=P{fila_ultimo}"
        ws[f"Q{fila}"] = f"=Q{fila_ultimo}"
        ws[f"R{fila}"] = f"=R{fila_ultimo}"
        ws[f"S{fila}"] = f"=S{fila_ultimo}"


# ---------------------------------------------------------------------------
# ESCRITURA EN PESTAÑA AGUA
# ---------------------------------------------------------------------------

def _escribir_fila_agua(ws, fila: int, factura: dict):
    f = factura
    ws[f"B{fila}"] = _fecha_excel(f.get("fecha_factura"))
    ws[f"B{fila}"].number_format = "DD/MM/YYYY"

    ws[f"D{fila}"] = _fecha_excel(f.get("fecha_inicio"))
    ws[f"D{fila}"].number_format = "DD/MM/YYYY"

    ws[f"F{fila}"] = _fecha_excel(f.get("fecha_fin"))
    ws[f"F{fila}"].number_format = "DD/MM/YYYY"

    # Col C: etiqueta de periodo tipo "ago-sep-20"
    try:
        _MES = ["ene","feb","mar","abr","may","jun","jul","ago","sep","oct","nov","dic"]
        fi, ff = str(f.get("fecha_inicio")), str(f.get("fecha_fin"))
        ws[f"C{fila}"] = f"{_MES[int(fi[5:7])-1]}-{_MES[int(ff[5:7])-1]}-{ff[2:4]}"
    except (TypeError, ValueError, IndexError):
        pass

    ws[f"T{fila}"] = f.get("importe_total", 0)
    ws[f"T{fila}"].number_format = "#,##0.00"

    # Si vienen desglosados
    if f.get("c_variable"):
        ws[f"V{fila}"] = f.get("c_variable")
        ws[f"V{fila}"].number_format = "#,##0.00"
    if f.get("c_fija"):
        ws[f"W{fila}"] = f.get("c_fija")
        ws[f"W{fila}"].number_format = "#,##0.00"


def _escribir_fila_suma_agua(ws, fila_suma: int, fila_ini: int, fila_fin: int):
    """
    Escribe la fila 'Suma ….' de un bloque AGUA.
    Si alguna factura trae c_variable/c_fija desglosados, se suman esas
    columnas directamente; para el resto (el caso normal hoy) se aplica el
    mismo reparto 55% variable / 45% fijo que usa el Excel de referencia
    (V=0.55*T, W=0.45*T), que es lo que referencian ANALISIS!28-29
    ('CONSUMO AGUA FRIA' / 'FIJO AGUA FRIA').
    """
    ws[f"B{fila_suma}"] = "Suma …."
    ws[f"T{fila_suma}"] = f"=SUM(T{fila_ini}:T{fila_fin})"
    ws[f"V{fila_suma}"] = f"=0.55*T{fila_suma}"
    ws[f"W{fila_suma}"] = f"=0.45*T{fila_suma}"
    for col in ["T", "V", "W"]:
        ws[f"{col}{fila_suma}"].number_format = "#,##0.00"


# ---------------------------------------------------------------------------
# ESCRITURA EN PESTAÑA LECTURAS ACS / CALEF
# ---------------------------------------------------------------------------

def _escribir_bloque_lecturas(ws, fila_inicio_datos: int,
                               lecturas: list[dict], tipo: str,
                               precio_unitario: float):
    """
    Escribe un conjunto de lecturas Metrigest.
    Cada lectura tiene: fecha, fecha_ant, val_ant, fecha_act, val_act, consumo.
    Alterna filas: fila de consumo variable (cuota_var) + fila de cuota fija.
    """
    fila = fila_inicio_datos
    cuota_fija_mes = 900.0 if tipo == "ACS" else 1500.0

    for lec in lecturas:
        consumo = lec.get("consumo", 0)

        # Fila de consumo variable
        ws[f"B{fila}"] = _fecha_excel(lec.get("fecha_act"))
        ws[f"B{fila}"].number_format = "DD/MM/YYYY"

        ws[f"C{fila}"] = _fecha_excel(lec.get("fecha_ant"))
        ws[f"C{fila}"].number_format = "DD/MM/YYYY"

        ws[f"F{fila}"] = lec.get("val_act", 0)
        ws[f"D{fila}"] = lec.get("val_ant", 0)

        if consumo and consumo > 0:
            cuota_var = round(consumo * precio_unitario, 2)
            ws[f"G{fila}"] = f"=H{fila}/L{fila}"
            ws[f"H{fila}"] = cuota_var
            ws[f"H{fila}"].number_format = "#,##0.00"
            ws[f"J{fila}"] = f"=SUM(H{fila}:I{fila})"
            ws[f"L{fila}"] = precio_unitario

        fila += 1

        # Fila de cuota fija (mes siguiente)
        ws[f"I{fila}"] = cuota_fija_mes
        ws[f"I{fila}"].number_format = "#,##0.00"
        ws[f"J{fila}"] = f"=SUM(H{fila}:I{fila})"
        fila += 1

    # Fila Suma
    fila_suma = fila
    ws[f"F{fila_suma}"] = "Suma …."
    fila_datos_fin = fila - 1
    ws[f"G{fila_suma}"] = f"=SUM(G{fila_inicio_datos}:G{fila_datos_fin})"
    ws[f"H{fila_suma}"] = f"=SUM(H{fila_inicio_datos}:H{fila_datos_fin})"
    ws[f"I{fila_suma}"] = f"=SUM(I{fila_inicio_datos}:I{fila_datos_fin})"
    ws[f"J{fila_suma}"] = f"=SUM(J{fila_inicio_datos}:J{fila_datos_fin})"
    ws[f"L{fila_suma}"] = precio_unitario

    return fila_suma + 1  # próxima fila libre


# ---------------------------------------------------------------------------
# ESCRITURA EN PESTAÑA OTROS GASTOS
# ---------------------------------------------------------------------------

def _actualizar_otros_gastos(ws, nombre_año: str,
                              facturas_mto: list[dict],
                              gastos_extra: list[dict]):
    """
    Actualiza la sección de resumen de Otros Gastos y añade reparaciones nuevas.
    """
    # Buscar la columna del año en la fila de cabecera (fila 5)
    col_año = None
    for col in range(1, 15):
        cell = ws.cell(row=OTROS_FILA_CABECERA_RESUMEN, column=col)
        if str(cell.value or "").strip() == nombre_año:
            col_año = col
            break

    # Si no existe el año, añadir nueva columna (o nueva sección)
    if col_año is None:
        # Encontrar la última columna con año
        ultima_col = 5
        for col in range(5, 15):
            cell = ws.cell(row=OTROS_FILA_CABECERA_RESUMEN, column=col)
            if cell.value and str(cell.value).strip().count("-") == 1:
                ultima_col = col
        col_año = ultima_col + 1
        ws.cell(row=OTROS_FILA_CABECERA_RESUMEN, column=col_año).value = nombre_año

    col_letra = get_column_letter(col_año)

    # Actualizar totales de resumen (filas 6-9)
    total_lecturas_acs  = sum(f["importe_total"] for f in facturas_mto if "ACS" in (f.get("notas") or "").upper())
    total_lecturas_cal  = sum(f["importe_total"] for f in facturas_mto if "CALEF" in (f.get("notas") or "").upper())
    total_mto_preventivo = sum(f["importe_total"] for f in facturas_mto
                               if "MANTENIMIENTO" in f.get("tipo_suministro", "").upper()
                               and "RIOS" in (f.get("proveedor") or "").upper())

    if total_lecturas_acs:
        ws[f"{col_letra}6"] = round(total_lecturas_acs, 2)
    if total_lecturas_cal:
        ws[f"{col_letra}7"] = round(total_lecturas_cal, 2)
    if total_mto_preventivo:
        ws[f"{col_letra}9"] = round(total_mto_preventivo, 2)


# ---------------------------------------------------------------------------
# FUNCIÓN AUXILIAR
# ---------------------------------------------------------------------------

def _abreviar_proveedor(nombre: str) -> str:
    """Abrevia el nombre del proveedor para la columna Nota."""
    mapa = {
        "BASER":        "BASER",
        "ENDESA":       "ENDESA",
        "ENERGÍA XXI":  "E.XXI",
        "RIOS":         "RIOS REN.",
        "GALP":         "GALP",
        "NATURGY":      "NATURGY",
        "IBERDROLA":    "IBERDROLA",
    }
    nombre_up = nombre.upper()
    for clave, abrev in mapa.items():
        if clave in nombre_up:
            return abrev
    return nombre[:12] if nombre else ""


def _limpiar_num_str(valor) -> float:
    """
    Convierte un valor de importe o consumo a float Python limpio.
    Acepta strings españoles ('1.234,56'), floats e ints directamente.
    Es imprescindible llamar a esta función antes de asignar valores a celdas
    Excel para que las fórmulas nativas de SUM, ANALISIS y RESUMEN funcionen.
    Sin esta conversión, openpyxl puede escribir strings en lugar de números
    y las fórmulas devuelven 0 o #VALUE!
    """
    if valor is None:
        return 0.0
    if isinstance(valor, (int, float)):
        return float(valor)
    texto = str(valor).strip().replace("€", "").replace(" ", "").replace("\xa0", "")
    # Formato español: punto como separador de miles, coma como decimal
    # Ej: "1.234,56" → quitar puntos → "1234,56" → coma→punto → "1234.56"
    if "." in texto and "," in texto and texto.index(".") < texto.index(","):
        texto = texto.replace(".", "")
    texto = texto.replace(",", ".")
    try:
        return float(texto)
    except ValueError:
        return 0.0


# ---------------------------------------------------------------------------
# VOLCADO DIRECTO — layout continuo con validación CUPS (Comunidad_607 y similares)
# ---------------------------------------------------------------------------

def volcar_facturas_a_excel(
    codigo_comunidad: str,
    tipo_suministro: str,
    datos_facturas: list[dict],
    cups_factura: str = None,
    ruta_base_excels: str = None,
) -> dict:
    """
    Vuelca facturas en la pestaña correspondiente del Excel Maestro de cualquier
    comunidad de Meditrade, construyendo la ruta de forma dinámica.

    Ruta resuelta: <ruta_base_excels>/Comunidad_{codigo_comunidad}.xlsx
    Si ruta_base_excels no se indica, se infiere como ../../Excels_Maestros/
    relativa a la ubicación de este módulo (core/ -> raiz -> Excels_Maestros/).

    Layout esperado de cada pestaña (todas las comunidades comparten el mismo
    formato de cabecera, por lo que la logica funciona de forma identica para
    Comunidad_575, 591, 607, 610, 615, 631, 644, etc.):
      Filas 1-4  -> metadatos fijos. C4 = CUPS de control de la hoja.
      Fila 7     -> cabecera de columnas.
      Filas 8-10 -> reservadas (formulas de balance u otras).
      Fila 11+   -> historico de facturas (primera fila vacia aqui).

    Columnas de datos:
      B = Fecha fra.    C = Fecha inicial    D = Fecha final
      E = Consumo       F = Termino fijo     G = Termino variable
      H = Importe TOTAL J = Nota (proveedor)

    Validacion de seguridad:
      Si cups_factura (o cups_o_referencia del primer dict) no coincide con el
      valor de la celda C4 del Excel que se abre en ese instante, se lanza
      CupsMismatchError SIN escribir ninguna fila. Esto protege contra mezclar
      facturas de distintas comunidades o suministros en el archivo equivocado.

    NOTA sobre formulas SUM:
      Esta funcion hace APPEND (no insert_rows), por lo que las referencias
      cruzadas de las hojas ANALISIS/RESUMEN no se desplazan. Si tus formulas
      SUM usan un rango fijo (ej. =SUM(H11:H100)), asegurate de que cubre
      suficientes filas futuras o define tablas Excel dinamicas.

    Args:
        codigo_comunidad: Codigo numerico de la comunidad (ej: '607', '644').
        tipo_suministro:  'ELECTRICIDAD' | 'GAS' | 'AGUA' | 'MANTENIMIENTO'.
        datos_facturas:   Lista de dicts de facturas (campos del modelo BD).
        cups_factura:     CUPS del PDF. Si None, se lee de
                          datos_facturas[0].get('cups_o_referencia').
        ruta_base_excels: Carpeta raiz de los Excels Maestros. Si None se
                          infiere automaticamente.

    Returns:
        dict: {ok, hoja, ruta_excel, fila_inicio, filas_escritas,
               cups_validado, errores}

    Raises:
        CupsMismatchError: CUPS del PDF != valor de celda C4.
        FileNotFoundError: El Excel de la comunidad no existe.
        KeyError:          La pestana de tipo_suministro no existe en el Excel.
    """
    # Construir ruta dinamica: Excels_Maestros/Comunidad_{codigo}.xlsx
    if ruta_base_excels:
        ruta_excel = os.path.join(
            ruta_base_excels, f"Comunidad_{codigo_comunidad}.xlsx"
        )
    else:
        _base = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        ruta_excel = os.path.join(
            _base, "Excels_Maestros", f"Comunidad_{codigo_comunidad}.xlsx"
        )

    if not os.path.exists(ruta_excel):
        raise FileNotFoundError(
            f"No existe el Excel para la comunidad '{codigo_comunidad}': {ruta_excel}"
        )
    if not datos_facturas:
        return {"ok": True, "hoja": tipo_suministro, "ruta_excel": ruta_excel,
                "fila_inicio": None, "filas_escritas": 0,
                "cups_validado": None, "errores": []}

    nombre_hoja = tipo_suministro.upper()

    # Posiciones fijas del layout Comunidad_607
    FILA_CUPS         = 4    # C4 = CUPS de control
    COL_CUPS          = 3
    FILA_DATOS_INICIO = 11   # primera fila de histórico

    # Columnas de datos (índices numéricos para ws.cell())
    COL_FECHA_FRA  = 2   # B
    COL_FECHA_INI  = 3   # C
    COL_FECHA_FIN  = 4   # D
    COL_CONSUMO    = 5   # E
    COL_T_FIJO     = 6   # F
    COL_T_VARIABLE = 7   # G
    COL_TOTAL      = 8   # H
    COL_NOTA       = 10  # J

    # Cargar sin data_only para preservar fórmulas existentes
    wb = load_workbook(ruta_excel)

    if nombre_hoja not in wb.sheetnames:
        raise KeyError(
            f"La pestaña '{nombre_hoja}' no existe en "
            f"{os.path.basename(ruta_excel)}. "
            f"Hojas disponibles: {wb.sheetnames}"
        )
    ws = wb[nombre_hoja]

    # ── Validación CUPS: C4 del Excel vs. CUPS del PDF ───────────────────────
    cups_en_excel  = str(ws.cell(row=FILA_CUPS, column=COL_CUPS).value or "").strip()
    cups_a_validar = (
        (cups_factura or "").strip()
        or str(datos_facturas[0].get("cups_o_referencia") or "").strip()
    )

    if cups_en_excel and cups_a_validar:
        if cups_en_excel.upper() != cups_a_validar.upper():
            raise CupsMismatchError(cups_a_validar, cups_en_excel, nombre_hoja)

    cups_validado = cups_en_excel or cups_a_validar or None

    # ── Buscar primera fila vacía a partir de FILA_DATOS_INICIO ──────────────
    fila_escritura = FILA_DATOS_INICIO
    while ws.cell(row=fila_escritura, column=COL_FECHA_FRA).value is not None:
        fila_escritura += 1
        if fila_escritura > FILA_DATOS_INICIO + 5000:
            raise RuntimeError(
                f"No se encontró fila vacía en '{nombre_hoja}' "
                f"tras buscar 5000 filas desde la fila {FILA_DATOS_INICIO}"
            )

    fila_inicio    = fila_escritura
    filas_escritas = 0
    errores: list[str] = []

    for factura in datos_facturas:
        try:
            fila = fila_escritura

            # B: fecha factura
            fecha_fra = _fecha_excel(factura.get("fecha_factura"))
            ws.cell(row=fila, column=COL_FECHA_FRA).value = fecha_fra
            if fecha_fra:
                ws.cell(row=fila, column=COL_FECHA_FRA).number_format = "DD/MM/YYYY"

            # C: fecha inicio período
            fecha_ini = _fecha_excel(factura.get("fecha_inicio"))
            ws.cell(row=fila, column=COL_FECHA_INI).value = fecha_ini
            if fecha_ini:
                ws.cell(row=fila, column=COL_FECHA_INI).number_format = "DD/MM/YYYY"

            # D: fecha fin período
            fecha_fin = _fecha_excel(factura.get("fecha_fin"))
            ws.cell(row=fila, column=COL_FECHA_FIN).value = fecha_fin
            if fecha_fin:
                ws.cell(row=fila, column=COL_FECHA_FIN).number_format = "DD/MM/YYYY"

            # E: consumo en kWh o m³ — float explícito para que SUM funcione
            consumo = _limpiar_num_str(
                factura.get("consumo_total")
                or factura.get("consumo_kwh")
                or factura.get("consumo_m3")
            )
            if consumo:
                ws.cell(row=fila, column=COL_CONSUMO).value = consumo
                ws.cell(row=fila, column=COL_CONSUMO).number_format = "#,##0.00"

            # F: término fijo — float explícito
            t_fijo = _limpiar_num_str(factura.get("termino_fijo"))
            if t_fijo:
                ws.cell(row=fila, column=COL_T_FIJO).value = t_fijo
                ws.cell(row=fila, column=COL_T_FIJO).number_format = "#,##0.00"

            # G: término variable — float explícito
            t_var = _limpiar_num_str(factura.get("termino_variable"))
            if t_var:
                ws.cell(row=fila, column=COL_T_VARIABLE).value = t_var
                ws.cell(row=fila, column=COL_T_VARIABLE).number_format = "#,##0.00"

            # H: importe total — fórmula si hay desglose, valor directo si no
            importe = _limpiar_num_str(factura.get("importe_total"))
            if t_fijo and t_var:
                # Fórmula: Excel la recalcula y los encadenados (ANALISIS, RESUMEN) suman bien
                ws.cell(row=fila, column=COL_TOTAL).value = f"=F{fila}+G{fila}"
            else:
                ws.cell(row=fila, column=COL_TOTAL).value = importe or None
            ws.cell(row=fila, column=COL_TOTAL).number_format = "#,##0.00"

            # J: nota (proveedor abreviado)
            ws.cell(row=fila, column=COL_NOTA).value = _abreviar_proveedor(
                factura.get("proveedor", "")
            )

            fila_escritura += 1
            filas_escritas += 1

        except Exception as exc:
            num_fra = factura.get("num_factura", "?")
            errores.append(f"Factura {num_fra}: {exc}")

    if filas_escritas > 0:
        wb.save(ruta_excel)

    return {
        "ok":             len(errores) == 0,
        "hoja":           nombre_hoja,
        "ruta_excel":     ruta_excel,
        "fila_inicio":    fila_inicio,
        "filas_escritas": filas_escritas,
        "cups_validado":  cups_validado,
        "errores":        errores,
    }


# ---------------------------------------------------------------------------
# FUNCIÓN PRINCIPAL
# ---------------------------------------------------------------------------

def actualizar_excel_maestro(ruta_excel: str, ruta_bd: str,
                              id_comunidad: int, id_periodo: int,
                              hacer_backup: bool = True) -> dict:
    """
    Vuelca los datos del periodo en el Excel Maestro.

    Args:
        ruta_excel:     Ruta al archivo Comunidad_XXX.xlsx
        ruta_bd:        Ruta a la BD SQLite
        id_comunidad:   ID de la comunidad
        id_periodo:     ID del periodo a volcar
        hacer_backup:   Si True, crea una copia .bak antes de modificar

    Returns:
        dict con: ok, pestañas_actualizadas, filas_escritas, errores
    """
    if not os.path.exists(ruta_excel):
        return {"ok": False, "error": f"No encontrado: {ruta_excel}"}

    # Backup de seguridad
    if hacer_backup:
        ruta_bak = ruta_excel.replace(".xlsx", f"_bak_{datetime.now():%Y%m%d_%H%M%S}.xlsx")
        shutil.copy2(ruta_excel, ruta_bak)

    # Leer datos de la BD
    con = sqlite3.connect(ruta_bd)
    con.row_factory = sqlite3.Row
    cur = con.cursor()

    periodo = cur.execute(
        "SELECT * FROM periodos WHERE id_periodo=?", (id_periodo,)
    ).fetchone()
    nombre_año = periodo["nombre"]  # ej: "2025-2026"

    facturas = cur.execute("""
        SELECT tipo_suministro, proveedor, fecha_factura, fecha_inicio, fecha_fin,
               dias_facturados, consumo_total, unidad_consumo,
               termino_fijo, termino_variable, impuestos, iva, importe_total,
               cups_o_referencia, notas
        FROM facturas
        WHERE id_comunidad=? AND id_periodo=?
        ORDER BY tipo_suministro, fecha_inicio
    """, (id_comunidad, id_periodo)).fetchall()
    facturas = [dict(f) for f in facturas]

    # ── Normalizar campos de consumo para los writers de Excel ──────────────
    # La BD almacena un solo campo consumo_total + unidad_consumo.
    # Los writers de cada pestaña esperan consumo_m3 / consumo_kwh por separado.
    for f in facturas:
        t = f.get("tipo_suministro", "")
        unidad = (f.get("unidad_consumo") or "").upper()
        total  = f.get("consumo_total") or 0.0
        if t == "GAS":
            if unidad == "M3":
                f["consumo_m3"]  = total
                f["consumo_kwh"] = 0.0
            else:                      # kWh (ej. Baser almacena kWh como primario)
                f["consumo_kwh"] = total
                f["consumo_m3"]  = 0.0
        elif t == "ELECTRICIDAD":
            f["consumo_kwh"] = total
        # AGUA: importe_total ya es el campo correcto, c_variable/c_fija no están en BD

    # Facturas Metrigest (lecturas de vecinos — vienen de archivos_procesados + datos)
    # Se leen de la tabla lecturas_vecino como resumen por periodo
    lecturas_acs = cur.execute("""
        SELECT lv.fecha_lectura as fecha_act, lv.valor_acumulado as val_act,
               lv.estado, p.codigo_vivienda, p.nombre_propietario,
               LAG(lv.valor_acumulado) OVER (PARTITION BY lv.id_propietario, lv.tipo
                   ORDER BY lv.fecha_lectura) as val_ant,
               LAG(lv.fecha_lectura) OVER (PARTITION BY lv.id_propietario, lv.tipo
                   ORDER BY lv.fecha_lectura) as fecha_ant
        FROM lecturas_vecino lv
        JOIN propietarios p ON p.id_propietario = lv.id_propietario
        WHERE p.id_comunidad=? AND lv.id_periodo=? AND lv.tipo='ACS'
        ORDER BY lv.fecha_lectura, p.codigo_vivienda
    """, (id_comunidad, id_periodo)).fetchall()
    lecturas_acs = [dict(l) for l in lecturas_acs if l["val_ant"] is not None]

    lecturas_cal = cur.execute("""
        SELECT lv.fecha_lectura as fecha_act, lv.valor_acumulado as val_act,
               lv.estado, p.codigo_vivienda, p.nombre_propietario,
               LAG(lv.valor_acumulado) OVER (PARTITION BY lv.id_propietario, lv.tipo
                   ORDER BY lv.fecha_lectura) as val_ant,
               LAG(lv.fecha_lectura) OVER (PARTITION BY lv.id_propietario, lv.tipo
                   ORDER BY lv.fecha_lectura) as fecha_ant
        FROM lecturas_vecino lv
        JOIN propietarios p ON p.id_propietario = lv.id_propietario
        WHERE p.id_comunidad=? AND lv.id_periodo=? AND lv.tipo='CALEFACCION'
        ORDER BY lv.fecha_lectura, p.codigo_vivienda
    """, (id_comunidad, id_periodo)).fetchall()
    lecturas_cal = [dict(l) for l in lecturas_cal if l["val_ant"] is not None]

    # Precios unitarios facturados (de config_suministro)
    config_acs = cur.execute(
        "SELECT * FROM config_suministro WHERE id_comunidad=? AND id_periodo=? AND tipo_suministro='ACS'",
        (id_comunidad, id_periodo)
    ).fetchone()
    config_cal = cur.execute(
        "SELECT * FROM config_suministro WHERE id_comunidad=? AND id_periodo=? AND tipo_suministro='CALEFACCION'",
        (id_comunidad, id_periodo)
    ).fetchone()

    precio_acs  = config_acs["precio_variable_facturado"]  if config_acs  else 6.0
    precio_cal  = config_cal["precio_variable_facturado"]  if config_cal  else 0.2
    precio_fijo_acs = config_acs["precio_fijo_facturado"] if config_acs else 7.5
    precio_fijo_cal = config_cal["precio_fijo_facturado"] if config_cal else 12.5

    # Nº de vecinos activos (para la cuota fija total; antes estaba cableado a 120)
    n_vecinos = cur.execute(
        "SELECT COUNT(*) FROM propietarios WHERE id_comunidad=? AND activo=1",
        (id_comunidad,)
    ).fetchone()[0] or 0

    con.close()

    # Abrir Excel
    wb = load_workbook(ruta_excel)

    pestañas_actualizadas = []
    filas_escritas = 0
    errores = []

    # ------------------------------------------------------------------
    # PESTAÑA GAS
    # ------------------------------------------------------------------
    facturas_gas = [f for f in facturas if f["tipo_suministro"] == "GAS"]
    if facturas_gas:
        try:
            ws = wb["GAS"]
            fila_bloque = _encontrar_bloque_año(ws, nombre_año)

            if fila_bloque is None:
                # Crear nuevo bloque: si es plantilla vacía, empieza justo bajo la cabecera
                ultimo_fila, _ = _ultimo_bloque_año(ws)
                if ultimo_fila == 0:
                    fila_bloque = GAS_FILA_CABECERA + 3   # fila 11
                else:
                    fila_ultima_suma = _encontrar_fila_suma(ws, ultimo_fila)
                    fila_bloque = (fila_ultima_suma or GAS_FILA_CABECERA + 3) + 4
                ws.cell(row=fila_bloque, column=1).value = nombre_año
                fila_datos_inicio = fila_bloque + 1
            else:
                fila_suma = _encontrar_fila_suma(ws, fila_bloque)
                fila_datos_inicio = fila_suma if fila_suma else fila_bloque + 1

            # Insertar filas para las nuevas facturas
            ws.insert_rows(fila_datos_inicio, amount=len(facturas_gas))

            for i, f in enumerate(facturas_gas):
                fila_actual = fila_datos_inicio + i
                _escribir_fila_gas(ws, fila_actual, f, i == 0)
                filas_escritas += 1

            # Actualizar fila Suma (ahora desplazada)
            fila_suma_nueva = fila_datos_inicio + len(facturas_gas)
            _escribir_fila_suma_gas(ws, fila_suma_nueva,
                                     fila_datos_inicio,
                                     fila_datos_inicio + len(facturas_gas) - 1)

            pestañas_actualizadas.append("GAS")
        except Exception as e:
            errores.append(f"GAS: {e}")

    # ------------------------------------------------------------------
    # PESTAÑA ELECTRICIDAD
    # ------------------------------------------------------------------
    facturas_elec = [f for f in facturas if f["tipo_suministro"] == "ELECTRICIDAD"]
    if facturas_elec:
        try:
            ws = wb["ELECTRICIDAD"]
            fila_bloque = _encontrar_bloque_año(ws, nombre_año)

            if fila_bloque is None:
                ultimo_fila, _ = _ultimo_bloque_año(ws)
                if ultimo_fila == 0:
                    fila_bloque = ELEC_FILA_CABECERA + 3   # fila 11
                else:
                    fila_suma_ant = _encontrar_fila_suma(ws, ultimo_fila)
                    fila_bloque = (fila_suma_ant or ELEC_FILA_CABECERA + 3) + 4
                ws.cell(row=fila_bloque, column=1).value = nombre_año
                fila_datos_inicio = fila_bloque + 1
            else:
                fila_suma = _encontrar_fila_suma(ws, fila_bloque)
                fila_datos_inicio = fila_suma if fila_suma else fila_bloque + 1

            ws.insert_rows(fila_datos_inicio, amount=len(facturas_elec))

            for i, f in enumerate(facturas_elec):
                _escribir_fila_elec(ws, fila_datos_inicio + i, f)
                filas_escritas += 1

            fila_suma_nueva = fila_datos_inicio + len(facturas_elec)
            _escribir_fila_suma_elec(ws, fila_suma_nueva,
                                      fila_datos_inicio,
                                      fila_datos_inicio + len(facturas_elec) - 1)

            pestañas_actualizadas.append("ELECTRICIDAD")
        except Exception as e:
            errores.append(f"ELECTRICIDAD: {e}")

    # ------------------------------------------------------------------
    # PESTAÑA AGUA
    # ------------------------------------------------------------------
    facturas_agua = [f for f in facturas if f["tipo_suministro"] == "AGUA"]
    if facturas_agua:
        try:
            ws = wb["AGUA"]
            fila_bloque = _encontrar_bloque_año(ws, nombre_año)

            if fila_bloque is None:
                ultimo_fila, _ = _ultimo_bloque_año(ws)
                if ultimo_fila == 0:
                    fila_bloque = AGUA_FILA_CABECERA + 2   # fila 14
                else:
                    fila_suma_ant = _encontrar_fila_suma(ws, ultimo_fila)
                    fila_bloque = (fila_suma_ant or AGUA_FILA_CABECERA + 2) + 4
                ws.cell(row=fila_bloque, column=1).value = nombre_año
                fila_datos_inicio = fila_bloque + 1
            else:
                fila_suma = _encontrar_fila_suma(ws, fila_bloque)
                fila_datos_inicio = fila_suma if fila_suma else fila_bloque + 1

            ws.insert_rows(fila_datos_inicio, amount=len(facturas_agua))

            for i, f in enumerate(facturas_agua):
                _escribir_fila_agua(ws, fila_datos_inicio + i, f)
                filas_escritas += 1

            fila_suma_nueva = fila_datos_inicio + len(facturas_agua)
            _escribir_fila_suma_agua(ws, fila_suma_nueva,
                                      fila_datos_inicio,
                                      fila_datos_inicio + len(facturas_agua) - 1)

            pestañas_actualizadas.append("AGUA")
        except Exception as e:
            errores.append(f"AGUA: {e}")

    # ------------------------------------------------------------------
    # PESTAÑA LECTURAS ACS M3
    # ------------------------------------------------------------------
    if lecturas_acs:
        try:
            ws = wb["LECTURAS ACS M3"]
            fila_bloque = _encontrar_bloque_año(ws, nombre_año)

            if fila_bloque is None:
                fin = _fin_datos_hoja(ws)
                fila_bloque = 8 if fin < 8 else fin + 3
                ws.cell(row=fila_bloque, column=1).value = nombre_año
                ws.cell(row=fila_bloque, column=2).value = "Días:"

            # Agrupar lecturas por fecha (cada fecha Metrigest = un lote)
            fechas_unicas = sorted(set(l["fecha_act"] for l in lecturas_acs))
            fila_datos = fila_bloque + 1
            fila_ini_blk = fila_datos
            for fecha in fechas_unicas:
                lecs_fecha = [l for l in lecturas_acs if l["fecha_act"] == fecha]
                # Si val_act < val_ant, el contador se sustituyó entre ambas
                # lecturas (un acumulado nunca baja) — esa diferencia no es
                # consumo real, así que se descarta en vez de sumar un
                # negativo que arrastraría "m3 consumidos" a valores absurdos.
                total_consumo = sum(l.get("consumo", 0) or
                                    max(0, l["val_act"] - l["val_ant"]) for l in lecs_fecha
                                    if l.get("val_ant") is not None)
                cuota_var = round(total_consumo * precio_acs, 2)
                fecha_ant = lecs_fecha[0].get("fecha_ant")
                meses_lote = _meses_entre(fecha_ant, fecha) if fecha_ant else 3

                # Fila de cuota variable (consumo del lote × precio)
                ws[f"B{fila_datos}"] = _fecha_excel(fecha)
                ws[f"B{fila_datos}"].number_format = "DD/MM/YYYY"
                if fecha_ant:
                    ws[f"C{fila_datos}"] = _fecha_excel(fecha_ant)
                    ws[f"C{fila_datos}"].number_format = "DD/MM/YYYY"
                ws[f"G{fila_datos}"] = f"=H{fila_datos}/L{fila_datos}"
                ws[f"H{fila_datos}"] = cuota_var
                ws[f"H{fila_datos}"].number_format = "#,##0.00"
                ws[f"J{fila_datos}"] = f"=SUM(H{fila_datos}:I{fila_datos})"
                ws[f"L{fila_datos}"] = precio_acs
                filas_escritas += 1
                fila_datos += 1

                # Fila de cuota fija: nº vecinos × precio fijo × meses del lote
                cuota_fija_total = round(n_vecinos * precio_fijo_acs * meses_lote, 2)
                ws[f"B{fila_datos}"] = _fecha_excel(fecha)
                ws[f"B{fila_datos}"].number_format = "DD/MM/YYYY"
                ws[f"I{fila_datos}"] = cuota_fija_total
                ws[f"I{fila_datos}"].number_format = "#,##0.00"
                ws[f"J{fila_datos}"] = f"=SUM(H{fila_datos}:I{fila_datos})"
                ws[f"M{fila_datos}"] = f"=I{fila_datos}/$C$4/{meses_lote}"
                fila_datos += 1

            # Días del bloque en la cabecera (lo referencia la hoja ANALISIS)
            try:
                from datetime import datetime as _dt
                _pf = lambda x: _dt.fromisoformat(str(x)[:10])
                _primera = min(_pf(l["fecha_ant"]) for l in lecturas_acs if l.get("fecha_ant"))
                _ultima  = max(_pf(l["fecha_act"]) for l in lecturas_acs)
                ws.cell(row=fila_bloque, column=3).value = (_ultima - _primera).days
            except (ValueError, TypeError):
                pass

            # Fila Suma del bloque
            ws[f"F{fila_datos}"] = "Suma …."
            for col in ("G", "H", "I", "J"):
                ws[f"{col}{fila_datos}"] = f"=SUM({col}{fila_ini_blk}:{col}{fila_datos - 1})"
                ws[f"{col}{fila_datos}"].number_format = "#,##0.00"
            ws[f"L{fila_datos}"] = precio_acs
            ws[f"M{fila_datos}"] = precio_fijo_acs

            pestañas_actualizadas.append("LECTURAS ACS M3")
        except Exception as e:
            errores.append(f"LECTURAS ACS M3: {e}")

    # ------------------------------------------------------------------
    # PESTAÑA LECTURAS CALEF KWH
    # ------------------------------------------------------------------
    if lecturas_cal:
        try:
            ws = wb["LECTURAS CALEF KWH"]
            fila_bloque = _encontrar_bloque_año(ws, nombre_año)

            if fila_bloque is None:
                fin = _fin_datos_hoja(ws)
                fila_bloque = 8 if fin < 8 else fin + 3
                ws.cell(row=fila_bloque, column=1).value = nombre_año
                ws.cell(row=fila_bloque, column=2).value = "DIAS:"

            fechas_unicas = sorted(set(l["fecha_act"] for l in lecturas_cal))
            fila_datos = fila_bloque + 1
            fila_ini_blk = fila_datos

            for fecha in fechas_unicas:
                lecs_fecha = [l for l in lecturas_cal if l["fecha_act"] == fecha]
                total_consumo = sum(
                    (l["val_act"] - l["val_ant"])
                    for l in lecs_fecha
                    if l.get("val_ant") is not None and l["val_act"] > l["val_ant"]
                )
                cuota_var = round(total_consumo * precio_cal, 2)
                fecha_ant = lecs_fecha[0].get("fecha_ant")
                meses_lote = _meses_entre(fecha_ant, fecha) if fecha_ant else 3

                ws[f"B{fila_datos}"] = _fecha_excel(fecha)
                ws[f"B{fila_datos}"].number_format = "DD/MM/YYYY"
                if fecha_ant:
                    ws[f"C{fila_datos}"] = _fecha_excel(fecha_ant)
                    ws[f"C{fila_datos}"].number_format = "DD/MM/YYYY"
                ws[f"G{fila_datos}"] = f"=H{fila_datos}/L{fila_datos}"
                ws[f"H{fila_datos}"] = cuota_var
                ws[f"H{fila_datos}"].number_format = "#,##0.00"
                ws[f"J{fila_datos}"] = f"=SUM(H{fila_datos}:I{fila_datos})"
                ws[f"L{fila_datos}"] = precio_cal
                filas_escritas += 1
                fila_datos += 1

                cuota_fija_total = round(n_vecinos * precio_fijo_cal * meses_lote, 2)
                ws[f"B{fila_datos}"] = _fecha_excel(fecha)
                ws[f"B{fila_datos}"].number_format = "DD/MM/YYYY"
                ws[f"I{fila_datos}"] = cuota_fija_total
                ws[f"I{fila_datos}"].number_format = "#,##0.00"
                ws[f"J{fila_datos}"] = f"=SUM(H{fila_datos}:I{fila_datos})"
                ws[f"M{fila_datos}"] = f"=I{fila_datos}/$C$4/{meses_lote}"
                fila_datos += 1

            # Días del bloque en la cabecera (lo referencia la hoja ANALISIS)
            try:
                from datetime import datetime as _dt
                _pf = lambda x: _dt.fromisoformat(str(x)[:10])
                _primera = min(_pf(l["fecha_ant"]) for l in lecturas_cal if l.get("fecha_ant"))
                _ultima  = max(_pf(l["fecha_act"]) for l in lecturas_cal)
                ws.cell(row=fila_bloque, column=3).value = (_ultima - _primera).days
            except (ValueError, TypeError):
                pass

            # Fila Suma del bloque
            ws[f"F{fila_datos}"] = "Suma …."
            for col in ("G", "H", "I", "J"):
                ws[f"{col}{fila_datos}"] = f"=SUM({col}{fila_ini_blk}:{col}{fila_datos - 1})"
                ws[f"{col}{fila_datos}"].number_format = "#,##0.00"
            ws[f"L{fila_datos}"] = precio_cal
            ws[f"M{fila_datos}"] = precio_fijo_cal

            pestañas_actualizadas.append("LECTURAS CALEF KWH")
        except Exception as e:
            errores.append(f"LECTURAS CALEF KWH: {e}")

    # ------------------------------------------------------------------
    # PESTAÑA OTROS GASTOS
    # ------------------------------------------------------------------
    facturas_mto = [f for f in facturas if f["tipo_suministro"] == "MANTENIMIENTO"]
    if facturas_mto:
        try:
            ws = wb["OTROS GASTOS"]
            _actualizar_otros_gastos(ws, nombre_año, facturas_mto, [])
            pestañas_actualizadas.append("OTROS GASTOS")
        except Exception as e:
            errores.append(f"OTROS GASTOS: {e}")

    # ------------------------------------------------------------------
    # HOJA ANALISIS — crear nueva si no existe
    # ------------------------------------------------------------------
    nombre_analisis_nuevo = f"ANALISIS {nombre_año[-2:]}-{str(int(nombre_año.split('-')[1][-2:])).zfill(2)}"
    partes_año = nombre_año.split("-")
    nombre_analisis_nuevo = f"ANALISIS {partes_año[0][-2:]}-{partes_año[1][-2:]}"

    if nombre_analisis_nuevo not in wb.sheetnames:
        try:
            # Detectar las filas Suma actuales en cada hoja para el nuevo periodo
            def _fila_suma_bloque(hoja_nombre, año):
                ws_h = wb[hoja_nombre]
                fila_bloque = _encontrar_bloque_año(ws_h, año)
                if fila_bloque:
                    return _encontrar_fila_suma_hoja(ws_h, fila_bloque)
                return None

            fila_suma_gas  = _fila_suma_bloque("GAS", nombre_año)
            fila_suma_elec = _fila_suma_bloque("ELECTRICIDAD", nombre_año)
            fila_suma_agua = _fila_suma_bloque("AGUA", nombre_año)

            # Para lecturas: buscar el bloque del año en LECTURAS ACS y CALEF
            ws_acs = wb["LECTURAS ACS M3"]
            fila_bloque_acs = _encontrar_bloque_año(ws_acs, nombre_año)
            fila_suma_acs   = _encontrar_fila_suma_hoja(ws_acs, fila_bloque_acs) if fila_bloque_acs else None

            ws_calef = wb["LECTURAS CALEF KWH"]
            fila_bloque_calef = _encontrar_bloque_año(ws_calef, nombre_año)
            fila_suma_calef   = _encontrar_fila_suma_hoja(ws_calef, fila_bloque_calef) if fila_bloque_calef else None

            filas_suma = {
                "gas":        fila_suma_gas,
                "elec":       fila_suma_elec,
                "agua":       fila_suma_agua,
                "acs_suma":   fila_suma_acs,
                "acs_ini":    fila_bloque_acs,
                "calef_suma": fila_suma_calef,
                "calef_ini":  fila_bloque_calef,
            }

            nombre_creado = crear_hoja_analisis(wb, nombre_año, filas_suma)
            if nombre_creado:
                pestañas_actualizadas.append(nombre_creado)

                # Actualizar RESUMEN
                actualizar_hoja_resumen(wb, nombre_año, nombre_creado, {
                    "acs_suma":   fila_suma_acs,
                    "acs_ini":    fila_bloque_acs,
                    "calef_suma": fila_suma_calef,
                    "calef_ini":  fila_bloque_calef,
                })
                pestañas_actualizadas.append("RESUMEN")

        except Exception as e:
            errores.append(f"ANALISIS/RESUMEN: {e}")

    # Reparto estacional ACS/CALEFACCION: se recalcula entero cada vez (barato,
    # y el periodo N+1 no existe todavía cuando se procesa el periodo N).
    try:
        recalcular_reparto_estacional_gas(wb)
        recalcular_reparto_estacional_elec(wb)
    except Exception as e:
        errores.append(f"Reparto estacional GAS/ELECTRICIDAD: {e}")

    # Guardar
    wb.save(ruta_excel)

    return {
        "ok": len(errores) == 0,
        "archivo": ruta_excel,
        "periodo": nombre_año,
        "pestañas_actualizadas": pestañas_actualizadas,
        "filas_escritas": filas_escritas,
        "backup": ruta_bak if hacer_backup else None,
        "errores": errores,
    }


# ---------------------------------------------------------------------------
# CREACIÓN DE HOJA ANALISIS Y ACTUALIZACIÓN DE RESUMEN/DATOS
# ---------------------------------------------------------------------------

def _encontrar_fila_suma_hoja(ws, fila_bloque_inicio: int) -> Optional[int]:
    """
    Busca la fila Suma de un bloque (puede estar en col B, F o G según la hoja).
    Busca hasta 80 filas después del inicio del bloque.
    """
    for i in range(fila_bloque_inicio + 1, fila_bloque_inicio + 80):
        for col in range(1, 8):
            cell = ws.cell(row=i, column=col)
            if cell.value and "Suma" in str(cell.value):
                return i
    return None


def crear_hoja_analisis(wb, nombre_año: str,
                        filas_suma: dict) -> str:
    """
    Crea una nueva hoja ANALISIS para el periodo indicado copiando la más reciente
    y actualizando todas las referencias a las filas Suma de las otras hojas.

    Args:
        wb:          Workbook abierto con openpyxl
        nombre_año:  Ej: '2024-2025'
        filas_suma:  Dict con las filas Suma de cada hoja para el nuevo periodo:
                     {
                       'gas':        98,   # fila Suma en GAS
                       'elec':       53,   # fila Suma en ELECTRICIDAD
                       'agua':       75,   # fila Suma en AGUA
                       'acs_suma':   73,   # fila Suma en LECTURAS ACS M3
                       'acs_ini':    54,   # fila inicio bloque ACS (para días)
                       'calef_suma': 82,   # fila Suma en LECTURAS CALEF KWH
                       'calef_ini':  63,   # fila inicio bloque CALEF
                     }

    Returns:
        Nombre de la nueva hoja creada
    """
    # Encontrar la hoja ANALISIS más reciente para copiar
    hojas_analisis = sorted([s for s in wb.sheetnames if s.startswith("ANALISIS")])
    if not hojas_analisis:
        return None
    # Preferir la hoja PLANTILLA (referencias conocidas y sin datos de otro año)
    if "ANALISIS PLANTILLA" in wb.sheetnames:
        plantilla_nombre = "ANALISIS PLANTILLA"
    else:
        plantilla_nombre = hojas_analisis[-1]
    plantilla = wb[plantilla_nombre]

    # Extraer año anterior de la plantilla para saber qué refs sustituir
    # Buscamos referencias a GAS!L en la plantilla para detectar la fila antigua
    fila_gas_antigua = None
    fila_elec_antigua = None
    fila_agua_antigua = None
    fila_acs_suma_antigua = None
    fila_acs_ini_antigua = None
    fila_calef_suma_antigua = None
    fila_calef_ini_antigua = None

    import re as _re
    for row in plantilla.iter_rows():
        for cell in row:
            v = str(cell.value or "")
            if "GAS!L" in v and fila_gas_antigua is None:
                m = _re.search(r"GAS!L(\d+)", v)
                if m:
                    fila_gas_antigua = int(m.group(1))
            if "ELECTRICIDAD!H" in v and fila_elec_antigua is None:
                m = _re.search(r"ELECTRICIDAD!H(\d+)", v)
                if m:
                    fila_elec_antigua = int(m.group(1))
            if "AGUA!V" in v and fila_agua_antigua is None:
                m = _re.search(r"AGUA!V(\d+)", v)
                if m:
                    fila_agua_antigua = int(m.group(1))
            if "LECTURAS ACS M3'!" in v and fila_acs_suma_antigua is None:
                # la fila Suma puede estar referenciada por G, H, I, L o M
                m = _re.search(r"ACS M3'![GHILM](\d+)", v)
                if m:
                    fila_acs_suma_antigua = int(m.group(1))
            if "LECTURAS ACS M3'!C" in v and fila_acs_ini_antigua is None:
                m = _re.search(r"ACS M3'!C(\d+)", v)
                if m:
                    fila_acs_ini_antigua = int(m.group(1))
            if "LECTURAS CALEF KWH'!" in v and fila_calef_suma_antigua is None:
                m = _re.search(r"CALEF KWH'![GHILM](\d+)", v)
                if m:
                    fila_calef_suma_antigua = int(m.group(1))
            if "LECTURAS CALEF KWH'!C" in v and fila_calef_ini_antigua is None:
                m = _re.search(r"CALEF KWH'!C(\d+)", v)
                if m:
                    fila_calef_ini_antigua = int(m.group(1))

    # Copiar la hoja plantilla
    nuevo_nombre = f"ANALISIS {nombre_año[-2:]}-{str(int(nombre_año[-2:])+1).zfill(2)}"
    # Formato correcto: "ANALISIS 24-25"
    partes = nombre_año.split("-")
    nuevo_nombre = f"ANALISIS {partes[0][-2:]}-{partes[1][-2:]}"

    nueva_hoja = wb.copy_worksheet(plantilla)
    nueva_hoja.title = nuevo_nombre
    nueva_hoja.sheet_state = "visible"   # la plantilla puede estar oculta

    # Mover la nueva hoja antes de RESUMEN
    idx_resumen = wb.sheetnames.index("RESUMEN")
    wb.move_sheet(nuevo_nombre, offset=idx_resumen - wb.sheetnames.index(nuevo_nombre))

    # Mapa de sustituciones: (texto_antiguo, texto_nuevo)
    sustituciones = []

    def _add(patron_antiguo, fila_ant, fila_nueva, cols):
        if fila_ant and fila_nueva:
            for col in cols:
                sustituciones.append(
                    (f"{patron_antiguo}!{col}{fila_ant}",
                     f"{patron_antiguo}!{col}{fila_nueva}")
                )

    _add("GAS",         fila_gas_antigua,    filas_suma.get("gas"),    ["L","V","W","X","Y"])
    _add("ELECTRICIDAD",fila_elec_antigua,   filas_suma.get("elec"),   ["H","P","Q","R","S"])
    _add("AGUA",        fila_agua_antigua,   filas_suma.get("agua"),   ["V","W"])

    # Lecturas ACS
    if fila_acs_suma_antigua and filas_suma.get("acs_suma"):
        for col in ["G","H","I","L","M"]:
            sustituciones.append((
                f"'LECTURAS ACS M3'!{col}{fila_acs_suma_antigua}",
                f"'LECTURAS ACS M3'!{col}{filas_suma['acs_suma']}"
            ))
    if fila_acs_ini_antigua and filas_suma.get("acs_ini"):
        sustituciones.append((
            f"'LECTURAS ACS M3'!C{fila_acs_ini_antigua}",
            f"'LECTURAS ACS M3'!C{filas_suma['acs_ini']}"
        ))

    # Lecturas CALEF
    if fila_calef_suma_antigua and filas_suma.get("calef_suma"):
        for col in ["G","H","I","L","M"]:
            sustituciones.append((
                f"'LECTURAS CALEF KWH'!{col}{fila_calef_suma_antigua}",
                f"'LECTURAS CALEF KWH'!{col}{filas_suma['calef_suma']}"
            ))
    if fila_calef_ini_antigua and filas_suma.get("calef_ini"):
        sustituciones.append((
            f"'LECTURAS CALEF KWH'!C{fila_calef_ini_antigua}",
            f"'LECTURAS CALEF KWH'!C{filas_suma['calef_ini']}"
        ))

    # Aplicar sustituciones en todas las celdas de la nueva hoja
    for row in nueva_hoja.iter_rows():
        for cell in row:
            if cell.value and isinstance(cell.value, str):
                nuevo_valor = cell.value
                for viejo, nuevo in sustituciones:
                    nuevo_valor = nuevo_valor.replace(viejo, nuevo)
                if nuevo_valor != cell.value:
                    cell.value = nuevo_valor

    # Si este periodo no tiene bloque en alguna hoja (ej. electricidad venía
    # incluida en las facturas de gas en 2020-21/2021-22), la sustitución de
    # arriba no tiene nada que reemplazar y la fórmula se quedaría apuntando
    # a la fila Suma antigua de la plantilla (de OTRO periodo). Mejor 0 que
    # una referencia falsa a datos de otro año.
    if not filas_suma.get("gas"):
        nueva_hoja["H24"] = 0
    if not filas_suma.get("elec"):
        nueva_hoja["H25"] = 0
    if not filas_suma.get("agua"):
        nueva_hoja["H28"] = 0
        nueva_hoja["H29"] = 0

    # Actualizar el título del periodo en la hoja (fila 4 col G apunta a DATOS!A6)
    # Añadir la línea del nuevo periodo en DATOS
    ws_datos = wb["DATOS"]
    # Buscar última fila de periodo en DATOS
    ultima_fila_periodo = None
    for i, row in enumerate(ws_datos.iter_rows(min_row=1, max_row=20, values_only=False)):
        fila = i + 1
        for cell in row:
            if cell.value and "EJERCICIO" in str(cell.value):
                ultima_fila_periodo = fila
    if ultima_fila_periodo:
        nueva_fila_datos = ultima_fila_periodo + 1
        ws_datos.cell(row=nueva_fila_datos, column=1).value = (
            f"PERIODO: EJERCICIO {nombre_año} (SEPTIEMBRE {nombre_año[:4]} - AGOSTO {nombre_año[5:]})"
        )

    return nuevo_nombre


def actualizar_hoja_resumen(wb, nombre_año: str,
                             nombre_analisis: str,
                             filas_suma_lecturas: dict):
    """
    Añade una nueva columna en la hoja RESUMEN para el nuevo periodo.

    Args:
        filas_suma_lecturas: {
            'acs_ini': fila inicio bloque ACS,
            'acs_suma': fila Suma ACS,
            'calef_ini': fila inicio CALEF (no usada aquí directamente)
        }
    """
    ws = wb["RESUMEN"]

    # Encontrar la última columna con año en fila 8
    fila_cabecera = 8
    ultima_col = None
    ultima_col_idx = 0
    for col in range(1, 20):
        cell = ws.cell(row=fila_cabecera, column=col)
        if cell.value and str(cell.value).strip().count("-") == 1:
            partes = str(cell.value).strip().split("-")
            if len(partes) == 2 and partes[0].isdigit():
                ultima_col = col
    if ultima_col is None:
        return

    # Si la última columna es el marcador de la plantilla (referencia a
    # 'ANALISIS PLANTILLA'), se sobreescribe en el sitio en lugar de añadir
    # otra columna duplicada.
    es_placeholder = False
    for r in range(fila_cabecera, min(ws.max_row, fila_cabecera + 60) + 1):
        v = ws.cell(row=r, column=ultima_col).value
        if isinstance(v, str) and "ANALISIS PLANTILLA" in v:
            es_placeholder = True
            break

    nueva_col = ultima_col if es_placeholder else ultima_col + 1
    if es_placeholder and ultima_col > 1:
        ultima_col = ultima_col  # las fórmulas se sustituyen en la misma columna
    nueva_col_letra = get_column_letter(nueva_col)
    ultima_col_letra = get_column_letter(ultima_col)

    # Cabecera: nombre del año — sobreescribir explícitamente
    ws.cell(row=fila_cabecera, column=nueva_col).value = nombre_año

    # Copiar las fórmulas de la columna anterior actualizando referencias
    import re as _re

    def _reescribir_formula(v: str) -> str:
        # Referencia al ANALISIS del nuevo periodo
        v = _re.sub(r"'ANALISIS [^']*'", f"'{nombre_analisis}'", v)
        v = _re.sub(r"ANALISIS \d\d-\d\d", nombre_analisis, v)
        v = v.replace("ANALISIS PLANTILLA", nombre_analisis)
        # Refs a LECTURAS ACS: fila Suma (G/H/I/L/M) y fila inicio (C, días)
        if filas_suma_lecturas.get("acs_suma"):
            v = _re.sub(
                r"'LECTURAS ACS M3'!([GHILM])(\d+)",
                lambda m: f"'LECTURAS ACS M3'!{m.group(1)}{filas_suma_lecturas['acs_suma']}",
                v)
        if filas_suma_lecturas.get("acs_ini"):
            v = _re.sub(
                r"'LECTURAS ACS M3'!C(\d+)",
                lambda m: f"'LECTURAS ACS M3'!C{filas_suma_lecturas['acs_ini']}",
                v)
        # Refs a LECTURAS CALEF: idem
        if filas_suma_lecturas.get("calef_suma"):
            v = _re.sub(
                r"'LECTURAS CALEF KWH'!([GHILM])(\d+)",
                lambda m: f"'LECTURAS CALEF KWH'!{m.group(1)}{filas_suma_lecturas['calef_suma']}",
                v)
        if filas_suma_lecturas.get("calef_ini"):
            v = _re.sub(
                r"'LECTURAS CALEF KWH'!C(\d+)",
                lambda m: f"'LECTURAS CALEF KWH'!C{filas_suma_lecturas['calef_ini']}",
                v)
        return v

    for row in ws.iter_rows(min_row=fila_cabecera, max_row=ws.max_row):
        for cell in row:
            if cell.row == fila_cabecera:
                continue  # la cabecera (año) ya se escribió explícitamente
            if cell.column == ultima_col and cell.value:
                nueva_celda = ws.cell(row=cell.row, column=nueva_col)
                if isinstance(cell.value, str) and cell.value.startswith("="):
                    nueva_celda.value = _reescribir_formula(str(cell.value))
                elif (isinstance(cell.value, str)
                      and _re.fullmatch(r"\d{4}-\d{4}", cell.value.strip())):
                    # Etiqueta de año suelta (no fórmula) en secciones que no
                    # son la cabecera principal (ej. "CALEF.", "CALEF. -
                    # RELACION DE GASTO"): sin esto se copiaba el año anterior
                    # sin cambiar, en vez de avanzar al nuevo periodo.
                    nueva_celda.value = nombre_año
                elif cell.value is not None:
                    nueva_celda.value = cell.value

    # Actualizar las fórmulas SUM de las filas de suma de columnas
    for row in ws.iter_rows(min_row=fila_cabecera, max_row=ws.max_row):
        for cell in row:
            if (cell.column < 4 and cell.value and isinstance(cell.value, str)
                    and "SUM(" in cell.value):
                # Ampliar el rango SUM para incluir la nueva columna
                import re as _re2
                m = _re2.search(r"SUM\(([A-Z]+)(\d+):([A-Z]+)(\d+)\)", cell.value)
                if m:
                    nueva_cell_value = cell.value.replace(
                        f"SUM({m.group(1)}{m.group(2)}:{m.group(3)}{m.group(4)})",
                        f"SUM({m.group(1)}{m.group(2)}:{nueva_col_letra}{m.group(4)})"
                    )
                    cell.value = nueva_cell_value


# ---------------------------------------------------------------------------
# CREACIÓN DE PLANTILLA EN BLANCO — una por comunidad nueva
# ---------------------------------------------------------------------------

def crear_plantilla_excel(ruta_destino: str, codigo: str = "", nombre: str = "") -> str:
    """
    Crea un Excel en blanco con la estructura correcta para una nueva comunidad.
    Incluye todas las hojas requeridas y cabeceras en las posiciones exactas
    que espera actualizar_excel_maestro(). Sin datos ni bloques de año.

    Args:
        ruta_destino:  Ruta completa del archivo a crear (Comunidad_XXX.xlsx)
        codigo:        Código de la comunidad (ej: '645')
        nombre:        Nombre de la comunidad (ej: 'CDAD. PROP. CALLE MAYOR 10')

    Returns:
        ruta_destino
    """
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    wb = openpyxl.Workbook()

    # ── Paleta ────────────────────────────────────────────────────────────────
    AZUL_MARINO  = "1A3A5C"
    AZUL_CLARO   = "D6EAF8"
    GRIS_CABEC   = "F2F2F2"

    def _estilo_titulo(ws, fila, texto):
        ws.cell(row=fila, column=1).value = texto
        ws.cell(row=fila, column=1).font = Font(
            bold=True, size=12, color="FFFFFF"
        )
        ws.cell(row=fila, column=1).fill = PatternFill(
            fill_type="solid", fgColor=AZUL_MARINO
        )

    def _cabecera_fila(ws, fila, cabeceras: list):
        """cabeceras: list of (col_letra, texto)"""
        for col_l, txt in cabeceras:
            from openpyxl.utils import column_index_from_string
            col = column_index_from_string(col_l)
            c = ws.cell(row=fila, column=col)
            c.value = txt
            c.font = Font(bold=True, size=9, color=AZUL_MARINO)
            c.fill = PatternFill(fill_type="solid", fgColor=AZUL_CLARO)
            c.alignment = Alignment(horizontal="center", wrap_text=True)

    def _ancho(ws, col_letra, ancho):
        ws.column_dimensions[col_letra].width = ancho

    # ─────────────────────────────────────────────────────────────────────────
    # HOJA GAS
    # ─────────────────────────────────────────────────────────────────────────
    ws = wb.active
    ws.title = "GAS"
    _estilo_titulo(ws, 2, f"FACTURAS SUMINISTRO GAS NATURAL — {nombre or codigo}")
    _cabecera_fila(ws, GAS_FILA_CABECERA, [
        ("B","Fecha fra."), ("C","Días"), ("D","Fecha ini"),
        ("E","Lec ini"),    ("F","Fecha fin"), ("G","Lec fin"),
        ("H","m³"),         ("I","kWh"),
        ("J","€fijo"),      ("K","€var"),      ("L","TOTAL"),
        ("M","Pago"),       ("N","Nota"),
        ("O","€/m³"),       ("P","Precio m³"), ("Q","€/kWh"), ("R","Precio kWh"),
    ])
    ws.row_dimensions[GAS_FILA_CABECERA].height = 30
    ws.freeze_panes = f"B{GAS_FILA_CABECERA + 1}"
    for col, ancho in [("A",12),("B",12),("C",7),("D",12),("E",10),
                        ("F",12),("G",10),("H",10),("I",10),("J",10),
                        ("K",10),("L",10),("M",12),("N",14)]:
        _ancho(ws, col, ancho)

    # ─────────────────────────────────────────────────────────────────────────
    # HOJA ELECTRICIDAD
    # ─────────────────────────────────────────────────────────────────────────
    ws = wb.create_sheet("ELECTRICIDAD")
    _estilo_titulo(ws, 2, f"FACTURAS SUMINISTRO ELECTRICIDAD — {nombre or codigo}")
    _cabecera_fila(ws, ELEC_FILA_CABECERA, [
        ("B","Fecha fra."), ("C","Fecha ini"), ("D","Fecha fin"),
        ("E","kWh"),        ("F","€fijo"),     ("G","€var"),
        ("H","TOTAL"),      ("I","Pago"),      ("J","Nota"),
        ("L","€/kWh"),      ("M","Precio kWh"),
    ])
    ws.row_dimensions[ELEC_FILA_CABECERA].height = 30
    ws.freeze_panes = f"B{ELEC_FILA_CABECERA + 1}"

    # ─────────────────────────────────────────────────────────────────────────
    # HOJA AGUA
    # ─────────────────────────────────────────────────────────────────────────
    ws = wb.create_sheet("AGUA")
    _estilo_titulo(ws, 2, f"FACTURAS SUMINISTRO AGUA — {nombre or codigo}")
    _cabecera_fila(ws, AGUA_FILA_CABECERA, [
        ("B","Fecha fra."), ("C","Periodo"), ("D","Fecha ini"),
        ("E","Lec ini"),    ("F","Fecha fin"), ("G","Lec fin"),
        ("H","m³"),         ("T","TOTAL IVA"), ("V","C.Variable"), ("W","C.Fija"),
    ])
    ws.row_dimensions[AGUA_FILA_CABECERA].height = 30

    # ─────────────────────────────────────────────────────────────────────────
    # HOJA OTROS GASTOS
    # ─────────────────────────────────────────────────────────────────────────
    ws = wb.create_sheet("OTROS GASTOS")
    _estilo_titulo(ws, 2, f"OTROS GASTOS Y MANTENIMIENTO — {nombre or codigo}")
    _cabecera_fila(ws, OTROS_FILA_CABECERA_RESUMEN, [
        ("A","Concepto"), ("B","€/mes"), ("C","Descripción"),
    ])
    ws.row_dimensions[OTROS_FILA_CABECERA_RESUMEN].height = 20

    # ─────────────────────────────────────────────────────────────────────────
    # HOJA LECTURAS ACS M3
    # ─────────────────────────────────────────────────────────────────────────
    ws = wb.create_sheet("LECTURAS ACS M3")
    _estilo_titulo(ws, 2, f"LECTURAS ACS (m³) — {nombre or codigo}")
    _cabecera_fila(ws, 6, [
        ("A","Ejercicio"), ("B","Fecha lect."), ("C","Fecha ant."),
        ("D","Lec ant."),  ("E","Fecha act."),  ("F","Lec act."),
        ("G","m³/€"),      ("H","Cuota var."),  ("I","Cuota fija"),
        ("J","TOTAL"),     ("L","€/m³"),
    ])

    # ─────────────────────────────────────────────────────────────────────────
    # HOJA LECTURAS CALEF KWH
    # ─────────────────────────────────────────────────────────────────────────
    ws = wb.create_sheet("LECTURAS CALEF KWH")
    _estilo_titulo(ws, 2, f"LECTURAS CALEFACCIÓN (kWh) — {nombre or codigo}")
    _cabecera_fila(ws, 6, [
        ("A","Ejercicio"), ("B","Fecha lect."), ("C","Fecha ant."),
        ("D","Lec ant."),  ("E","Fecha act."),  ("F","Lec act."),
        ("G","kWh/€"),     ("H","Cuota var."),  ("I","Cuota fija"),
        ("J","TOTAL"),     ("L","€/kWh"),
    ])

    # ─────────────────────────────────────────────────────────────────────────
    # HOJA DATOS
    # ─────────────────────────────────────────────────────────────────────────
    ws = wb.create_sheet("DATOS")
    ws["A1"] = f"COMUNIDAD: {nombre or codigo}"
    ws["A1"].font = Font(bold=True, size=11, color=AZUL_MARINO)
    ws["A2"] = f"CÓDIGO: {codigo}"
    ws["A3"] = f"FECHA ALTA: {datetime.now().strftime('%d/%m/%Y')}"
    ws["A5"] = "EJERCICIO ACTIVO:"
    ws["A5"].font = Font(bold=True)

    # ─────────────────────────────────────────────────────────────────────────
    # HOJA RESUMEN
    # ─────────────────────────────────────────────────────────────────────────
    ws = wb.create_sheet("RESUMEN")
    _estilo_titulo(ws, 2, f"RESUMEN ANUAL — {nombre or codigo}")
    _cabecera_fila(ws, 8, [
        ("A","Concepto"), ("B","Unidad"), ("C","Total"),
    ])

    wb.save(ruta_destino)
    return ruta_destino


# ---------------------------------------------------------------------------
# EJECUCIÓN DIRECTA — test
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    import sys
    import shutil

    ruta_original = "/mnt/user-data/uploads/Comunidad_644.xlsx"
    ruta_test     = "/tmp/Comunidad_644_TEST.xlsx"
    shutil.copy2(ruta_original, ruta_test)

    # Usar la BD de test del motor_reparto
    RUTA_BD = "/tmp/test_motor/gestion.db"

    if not os.path.exists(RUTA_BD):
        print("❌ Ejecuta primero motor_reparto.py para crear la BD de test.")
        sys.exit(1)

    print("Para regenerar un Excel Maestro usa:  python excel_generator.py --comunidad 644")
