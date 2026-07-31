"""
importar_excel_maestro.py
=========================
Lee facturas del Excel Maestro (Comunidad_XXX.xlsx) para un periodo concreto
e inserta las que falten en la BD SQLite.

Útil cuando el Excel ya tiene datos históricos y la BD está vacía (primer uso,
restauración de BD, etc.). Tras la importación, el flujo normal continúa:
Calcular Reparto → Generar Cartas.

USO:
    python importar_excel_maestro.py --excel Excels_Maestros/Comunidad_644.xlsx
                                     --bd data/gestion.db
                                     --comunidad 1
                                     --periodo 1
                                     --año 2024-2025
"""

import re
import sys
import sqlite3
import argparse
from datetime import datetime, date
from pathlib import Path
from typing import Optional

sys.path.insert(0, str(Path(__file__).parent))
from gestor_bd import conectar, insertar_factura

try:
    import openpyxl
    from openpyxl.utils import column_index_from_string
except ImportError:
    openpyxl = None


# ---------------------------------------------------------------------------
# UTILIDADES
# ---------------------------------------------------------------------------

def _fecha_celda(valor) -> Optional[str]:
    """Convierte un valor de celda Excel (date, datetime, string) a 'YYYY-MM-DD'."""
    if valor is None:
        return None
    if isinstance(valor, (datetime, date)):
        v = valor.date() if isinstance(valor, datetime) else valor
        return v.strftime("%Y-%m-%d")
    s = str(valor).strip()
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d.%m.%Y"):
        try:
            return datetime.strptime(s[:10], fmt).strftime("%Y-%m-%d")
        except (ValueError, TypeError):
            continue
    return None


def _float_celda(valor) -> float:
    """Convierte un valor de celda a float, devuelve 0 si no es numérico."""
    if valor is None:
        return 0.0
    try:
        return float(valor)
    except (TypeError, ValueError):
        return 0.0


def _str_celda(valor) -> str:
    if valor is None:
        return ""
    return str(valor).strip()


def _es_año(texto: str) -> bool:
    """Devuelve True si el texto tiene patrón YYYY-YYYY (ej: '2024-2025')."""
    return bool(re.match(r'^\d{4}-\d{4}$', str(texto or "").strip()))


def _es_suma(texto: str) -> bool:
    return "Suma" in str(texto or "")


# ---------------------------------------------------------------------------
# LECTURA DE BLOQUES DE AÑO
# ---------------------------------------------------------------------------

def _encontrar_bloque_año(ws, nombre_año: str):
    """
    Devuelve (fila_inicio, fila_fin_exclusive) del bloque nombre_año en col A.
    fila_fin apunta a la fila 'Suma ….' (inclusive) o None si no se encuentra.
    """
    fila_inicio = None
    for row in ws.iter_rows(min_col=1, max_col=1):
        cell = row[0]
        val  = _str_celda(cell.value)
        if val == nombre_año:
            fila_inicio = cell.row
            break

    if fila_inicio is None:
        return None, None

    # Buscar fila Suma: en col B o F (según la hoja) desde fila_inicio+1
    for i in range(fila_inicio + 1, fila_inicio + 80):
        for col in [2, 3, 5, 6]:  # cols B, C, E, F
            c = ws.cell(row=i, column=col)
            if c.value and _es_suma(str(c.value)):
                return fila_inicio, i  # incluye la fila Suma

    return fila_inicio, fila_inicio + 40  # fallback


# ---------------------------------------------------------------------------
# LECTORES POR HOJA
# ---------------------------------------------------------------------------

def _leer_gas(ws, fila_ini: int, fila_fin: int) -> list[dict]:
    """Lee filas de datos GAS entre fila_ini+1 y fila_fin-1 (excluye encabezado y Suma)."""
    facturas = []
    for fila in range(fila_ini + 1, fila_fin):
        b = ws.cell(row=fila, column=2).value   # fecha_factura
        l = ws.cell(row=fila, column=12).value  # importe_total (col L)

        if b is None or _es_suma(str(b)):
            continue
        fecha_f = _fecha_celda(b)
        if fecha_f is None:
            continue

        consumo_m3  = _float_celda(ws.cell(row=fila, column=8).value)   # H
        consumo_kwh = _float_celda(ws.cell(row=fila, column=9).value)   # I
        importe     = _float_celda(l)

        if importe == 0:
            continue

        # consumo_total: preferir kWh si existe, sino m3
        consumo_total = consumo_kwh if consumo_kwh else consumo_m3
        unidad        = "kWh" if consumo_kwh else ("m3" if consumo_m3 else None)

        nota = _str_celda(ws.cell(row=fila, column=14).value)  # N

        facturas.append({
            "tipo_suministro":  "GAS",
            "proveedor":        nota or "GAS",
            "fecha_factura":    fecha_f,
            "fecha_inicio":     _fecha_celda(ws.cell(row=fila, column=4).value),   # D
            "fecha_fin":        _fecha_celda(ws.cell(row=fila, column=6).value),   # F
            "dias_facturados":  int(_float_celda(ws.cell(row=fila, column=3).value)) or None,
            "consumo_total":    consumo_total,
            "unidad_consumo":   unidad,
            "termino_fijo":     _float_celda(ws.cell(row=fila, column=10).value),  # J
            "termino_variable": _float_celda(ws.cell(row=fila, column=11).value),  # K
            "importe_total":    importe,
            "notas":            nota,
        })
    return facturas


def _leer_electricidad(ws, fila_ini: int, fila_fin: int) -> list[dict]:
    facturas = []
    for fila in range(fila_ini + 1, fila_fin):
        b = ws.cell(row=fila, column=2).value   # fecha_factura
        h = ws.cell(row=fila, column=8).value   # importe_total (col H)

        if b is None or _es_suma(str(b)):
            continue
        fecha_f = _fecha_celda(b)
        if fecha_f is None:
            continue

        importe = _float_celda(h)
        if importe == 0:
            continue

        nota = _str_celda(ws.cell(row=fila, column=10).value)  # J

        facturas.append({
            "tipo_suministro":  "ELECTRICIDAD",
            "proveedor":        nota or "ELECTRICIDAD",
            "fecha_factura":    fecha_f,
            "fecha_inicio":     _fecha_celda(ws.cell(row=fila, column=3).value),  # C
            "fecha_fin":        _fecha_celda(ws.cell(row=fila, column=4).value),  # D
            "dias_facturados":  None,
            "consumo_total":    _float_celda(ws.cell(row=fila, column=5).value),  # E
            "unidad_consumo":   "kWh",
            "termino_fijo":     _float_celda(ws.cell(row=fila, column=6).value),  # F
            "termino_variable": _float_celda(ws.cell(row=fila, column=7).value),  # G
            "importe_total":    importe,
            "notas":            nota,
        })
    return facturas


def _leer_agua(ws, fila_ini: int, fila_fin: int) -> list[dict]:
    facturas = []
    for fila in range(fila_ini + 1, fila_fin):
        b  = ws.cell(row=fila, column=2).value   # fecha_factura (col B)
        t  = ws.cell(row=fila, column=20).value  # importe_total (col T)

        if b is None or _es_suma(str(b)):
            continue
        fecha_f = _fecha_celda(b)
        if fecha_f is None:
            continue

        importe = _float_celda(t)
        if importe == 0:
            continue

        facturas.append({
            "tipo_suministro":  "AGUA",
            "proveedor":        "AGUA",
            "fecha_factura":    fecha_f,
            "fecha_inicio":     _fecha_celda(ws.cell(row=fila, column=4).value),  # D
            "fecha_fin":        _fecha_celda(ws.cell(row=fila, column=6).value),  # F
            "dias_facturados":  None,
            "consumo_total":    _float_celda(ws.cell(row=fila, column=8).value),  # H
            "unidad_consumo":   "m3",
            "termino_fijo":     0.0,
            "termino_variable": 0.0,
            "importe_total":    importe,
            "notas":            None,
        })
    return facturas


def _leer_otros(ws, fila_ini: int, fila_fin: int) -> list[dict]:
    """Lee facturas de OTROS GASTOS / MANTENIMIENTO."""
    facturas = []
    for fila in range(fila_ini + 1, fila_fin):
        b = ws.cell(row=fila, column=2).value

        if b is None or _es_suma(str(b)):
            continue
        fecha_f = _fecha_celda(b)
        if fecha_f is None:
            # Puede haber descripción en col A/B/C sin fecha
            continue

        # Col L (12) suele tener el importe en OTROS GASTOS
        importe = _float_celda(ws.cell(row=fila, column=12).value)
        if importe == 0:
            importe = _float_celda(ws.cell(row=fila, column=8).value)
        if importe == 0:
            continue

        nota = _str_celda(ws.cell(row=fila, column=14).value) or \
               _str_celda(ws.cell(row=fila, column=3).value)

        facturas.append({
            "tipo_suministro":  "MANTENIMIENTO",
            "proveedor":        nota or "MANTENIMIENTO",
            "fecha_factura":    fecha_f,
            "fecha_inicio":     None,
            "fecha_fin":        None,
            "dias_facturados":  None,
            "consumo_total":    0.0,
            "unidad_consumo":   None,
            "termino_fijo":     0.0,
            "termino_variable": 0.0,
            "importe_total":    importe,
            "notas":            nota,
        })
    return facturas


# ---------------------------------------------------------------------------
# FUNCIÓN PRINCIPAL
# ---------------------------------------------------------------------------

HOJAS_CONFIG = {
    "GAS":          _leer_gas,
    "ELECTRICIDAD": _leer_electricidad,
    "AGUA":         _leer_agua,
    "OTROS GASTOS": _leer_otros,
}


def importar_excel_a_bd(ruta_excel: str, ruta_bd: str,
                        id_comunidad: int, id_periodo: int,
                        nombre_año: str,
                        verbose: bool = True) -> dict:
    """
    Lee las facturas del año nombre_año desde el Excel Maestro e inserta
    en la BD las que aún no existan.

    Args:
        ruta_excel:   Ruta al Comunidad_XXX.xlsx
        ruta_bd:      Ruta a gestion.db
        id_comunidad: ID de la comunidad en BD
        id_periodo:   ID del periodo en BD
        nombre_año:   Nombre del año/periodo (ej: '2024-2025')
        verbose:      Imprimir progreso

    Returns:
        {ok, facturas_importadas, duplicadas, errores, detalle_por_hoja}
    """
    if openpyxl is None:
        return {"ok": False, "error": "openpyxl no instalado"}

    if not Path(ruta_excel).exists():
        return {"ok": False, "error": f"Excel no encontrado: {ruta_excel}"}

    if verbose:
        print(f"\n=== IMPORTAR DESDE EXCEL: {Path(ruta_excel).name} — {nombre_año} ===")

    wb = openpyxl.load_workbook(ruta_excel, data_only=True)

    con = conectar(ruta_bd)

    total_importadas = 0
    total_duplicadas = 0
    errores_globales = []
    detalle = {}

    for nombre_hoja, lector in HOJAS_CONFIG.items():
        if nombre_hoja not in wb.sheetnames:
            if verbose:
                print(f"  ⚠️  Hoja '{nombre_hoja}' no existe — ignorando")
            continue

        ws = wb[nombre_hoja]
        fila_ini, fila_fin = _encontrar_bloque_año(ws, nombre_año)

        if fila_ini is None:
            if verbose:
                print(f"  ⚠️  Hoja '{nombre_hoja}': no se encontró bloque '{nombre_año}'")
            detalle[nombre_hoja] = {"importadas": 0, "duplicadas": 0, "errores": 0,
                                    "nota": f"bloque '{nombre_año}' no encontrado"}
            continue

        filas_datos = lector(ws, fila_ini, fila_fin)
        importadas = 0
        duplicadas = 0
        errores    = 0

        for f in filas_datos:
            f["id_comunidad"] = id_comunidad
            f["id_periodo"]   = id_periodo
            f["num_factura"]  = (
                f"{f['tipo_suministro']}_{f.get('fecha_factura','?')}"
                f"_{f.get('importe_total',0):.2f}"
            )
            f["cups_o_referencia"] = f["tipo_suministro"]
            f["archivo_origen"]    = Path(ruta_excel).name

            try:
                id_fac = insertar_factura(con, f)
                if id_fac:
                    importadas += 1
                    if verbose:
                        print(f"    ✅ {nombre_hoja:<15} {f.get('fecha_factura','?'):<12} "
                              f"{f.get('importe_total',0):>10.2f} €")
                else:
                    duplicadas += 1
            except Exception as e:
                errores += 1
                errores_globales.append(f"{nombre_hoja}: {e}")
                if verbose:
                    print(f"    ❌ {nombre_hoja}: {e}")

        total_importadas += importadas
        total_duplicadas += duplicadas
        detalle[nombre_hoja] = {"importadas": importadas,
                                 "duplicadas": duplicadas,
                                 "errores": errores}
        if verbose:
            print(f"  📊 {nombre_hoja:<18} → {importadas} nuevas, "
                  f"{duplicadas} dup, {errores} err")

    con.close()
    wb.close()

    ok = total_importadas > 0 or total_duplicadas > 0
    return {
        "ok":                 ok,
        "facturas_importadas": total_importadas,
        "duplicadas":          total_duplicadas,
        "errores":             errores_globales,
        "detalle_por_hoja":    detalle,
    }


# ---------------------------------------------------------------------------
# EJECUCIÓN DIRECTA
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Importa facturas desde Excel Maestro a BD")
    parser.add_argument("--excel",    required=True, help="Ruta al Excel Maestro")
    parser.add_argument("--bd",       required=True, help="Ruta a gestion.db")
    parser.add_argument("--comunidad", required=True, type=int, help="ID comunidad en BD")
    parser.add_argument("--periodo",  required=True, type=int,  help="ID periodo en BD")
    parser.add_argument("--año",      required=True, help="Nombre del año (ej: 2024-2025)")
    args = parser.parse_args()

    resultado = importar_excel_a_bd(
        ruta_excel   = args.excel,
        ruta_bd      = args.bd,
        id_comunidad = args.comunidad,
        id_periodo   = args.periodo,
        nombre_año   = args.año,
        verbose      = True,
    )

    print(f"\n{'='*50}")
    print(f"Resultado:  {'✅ OK' if resultado['ok'] else '⚠️  Sin datos'}")
    print(f"Importadas: {resultado['facturas_importadas']}")
    print(f"Duplicadas: {resultado['duplicadas']}")
    if resultado['errores']:
        print(f"Errores:    {len(resultado['errores'])}")
        for e in resultado['errores'][:5]:
            print(f"  - {e}")
