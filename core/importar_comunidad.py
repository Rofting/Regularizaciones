"""
importar_comunidad.py
=====================
Importa a la BD todos los datos de una comunidad desde sus archivos fuente:
  - Vecinos y coeficientes (desde el Excel CAL-ACS)
  - Lecturas históricas ACS y Calefacción (desde el mismo Excel)
  - Periodos de regularización (detectados automáticamente)
  - Configuración de precios por periodo

USO (desde carpeta core/):
    python importar_comunidad.py --comunidad 644
    python importar_comunidad.py --comunidad 644 --excel "../data_fuente/644_CAL-ACS_2025.xlsx"

RESULTADO:
    - Tabla propietarios: 120 vecinos con coeficiente
    - Tabla periodos: PER1, PER2, PER3, PER4 (2020-2024)
    - Tabla lecturas_vecino: todas las lecturas ACS y CALEF históricas
    - Tabla config_suministro: precios por periodo
"""

import sys
import os
import re
import sqlite3
import argparse
from datetime import datetime, date
from pathlib import Path

import openpyxl

sys.path.insert(0, str(Path(__file__).parent))
from gestor_bd import conectar, crear_bd, obtener_o_crear_comunidad, obtener_o_crear_periodo

# ---------------------------------------------------------------------------
# RUTAS POR DEFECTO
# ---------------------------------------------------------------------------
BASE_DIR     = Path(__file__).parent.parent
RUTA_BD      = BASE_DIR / "data" / "gestion.db"
RUTA_EXCELS  = BASE_DIR / "data_fuente"

# Mapas de periodos de regularización
# Cada periodo tiene: nombre, lecturas_ACS (col_ini, col_fin), lecturas_CALEF (col_ini, col_fin)
# Las columnas son 0-based
PERIODOS_644 = [
    {
        "nombre":       "2020-2021",
        "fecha_inicio": "2020-06-01",
        "fecha_fin":    "2021-08-31",
        "col_acs_ini":  5,   # jun-20
        "col_acs_fin":  10,  # sep-21 (índice 10 = col K)
        "col_cal_ini":  36,  # may-20
        "col_cal_fin":  43,  # nov-21
        "precio_acs_var":  6.00, "precio_acs_fijo":  7.50,
        "precio_cal_var":  0.18, "precio_cal_fijo": 12.50,
        "pct_acs": 0.511, "pct_cal": 0.489,
    },
    {
        "nombre":       "2021-2022",
        "fecha_inicio": "2021-09-01",
        "fecha_fin":    "2022-08-31",
        "col_acs_ini":  10,  # sep-21
        "col_acs_fin":  17,  # dic-22
        "col_cal_ini":  43,  # nov-21
        "col_cal_fin":  51,  # dic-22
        "precio_acs_var":  6.00, "precio_acs_fijo":  7.50,
        "precio_cal_var":  0.18, "precio_cal_fijo": 12.50,
        "pct_acs": 0.511, "pct_cal": 0.489,
    },
    {
        "nombre":       "2022-2023",
        "fecha_inicio": "2022-09-01",
        "fecha_fin":    "2023-08-31",
        "col_acs_ini":  17,  # dic-22
        "col_acs_fin":  24,  # dic-23
        "col_cal_ini":  51,  # dic-22
        "col_cal_fin":  56,  # dic-23
        "precio_acs_var":  6.00, "precio_acs_fijo":  7.50,
        "precio_cal_var":  0.18, "precio_cal_fijo": 12.50,
        "pct_acs": 0.511, "pct_cal": 0.489,
    },
    {
        "nombre":       "2023-2024",
        "fecha_inicio": "2023-09-01",
        "fecha_fin":    "2024-08-31",
        "col_acs_ini":  24,  # dic-23
        "col_acs_fin":  29,  # dic-24
        "col_cal_ini":  56,  # dic-23
        "col_cal_fin":  61,  # dic-24
        "precio_acs_var":  6.00, "precio_acs_fijo":  7.50,
        "precio_cal_var":  0.20, "precio_cal_fijo": 12.50,
        "pct_acs": 0.511, "pct_cal": 0.489,
    },
    {
        "nombre":       "2024-2025",
        "fecha_inicio": "2024-09-01",
        "fecha_fin":    None,    # abierto
        "col_acs_ini":  29,  # dic-24
        "col_acs_fin":  29,  # misma — sólo tenemos la lectura inicial
        "col_cal_ini":  61,  # dic-24
        "col_cal_fin":  61,
        "precio_acs_var":  6.00, "precio_acs_fijo":  7.50,
        "precio_cal_var":  0.20, "precio_cal_fijo": 12.50,
        "pct_acs": 0.511, "pct_cal": 0.489,
    },
]


# ---------------------------------------------------------------------------
# PARSEAR FECHA DE CABECERA DEL EXCEL
# ---------------------------------------------------------------------------
def _parsear_fecha_col(valor) -> str | None:
    """Convierte el valor de la cabecera de columna a YYYY-MM-DD."""
    if valor is None:
        return None
    sv = str(valor).strip()

    # Ya es datetime
    if hasattr(valor, 'year'):
        return valor.strftime("%Y-%m-%d")

    # Formato 'YYYY-MM-DD HH:MM:SS'
    m = re.match(r'(\d{4}-\d{2}-\d{2})', sv)
    if m:
        return m.group(1)

    # Formato 'mes-YY' → convertir a fecha del primer día del mes
    meses = {
        'ene': 1, 'feb': 2, 'mar': 3, 'abr': 4, 'may': 5, 'jun': 6,
        'jul': 7, 'ago': 8, 'sep': 9, 'oct': 10, 'nov': 11, 'dic': 12
    }
    m = re.match(r'([a-z]{3})-(\d{2,4})', sv, re.IGNORECASE)
    if m:
        mes_str = m.group(1).lower()[:3]
        anio_str = m.group(2)
        mes = meses.get(mes_str)
        if mes:
            anio = int(anio_str)
            if anio < 100:
                anio += 2000
            return f"{anio:04d}-{mes:02d}-01"

    return None


# ---------------------------------------------------------------------------
# IMPORTAR VECINOS
# ---------------------------------------------------------------------------
def importar_vecinos(con: sqlite3.Connection, ws, id_comunidad: int) -> dict[int, int]:
    """
    Lee los vecinos del Excel y los inserta en propietarios.
    Devuelve {fila_excel: id_propietario}.
    """
    mapa = {}
    insertados = 0
    actualizados = 0

    for fila_idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True)):
        cod = row[0]  # columna A: código numérico
        if cod is None:
            continue

        vivienda = str(row[1] or "").strip()  # col B
        nombre   = str(row[2] or "").strip()  # col C
        coef_raw = row[3]                      # col D: coeficiente en enteros

        if not vivienda or not nombre:
            continue

        # Coeficiente — viene como string '0.676' o número
        try:
            coef = float(str(coef_raw).replace(',', '.')) if coef_raw else 0.0
        except ValueError:
            coef = 0.0

        fila_excel = fila_idx + 2

        # Buscar si ya existe
        existente = con.execute(
            "SELECT id_propietario FROM propietarios WHERE id_comunidad=? AND codigo_vivienda=?",
            (id_comunidad, vivienda)
        ).fetchone()

        if existente:
            con.execute(
                "UPDATE propietarios SET nombre_propietario=?, coeficiente=? WHERE id_propietario=?",
                (nombre, coef, existente["id_propietario"])
            )
            mapa[fila_excel] = existente["id_propietario"]
            actualizados += 1
        else:
            cur = con.execute(
                """INSERT INTO propietarios
                   (id_comunidad, codigo_vivienda, nombre_propietario, coeficiente,
                    estado_contador_acs, estado_contador_cal,
                    metodo_lectura_acs, metodo_lectura_cal)
                   VALUES (?,?,?,?,?,?,?,?)""",
                (id_comunidad, vivienda, nombre, coef,
                 "ok", "ok", "real", "real")
            )
            mapa[fila_excel] = cur.lastrowid
            insertados += 1

    con.commit()
    print(f"  Vecinos: {insertados} insertados, {actualizados} actualizados")
    return mapa


# ---------------------------------------------------------------------------
# IMPORTAR LECTURAS
# ---------------------------------------------------------------------------
def importar_lecturas_tipo(con: sqlite3.Connection, ws, id_comunidad: int,
                            mapa_vecinos: dict[int, int],
                            cabeceras: list, tipo: str,
                            col_inicio: int, col_fin: int,
                            id_periodo: int = None) -> int:
    """
    Importa lecturas acumuladas de un tipo (ACS o CALEFACCION) para un rango de columnas.
    Devuelve número de lecturas insertadas.
    """
    insertadas = 0
    duplicadas = 0

    # Filtrar sólo las columnas del rango que tienen fecha válida
    cols_con_fecha = [
        (col_idx, fecha_str)
        for col_idx, fecha_str in cabeceras
        if col_inicio <= col_idx <= col_fin and fecha_str
    ]

    for fila_excel, id_prop in mapa_vecinos.items():
        row = list(ws.iter_rows(min_row=fila_excel, max_row=fila_excel, values_only=True))[0]

        for col_idx, fecha_str in cols_con_fecha:
            if col_idx >= len(row):
                continue
            val = row[col_idx]

            # Ignorar valores nulos, fórmulas o no numéricos
            if val is None:
                continue
            sv = str(val).strip()
            if not sv or sv.startswith('=') or sv == '-':
                continue
            try:
                valor_acumulado = float(sv.replace(',', '.'))
            except ValueError:
                continue

            # Determinar estado de la lectura
            # Si el valor es igual al anterior, puede ser contador parado
            estado = "real"

            try:
                con.execute("""
                    INSERT INTO lecturas_vecino
                        (id_propietario, id_periodo, tipo, fecha_lectura,
                         valor_acumulado, estado, fuente)
                    VALUES (?,?,?,?,?,?,?)
                """, (id_prop, id_periodo, tipo, fecha_str,
                      valor_acumulado, estado, "excel_historico"))
                insertadas += 1
            except sqlite3.IntegrityError:
                duplicadas += 1  # ya existía

    con.commit()
    return insertadas


# ---------------------------------------------------------------------------
# FUNCIÓN PRINCIPAL
# ---------------------------------------------------------------------------
def importar_comunidad(codigo: str, nombre: str, cif: str,
                        ruta_excel_lecturas: str,
                        ruta_bd: str,
                        solo_vecinos: bool = False) -> dict:
    """
    Importa todos los datos de una comunidad a la BD.

    Args:
        codigo:               Código de comunidad (ej: '644')
        nombre:               Nombre completo
        cif:                  CIF de la comunidad (ej: 'H99258139')
        ruta_excel_lecturas:  Ruta al Excel CAL-ACS con lecturas históricas
        ruta_bd:              Ruta a gestion.db
        solo_vecinos:         Si True, solo importa vecinos (sin lecturas)
    """
    print(f"\n{'='*55}")
    print(f"  Importando comunidad {codigo}: {nombre}")
    print(f"{'='*55}\n")

    if not os.path.exists(ruta_bd):
        crear_bd(ruta_bd)
        print(f"  BD creada en: {ruta_bd}")

    con = conectar(ruta_bd)

    # 1. Registrar comunidad
    id_com = obtener_o_crear_comunidad(con, codigo, nombre, cif=cif)
    print(f"  Comunidad registrada (id={id_com})")

    # 2. Cargar Excel
    if not os.path.exists(ruta_excel_lecturas):
        print(f"  ⚠️  Excel no encontrado: {ruta_excel_lecturas}")
        con.close()
        return {"ok": False, "error": "Excel no encontrado"}

    wb = openpyxl.load_workbook(ruta_excel_lecturas, read_only=True, data_only=True)
    ws = wb.active
    print(f"  Excel cargado: {Path(ruta_excel_lecturas).name}")

    # 3. Importar vecinos
    print("\n  [1/3] Importando vecinos...")
    mapa_vecinos = importar_vecinos(con, ws, id_com)
    print(f"  → {len(mapa_vecinos)} vecinos en memoria")

    if solo_vecinos:
        con.close()
        return {"ok": True, "vecinos": len(mapa_vecinos)}

    # 4. Leer cabeceras de fechas
    row1 = list(ws.iter_rows(min_row=1, max_row=1, values_only=True))[0]
    cabeceras_acs  = []
    cabeceras_cal  = []

    for i, v in enumerate(row1):
        if v is None:
            continue
        fecha = _parsear_fecha_col(v)
        if not fecha:
            continue
        if 5 <= i <= 30:     # cols F a AE: ACS
            cabeceras_acs.append((i, fecha))
        elif 36 <= i <= 62:  # cols AK a BJ: CALEF
            cabeceras_cal.append((i, fecha))

    print(f"\n  [2/3] Importando lecturas históricas...")
    print(f"  Fechas ACS disponibles:  {len(cabeceras_acs)}")
    print(f"  Fechas CALEF disponibles: {len(cabeceras_cal)}")

    total_lecturas_acs  = 0
    total_lecturas_cal  = 0

    # 5. Crear periodos e importar lecturas por periodo
    for periodo_cfg in PERIODOS_644:
        nombre_per = periodo_cfg["nombre"]
        fecha_ini  = periodo_cfg["fecha_inicio"]
        fecha_fin  = periodo_cfg.get("fecha_fin")

        id_per = obtener_o_crear_periodo(con, id_com, nombre_per, fecha_ini)

        # Actualizar fecha_fin y estado
        estado = "abierto" if fecha_fin is None else "cerrado"
        con.execute(
            "UPDATE periodos SET fecha_fin=?, estado=? WHERE id_periodo=?",
            (fecha_fin, estado, id_per)
        )
        con.commit()

        # Config de suministro para este periodo
        for tipo_s, precio_var, precio_fijo, pct_a, pct_c in [
            ("ACS",         periodo_cfg["precio_acs_var"], periodo_cfg["precio_acs_fijo"],
             periodo_cfg["pct_acs"], periodo_cfg["pct_cal"]),
            ("CALEFACCION", periodo_cfg["precio_cal_var"], periodo_cfg["precio_cal_fijo"],
             periodo_cfg["pct_acs"], periodo_cfg["pct_cal"]),
        ]:
            try:
                con.execute("""
                    INSERT OR IGNORE INTO config_suministro
                        (id_comunidad, id_periodo, tipo_suministro, metodo_reparto,
                         precio_variable_facturado, precio_fijo_facturado,
                         pct_acs, pct_calefaccion)
                    VALUES (?,?,?,?,?,?,?,?)
                """, (id_com, id_per, tipo_s, "contador",
                      precio_var, precio_fijo, pct_a, pct_c))
            except sqlite3.IntegrityError:
                pass
        con.commit()

        # Lecturas ACS del periodo
        n_acs = importar_lecturas_tipo(
            con, ws, id_com, mapa_vecinos,
            cabeceras_acs, "ACS",
            periodo_cfg["col_acs_ini"],
            periodo_cfg["col_acs_fin"],
            id_periodo=id_per
        )

        # Lecturas CALEF del periodo
        n_cal = importar_lecturas_tipo(
            con, ws, id_com, mapa_vecinos,
            cabeceras_cal, "CALEFACCION",
            periodo_cfg["col_cal_ini"],
            periodo_cfg["col_cal_fin"],
            id_periodo=id_per
        )

        total_lecturas_acs += n_acs
        total_lecturas_cal += n_cal
        print(f"  Periodo {nombre_per}: {n_acs} lecturas ACS + {n_cal} lecturas CALEF")

    print(f"\n  [3/3] Resumen final:")
    print(f"  Vecinos importados:     {len(mapa_vecinos)}")
    print(f"  Lecturas ACS total:     {total_lecturas_acs}")
    print(f"  Lecturas CALEF total:   {total_lecturas_cal}")
    print(f"  Periodos registrados:   {len(PERIODOS_644)}")

    con.close()
    return {
        "ok": True,
        "id_comunidad":  id_com,
        "vecinos":       len(mapa_vecinos),
        "lecturas_acs":  total_lecturas_acs,
        "lecturas_cal":  total_lecturas_cal,
        "periodos":      len(PERIODOS_644),
    }


# ---------------------------------------------------------------------------
# EJECUCIÓN DIRECTA
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Importa datos de una comunidad a la BD")
    parser.add_argument("--comunidad",    default="644",  help="Código de comunidad")
    parser.add_argument("--nombre",       default="El Séptimo Sello 14-22 / El Tambor de Hojalata 13-21")
    parser.add_argument("--cif",          default="H99258139")
    parser.add_argument("--excel",        default=None,   help="Ruta al Excel CAL-ACS")
    parser.add_argument("--bd",           default=None,   help="Ruta a gestion.db")
    parser.add_argument("--solo-vecinos", action="store_true", help="Solo importar vecinos")
    args = parser.parse_args()

    # Rutas por defecto
    ruta_bd = args.bd or str(RUTA_BD)
    ruta_excel = args.excel or str(RUTA_EXCELS / f"{args.comunidad}_CAL-ACS_2025.xlsx")

    resultado = importar_comunidad(
        codigo=args.comunidad,
        nombre=args.nombre,
        cif=args.cif,
        ruta_excel_lecturas=ruta_excel,
        ruta_bd=ruta_bd,
        solo_vecinos=args.solo_vecinos,
    )

    if resultado["ok"]:
        print(f"\n✅ Importación completada correctamente")
    else:
        print(f"\n❌ Error: {resultado.get('error')}")
