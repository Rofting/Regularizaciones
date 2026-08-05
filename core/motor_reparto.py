"""
motor_reparto.py
================
Calcula el importe REAL que debe pagar cada vecino en un periodo,
comparándolo con lo que ya pagó (cuota facturada), y guarda la diferencia
en la tabla 'repartos' de la BD.

FLUJO:
    1. Lee todas las facturas del periodo (gas, luz, agua, mantenimiento)
    2. Lee los gastos extra con su amortización correspondiente
    3. Reparte cada concepto entre los vecinos según el método configurado
    4. Para ACS y Calefacción: lee lecturas individuales de contadores
    5. Calcula coste_real y lo compara con importe_cobrado
    6. Escribe el resultado en la tabla 'repartos'

MÉTODOS DE REPARTO (configurados en config_suministro):
    'coeficiente'   → proporcional al coeficiente de participación de cada vecino
    'contador'      → proporcional al consumo real del contador individual
    'cuota_fija'    → importe fijo igual para todos
    'partes_iguales'→ total / número de vecinos activos

CASOS ESPECIALES DE LECTURAS:
    estado='real'           → consumo = lectura_final - lectura_inicial
    estado='estimado'       → ídem, pero se marca como estimado en el reparto
    estado='sin_lectura'    → se usa promedio del resto de vecinos o coeficiente
    estado='contador_averiado' → igual que sin_lectura
    valor repetido (sin cambio) → consumo = 0 (piso vacío o contador parado)
"""

import sqlite3
from datetime import datetime
from typing import Optional


# ---------------------------------------------------------------------------
# CONSTANTES
# ---------------------------------------------------------------------------

METODOS_VALIDOS = ("coeficiente", "contador", "cuota_fija", "partes_iguales")

# Porcentajes de distribución de gastos comunes entre ACS y Calefacción
# cuando NO hay configuración explícita en config_suministro
PCT_ACS_DEFAULT    = 0.50
PCT_CALEF_DEFAULT  = 0.50

# Gastos que se reparten 50/50 entre ACS y Calefacción por defecto
GASTOS_MIXTOS = ("MANTENIMIENTO", "LECTURA_ACS", "LECTURA_CALEF")


# ---------------------------------------------------------------------------
# CONSULTAS A LA BD
# ---------------------------------------------------------------------------

def _obtener_periodo(cur, id_periodo: int) -> dict:
    row = cur.execute(
        "SELECT * FROM periodos WHERE id_periodo = ?", (id_periodo,)
    ).fetchone()
    if not row:
        raise ValueError(f"Periodo {id_periodo} no encontrado")
    return dict(row)


def _obtener_vecinos_activos(cur, id_comunidad: int) -> list[dict]:
    rows = cur.execute("""
        SELECT id_propietario, codigo_vivienda, nombre_propietario,
               coeficiente, estado_contador_acs, estado_contador_cal,
               metodo_lectura_acs, metodo_lectura_cal
        FROM propietarios
        WHERE id_comunidad = ? AND activo = 1
        ORDER BY codigo_vivienda
    """, (id_comunidad,)).fetchall()
    return [dict(r) for r in rows]


def _obtener_facturas_periodo(cur, id_comunidad: int, id_periodo: int) -> list[dict]:
    rows = cur.execute("""
        SELECT tipo_suministro, proveedor, importe_total,
               termino_fijo, termino_variable, impuestos, iva,
               consumo_total, unidad_consumo, fecha_inicio, fecha_fin
        FROM facturas
        WHERE id_comunidad = ? AND id_periodo = ?
        ORDER BY tipo_suministro, fecha_inicio
    """, (id_comunidad, id_periodo)).fetchall()
    return [dict(r) for r in rows]


def _obtener_gastos_extra_periodo(cur, id_comunidad: int,
                                   fecha_inicio: str, fecha_fin: str) -> list[dict]:
    """
    Devuelve los gastos extra cuyo periodo de amortización incluye este ejercicio.
    Un gasto con fecha=2022-01-01 y años_amortizacion=5 aplica a 2022,23,24,25,26.
    """
    rows = cur.execute("""
        SELECT id_gasto, descripcion, fecha, importe_total,
               años_amortizacion, tipo_gasto, servicio_afectado
        FROM gastos_extra
        WHERE id_comunidad = ? AND activo = 1
    """, (id_comunidad,)).fetchall()

    gastos_aplicables = []
    año_periodo = int(fecha_inicio[:4])

    for r in rows:
        g = dict(r)
        año_gasto = int(g["fecha"][:4])
        años_amort = max(1, g["años_amortizacion"])
        años_aplica = range(año_gasto, año_gasto + años_amort)
        if año_periodo in años_aplica:
            g["importe_anual"] = round(g["importe_total"] / años_amort, 6)
            gastos_aplicables.append(g)

    return gastos_aplicables


def _obtener_config_suministro(cur, id_comunidad: int, id_periodo: int,
                                tipo: str) -> Optional[dict]:
    row = cur.execute("""
        SELECT * FROM config_suministro
        WHERE id_comunidad = ? AND id_periodo = ? AND tipo_suministro = ?
    """, (id_comunidad, id_periodo, tipo)).fetchone()
    return dict(row) if row else None


def _dias_entre(fecha_a: str, fecha_b: str) -> int:
    """
    Días naturales entre dos fechas en formato YYYY-MM-DD.
    Siempre devuelve valor positivo; si alguna fecha falta devuelve 0.
    """
    if not fecha_a or not fecha_b:
        return 0
    try:
        from datetime import date as _date
        a = _date.fromisoformat(fecha_a[:10])
        b = _date.fromisoformat(fecha_b[:10])
        return abs((b - a).days)
    except ValueError:
        return 0


def _interpolar_lectura(val_ant: float, fecha_ant: str,
                         val_pos: float, fecha_pos: str,
                         fecha_objetivo: str) -> float:
    """
    Interpola linealmente el valor de contador en fecha_objetivo
    entre dos lecturas reales (anterior y posterior).
    Se usa cuando la lectura no cae exactamente el día de inicio/fin del periodo.

    Ejemplo:
        val_ant=1000  fecha_ant=2024-08-25
        val_pos=1060  fecha_pos=2024-09-05
        fecha_objetivo=2024-09-01  → 1000 + 60 * (7/11) ≈ 1038.18
    """
    d_total = _dias_entre(fecha_ant, fecha_pos)
    if d_total == 0:
        return val_ant
    d_objetivo = _dias_entre(fecha_ant, fecha_objetivo)
    fraccion = min(1.0, max(0.0, d_objetivo / d_total))
    return round(val_ant + (val_pos - val_ant) * fraccion, 4)


def _detectar_reinicio_contador(cur, id_propietario: int, tipo: str,
                                 fecha_inicio: str, fecha_fin: str) -> Optional[dict]:
    """
    Un contador acumulado nunca debería bajar de valor con el tiempo. Si baja,
    es que el aparato físico se sustituyó y el nuevo arrancó desde ~0 (caso
    real detectado en la 644: cientos/miles de litros acumulados pasan a un
    valor de un solo dígito de un periodo a otro, en decenas de viviendas a
    la vez).

    Devuelve el primer salto descendente encontrado dentro de
    [fecha_inicio, fecha_fin], o None si no hay ninguno.
    """
    filas = cur.execute("""
        SELECT fecha_lectura, valor_acumulado
        FROM lecturas_vecino
        WHERE id_propietario=? AND tipo=? AND fecha_lectura BETWEEN ? AND ?
        ORDER BY fecha_lectura
    """, (id_propietario, tipo, fecha_inicio, fecha_fin)).fetchall()

    for i in range(1, len(filas)):
        if filas[i]["valor_acumulado"] < filas[i - 1]["valor_acumulado"]:
            return {
                "fecha_antes":   filas[i - 1]["fecha_lectura"],
                "valor_antes":   filas[i - 1]["valor_acumulado"],
                "fecha_despues": filas[i]["fecha_lectura"],
                "valor_despues": filas[i]["valor_acumulado"],
            }
    return None


def _obtener_lecturas_vecino(cur, id_propietario: int, tipo: str,
                              fecha_inicio: str, fecha_fin: str) -> dict:
    """
    Devuelve lecturas inicial y final del vecino para el rango exacto del periodo.

    Estrategia de búsqueda (por orden de preferencia):
      1. Lectura exacta el día de inicio / el día de fin → uso directo
      2. Lectura el mismo mes (cualquier día) → interpolación lineal con la
         lectura inmediatamente anterior y la inmediatamente posterior
      3. Lectura en el rango ampliado ±45 días → interpolación lineal
      4. Sin lecturas → None (el motor usará promedio de la comunidad)

    La función devuelve siempre valores INTERPOLADOS al día exacto del periodo,
    no los valores brutos del contador, para que el consumo calculado sea
    correcto aunque las lecturas se tomaran el día 1 o el día 28/31.

    Si el contador se sustituyó dentro del periodo (ver
    _detectar_reinicio_contador), "final" se sustituye por un valor virtual
    en escala continua (inicial + consumo real de los dos tramos), para que
    consumo = final − inicial siga siendo correcto aguas abajo sin tener que
    tocar todo el código que ya asume esa resta directa.
    """
    # ── LECTURA INICIAL ──────────────────────────────────────────────────────
    # Lectura exacta en fecha_inicio
    lec_ini_exacta = cur.execute("""
        SELECT fecha_lectura, valor_acumulado, estado
        FROM lecturas_vecino
        WHERE id_propietario=? AND tipo=? AND fecha_lectura=?
        ORDER BY fecha_lectura DESC LIMIT 1
    """, (id_propietario, tipo, fecha_inicio)).fetchone()

    if lec_ini_exacta:
        inicial = dict(lec_ini_exacta)
        inicial["interpolada"] = False
    else:
        # Lectura justo antes del inicio (la más reciente ≤ fecha_inicio)
        antes = cur.execute("""
            SELECT fecha_lectura, valor_acumulado, estado
            FROM lecturas_vecino
            WHERE id_propietario=? AND tipo=? AND fecha_lectura<?
            ORDER BY fecha_lectura DESC LIMIT 1
        """, (id_propietario, tipo, fecha_inicio)).fetchone()

        # Lectura justo después del inicio (la más próxima > fecha_inicio)
        despues = cur.execute("""
            SELECT fecha_lectura, valor_acumulado, estado
            FROM lecturas_vecino
            WHERE id_propietario=? AND tipo=? AND fecha_lectura>?
            ORDER BY fecha_lectura ASC LIMIT 1
        """, (id_propietario, tipo, fecha_inicio)).fetchone()

        # Si el contador se sustituyó entre "antes" y "despues", interpolar
        # entre ambos mezclaría dos aparatos distintos y daría un valor sin
        # sentido (ej. entre 1023 del contador viejo y 10 del nuevo). En ese
        # caso se descarta la interpolación y se cae a las ramas de abajo
        # (usar solo "antes" si está lo bastante cerca, si no, sin lectura).
        hay_reinicio_entre = bool(antes and despues and _detectar_reinicio_contador(
            cur, id_propietario, tipo, antes["fecha_lectura"], despues["fecha_lectura"]
        ))

        if antes and despues and not hay_reinicio_entre:
            val_interp = _interpolar_lectura(
                antes["valor_acumulado"], antes["fecha_lectura"],
                despues["valor_acumulado"], despues["fecha_lectura"],
                fecha_inicio
            )
            inicial = {
                "fecha_lectura":   fecha_inicio,
                "valor_acumulado": val_interp,
                "estado":          "interpolado",
                "interpolada":     True,
                "referencia":      f"{antes['fecha_lectura']}→{despues['fecha_lectura']}",
            }
        elif antes:
            # Solo lectura anterior: la usamos tal cual (puede haber desfase pequeño)
            d = _dias_entre(antes["fecha_lectura"], fecha_inicio)
            if d <= 45:
                inicial = {
                    "fecha_lectura":   fecha_inicio,
                    "valor_acumulado": antes["valor_acumulado"],
                    "estado":          antes["estado"],
                    "interpolada":     True,
                    "referencia":      f"aprox desde {antes['fecha_lectura']} ({d}d)",
                }
            else:
                inicial = None
        else:
            inicial = None

    # ── LECTURA FINAL ─────────────────────────────────────────────────────────
    lec_fin_exacta = cur.execute("""
        SELECT fecha_lectura, valor_acumulado, estado
        FROM lecturas_vecino
        WHERE id_propietario=? AND tipo=? AND fecha_lectura=?
        ORDER BY fecha_lectura DESC LIMIT 1
    """, (id_propietario, tipo, fecha_fin)).fetchone()

    if lec_fin_exacta:
        final = dict(lec_fin_exacta)
        final["interpolada"] = False
    else:
        antes_fin = cur.execute("""
            SELECT fecha_lectura, valor_acumulado, estado
            FROM lecturas_vecino
            WHERE id_propietario=? AND tipo=? AND fecha_lectura<?
            ORDER BY fecha_lectura DESC LIMIT 1
        """, (id_propietario, tipo, fecha_fin)).fetchone()

        despues_fin = cur.execute("""
            SELECT fecha_lectura, valor_acumulado, estado
            FROM lecturas_vecino
            WHERE id_propietario=? AND tipo=? AND fecha_lectura>?
            ORDER BY fecha_lectura ASC LIMIT 1
        """, (id_propietario, tipo, fecha_fin)).fetchone()

        hay_reinicio_entre_fin = bool(antes_fin and despues_fin and _detectar_reinicio_contador(
            cur, id_propietario, tipo, antes_fin["fecha_lectura"], despues_fin["fecha_lectura"]
        ))

        if antes_fin and despues_fin and not hay_reinicio_entre_fin:
            val_interp = _interpolar_lectura(
                antes_fin["valor_acumulado"], antes_fin["fecha_lectura"],
                despues_fin["valor_acumulado"], despues_fin["fecha_lectura"],
                fecha_fin
            )
            final = {
                "fecha_lectura":   fecha_fin,
                "valor_acumulado": val_interp,
                "estado":          "interpolado",
                "interpolada":     True,
                "referencia":      f"{antes_fin['fecha_lectura']}→{despues_fin['fecha_lectura']}",
            }
        elif antes_fin:
            d = _dias_entre(antes_fin["fecha_lectura"], fecha_fin)
            if d <= 45:
                final = {
                    "fecha_lectura":   fecha_fin,
                    "valor_acumulado": antes_fin["valor_acumulado"],
                    "estado":          antes_fin["estado"],
                    "interpolada":     True,
                    "referencia":      f"aprox desde {antes_fin['fecha_lectura']} ({d}d)",
                }
            else:
                final = None
        else:
            final = None

    # ── Contador sustituido dentro del periodo ───────────────────────────────
    # Si el valor bajó en algún punto entre fecha_inicio y fecha_fin, "final"
    # crudo (contador nuevo) sería menor que "inicial" (contador viejo) y el
    # consumo saldría en 0 o negativo. Se sustituye "final" por un valor
    # virtual en escala continua: inicial + consumo real (tramo del contador
    # viejo hasta la sustitución + tramo del contador nuevo hasta el final).
    if inicial and final:
        reinicio = _detectar_reinicio_contador(cur, id_propietario, tipo, fecha_inicio, fecha_fin)
        if reinicio:
            consumo_tramo1 = max(0.0, reinicio["valor_antes"] - inicial["valor_acumulado"])
            consumo_tramo2 = max(0.0, final["valor_acumulado"] - reinicio["valor_despues"])
            consumo_total = round(consumo_tramo1 + consumo_tramo2, 4)
            final = {
                "fecha_lectura":   fecha_fin,
                "valor_acumulado": round(inicial["valor_acumulado"] + consumo_total, 4),
                "estado":          "real",
                "interpolada":     True,
                "referencia": (
                    f"contador sustituido entre {reinicio['fecha_antes']} "
                    f"({reinicio['valor_antes']}) y {reinicio['fecha_despues']} "
                    f"({reinicio['valor_despues']}); consumo real={consumo_total}"
                ),
            }

    return {"inicial": inicial, "final": final}


def _obtener_importe_cobrado(cur, id_propietario: int, id_periodo: int,
                              tipo: str) -> float:
    """
    Lee la cuota ya facturada al vecino en este periodo y suministro.
    Viene de la tabla repartos si ya fue introducida manualmente,
    o se recalcula desde config_suministro × meses × lecturas.
    Por ahora devuelve 0 si no hay dato previo (se rellenará después
    desde el proceso de importación de cuotas cobradas).
    """
    row = cur.execute("""
        SELECT importe_cobrado FROM repartos
        WHERE id_propietario = ? AND id_periodo = ? AND tipo_suministro = ?
    """, (id_propietario, id_periodo, tipo)).fetchone()
    return row["importe_cobrado"] if row else 0.0


# ---------------------------------------------------------------------------
# CÁLCULO DEL REPARTO
# ---------------------------------------------------------------------------

def _calcular_consumo_vecino(lecturas: dict, vecino: dict,
                              todos_consumos: list[float]) -> tuple[float, float, float, str]:
    """
    Calcula el consumo individual del vecino.
    Devuelve (lectura_inicial, lectura_final, consumo, nota).
    Las lecturas pueden ser valores exactos o interpolados al día del periodo.
    """
    ini = lecturas.get("inicial")
    fin = lecturas.get("final")

    if not ini or not fin:
        consumos_reales = [c for c in todos_consumos if c > 0]
        promedio = round(sum(consumos_reales) / len(consumos_reales), 4) if consumos_reales else 0.0
        return 0.0, 0.0, promedio, "estimado_promedio"

    val_ini = ini["valor_acumulado"]
    val_fin = fin["valor_acumulado"]
    consumo = max(0.0, round(val_fin - val_ini, 4))

    ini_interp = ini.get("interpolada", False)
    fin_interp = fin.get("interpolada", False)
    estado_fin = fin.get("estado", "real")

    if ini_interp or fin_interp:
        ref_ini = ini.get("referencia", "")
        ref_fin = fin.get("referencia", "")
        nota = f"interpolado({ref_ini}|{ref_fin})" if ref_ini or ref_fin else "interpolado"
    elif consumo == 0 and val_ini == val_fin and val_ini > 0:
        nota = "contador_sin_variacion"
    elif estado_fin in ("estimado",):
        nota = "estimado"
    else:
        nota = "real"

    return val_ini, val_fin, consumo, nota


def _repartir_por_coeficiente(vecinos: list[dict], importe_total: float) -> dict[int, float]:
    """Reparte proporcionalmente al coeficiente de participación."""
    suma_coef = sum(v["coeficiente"] for v in vecinos)
    if suma_coef == 0:
        return {v["id_propietario"]: 0.0 for v in vecinos}
    return {
        v["id_propietario"]: round(importe_total * v["coeficiente"] / suma_coef, 6)
        for v in vecinos
    }


def _repartir_por_contador(vecinos_consumos: dict[int, float],
                            importe_total: float) -> dict[int, float]:
    """Reparte proporcionalmente al consumo real del contador."""
    total_consumo = sum(vecinos_consumos.values())
    if total_consumo == 0:
        n = len(vecinos_consumos)
        return {k: round(importe_total / n, 6) for k in vecinos_consumos} if n else {}
    return {
        k: round(importe_total * consumo / total_consumo, 6)
        for k, consumo in vecinos_consumos.items()
    }


def _repartir_partes_iguales(vecinos: list[dict], importe_total: float) -> dict[int, float]:
    n = len(vecinos)
    if n == 0:
        return {}
    cuota = round(importe_total / n, 6)
    return {v["id_propietario"]: cuota for v in vecinos}


# ---------------------------------------------------------------------------
# FUNCIÓN PRINCIPAL
# ---------------------------------------------------------------------------

def calcular_reparto(con: sqlite3.Connection, id_comunidad: int,
                     id_periodo: int, sobrescribir: bool = False) -> dict:
    """
    Calcula el reparto completo de un periodo para una comunidad.

    Args:
        con:            Conexión SQLite activa (con row_factory = sqlite3.Row)
        id_comunidad:   ID de la comunidad
        id_periodo:     ID del periodo a calcular
        sobrescribir:   Si True, borra repartos previos y recalcula desde cero

    Returns:
        dict con resumen: vecinos_procesados, total_calculado, errores, avisos
    """
    cur = con.cursor()
    periodo = _obtener_periodo(cur, id_periodo)
    vecinos = _obtener_vecinos_activos(cur, id_comunidad)
    facturas = _obtener_facturas_periodo(cur, id_comunidad, id_periodo)
    gastos_extra = _obtener_gastos_extra_periodo(
        cur, id_comunidad, periodo["fecha_inicio"], periodo["fecha_fin"] or ""
    )

    if not vecinos:
        return {"ok": False, "error": "No hay vecinos activos en esta comunidad"}
    if not facturas and not gastos_extra:
        return {"ok": False, "error": "No hay facturas ni gastos extra para este periodo"}

    suma_coef_total = sum(v["coeficiente"] for v in vecinos)
    errores = []
    avisos = []
    registros_guardados = 0

    # ------------------------------------------------------------------
    # PASO 1: Agrupar facturas por tipo de suministro
    # ------------------------------------------------------------------
    totales_por_tipo: dict[str, float] = {}
    for f in facturas:
        t = f["tipo_suministro"]
        totales_por_tipo[t] = totales_por_tipo.get(t, 0.0) + f["importe_total"]

    # ------------------------------------------------------------------
    # PASO 2: Añadir gastos extra amortizados
    # ------------------------------------------------------------------
    gastos_acs    = 0.0
    gastos_calef  = 0.0
    gastos_comunes = 0.0

    for g in gastos_extra:
        servicio = (g["servicio_afectado"] or "AMBOS").upper()
        importe = g["importe_anual"]
        if servicio == "ACS":
            gastos_acs += importe
        elif servicio == "CALEFACCION":
            gastos_calef += importe
        else:  # AMBOS o None → 50/50
            gastos_acs   += importe * 0.5
            gastos_calef += importe * 0.5

    # ------------------------------------------------------------------
    # PASO 3: Calcular porcentajes ACS/Calef del combustible (Gas)
    # ------------------------------------------------------------------
    config_acs   = _obtener_config_suministro(cur, id_comunidad, id_periodo, "ACS")
    config_calef = _obtener_config_suministro(cur, id_comunidad, id_periodo, "CALEFACCION")

    pct_acs  = config_acs["pct_acs"]   if config_acs   else PCT_ACS_DEFAULT
    pct_cal  = config_calef["pct_calefaccion"] if config_calef else PCT_CALEF_DEFAULT

    # Coste total de Gas dividido entre ACS y Calefacción
    total_gas = totales_por_tipo.get("GAS", 0.0)
    total_luz = totales_por_tipo.get("ELECTRICIDAD", 0.0)
    total_agua = totales_por_tipo.get("AGUA", 0.0)
    total_mantenimiento = totales_por_tipo.get("MANTENIMIENTO", 0.0)

    coste_combustible_acs  = round(total_gas * pct_acs  + total_luz * pct_acs,  4)
    coste_combustible_cal  = round(total_gas * pct_cal  + total_luz * pct_cal,  4)

    # Mantenimiento y lecturas: 50/50 por defecto
    coste_mto_acs  = round(total_mantenimiento * 0.5, 4)
    coste_mto_cal  = round(total_mantenimiento * 0.5, 4)

    # Agua: 100% ACS
    coste_agua_acs = round(total_agua, 4)

    # Coste total por servicio
    coste_total_acs  = round(coste_combustible_acs  + coste_agua_acs  + coste_mto_acs  + gastos_acs,  4)
    coste_total_cal  = round(coste_combustible_cal  + coste_mto_cal  + gastos_calef, 4)

    # ------------------------------------------------------------------
    # PASO 4: Leer consumos individuales de contadores
    # ------------------------------------------------------------------
    consumos_acs:  dict[int, float] = {}
    consumos_cal:  dict[int, float] = {}
    lecturas_raw:  dict[int, dict]  = {}

    for v in vecinos:
        pid = v["id_propietario"]
        lecturas_raw[pid] = {
            "ACS":  _obtener_lecturas_vecino(cur, pid, "ACS",
                                              periodo["fecha_inicio"],
                                              periodo["fecha_fin"] or "9999-12-31"),
            "CALEFACCION": _obtener_lecturas_vecino(cur, pid, "CALEFACCION",
                                                     periodo["fecha_inicio"],
                                                     periodo["fecha_fin"] or "9999-12-31"),
        }

    # Primera pasada: calcular consumos reales
    for v in vecinos:
        pid = v["id_propietario"]
        _, _, c_acs, _  = _calcular_consumo_vecino(lecturas_raw[pid]["ACS"], v, [])
        _, _, c_cal, _  = _calcular_consumo_vecino(lecturas_raw[pid]["CALEFACCION"], v, [])
        consumos_acs[pid] = c_acs
        consumos_cal[pid] = c_cal

    # Segunda pasada: reemplazar estimados con promedio de los reales
    media_acs = _promedio_sin_cero(list(consumos_acs.values()))
    media_cal = _promedio_sin_cero(list(consumos_cal.values()))

    for v in vecinos:
        pid = v["id_propietario"]
        lec = lecturas_raw[pid]
        if lec["ACS"]["inicial"] is None or lec["ACS"]["final"] is None:
            consumos_acs[pid] = media_acs
        if lec["CALEFACCION"]["inicial"] is None or lec["CALEFACCION"]["final"] is None:
            consumos_cal[pid] = media_cal

    # ------------------------------------------------------------------
    # PASO 5: Calcular coste real por vecino
    # ------------------------------------------------------------------

    # Reparto ACS por contador
    reparto_acs  = _repartir_por_contador(consumos_acs, coste_total_acs)

    # Reparto Calefacción por contador
    reparto_cal  = _repartir_por_contador(consumos_cal, coste_total_cal)

    # Agua fría: por coeficiente (no hay contadores individuales de agua fría)
    config_agua = _obtener_config_suministro(cur, id_comunidad, id_periodo, "AGUA")
    metodo_agua = config_agua["metodo_reparto"] if config_agua else "coeficiente"
    if metodo_agua == "coeficiente":
        reparto_agua_fria = _repartir_por_coeficiente(vecinos, coste_agua_acs)
    else:
        reparto_agua_fria = _repartir_partes_iguales(vecinos, coste_agua_acs)

    # Luz: por coeficiente
    reparto_luz = _repartir_por_coeficiente(vecinos, total_luz * (1 - pct_acs - pct_cal + 1))
    # (La luz ya está repartida dentro de coste_combustible_acs/cal; aquí no se cobra aparte)

    # ------------------------------------------------------------------
    # PASO 6: Guardar resultados en tabla repartos
    # ------------------------------------------------------------------
    fecha_calculo = datetime.now().isoformat()

    for v in vecinos:
        pid = v["id_propietario"]
        lec = lecturas_raw[pid]

        for tipo_serv, reparto, lec_tipo in [
            ("ACS",         reparto_acs,  "ACS"),
            ("CALEFACCION", reparto_cal,  "CALEFACCION"),
        ]:
            importe_real   = reparto.get(pid, 0.0)
            importe_cobrado = _obtener_importe_cobrado(cur, pid, id_periodo, tipo_serv)
            diferencia     = round(importe_real - importe_cobrado, 4)

            lec_datos = lec[lec_tipo]
            ini = lec_datos["inicial"]
            fin = lec_datos["final"]

            lec_inicial = ini["valor_acumulado"] if ini else 0.0
            lec_final   = fin["valor_acumulado"] if fin else 0.0
            consumo     = max(0.0, round(lec_final - lec_inicial, 4))

            _, _, consumo_calc, nota = _calcular_consumo_vecino(
                lec_datos, v,
                list(consumos_acs.values()) if tipo_serv == "ACS"
                else list(consumos_cal.values())
            )

            # Upsert: insertar o actualizar si sobrescribir=True
            existente = cur.execute("""
                SELECT id_reparto FROM repartos
                WHERE id_propietario = ? AND id_periodo = ? AND tipo_suministro = ?
            """, (pid, id_periodo, tipo_serv)).fetchone()

            if existente and not sobrescribir:
                avisos.append(f"Reparto existente para {v['codigo_vivienda']} {tipo_serv} — omitido")
                continue

            if existente and sobrescribir:
                cur.execute("""
                    UPDATE repartos SET
                        lectura_inicial=?, lectura_final=?, consumo_real=?,
                        importe_cobrado=?, importe_real=?, diferencia=?,
                        estado='calculado', fecha_calculo=?, notas=?
                    WHERE id_propietario=? AND id_periodo=? AND tipo_suministro=?
                """, (lec_inicial, lec_final, consumo_calc,
                      importe_cobrado, round(importe_real, 4), diferencia,
                      fecha_calculo, nota,
                      pid, id_periodo, tipo_serv))
            else:
                cur.execute("""
                    INSERT INTO repartos
                        (id_propietario, id_periodo, tipo_suministro,
                         lectura_inicial, lectura_final, consumo_real,
                         importe_cobrado, importe_real, diferencia,
                         estado, fecha_calculo, notas)
                    VALUES (?,?,?,?,?,?,?,?,?,?,?,?)
                """, (pid, id_periodo, tipo_serv,
                      lec_inicial, lec_final, consumo_calc,
                      importe_cobrado, round(importe_real, 4), diferencia,
                      "calculado", fecha_calculo, nota))

            registros_guardados += 1

    con.commit()

    # ------------------------------------------------------------------
    # Resumen
    # ------------------------------------------------------------------
    total_acs_calculado  = sum(reparto_acs.values())
    total_cal_calculado  = sum(reparto_cal.values())

    return {
        "ok": True,
        "periodo": periodo["nombre"],
        "vecinos_procesados": len(vecinos),
        "registros_guardados": registros_guardados,
        "resumen_costes": {
            "gas_total":          round(total_gas, 2),
            "luz_total":          round(total_luz, 2),
            "agua_total":         round(total_agua, 2),
            "mantenimiento_total": round(total_mantenimiento, 2),
            "gastos_extra_acs":   round(gastos_acs, 2),
            "gastos_extra_cal":   round(gastos_calef, 2),
            "coste_total_acs":    round(coste_total_acs, 2),
            "coste_total_cal":    round(coste_total_cal, 2),
            "pct_acs":            round(pct_acs * 100, 1),
            "pct_cal":            round(pct_cal * 100, 1),
        },
        "resumen_repartos": {
            "total_acs_distribuido":  round(total_acs_calculado, 2),
            "total_cal_distribuido":  round(total_cal_calculado, 2),
            "media_acs_vecino":       round(total_acs_calculado / len(vecinos), 2) if vecinos else 0,
            "media_cal_vecino":       round(total_cal_calculado / len(vecinos), 2) if vecinos else 0,
        },
        "errores": errores,
        "avisos": avisos,
    }


def _promedio_sin_cero(valores: list[float]) -> float:
    """Promedio excluyendo ceros (vecinos vacíos o sin lectura)."""
    reales = [v for v in valores if v > 0]
    return round(sum(reales) / len(reales), 4) if reales else 0.0


# ---------------------------------------------------------------------------
# FUNCIONES DE CONSULTA — usadas por excel_writer y carta_writer
# ---------------------------------------------------------------------------

def obtener_repartos_vecino(con: sqlite3.Connection, id_propietario: int,
                             id_periodo: int) -> list[dict]:
    """Devuelve todos los conceptos del reparto de un vecino en un periodo."""
    rows = con.execute("""
        SELECT r.tipo_suministro, r.lectura_inicial, r.lectura_final,
               r.consumo_real, r.importe_cobrado, r.importe_real,
               r.diferencia, r.estado, r.notas,
               p.nombre_propietario, p.codigo_vivienda, p.coeficiente,
               per.nombre as periodo_nombre
        FROM repartos r
        JOIN propietarios p ON p.id_propietario = r.id_propietario
        JOIN periodos per ON per.id_periodo = r.id_periodo
        WHERE r.id_propietario = ? AND r.id_periodo = ?
        ORDER BY r.tipo_suministro
    """, (id_propietario, id_periodo)).fetchall()
    return [dict(r) for r in rows]


def obtener_resumen_periodo(con: sqlite3.Connection, id_comunidad: int,
                             id_periodo: int) -> list[dict]:
    """Devuelve el resumen de todos los vecinos para generar el Excel o las cartas."""
    rows = con.execute("""
        SELECT p.codigo_vivienda, p.nombre_propietario, p.coeficiente,
               r.tipo_suministro,
               r.lectura_inicial, r.lectura_final, r.consumo_real,
               r.importe_cobrado, r.importe_real, r.diferencia,
               r.estado, r.notas
        FROM repartos r
        JOIN propietarios p ON p.id_propietario = r.id_propietario
        WHERE p.id_comunidad = ? AND r.id_periodo = ?
        ORDER BY p.codigo_vivienda, r.tipo_suministro
    """, (id_comunidad, id_periodo)).fetchall()
    return [dict(r) for r in rows]


def obtener_total_vecino(con: sqlite3.Connection, id_propietario: int,
                          id_periodo: int) -> dict:
    """
    Agrega todos los conceptos del vecino en un solo total.
    Devuelve: {cobrado_total, real_total, diferencia_total, desglose: [...]}
    """
    repartos = obtener_repartos_vecino(con, id_propietario, id_periodo)
    cobrado  = sum(r["importe_cobrado"] for r in repartos)
    real     = sum(r["importe_real"]    for r in repartos)
    dif      = sum(r["diferencia"]      for r in repartos)
    return {
        "cobrado_total":    round(cobrado, 2),
        "real_total":       round(real, 2),
        "diferencia_total": round(dif, 2),
        "desglose":         repartos,
    }


# ---------------------------------------------------------------------------
# EJECUCIÓN DIRECTA — test con BD de prueba
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    import sys
    import os
    sys.path.insert(0, os.path.dirname(__file__))

    from gestor_bd import crear_bd, conectar, obtener_o_crear_comunidad, obtener_o_crear_periodo

    RUTA_TEST = "/tmp/test_motor/gestion.db"
    os.makedirs("/tmp/test_motor", exist_ok=True)
    crear_bd(RUTA_TEST)

    con = conectar(RUTA_TEST)

    # Insertar datos de prueba mínimos
    id_com = obtener_o_crear_comunidad(con, "644", "Séptimo Sello", "H99258139")
    id_per = obtener_o_crear_periodo(con, id_com, "2023-2024", "2023-09-01")
    con.execute("UPDATE periodos SET fecha_fin='2024-08-31' WHERE id_periodo=?", (id_per,))

    # 3 vecinos de prueba
    vecinos_test = [
        ("SS22 BAJO IZDA", "SILVIA MORER",  0.676),
        ("SS22 BAJO DCHA", "JAVIER REDONDO", 0.752),
        ("SS22 1 IZDA",    "JESUS LORENZO",  0.752),
    ]
    ids_v = []
    for cod, nom, coef in vecinos_test:
        cur2 = con.execute(
            "INSERT OR IGNORE INTO propietarios (id_comunidad, codigo_vivienda, nombre_propietario, coeficiente) VALUES (?,?,?,?)",
            (id_com, cod, nom, coef)
        )
        ids_v.append(con.execute(
            "SELECT id_propietario FROM propietarios WHERE id_comunidad=? AND codigo_vivienda=?",
            (id_com, cod)
        ).fetchone()["id_propietario"])
    con.commit()

    # Facturas de prueba
    for tipo, imp in [("GAS", 61415.90), ("ELECTRICIDAD", 0.0), ("AGUA", 4298.63), ("MANTENIMIENTO", 3867.26)]:
        if imp > 0:
            con.execute("""INSERT OR IGNORE INTO facturas
                (id_comunidad, id_periodo, tipo_suministro, proveedor, num_factura,
                 cups_o_referencia, fecha_factura, importe_total)
                VALUES (?,?,?,?,?,?,?,?)""",
                (id_com, id_per, tipo, "TEST", f"TEST-{tipo}",
                 "CUPS_TEST", "2024-01-01", imp))
    con.commit()

    # Config: 51% ACS, 49% Calef para el gas
    con.execute("""INSERT OR IGNORE INTO config_suministro
        (id_comunidad, id_periodo, tipo_suministro, metodo_reparto,
         precio_variable_facturado, precio_fijo_facturado,
         pct_acs, pct_calefaccion)
        VALUES (?,?,?,?,?,?,?,?)""",
        (id_com, id_per, "ACS", "contador", 6.0, 7.5, 0.511, 0.489))
    con.execute("""INSERT OR IGNORE INTO config_suministro
        (id_comunidad, id_periodo, tipo_suministro, metodo_reparto,
         precio_variable_facturado, precio_fijo_facturado,
         pct_acs, pct_calefaccion)
        VALUES (?,?,?,?,?,?,?,?)""",
        (id_com, id_per, "CALEFACCION", "contador", 0.20, 12.5, 0.511, 0.489))
    con.commit()

    # Lecturas de prueba (ACS en m3 acumulados)
    for pid, val_ini, val_fin in [
        (ids_v[0], 29.0, 32.0),   # 3 m3
        (ids_v[1], 188.0, 196.0), # 8 m3
        (ids_v[2], 207.0, 231.0), # 24 m3
    ]:
        con.execute("""INSERT OR IGNORE INTO lecturas_vecino
            (id_propietario, id_periodo, tipo, fecha_lectura, valor_acumulado, estado, fuente)
            VALUES (?,?,?,?,?,?,?)""",
            (pid, id_per, "ACS", "2023-09-01", val_ini, "real", "test"))
        con.execute("""INSERT OR IGNORE INTO lecturas_vecino
            (id_propietario, id_periodo, tipo, fecha_lectura, valor_acumulado, estado, fuente)
            VALUES (?,?,?,?,?,?,?)""",
            (pid, id_per, "ACS", "2024-08-31", val_fin, "real", "test"))

    # Lecturas Calef (kWh acumulados)
    for pid, val_ini, val_fin in [
        (ids_v[0], 17371.0, 18455.0),  # 1084 kWh
        (ids_v[1], 28210.0, 28210.0),  # 0 kWh — contador parado
        (ids_v[2], 23196.0, 23293.0),  # 97 kWh
    ]:
        con.execute("""INSERT OR IGNORE INTO lecturas_vecino
            (id_propietario, id_periodo, tipo, fecha_lectura, valor_acumulado, estado, fuente)
            VALUES (?,?,?,?,?,?,?)""",
            (pid, id_per, "CALEFACCION", "2023-09-01", val_ini, "real", "test"))
        con.execute("""INSERT OR IGNORE INTO lecturas_vecino
            (id_propietario, id_periodo, tipo, fecha_lectura, valor_acumulado, estado, fuente)
            VALUES (?,?,?,?,?,?,?)""",
            (pid, id_per, "CALEFACCION", "2024-08-31", val_fin, "real", "test"))
    con.commit()

    # Gasto extra con amortización
    con.execute("""INSERT OR IGNORE INTO gastos_extra
        (id_comunidad, descripcion, fecha, importe_total, años_amortizacion,
         tipo_gasto, servicio_afectado)
        VALUES (?,?,?,?,?,?,?)""",
        (id_com, "Rep. Sala Calderas bomba", "2022-01-01", 3619.45, 5, "REPARACION", "AMBOS"))
    con.commit()

    # EJECUTAR EL MOTOR
    print("\n=== EJECUTANDO MOTOR DE REPARTO ===\n")
    resultado = calcular_reparto(con, id_com, id_per, sobrescribir=True)

    if resultado["ok"]:
        print(f"Periodo: {resultado['periodo']}")
        print(f"Vecinos: {resultado['vecinos_procesados']}")
        print(f"Registros guardados: {resultado['registros_guardados']}")
        print()
        print("COSTES TOTALES:")
        for k, v in resultado["resumen_costes"].items():
            print(f"  {k:<30} {v}")
        print()
        print("DISTRIBUCIÓN:")
        for k, v in resultado["resumen_repartos"].items():
            print(f"  {k:<35} {v}")
        print()
        print("DETALLE POR VECINO:")
        print(f"  {'Vivienda':<20} {'Suministro':<14} {'Consumo':>10} {'Cobrado':>10} {'Real':>10} {'Diferencia':>12}")
        print("  " + "-"*78)
        for v, pid in zip(vecinos_test, ids_v):
            total = obtener_total_vecino(con, pid, id_per)
            for d in total["desglose"]:
                print(f"  {d['codigo_vivienda']:<20} {d['tipo_suministro']:<14} "
                      f"{d['consumo_real']:>10.2f} {d['importe_cobrado']:>10.2f} "
                      f"{d['importe_real']:>10.2f} {d['diferencia']:>12.2f}  [{d['notas']}]")
            print(f"  {'':>34} {'TOTAL':>10} {total['cobrado_total']:>10.2f} "
                  f"{total['real_total']:>10.2f} {total['diferencia_total']:>12.2f}")
            print()
    else:
        print(f"ERROR: {resultado['error']}")

    if resultado.get("errores"):
        print("ERRORES:", resultado["errores"])
    if resultado.get("avisos"):
        print("AVISOS:", resultado["avisos"])

    con.close()
