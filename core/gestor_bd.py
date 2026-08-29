"""
gestor_bd.py
============
Crea y gestiona la base de datos SQLite de un despacho.
Una BD por despacho — nunca hay datos de comunidades distintas mezclados.

USO:
    python gestor_bd.py                        # crea/verifica la BD en la ruta por defecto
    python gestor_bd.py --ruta C:/MiDespacho   # crea la BD en otra carpeta

TABLAS:
    comunidades         Cada comunidad gestionada por el despacho
    propietarios        Vecinos de cada comunidad (con coeficiente y estado)
    facturas            Facturas de suministros (Gas, Luz, Agua, Mantenimiento)
    lecturas_vecino     Lecturas acumuladas de contador por vecino y fecha
    periodos            Ejercicios de regularización (ej: 2023-2024)
    gastos_extra        Reparaciones y otros gastos con amortización
    repartos            Resultado del cálculo de cuota por vecino por periodo
    config_suministro   Método de reparto y precios por suministro y comunidad
"""

import sqlite3
import os
import argparse
from datetime import datetime

import db_migrations

RUTA_BD_DEFAULT = os.path.join(
    os.path.expanduser("~"),
    "GestionFincas", "data", "gestion.db"
)


# ---------------------------------------------------------------------------
# CREACIÓN DE TABLAS
# ---------------------------------------------------------------------------

TABLAS = [

    # ------------------------------------------------------------------
    # 1. COMUNIDADES
    # Una fila por comunidad gestionada por el despacho.
    # cif: NIF/CIF de la comunidad — aparece en TODAS las facturas.
    #      Es el identificador automático más fiable para vincular PDFs.
    # ------------------------------------------------------------------
    """
    CREATE TABLE IF NOT EXISTS comunidades (
        id_comunidad    INTEGER PRIMARY KEY AUTOINCREMENT,
        codigo          TEXT    NOT NULL UNIQUE,   -- ej: '644'
        nombre          TEXT    NOT NULL,           -- nombre completo
        cif             TEXT    UNIQUE,             -- ej: 'H99258139' — usado para autodetectar comunidad en PDFs
        direccion       TEXT,
        num_viviendas   INTEGER DEFAULT 0,
        activa          INTEGER DEFAULT 1,          -- 1=activa, 0=baja
        fecha_alta      TEXT    DEFAULT (date('now')),
        notas           TEXT
    )
    """,

    # ------------------------------------------------------------------
    # 2. PROPIETARIOS
    # Cada vecino de cada comunidad.
    # estado_contador: 'ok' | 'averiado' | 'sin_contador' | 'vacio'
    # metodo_lectura_acs / calef: 'real' | 'estimado' | 'coeficiente' | 'fijo'
    # ------------------------------------------------------------------
    """
    CREATE TABLE IF NOT EXISTS propietarios (
        id_propietario      INTEGER PRIMARY KEY AUTOINCREMENT,
        id_comunidad        INTEGER NOT NULL REFERENCES comunidades(id_comunidad),
        codigo_vivienda     TEXT    NOT NULL,   -- ej: 'SS22 BAJO IZDA'
        nombre_propietario  TEXT    NOT NULL,
        coeficiente         REAL    NOT NULL DEFAULT 0.0,   -- enteros de participación
        activo              INTEGER DEFAULT 1,
        estado_contador_acs TEXT    DEFAULT 'ok',
        estado_contador_cal TEXT    DEFAULT 'ok',
        metodo_lectura_acs  TEXT    DEFAULT 'real',
        metodo_lectura_cal  TEXT    DEFAULT 'real',
        notas               TEXT,
        UNIQUE (id_comunidad, codigo_vivienda)
    )
    """,

    # ------------------------------------------------------------------
    # 3. PERIODOS
    # Cada ejercicio de regularización.
    # fecha_inicio / fecha_fin: última lectura anterior / última lectura cierre.
    # estado: 'abierto' | 'cerrado' | 'regularizado'
    # ------------------------------------------------------------------
    """
    CREATE TABLE IF NOT EXISTS periodos (
        id_periodo      INTEGER PRIMARY KEY AUTOINCREMENT,
        id_comunidad    INTEGER NOT NULL REFERENCES comunidades(id_comunidad),
        nombre          TEXT    NOT NULL,   -- ej: '2023-2024'
        fecha_inicio    TEXT    NOT NULL,   -- fecha lectura inicial del periodo
        fecha_fin       TEXT,               -- fecha lectura final (NULL si abierto)
        estado          TEXT    DEFAULT 'abierto',
        notas           TEXT,
        UNIQUE (id_comunidad, nombre)
    )
    """,

    # ------------------------------------------------------------------
    # 4. FACTURAS
    # Una fila por factura recibida (gas, luz, agua, mantenimiento).
    # num_factura + cups_o_referencia = clave anti-duplicados.
    # tipo_suministro: 'GAS' | 'ELECTRICIDAD' | 'AGUA' | 'MANTENIMIENTO'
    # ------------------------------------------------------------------
    """
    CREATE TABLE IF NOT EXISTS facturas (
        id_factura          INTEGER PRIMARY KEY AUTOINCREMENT,
        id_comunidad        INTEGER NOT NULL REFERENCES comunidades(id_comunidad),
        id_periodo          INTEGER REFERENCES periodos(id_periodo),
        tipo_suministro     TEXT    NOT NULL,
        proveedor           TEXT,
        cups_o_referencia   TEXT,
        num_factura         TEXT,
        fecha_factura       TEXT,
        fecha_inicio        TEXT,
        fecha_fin           TEXT,
        dias_facturados     INTEGER,
        consumo_total       REAL    DEFAULT 0.0,   -- kWh, m3, etc.
        unidad_consumo      TEXT,                  -- 'kWh' | 'm3'
        termino_fijo        REAL    DEFAULT 0.0,
        termino_variable    REAL    DEFAULT 0.0,
        impuestos           REAL    DEFAULT 0.0,
        iva                 REAL    DEFAULT 0.0,
        importe_total       REAL    NOT NULL,
        archivo_origen      TEXT,   -- nombre del PDF del que se extrajo
        fecha_ingesta       TEXT    DEFAULT (datetime('now')),
        notas               TEXT,
        -- Anti-duplicados: misma factura no puede entrar dos veces
        UNIQUE (id_comunidad, num_factura, cups_o_referencia)
    )
    """,

    # ------------------------------------------------------------------
    # 5. LECTURAS_VECINO
    # Lecturas ACUMULADAS del contador de cada vecino en cada fecha.
    # valor es el odómetro (va siempre subiendo), no el consumo del periodo.
    # estado: 'real' | 'estimado' | 'sin_lectura' | 'contador_averiado'
    # tipo: 'ACS' | 'CALEFACCION'
    # ------------------------------------------------------------------
    """
    CREATE TABLE IF NOT EXISTS lecturas_vecino (
        id_lectura          INTEGER PRIMARY KEY AUTOINCREMENT,
        id_propietario      INTEGER NOT NULL REFERENCES propietarios(id_propietario),
        id_periodo          INTEGER REFERENCES periodos(id_periodo),
        tipo                TEXT    NOT NULL,   -- 'ACS' | 'CALEFACCION'
        fecha_lectura       TEXT    NOT NULL,
        valor_acumulado     REAL    NOT NULL,   -- valor del contador en esa fecha
        estado              TEXT    DEFAULT 'real',
        metodo_estimacion   TEXT,   -- si estado='estimado': 'promedio' | 'calculo_inverso' | 'manual'
        fuente              TEXT,   -- 'metrigest' | 'manual' | 'importado'
        notas               TEXT,
        UNIQUE (id_propietario, tipo, fecha_lectura)
    )
    """,

    # ------------------------------------------------------------------
    # 6. GASTOS_EXTRA
    # Reparaciones y otros gastos con amortización (máx. 5 años).
    # El sistema calcula el importe_por_ejercicio = importe_total / años_amortizacion.
    # ------------------------------------------------------------------
    """
    CREATE TABLE IF NOT EXISTS gastos_extra (
        id_gasto            INTEGER PRIMARY KEY AUTOINCREMENT,
        id_comunidad        INTEGER NOT NULL REFERENCES comunidades(id_comunidad),
        descripcion         TEXT    NOT NULL,
        fecha               TEXT    NOT NULL,
        importe_total       REAL    NOT NULL,
        años_amortizacion   INTEGER DEFAULT 1,   -- 1 = cargo directo, 2-5 = amortizado
        tipo_gasto          TEXT    DEFAULT 'REPARACION',  -- 'REPARACION' | 'MANTENIMIENTO' | 'OTRO'
        servicio_afectado   TEXT,   -- 'ACS' | 'CALEFACCION' | 'AMBOS' | NULL=general
        activo              INTEGER DEFAULT 1,
        notas               TEXT
    )
    """,

    # ------------------------------------------------------------------
    # 7. CONFIG_SUMINISTRO
    # Método de reparto y precios facturados por suministro y periodo.
    # Esto es lo que se cobra a los vecinos durante el ejercicio.
    # metodo_reparto: 'coeficiente' | 'partes_iguales' | 'contador' | 'cuota_fija'
    # ------------------------------------------------------------------
    """
    CREATE TABLE IF NOT EXISTS config_suministro (
        id_config           INTEGER PRIMARY KEY AUTOINCREMENT,
        id_comunidad        INTEGER NOT NULL REFERENCES comunidades(id_comunidad),
        id_periodo          INTEGER NOT NULL REFERENCES periodos(id_periodo),
        tipo_suministro     TEXT    NOT NULL,   -- 'ACS' | 'CALEFACCION' | 'AGUA' | 'LUZ' | 'GAS'
        metodo_reparto      TEXT    NOT NULL DEFAULT 'coeficiente',
        precio_variable_facturado   REAL DEFAULT 0.0,  -- €/m3 o €/kWh cobrado al vecino
        precio_fijo_facturado       REAL DEFAULT 0.0,  -- €/mes cobrado al vecino
        precio_variable_real        REAL DEFAULT 0.0,  -- €/m3 o €/kWh real (post-regularización)
        precio_fijo_real            REAL DEFAULT 0.0,  -- €/mes real (post-regularización)
        pct_acs                     REAL DEFAULT 0.5,  -- % del gasto de combustible que va a ACS
        pct_calefaccion             REAL DEFAULT 0.5,  -- % del gasto de combustible que va a Calef
        UNIQUE (id_comunidad, id_periodo, tipo_suministro)
    )
    """,

    # ------------------------------------------------------------------
    # 8. REPARTOS
    # Resultado final del cálculo: cuánto debe pagar o le devolvemos
    # a cada vecino en cada periodo y suministro.
    # estado: 'calculado' | 'revisado' | 'carta_generada' | 'cobrado'
    # ------------------------------------------------------------------
    """
    CREATE TABLE IF NOT EXISTS repartos (
        id_reparto          INTEGER PRIMARY KEY AUTOINCREMENT,
        id_propietario      INTEGER NOT NULL REFERENCES propietarios(id_propietario),
        id_periodo          INTEGER NOT NULL REFERENCES periodos(id_periodo),
        tipo_suministro     TEXT    NOT NULL,
        lectura_inicial     REAL    DEFAULT 0.0,
        lectura_final       REAL    DEFAULT 0.0,
        consumo_real        REAL    DEFAULT 0.0,
        importe_cobrado     REAL    DEFAULT 0.0,   -- lo que pagó durante el ejercicio
        importe_real        REAL    DEFAULT 0.0,   -- lo que debería haber pagado
        diferencia          REAL    DEFAULT 0.0,   -- real - cobrado (neg = devolver, pos = cobrar)
        estado              TEXT    DEFAULT 'calculado',
        fecha_calculo       TEXT    DEFAULT (datetime('now')),
        notas               TEXT,
        UNIQUE (id_propietario, id_periodo, tipo_suministro)
    )
    """,

    # ------------------------------------------------------------------
    # 9. ARCHIVOS_PROCESADOS
    # Registro de PDFs ya ingeridos para no procesarlos dos veces.
    # ------------------------------------------------------------------
    """
    CREATE TABLE IF NOT EXISTS archivos_procesados (
        id_archivo      INTEGER PRIMARY KEY AUTOINCREMENT,
        nombre_archivo  TEXT    NOT NULL UNIQUE,
        hash_md5        TEXT,
        fecha_proceso   TEXT    DEFAULT (datetime('now')),
        id_factura      INTEGER REFERENCES facturas(id_factura),
        resultado       TEXT    DEFAULT 'ok',   -- 'ok' | 'error' | 'omitido'
        notas           TEXT
    )
    """,
]


INDICES = [
    "CREATE INDEX IF NOT EXISTS idx_facturas_comunidad  ON facturas(id_comunidad)",
    "CREATE INDEX IF NOT EXISTS idx_facturas_periodo    ON facturas(id_periodo)",
    "CREATE INDEX IF NOT EXISTS idx_lecturas_propietario ON lecturas_vecino(id_propietario)",
    "CREATE INDEX IF NOT EXISTS idx_lecturas_tipo_fecha  ON lecturas_vecino(tipo, fecha_lectura)",
    "CREATE INDEX IF NOT EXISTS idx_repartos_periodo    ON repartos(id_periodo)",
    "CREATE INDEX IF NOT EXISTS idx_propietarios_comunidad ON propietarios(id_comunidad)",
]


# ---------------------------------------------------------------------------
# FUNCIONES PÚBLICAS — usadas por los demás módulos
# ---------------------------------------------------------------------------

def crear_bd(ruta_bd: str) -> None:
    """Crea todas las tablas e índices. Es seguro llamarlo varias veces (IF NOT EXISTS)."""
    os.makedirs(os.path.dirname(ruta_bd), exist_ok=True)
    with sqlite3.connect(ruta_bd) as con:
        con.execute("PRAGMA foreign_keys = ON")
        con.execute("PRAGMA journal_mode = DELETE")  # compatible con sistemas de ficheros montados
        for sql in TABLAS:
            con.execute(sql)
        for sql in INDICES:
            con.execute(sql)
        aplicar_migraciones(con)
        con.commit()
    print(f"✅ BD lista en: {ruta_bd}")


def aplicar_migraciones(con: sqlite3.Connection) -> int:
    """Actualiza una conexión existente al esquema más reciente."""
    return db_migrations.migrate(con)


def conectar(ruta_bd: str) -> sqlite3.Connection:
    """Devuelve una conexión con foreign keys activadas. El caller hace .close()."""
    if not os.path.exists(ruta_bd):
        raise FileNotFoundError(
            f"No existe la BD en '{ruta_bd}'. Ejecuta primero: python gestor_bd.py"
        )
    con = sqlite3.connect(ruta_bd)
    con.execute("PRAGMA foreign_keys = ON")
    con.row_factory = sqlite3.Row   # acceso por nombre de columna: fila['nombre']
    return con


def obtener_o_crear_comunidad(con: sqlite3.Connection, codigo: str, nombre: str,
                               cif: str = None) -> int:
    """Devuelve id_comunidad. La crea si no existe."""
    fila = con.execute(
        "SELECT id_comunidad FROM comunidades WHERE codigo = ?", (codigo,)
    ).fetchone()
    if fila:
        # Actualizar CIF si se proporciona y no estaba
        if cif:
            con.execute(
                "UPDATE comunidades SET cif = ? WHERE codigo = ? AND cif IS NULL",
                (cif, codigo)
            )
            con.commit()
        return fila["id_comunidad"]
    cur = con.execute(
        "INSERT INTO comunidades (codigo, nombre, cif) VALUES (?, ?, ?)",
        (codigo, nombre, cif)
    )
    con.commit()
    return cur.lastrowid


def buscar_comunidad_por_cif(con: sqlite3.Connection, cif: str) -> dict | None:
    """
    Busca una comunidad por su CIF/NIF.
    Devuelve {'id_comunidad': X, 'codigo': '644', 'nombre': '...'} o None.
    Usado para autodetectar la comunidad a partir del CIF que aparece en el PDF.
    """
    fila = con.execute(
        "SELECT id_comunidad, codigo, nombre FROM comunidades WHERE cif = ?",
        (cif.upper(),)
    ).fetchone()
    if fila:
        return dict(fila)
    return None


def obtener_o_crear_periodo(con: sqlite3.Connection, id_comunidad: int,
                             nombre: str, fecha_inicio: str) -> int:
    """Devuelve id_periodo. Lo crea si no existe."""
    fila = con.execute(
        "SELECT id_periodo FROM periodos WHERE id_comunidad=? AND nombre=?",
        (id_comunidad, nombre)
    ).fetchone()
    if fila:
        return fila["id_periodo"]
    cur = con.execute(
        "INSERT INTO periodos (id_comunidad, nombre, fecha_inicio) VALUES (?,?,?)",
        (id_comunidad, nombre, fecha_inicio)
    )
    con.commit()
    return cur.lastrowid


def insertar_factura(con: sqlite3.Connection, datos: dict) -> int | None:
    """
    Inserta una factura. Devuelve id_factura o None si ya existía (duplicado).
    'datos' debe tener al menos: id_comunidad, tipo_suministro, importe_total.
    Anti-duplicados por (id_comunidad, num_factura, cups_o_referencia).
    """
    try:
        cur = con.execute("""
            INSERT INTO facturas (
                id_comunidad, id_periodo, tipo_suministro, proveedor,
                cups_o_referencia, num_factura, fecha_factura,
                fecha_inicio, fecha_fin, dias_facturados,
                consumo_total, unidad_consumo,
                termino_fijo, termino_variable, impuestos, iva,
                importe_total, archivo_origen, notas
            ) VALUES (
                :id_comunidad, :id_periodo, :tipo_suministro, :proveedor,
                :cups_o_referencia, :num_factura, :fecha_factura,
                :fecha_inicio, :fecha_fin, :dias_facturados,
                :consumo_total, :unidad_consumo,
                :termino_fijo, :termino_variable, :impuestos, :iva,
                :importe_total, :archivo_origen, :notas
            )
        """, {
            "id_comunidad":      datos.get("id_comunidad"),
            "id_periodo":        datos.get("id_periodo"),
            "tipo_suministro":   datos.get("tipo_suministro"),
            "proveedor":         datos.get("proveedor"),
            "cups_o_referencia": datos.get("cups_o_referencia"),
            "num_factura":       datos.get("num_factura"),
            "fecha_factura":     datos.get("fecha_factura"),
            "fecha_inicio":      datos.get("fecha_inicio"),
            "fecha_fin":         datos.get("fecha_fin"),
            "dias_facturados":   datos.get("dias_facturados"),
            "consumo_total":     datos.get("consumo_total", 0.0),
            "unidad_consumo":    datos.get("unidad_consumo"),
            "termino_fijo":      datos.get("termino_fijo", 0.0),
            "termino_variable":  datos.get("termino_variable", 0.0),
            "impuestos":         datos.get("impuestos", 0.0),
            "iva":               datos.get("iva", 0.0),
            "importe_total":     datos.get("importe_total"),
            "archivo_origen":    datos.get("archivo_origen"),
            "notas":             datos.get("notas"),
        })
        con.commit()
        return cur.lastrowid
    except sqlite3.IntegrityError:
        return None  # duplicado — ya estaba en la BD


def insertar_lectura_vecino(con: sqlite3.Connection, datos: dict) -> int | None:
    """
    Inserta una lectura acumulada de contador.
    Anti-duplicados por (id_propietario, tipo, fecha_lectura).
    """
    try:
        cur = con.execute("""
            INSERT INTO lecturas_vecino (
                id_propietario, id_periodo, tipo,
                fecha_lectura, valor_acumulado,
                estado, metodo_estimacion, fuente, notas
            ) VALUES (
                :id_propietario, :id_periodo, :tipo,
                :fecha_lectura, :valor_acumulado,
                :estado, :metodo_estimacion, :fuente, :notas
            )
        """, {
            "id_propietario":    datos.get("id_propietario"),
            "id_periodo":        datos.get("id_periodo"),
            "tipo":              datos.get("tipo"),
            "fecha_lectura":     datos.get("fecha_lectura"),
            "valor_acumulado":   datos.get("valor_acumulado"),
            "estado":            datos.get("estado", "real"),
            "metodo_estimacion": datos.get("metodo_estimacion"),
            "fuente":            datos.get("fuente", "manual"),
            "notas":             datos.get("notas"),
        })
        con.commit()
        return cur.lastrowid
    except sqlite3.IntegrityError:
        return None


def marcar_archivo_procesado(con: sqlite3.Connection, nombre: str,
                              hash_md5: str = None, id_factura: int = None,
                              resultado: str = "ok", notas: str = None) -> None:
    """Registra un PDF como procesado para no volver a ingerirlo."""
    try:
        con.execute("""
            INSERT INTO archivos_procesados (nombre_archivo, hash_md5, id_factura, resultado, notas)
            VALUES (?, ?, ?, ?, ?)
        """, (nombre, hash_md5, id_factura, resultado, notas))
        con.commit()
    except sqlite3.IntegrityError:
        pass  # ya estaba registrado


def archivo_ya_procesado(con: sqlite3.Connection, nombre: str) -> bool:
    """Devuelve True si el PDF ya fue procesado anteriormente."""
    fila = con.execute(
        "SELECT 1 FROM archivos_procesados WHERE nombre_archivo = ? AND resultado = 'ok'",
        (nombre,)
    ).fetchone()
    return fila is not None


def resumen_bd(ruta_bd: str) -> None:
    """Imprime un resumen de lo que hay en la BD."""
    with conectar(ruta_bd) as con:
        tablas = [
            "comunidades", "propietarios", "periodos", "facturas",
            "lecturas_vecino", "gastos_extra", "config_suministro",
            "repartos", "archivos_procesados"
        ]
        print(f"\n{'─'*45}")
        print(f"  Resumen de la BD: {ruta_bd}")
        print(f"{'─'*45}")
        for t in tablas:
            n = con.execute(f"SELECT COUNT(*) FROM {t}").fetchone()[0]
            print(f"  {t:<30} {n:>6} filas")
        print(f"{'─'*45}\n")


# ---------------------------------------------------------------------------
# EJECUCIÓN DIRECTA — crea la BD
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Crea la BD de GestionFincas")
    parser.add_argument(
        "--ruta", default=os.path.dirname(RUTA_BD_DEFAULT),
        help=f"Carpeta donde crear la BD (por defecto: {os.path.dirname(RUTA_BD_DEFAULT)})"
    )
    args = parser.parse_args()

    ruta_bd = os.path.join(args.ruta, "gestion.db")
    crear_bd(ruta_bd)
    resumen_bd(ruta_bd)
    print("Listo. Las demás funciones de este módulo se importan desde otros scripts.")
