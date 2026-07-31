"""
pipeline.py
===========
"PROCESAR TODO" — flujo completo desatendido, sin GUI.

Hace en un solo paso todo el trabajo:
  1. Ingesta la carpeta entrada/ (y opcionalmente otra carpeta extra):
       - Facturas PDF        → identifica proveedor, comunidad (nombre/CIF/CUPS)
                               y periodo por fechas → inserta en BD
       - Resúmenes Metrigest → lecturas de vecinos → BD
       - Excel CAL-ACS       → vecinos + lecturas históricas → BD
       - Lo que no reconoce  → cuarentena/ con el motivo en cuarentena/motivos.log
  2. Recalcula el reparto de cada (comunidad, periodo) afectado
  3. Regenera el Excel Maestro completo de cada comunidad afectada
  4. Genera las cartas de los periodos afectados

USO:
    python pipeline.py                          # procesa todo lo pendiente
    python pipeline.py --comunidad 644          # limita a una comunidad
    python pipeline.py --carpeta "C:\\ruta"     # ingesta también esa carpeta
    python pipeline.py --sin-cartas             # no genera cartas
    python pipeline.py --sin-reparto            # solo registra, no calcula

Pensado para lanzarse desde la GUI, el Programador de tareas de Windows
o Power Automate (después del RPA de correo).
"""

import os
import re
import sys
import json
import time
import shutil
import sqlite3
import argparse
from datetime import datetime, date
from pathlib import Path


# ---------------------------------------------------------------------------
# MOVIMIENTO SEGURO DE ARCHIVOS
# ---------------------------------------------------------------------------
def mover_seguro(origen, destino, log=None, intentos: int = 3,
                 espera: float = 1.2) -> bool:
    """
    Mueve un archivo con reintentos. En Windows, si el archivo está abierto
    en otro programa (p. ej. Excel), shutil.move lanza PermissionError
    [WinError 32]. En ese caso:
      1. Reintenta hasta `intentos` veces (por si era un bloqueo momentáneo)
      2. Si sigue bloqueado, lo COPIA al destino y avisa con un mensaje claro
         (el original queda en su sitio; el proceso NO se rompe)
    Devuelve True si el original quedó movido, False si solo se pudo copiar.
    """
    origen, destino = str(origen), str(destino)
    nombre = Path(origen).name
    for _ in range(intentos):
        try:
            shutil.move(origen, destino)
            return True
        except (PermissionError, OSError):
            time.sleep(espera)
    # Plan B: copiar sin borrar el original
    try:
        shutil.copy2(origen, destino)
        if log:
            log(f"  ⚠️  «{nombre}» está abierto en otro programa (¿Excel?).", "aviso")
            log(f"     Se ha copiado a procesados/, pero el original sigue en entrada/.", "aviso")
            log(f"     👉 Cierra el archivo en Excel y bórralo de entrada/ (o vuelve a Procesar).", "aviso")
    except Exception:
        if log:
            log(f"  ❌ No se pudo mover «{nombre}»: está abierto en otro programa.", "error")
            log(f"     👉 Cierra el archivo (Excel/visor de PDF) y pulsa de nuevo el botón.", "error")
    return False

sys.path.insert(0, str(Path(__file__).parent))
import gestor_bd
import lector_pdf
import excel_generator

try:
    import importar_lecturas_metrigest
except ImportError:
    importar_lecturas_metrigest = None
try:
    import importar_comunidad
except ImportError:
    importar_comunidad = None
try:
    import motor_reparto
except ImportError:
    motor_reparto = None
try:
    import carta_writer
except ImportError:
    carta_writer = None

BASE_DIR = Path(__file__).parent.parent


# ---------------------------------------------------------------------------
# RUTAS — por defecto relativas al proyecto; se pueden sobrescribir en
# config/rutas.json para instalar el sistema en otro despacho sin tocar código.
# ---------------------------------------------------------------------------
def cargar_rutas() -> dict:
    rutas = {
        "bd":          BASE_DIR / "data" / "gestion.db",
        "entrada":     BASE_DIR / "entrada",
        "procesados":  BASE_DIR / "procesados",
        "cuarentena":  BASE_DIR / "cuarentena",
        "excels":      BASE_DIR / "Excels_Maestros",
        "plantilla_cartas": BASE_DIR / "plantillas" / "Plantilla_Cartas.docx",
        "cartas":      BASE_DIR / "salidas" / "cartas",
        "logs":        BASE_DIR / "salidas" / "logs",
        "proveedores": BASE_DIR / "config" / "proveedores.json",
    }
    ruta_config = BASE_DIR / "config" / "rutas.json"
    if ruta_config.exists():
        try:
            overrides = json.loads(ruta_config.read_text(encoding="utf-8"))
            for k, v in overrides.items():
                if k in rutas and v:
                    rutas[k] = Path(v)
        except (json.JSONDecodeError, OSError):
            pass
    for k in ("entrada", "procesados", "cuarentena", "excels", "cartas", "logs"):
        rutas[k].mkdir(parents=True, exist_ok=True)
    return rutas


# ---------------------------------------------------------------------------
# LOG — a consola (o callback de la GUI) y a archivo
# ---------------------------------------------------------------------------
class Registro:
    def __init__(self, carpeta_logs: Path, callback=None):
        self.lineas = []
        self.callback = callback
        self.ruta = carpeta_logs / f"pipeline_{datetime.now():%Y%m%d_%H%M%S}.txt"

    def __call__(self, mensaje: str, tipo: str = "neutro"):
        self.lineas.append(mensaje)
        if self.callback:
            self.callback(mensaje, tipo)
        else:
            print(mensaje)

    def guardar(self):
        try:
            self.ruta.write_text("\n".join(self.lineas), encoding="utf-8")
        except OSError:
            pass


# ---------------------------------------------------------------------------
# DETECCIÓN DE COMUNIDAD Y PERIODO
# ---------------------------------------------------------------------------
RE_CODIGO = re.compile(r'(?<![A-Za-z])(\d{3,6})(?![A-Za-z0-9])')


def detectar_codigo_en_nombre(nombre_archivo: str) -> str | None:
    m = RE_CODIGO.search(nombre_archivo)
    return m.group(1) if m else None


def _parsear_fecha(valor) -> date | None:
    """Acepta ISO (2026-06-25), DD/MM/YYYY, DD/MM/YY, DD.MM.YYYY, DD-MM-YYYY."""
    if not valor:
        return None
    s = str(valor).strip()[:10]
    try:
        return date.fromisoformat(s)
    except ValueError:
        pass
    m = re.match(r'(\d{1,2})[/.\-](\d{1,2})[/.\-](\d{2,4})', s)
    if not m:
        return None
    d, mth, a = int(m.group(1)), int(m.group(2)), int(m.group(3))
    if a < 100:
        a += 2000
    try:
        return date(a, mth, d)
    except ValueError:
        return None


def detectar_o_crear_periodo(con, datos: dict, id_comunidad: int, log) -> int | None:
    """
    Busca el periodo de la comunidad que solapa con las fechas de la factura.
    Si no existe, lo crea con año fiscal sep→ago. (Versión headless de la
    lógica que estaba en app.py.)
    """
    fecha = (_parsear_fecha(datos.get("fecha_inicio")) or
             _parsear_fecha(datos.get("fecha_factura")) or
             _parsear_fecha(datos.get("fecha_fin")))
    if not fecha:
        return None

    rows = con.execute(
        "SELECT id_periodo, fecha_inicio, fecha_fin FROM periodos "
        "WHERE id_comunidad=? AND fecha_inicio IS NOT NULL AND fecha_fin IS NOT NULL",
        (id_comunidad,)
    ).fetchall()
    for r in rows:
        fi = _parsear_fecha(r["fecha_inicio"])
        ff = _parsear_fecha(r["fecha_fin"])
        if fi and ff and fi <= fecha <= ff:
            return r["id_periodo"]

    # Crear periodo automático (año fiscal sep→ago)
    if fecha.month >= 9:
        anio_ini = fecha.year
    else:
        anio_ini = fecha.year - 1
    f_ini, f_fin = date(anio_ini, 9, 1), date(anio_ini + 1, 8, 31)
    nombre_auto = f"{anio_ini}-{anio_ini + 1}"
    try:
        id_per = gestor_bd.obtener_o_crear_periodo(
            con, id_comunidad, nombre_auto, str(f_ini))
        con.execute(
            "UPDATE periodos SET fecha_fin=?, estado='abierto' WHERE id_periodo=?",
            (str(f_fin), id_per))
        con.commit()
        log(f"     📅 Periodo '{nombre_auto}' creado automáticamente")
        return id_per
    except Exception as e:
        log(f"     ⚠️ No se pudo crear periodo automático: {e}", "error")
        return None


# ---------------------------------------------------------------------------
# CUARENTENA
# ---------------------------------------------------------------------------
def a_cuarentena(ruta_archivo: Path, motivo: str, rutas: dict, log):
    destino = rutas["cuarentena"] / ruta_archivo.name
    n = 1
    while destino.exists():
        destino = rutas["cuarentena"] / f"{ruta_archivo.stem}_{n}{ruta_archivo.suffix}"
        n += 1
    if not mover_seguro(ruta_archivo, destino, log):
        return
    try:
        with open(rutas["cuarentena"] / "motivos.log", "a", encoding="utf-8") as f:
            f.write(f"{datetime.now():%Y-%m-%d %H:%M} | {ruta_archivo.name} | {motivo}\n")
    except OSError:
        pass
    log(f"  🔶 {ruta_archivo.name} → cuarentena ({motivo})", "aviso")


# ---------------------------------------------------------------------------
# INGESTA DE ARCHIVOS
# ---------------------------------------------------------------------------
def _ingerir_pdf(ruta: Path, con, rutas: dict, log, tocados: dict,
                 contadores: dict):
    nombre = ruta.name
    if gestor_bd.archivo_ya_procesado(con, nombre):
        log(f"  ⏭  {nombre} — ya procesado", "neutro")
        mover_seguro(ruta, rutas["procesados"] / _nombre_libre(rutas["procesados"], nombre), log)
        contadores["duplicados"] += 1
        return

    codigo = detectar_codigo_en_nombre(nombre)
    resultado = lector_pdf.procesar_archivo(
        str(ruta), codigo,
        con_bd=con,
        ruta_proveedores=str(rutas["proveedores"]),
    )

    # lector_pdf puede renombrar el archivo ("644_" + nombre): seguir la ruta real
    nombre_nuevo = resultado.get("nombre_archivo") or nombre
    if nombre_nuevo != nombre and (ruta.parent / nombre_nuevo).exists():
        ruta = ruta.parent / nombre_nuevo
        nombre = nombre_nuevo

    if resultado["ok"] and resultado["tipo"] == "FACTURA":
        datos = resultado["datos"]
        codigo_final = resultado.get("codigo_comunidad") or codigo
        row = con.execute(
            "SELECT id_comunidad FROM comunidades WHERE codigo=?",
            (str(codigo_final),)).fetchone() if codigo_final else None
        if not row:
            a_cuarentena(ruta, "comunidad no identificada (ni nombre ni CIF)", rutas, log)
            contadores["errores"] += 1
            return
        id_com = row["id_comunidad"]
        datos["id_comunidad"] = id_com
        datos["id_periodo"] = detectar_o_crear_periodo(con, datos, id_com, log)

        id_fac = gestor_bd.insertar_factura(con, datos)
        if id_fac:
            gestor_bd.marcar_archivo_procesado(con, nombre, resultado["hash_md5"], id_fac)
            mover_seguro(ruta, rutas["procesados"] / _nombre_libre(rutas["procesados"], nombre), log)
            tipo = datos.get("tipo_suministro", "?")
            imp  = datos.get("importe_total", 0) or 0
            log(f"  ✅ {nombre[:45]:<45} {tipo:<14} {imp:>9.2f} €", "ok")
            contadores["facturas"] += 1
            tocados.setdefault(id_com, set())
            if datos.get("id_periodo"):
                tocados[id_com].add(datos["id_periodo"])
        else:
            log(f"  ⏭  {nombre} — duplicado en BD", "neutro")
            mover_seguro(ruta, rutas["procesados"] / _nombre_libre(rutas["procesados"], nombre), log)
            contadores["duplicados"] += 1

    elif resultado["ok"] and resultado["tipo"] == "LECTURA_METRIGEST":
        if importar_lecturas_metrigest is None:
            a_cuarentena(ruta, "módulo importar_lecturas_metrigest no disponible", rutas, log)
            contadores["errores"] += 1
            return
        res = importar_lecturas_metrigest.importar_lecturas_pdf(
            ruta_pdf=str(ruta),
            ruta_bd=str(rutas["bd"]),
            mover_procesado=True,
            ruta_proveedores=str(rutas["proveedores"]),
            verbose=False,
        )
        if res.get("ok"):
            ins = res.get("lecturas_insertadas", 0)
            log(f"  📊 {nombre[:45]:<45} LECTURAS — {ins} insertadas", "ok")
            contadores["lecturas"] += 1
            id_com = res.get("id_comunidad")
            id_per = res.get("id_periodo")
            if id_com and id_per:
                tocados.setdefault(id_com, set()).add(id_per)
            elif id_com:
                tocados.setdefault(id_com, set())
        else:
            a_cuarentena(ruta, f"lecturas Metrigest: {res.get('error', '?')}", rutas, log)
            contadores["errores"] += 1

    elif resultado["ok"]:
        # Otros documentos reconocidos (ej: JUSTIFICANTE_PAGO): archivar sin registrar
        tipo_doc = resultado.get("tipo", "DOCUMENTO")
        gestor_bd.marcar_archivo_procesado(con, nombre, resultado["hash_md5"],
                                           resultado="ok", notas=tipo_doc)
        mover_seguro(ruta, rutas["procesados"] / _nombre_libre(rutas["procesados"], nombre), log)
        log(f"  📎 {nombre[:45]:<45} {tipo_doc} — archivado", "info")
        contadores["otros"] = contadores.get("otros", 0) + 1
    else:
        motivo = resultado.get("motivo", "?")
        detalle = (resultado.get("detalle") or "")[:80]
        a_cuarentena(ruta, f"{motivo} {detalle}".strip(), rutas, log)
        contadores["errores"] += 1


def _ingerir_excel_lecturas(ruta: Path, con, rutas: dict, log, tocados: dict,
                            contadores: dict):
    """Excel CAL-ACS exportado de Gesfincas: vecinos + lecturas históricas."""
    nombre = ruta.name
    if importar_comunidad is None:
        a_cuarentena(ruta, "módulo importar_comunidad no disponible", rutas, log)
        contadores["errores"] += 1
        return
    codigo = detectar_codigo_en_nombre(nombre)
    if not codigo:
        a_cuarentena(ruta, "Excel de lecturas sin código de comunidad en el nombre", rutas, log)
        contadores["errores"] += 1
        return
    row = con.execute(
        "SELECT nombre, cif FROM comunidades WHERE codigo=?", (codigo,)).fetchone()
    if not row:
        a_cuarentena(ruta, f"comunidad {codigo} no está en la BD (sincroniza el listado)", rutas, log)
        contadores["errores"] += 1
        return

    log(f"  📥 {nombre} — importando vecinos y lecturas (comunidad {codigo})…", "info")
    try:
        res = importar_comunidad.importar_comunidad(
            codigo=codigo,
            nombre=row["nombre"] or "",
            cif=row["cif"] or "",
            ruta_excel_lecturas=str(ruta),
            ruta_bd=str(rutas["bd"]),
        )
    except Exception as e:
        a_cuarentena(ruta, f"error importando CAL-ACS: {e}", rutas, log)
        contadores["errores"] += 1
        return

    if res.get("ok"):
        mover_seguro(ruta, rutas["procesados"] / _nombre_libre(rutas["procesados"], nombre), log)
        log(f"  ✅ {nombre} importado", "ok")
        contadores["lecturas"] += 1
        row_id = con.execute(
            "SELECT id_comunidad FROM comunidades WHERE codigo=?", (codigo,)).fetchone()
        if row_id:
            tocados.setdefault(row_id["id_comunidad"], set())
    else:
        a_cuarentena(ruta, f"CAL-ACS: {res.get('error', '?')}", rutas, log)
        contadores["errores"] += 1


def _nombre_libre(carpeta: Path, nombre: str) -> str:
    destino = carpeta / nombre
    if not destino.exists():
        return nombre
    stem, suf = Path(nombre).stem, Path(nombre).suffix
    n = 1
    while (carpeta / f"{stem}_{n}{suf}").exists():
        n += 1
    return f"{stem}_{n}{suf}"


def ingerir_entrada(rutas: dict, log, filtro_codigo: str = None) -> tuple[dict, dict]:
    """
    Procesa todos los archivos de entrada/.
    Returns:
        (tocados, contadores)
        tocados: {id_comunidad: {id_periodo, ...}}
    """
    contadores = {"facturas": 0, "lecturas": 0, "duplicados": 0, "errores": 0}
    tocados: dict[int, set] = {}

    archivos = sorted(p for p in rutas["entrada"].iterdir() if p.is_file())
    if filtro_codigo:
        archivos = [p for p in archivos
                    if detectar_codigo_en_nombre(p.name) in (filtro_codigo, None)]
    if not archivos:
        log("📂 Carpeta entrada/ vacía — nada que ingerir", "neutro")
        return tocados, contadores

    log(f"📂 {len(archivos)} archivo(s) en entrada/", "info")
    con = gestor_bd.conectar(str(rutas["bd"]))
    try:
        for ruta in archivos:
            ext = ruta.suffix.lower()
            if ext == ".pdf" or ext == ".zip":
                _ingerir_pdf(ruta, con, rutas, log, tocados, contadores)
            elif ext in (".xlsx", ".xls") and "cal-acs" in ruta.name.lower().replace("_", "-").replace(" ", "-"):
                _ingerir_excel_lecturas(ruta, con, rutas, log, tocados, contadores)
            elif ext in (".xlsx", ".xls"):
                # Excel sin patrón CAL-ACS: probar igualmente si tiene código
                _ingerir_excel_lecturas(ruta, con, rutas, log, tocados, contadores)
            else:
                a_cuarentena(ruta, f"extensión no soportada ({ext})", rutas, log)
                contadores["errores"] += 1
    finally:
        con.close()
    return tocados, contadores


# ---------------------------------------------------------------------------
# PROCESAR TODO
# ---------------------------------------------------------------------------
def procesar_todo(codigo_comunidad: str = None,
                  carpeta_extra: str = None,
                  hacer_reparto: bool = True,
                  hacer_cartas: bool = True,
                  log_callback=None) -> dict:
    """
    Flujo completo. Devuelve dict resumen.
    """
    rutas = cargar_rutas()
    log = Registro(rutas["logs"], callback=log_callback)
    resumen = {"ok": True, "contadores": {}, "comunidades": [], "errores": [],
               "log": str(log.ruta)}

    log(f"━━━ PROCESAR TODO — {datetime.now():%d/%m/%Y %H:%M} ━━━", "titulo")

    # 0. Copiar archivos de una carpeta extra (histórico ya descargado)
    if carpeta_extra:
        origen = Path(carpeta_extra)
        if origen.is_dir():
            copiados = 0
            for p in origen.iterdir():
                if p.is_file() and p.suffix.lower() in (".pdf", ".zip", ".xlsx", ".xls"):
                    destino = rutas["entrada"] / _nombre_libre(rutas["entrada"], p.name)
                    shutil.copy2(str(p), str(destino))
                    copiados += 1
            log(f"📥 {copiados} archivo(s) copiados de {origen}", "info")
        else:
            log(f"⚠️ Carpeta extra no encontrada: {carpeta_extra}", "aviso")

    # 1. Ingesta
    tocados, contadores = ingerir_entrada(rutas, log, filtro_codigo=codigo_comunidad)
    resumen["contadores"] = contadores

    # Si se pidió una comunidad concreta, incluirla aunque no haya archivos nuevos
    con = sqlite3.connect(str(rutas["bd"]))
    con.row_factory = sqlite3.Row
    if codigo_comunidad:
        row = con.execute("SELECT id_comunidad FROM comunidades WHERE codigo=?",
                          (str(codigo_comunidad),)).fetchone()
        if row:
            tocados.setdefault(row["id_comunidad"], set())
        else:
            log(f"❌ Comunidad {codigo_comunidad} no existe en la BD", "error")
            resumen["ok"] = False
            con.close()
            log.guardar()
            return resumen

    if not tocados:
        log("\nNada que procesar.", "neutro")
        con.close()
        log.guardar()
        return resumen

    # 2-4. Por cada comunidad afectada
    for id_com, periodos_tocados in tocados.items():
        com = con.execute(
            "SELECT codigo, nombre FROM comunidades WHERE id_comunidad=?",
            (id_com,)).fetchone()
        codigo = com["codigo"]
        log(f"\n━━━ Comunidad {codigo} — {com['nombre'] or ''} ━━━", "titulo")

        # Si no hay periodos concretos afectados, usar los abiertos
        if not periodos_tocados:
            abiertos = con.execute(
                "SELECT id_periodo FROM periodos WHERE id_comunidad=? "
                "AND COALESCE(estado,'abierto')!='cerrado'", (id_com,)).fetchall()
            periodos_tocados = {r["id_periodo"] for r in abiertos}

        # 2. Reparto por periodo
        if hacer_reparto and motor_reparto:
            for id_per in sorted(periodos_tocados):
                nom_per = con.execute(
                    "SELECT nombre FROM periodos WHERE id_periodo=?",
                    (id_per,)).fetchone()
                nom_per = nom_per["nombre"] if nom_per else str(id_per)
                con_r = gestor_bd.conectar(str(rutas["bd"]))
                try:
                    r = motor_reparto.calcular_reparto(
                        con=con_r, id_comunidad=id_com,
                        id_periodo=id_per, sobrescribir=True)
                finally:
                    con_r.close()
                if r.get("ok"):
                    log(f"  🧮 Reparto {nom_per}: {r.get('vecinos_procesados', 0)} vecinos, "
                        f"{r.get('registros_guardados', 0)} registros", "ok")
                else:
                    msg = f"Reparto {codigo}/{nom_per}: {r.get('error', '?')}"
                    log(f"  ⚠️ {msg}", "aviso")
                    resumen["errores"].append(msg)

        # 3. Regenerar Excel completo
        r_x = excel_generator.regenerar_excel_comunidad(
            codigo, ruta_bd=str(rutas["bd"]),
            ruta_excels=str(rutas["excels"]), log=log)
        if not r_x["ok"]:
            resumen["errores"].extend(r_x["errores"])
            resumen["ok"] = False

        # 4. Cartas por periodo
        if hacer_cartas and carta_writer and rutas["plantilla_cartas"].exists():
            for id_per in sorted(periodos_tocados):
                nom_per = con.execute(
                    "SELECT nombre FROM periodos WHERE id_periodo=?",
                    (id_per,)).fetchone()
                if not nom_per:
                    continue
                nom_per = nom_per["nombre"]
                carpeta_salida = rutas["cartas"] / str(codigo) / nom_per
                carpeta_salida.mkdir(parents=True, exist_ok=True)
                r_c = carta_writer.generar_todas_las_cartas(
                    ruta_bd=str(rutas["bd"]),
                    id_comunidad=id_com,
                    nombre_periodo=nom_per,
                    ruta_plantilla=str(rutas["plantilla_cartas"]),
                    carpeta_salida=str(carpeta_salida))
                if r_c.get("ok"):
                    log(f"  ✉️ {r_c.get('cartas_generadas', 0)} cartas → {carpeta_salida}", "ok")
                else:
                    msg = f"Cartas {codigo}/{nom_per}: {r_c.get('error', '?')}"
                    log(f"  ⚠️ {msg}", "aviso")
                    resumen["errores"].append(msg)

        resumen["comunidades"].append(codigo)

    con.close()

    c = contadores
    log(f"\n━━━ RESUMEN ━━━", "titulo")
    log(f"  Facturas: {c['facturas']}  |  Lecturas: {c['lecturas']}  |  "
        f"Duplicados: {c['duplicados']}  |  Errores/cuarentena: {c['errores']}")
    log(f"  Comunidades actualizadas: {', '.join(resumen['comunidades']) or 'ninguna'}")
    if resumen["errores"]:
        log(f"  ⚠️ {len(resumen['errores'])} incidencia(s) — revisa el log y cuarentena/")
    log.guardar()
    return resumen


# ---------------------------------------------------------------------------
if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Procesar todo (sin GUI)")
    parser.add_argument("--comunidad", default=None, help="Limitar a un código (ej: 644)")
    parser.add_argument("--carpeta", default=None, help="Carpeta extra a ingerir")
    parser.add_argument("--sin-cartas", action="store_true")
    parser.add_argument("--sin-reparto", action="store_true")
    args = parser.parse_args()

    r = procesar_todo(
        codigo_comunidad=args.comunidad,
        carpeta_extra=args.carpeta,
        hacer_reparto=not args.sin_reparto,
        hacer_cartas=not args.sin_cartas,
    )
    sys.exit(0 if r["ok"] else 1)
