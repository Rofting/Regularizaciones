"""
importar_lecturas_metrigest.py
==============================
Importa lecturas de contadores desde PDFs de Metrigest directamente a la BD.
Detecta automáticamente la comunidad por CIF o nombre en el PDF.
Actualiza también la pestaña de LECTURAS en el Excel Maestro.

USO:
    python importar_lecturas_metrigest.py --pdf "../entrada/644_METRIGEST_ACS.pdf"
    python importar_lecturas_metrigest.py --carpeta "../entrada/"  # procesa todos los Metrigest

RESULTADO:
    - Inserta lecturas en tabla lecturas_vecino (fecha_anterior + fecha_actual)
    - Marca el archivo como procesado en archivos_procesados
    - Mueve el archivo a ../procesados/
    - Opcional: actualiza pestaña LECTURAS en Excel Maestro
"""

import os
import sys
import re
import sqlite3
import shutil
import argparse
from pathlib import Path
from datetime import datetime

sys.path.insert(0, str(Path(__file__).parent))
from gestor_bd import conectar, crear_bd, marcar_archivo_procesado, archivo_ya_procesado
from lector_pdf import procesar_archivo

BASE_DIR        = Path(__file__).parent.parent
RUTA_BD         = BASE_DIR / "data" / "gestion.db"
RUTA_ENTRADA    = BASE_DIR / "entrada"
RUTA_PROCESADOS = BASE_DIR / "procesados"
RUTA_PROVEEDORES= BASE_DIR / "config" / "proveedores.json"


# ---------------------------------------------------------------------------
# IMPORTAR LECTURAS DE UN RESUMEN METRIGEST
# ---------------------------------------------------------------------------

def importar_lecturas_pdf(ruta_pdf: str, ruta_bd: str,
                           mover_procesado: bool = True,
                           ruta_proveedores: str = None,
                           verbose: bool = True) -> dict:
    """
    Procesa un PDF Metrigest e importa todas las lecturas a la BD.

    Args:
        ruta_pdf:        Ruta al PDF de Metrigest
        ruta_bd:         Ruta a gestion.db
        mover_procesado: Si True, mueve el PDF a ../procesados/ al terminar
        verbose:         Imprimir progreso

    Returns:
        {ok, tipo, lecturas_insertadas, lecturas_duplicadas, errores, vecinos_no_encontrados}
    """
    nombre = os.path.basename(ruta_pdf)
    ruta_prov = ruta_proveedores or str(BASE_DIR / "config" / "proveedores.json")

    if not os.path.exists(ruta_bd):
        return {"ok": False, "error": "BD no encontrada. Ejecuta importar_comunidad.py primero"}

    con = conectar(ruta_bd)

    # ¿Ya procesado?
    if archivo_ya_procesado(con, nombre):
        if verbose:
            print(f"  ⏭️  Ya procesado: {nombre}")
        con.close()
        return {"ok": True, "ya_procesado": True, "lecturas_insertadas": 0}

    # Detectar código de comunidad desde el nombre del archivo
    # Convención: '644_METRIGEST_...' o 'METRIGEST_644_...'
    import re as _re
    codigo_desde_nombre = None
    m_cod = _re.match(r'^(\d{3,4})[_\-\s]', nombre, _re.IGNORECASE)
    if m_cod:
        codigo_desde_nombre = m_cod.group(1)

    # Procesar el PDF
    resultado = procesar_archivo(ruta_pdf, codigo_desde_nombre, con_bd=con,
                                  ruta_proveedores=ruta_prov)

    if not resultado.get("ok") or resultado.get("tipo") != "LECTURA_METRIGEST":
        con.close()
        return {
            "ok": False,
            "error": resultado.get("motivo", "No es un archivo Metrigest"),
            "detalle": resultado.get("detalle", "")
        }

    datos    = resultado["datos"]
    tipo     = datos["tipo"]             # 'ACS' o 'CALEFACCION'
    vecinos  = datos["vecinos"]          # lista de dicts
    cabecera = datos.get("cabecera", {})
    codigo_comunidad = resultado.get("codigo_comunidad")

    if verbose:
        print(f"\n  Comunidad: {codigo_comunidad}")
        print(f"  Tipo:      {tipo}")
        print(f"  Periodo:   {cabecera.get('fecha_inicio')} → {cabecera.get('fecha_fin')}")
        print(f"  Vecinos:   {len(vecinos)}")

    # Buscar la comunidad en BD
    if codigo_comunidad:
        com = con.execute(
            "SELECT id_comunidad, nombre FROM comunidades WHERE codigo=?",
            (codigo_comunidad,)
        ).fetchone()
    else:
        # Buscar por CIF
        cif = resultado.get("cif_pdf")
        com = con.execute(
            "SELECT id_comunidad, nombre FROM comunidades WHERE cif=?",
            (cif,)
        ).fetchone() if cif else None

    if not com:
        con.close()
        return {"ok": False, "error": f"Comunidad no encontrada en BD (código={codigo_comunidad})"}

    id_comunidad = com["id_comunidad"]

    # Obtener o crear periodo — usar las fechas del PDF
    fecha_ini = cabecera.get("fecha_inicio", "")
    fecha_fin = cabecera.get("fecha_fin", "")

    # Buscar periodo existente que incluya estas fechas
    # O el periodo abierto más reciente
    periodo = con.execute("""
        SELECT id_periodo, nombre FROM periodos
        WHERE id_comunidad=? AND estado='abierto'
        ORDER BY fecha_inicio DESC LIMIT 1
    """, (id_comunidad,)).fetchone()

    if not periodo:
        # Crear periodo automáticamente
        nombre_per = _inferir_nombre_periodo(fecha_ini)
        con.execute("""
            INSERT OR IGNORE INTO periodos
                (id_comunidad, nombre, fecha_inicio, fecha_fin, estado)
            VALUES (?,?,?,?,?)
        """, (id_comunidad, nombre_per, fecha_ini, fecha_fin, "abierto"))
        con.commit()
        periodo = con.execute(
            "SELECT id_periodo, nombre FROM periodos WHERE id_comunidad=? AND nombre=?",
            (id_comunidad, nombre_per)
        ).fetchone()

    id_periodo = periodo["id_periodo"]
    if verbose:
        print(f"  Periodo BD: {periodo['nombre']} (id={id_periodo})")

    # Construir mapa vivienda→id_propietario
    propietarios = con.execute(
        "SELECT id_propietario, codigo_vivienda, nombre_propietario FROM propietarios WHERE id_comunidad=?",
        (id_comunidad,)
    ).fetchall()

    # Mapa con variantes de búsqueda: exacto, normalizado y por nombre
    mapa_vivienda = {}
    mapa_nombre   = {}
    for p in propietarios:
        cod_norm = _normalizar_vivienda(p["codigo_vivienda"])
        mapa_vivienda[cod_norm] = p["id_propietario"]
        # También por nombre (primeras palabras)
        nom_clave = p["nombre_propietario"].upper()[:20]
        mapa_nombre[nom_clave] = p["id_propietario"]

    insertadas    = 0
    duplicadas    = 0
    no_encontrados = []

    for vecino in vecinos:
        vivienda_pdf = vecino["vivienda"]
        nombre_pdf   = vecino["nombre"]
        tipo_lec     = vecino["tipo"]     # 'ACS' o 'CALEFACCION'

        # Buscar propietario: primero por vivienda, luego por nombre
        cod_norm = _normalizar_vivienda(vivienda_pdf)
        id_prop  = mapa_vivienda.get(cod_norm)

        if not id_prop:
            # Intentar por nombre
            nom_key = nombre_pdf.upper()[:20]
            id_prop = mapa_nombre.get(nom_key)

        if not id_prop:
            no_encontrados.append(vivienda_pdf)
            continue

        # Insertar lectura ANTERIOR (val_ant en fecha_ant)
        if vecino.get("fecha_ant") and vecino.get("val_ant") is not None:
            try:
                con.execute("""
                    INSERT INTO lecturas_vecino
                        (id_propietario, id_periodo, tipo, fecha_lectura,
                         valor_acumulado, estado, fuente)
                    VALUES (?,?,?,?,?,?,?)
                """, (id_prop, id_periodo, tipo_lec,
                      vecino["fecha_ant"], vecino["val_ant"],
                      "real", "metrigest_pdf"))
                insertadas += 1
            except sqlite3.IntegrityError:
                duplicadas += 1

        # Insertar lectura ACTUAL (val_act en fecha_act)
        if vecino.get("fecha_act") and vecino.get("val_act") is not None:
            try:
                con.execute("""
                    INSERT INTO lecturas_vecino
                        (id_propietario, id_periodo, tipo, fecha_lectura,
                         valor_acumulado, estado, fuente)
                    VALUES (?,?,?,?,?,?,?)
                """, (id_prop, id_periodo, tipo_lec,
                      vecino["fecha_act"], vecino["val_act"],
                      "real", "metrigest_pdf"))
                insertadas += 1
            except sqlite3.IntegrityError:
                duplicadas += 1

    con.commit()

    # Marcar archivo como procesado
    marcar_archivo_procesado(con, nombre, resultado.get("hash_md5"),
                              resultado="ok",
                              notas=f"METRIGEST {tipo} {fecha_ini}→{fecha_fin} {len(vecinos)}vec")
    con.close()

    # Mover a procesados
    if mover_procesado and os.path.exists(ruta_pdf):
        dest_dir = Path(ruta_pdf).parent.parent / "procesados"
        dest_dir.mkdir(exist_ok=True)
        shutil.move(ruta_pdf, str(dest_dir / nombre))

    if verbose:
        print(f"  ✅ Insertadas: {insertadas}  |  Duplicadas: {duplicadas}  |  No encontrados: {len(no_encontrados)}")
        if no_encontrados[:5]:
            print(f"  ⚠️  Sin match: {no_encontrados[:5]}")

    return {
        "ok": True,
        "tipo": tipo,
        "periodo": periodo["nombre"],
        "vecinos_total": len(vecinos),
        "lecturas_insertadas": insertadas,
        "lecturas_duplicadas": duplicadas,
        "vecinos_no_encontrados": no_encontrados,
    }


# ---------------------------------------------------------------------------
# UTILIDADES
# ---------------------------------------------------------------------------

def _normalizar_vivienda(codigo: str) -> str:
    """
    Normaliza el código de vivienda para comparación robusta.
    'SS22 BAJO IZDA' → 'SS22BAJOIZDA'
    'SS22 1 IZDA' → 'SS221IZDA'
    """
    # Quitar espacios, puntos, º
    return re.sub(r'[\s\.º°]', '', codigo.upper())


def _inferir_nombre_periodo(fecha_inicio: str) -> str:
    """
    Infiere el nombre del periodo a partir de la fecha de inicio.
    '2025-08-25' → '2024-2025'
    '2026-01-30' → '2025-2026'
    """
    try:
        año = int(fecha_inicio[:4])
        mes = int(fecha_inicio[5:7])
        # Si la lectura es en la primera mitad del año, pertenece al ejercicio anterior
        if mes <= 8:
            return f"{año-1}-{año}"
        else:
            return f"{año}-{año+1}"
    except (ValueError, IndexError):
        return "2024-2025"


# ---------------------------------------------------------------------------
# PROCESAR CARPETA COMPLETA
# ---------------------------------------------------------------------------

def procesar_carpeta(carpeta: str, ruta_bd: str,
                     ruta_proveedores: str = None) -> dict:
    """
    Procesa todos los PDFs Metrigest en una carpeta.
    Sólo procesa archivos que contengan 'METRIGEST' o 'Resumen' en el nombre.
    """
    archivos = [
        f for f in os.listdir(carpeta)
        if f.lower().endswith(".pdf")
        and ("metrigest" in f.lower() or "resumen" in f.lower())
    ]

    if not archivos:
        print(f"No se encontraron PDFs Metrigest en {carpeta}")
        return {"ok": True, "procesados": 0}

    print(f"\n📂 {len(archivos)} PDF(s) Metrigest en {carpeta}\n")

    total_insertadas = 0
    total_errores = 0

    for nombre in sorted(archivos):
        ruta = os.path.join(carpeta, nombre)
        print(f"  → {nombre}")
        resultado = importar_lecturas_pdf(ruta, ruta_bd,
                                           ruta_proveedores=ruta_proveedores)
        if resultado.get("ok"):
            total_insertadas += resultado.get("lecturas_insertadas", 0)
        else:
            print(f"     ❌ {resultado.get('error', '')}")
            total_errores += 1

    print(f"\n  Total lecturas insertadas: {total_insertadas}")
    print(f"  Total errores: {total_errores}")
    return {"ok": total_errores == 0, "lecturas_insertadas": total_insertadas}


# ---------------------------------------------------------------------------
# EJECUCIÓN DIRECTA
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Importa lecturas Metrigest a la BD"
    )
    parser.add_argument("--pdf",      help="Ruta a un PDF Metrigest específico")
    parser.add_argument("--carpeta",  help="Carpeta con PDFs Metrigest")
    parser.add_argument("--bd",       default=str(RUTA_BD), help="Ruta a gestion.db")
    parser.add_argument("--no-mover", action="store_true",
                        help="No mover el PDF a procesados/")
    args = parser.parse_args()

    if args.pdf:
        print(f"\n=== IMPORTANDO LECTURAS METRIGEST ===")
        resultado = importar_lecturas_pdf(
            ruta_pdf=args.pdf,
            ruta_bd=args.bd,
            mover_procesado=not args.no_mover,
        )
        if resultado.get("ok"):
            print(f"\n✅ OK — {resultado.get('lecturas_insertadas', 0)} lecturas insertadas")
        else:
            print(f"\n❌ Error: {resultado.get('error', '')}")

    elif args.carpeta:
        procesar_carpeta(args.carpeta, args.bd)

    else:
        # Por defecto: procesar la carpeta entrada/
        procesar_carpeta(str(RUTA_ENTRADA), args.bd)
