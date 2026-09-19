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
import math
from pathlib import Path
from datetime import datetime

sys.path.insert(0, str(Path(__file__).parent))
from gestor_bd import conectar, crear_bd, marcar_archivo_procesado, archivo_ya_procesado
from lector_pdf import procesar_archivo
import document_review

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
    conflictos     = []

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
            ok, conflicto = _insertar_o_detectar_conflicto(
                con, id_prop, id_periodo, tipo_lec,
                vecino["fecha_ant"], vecino["val_ant"], vivienda_pdf)
            if ok:
                insertadas += 1
            elif conflicto:
                conflictos.append(conflicto)
            else:
                duplicadas += 1

        # Insertar lectura ACTUAL (val_act en fecha_act)
        if vecino.get("fecha_act") and vecino.get("val_act") is not None:
            ok, conflicto = _insertar_o_detectar_conflicto(
                con, id_prop, id_periodo, tipo_lec,
                vecino["fecha_act"], vecino["val_act"], vivienda_pdf)
            if ok:
                insertadas += 1
            elif conflicto:
                conflictos.append(conflicto)
            else:
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
        if conflictos:
            print(f"  🔶 CONFLICTOS: {len(conflictos)} lectura(s) donde este informe no coincide con "
                  f"lo que ya había en la BD para la misma fecha — revisar a mano, no se sobrescribió nada:")
            for c in conflictos[:10]:
                print(f"       {c['vivienda']} {c['tipo']} {c['fecha']}: "
                      f"BD={c['valor_en_bd']} vs este informe={c['valor_en_este_informe']}")

    return {
        "ok": True,
        "tipo": tipo,
        "periodo": periodo["nombre"],
        "vecinos_total": len(vecinos),
        "lecturas_insertadas": insertadas,
        "lecturas_duplicadas": duplicadas,
        "vecinos_no_encontrados": no_encontrados,
        "conflictos": conflictos,
    }


# ---------------------------------------------------------------------------
# UTILIDADES
# ---------------------------------------------------------------------------

def apply_confirmed_readings(
    con: sqlite3.Connection,
    *,
    case_id: int,
    document_id: int,
    community_id: int,
    period_id: int,
    source_path: str,
    readings: list[dict],
) -> bool:
    """Aplica lecturas ya confirmadas sin confirmar ni reemplazar datos dudosos.

    Devuelve ``True`` cuando la fuente sigue pendiente de revisión (por ejemplo,
    por un reinicio de contador). El llamador conserva la transacción completa.
    """
    owners = {
        _normalizar_vivienda(row["codigo_vivienda"]): row
        for row in con.execute(
            "SELECT id_propietario,codigo_vivienda FROM propietarios WHERE id_comunidad=?",
            (community_id,),
        )
    }
    pending_review = False

    for reading in readings:
        if not isinstance(reading, dict):
            raise ValueError("Cada lectura confirmada debe ser un objeto")
        property_code = str(reading.get("vivienda") or "").strip()
        service = str(reading.get("tipo") or "").strip().upper()
        if not property_code or service not in {"ACS", "CALEFACCION"}:
            raise ValueError("La lectura confirmada debe incluir vivienda y tipo válidos")
        owner = owners.get(_normalizar_vivienda(property_code))
        if owner is None:
            _create_reading_issue(
                con, case_id, document_id, "UNMATCHED_OWNER",
                f"reading.{property_code}.{service}",
                "La vivienda de la lectura no existe entre los propietarios",
            )
            pending_review = True
            continue
        owner_id = owner["id_propietario"]

        initial_date = _confirmed_date(reading.get("fecha_ant"))
        final_date = _confirmed_date(reading.get("fecha_act"))
        if final_date <= initial_date:
            raise ValueError("Las fechas de lectura no forman un intervalo válido")
        initial_value = _confirmed_number(reading.get("val_ant"))
        final_value = _confirmed_number(reading.get("val_act"))
        field_name = f"reading.{owner['codigo_vivienda']}.{service}"
        reset_status = _counter_reset_status(con, document_id, field_name, initial_date, final_date)

        initial_observation = record_reading_observation(
            con, owner_id=owner_id, service=service, reading_date=initial_date,
            observed_value=initial_value, source_path=source_path, document_id=document_id,
        )
        if initial_value == 0:
            if not apply_zero_carry_forward(
                con, owner_id=owner_id, service=service, reading_date=initial_date,
                period_id=period_id, observation_id=initial_observation, source_path=source_path,
            ):
                _create_reading_issue(
                    con, case_id, document_id, "READING_ZERO_REVIEW", field_name,
                    "La lectura cero no tiene una lectura anterior fiable para conservar",
                    "0",
                )
                pending_review = True
                continue
        else:
            initial_ok = _insert_confirmed_reading(
                con, owner_id, period_id, service, initial_date, initial_value,
                "real", source_path,
            )
            _link_observation_to_effective_reading(
                con, observation_id=initial_observation, owner_id=owner_id,
                service=service, reading_date=initial_date,
                status="observed" if initial_ok else "conflict",
            )
            if not initial_ok:
                _create_reading_issue(
                    con, case_id, document_id, "READING_CONFLICT", field_name,
                    "La lectura confirmada contradice una lectura canónica existente",
                )
                pending_review = True
                continue

        final_observation = record_reading_observation(
            con, owner_id=owner_id, service=service, reading_date=final_date,
            observed_value=final_value, source_path=source_path, document_id=document_id,
        )
        if final_value == 0:
            if not apply_zero_carry_forward(
                con, owner_id=owner_id, service=service, reading_date=final_date,
                period_id=period_id, observation_id=final_observation, source_path=source_path,
            ):
                _create_reading_issue(
                    con, case_id, document_id, "READING_ZERO_REVIEW", field_name,
                    "La lectura cero no tiene una lectura anterior fiable para conservar",
                    "0",
                )
                pending_review = True
            continue

        if final_value < initial_value:
            if reset_status == "resolved":
                # La aprobación ya sustituyó el valor final por una lectura virtual.
                continue
            if reset_status == "open":
                pending_review = True
                continue
            final_ok = _insert_confirmed_reading(
                con, owner_id, period_id, service, final_date, final_value,
                "contador_averiado", source_path,
                "lectura inferior a la anterior; pendiente de estimación aprobada",
            )
            _link_observation_to_effective_reading(
                con, observation_id=final_observation, owner_id=owner_id,
                service=service, reading_date=final_date,
                status="observed" if final_ok else "conflict",
            )
            if not final_ok:
                _create_reading_issue(
                    con, case_id, document_id, "READING_CONFLICT", field_name,
                    "La lectura confirmada contradice una lectura canónica existente",
                )
            else:
                counter_field = (
                    "estado_contador_acs" if service == "ACS" else "estado_contador_cal"
                )
                con.execute(
                    f"UPDATE propietarios SET {counter_field}='averiado' WHERE id_propietario=?",
                    (owner_id,),
                )
                issue = _create_reading_issue(
                    con, case_id, document_id, "COUNTER_RESET",
                    f"{field_name}|{initial_date}/{final_date}",
                    "El contador disminuye y requiere una estimación aprobada",
                    f"{initial_value} -> {final_value}",
                )
                targets = con.execute("""SELECT id_lectura,fecha_lectura FROM lecturas_vecino
                    WHERE id_propietario=? AND tipo=? AND fecha_lectura IN (?,?)""",
                    (owner_id, service, initial_date, final_date)).fetchall()
                by_date = {row['fecha_lectura']: row['id_lectura'] for row in targets}
                con.execute("""INSERT OR IGNORE INTO counter_reset_targets
                    (id_issue,initial_reading_id,final_reading_id) VALUES (?,?,?)""",
                    (issue.id_issue, by_date[initial_date], by_date[final_date]))
            pending_review = True
            continue

        final_ok = _insert_confirmed_reading(
            con, owner_id, period_id, service, final_date, final_value,
            "real", source_path,
        )
        _link_observation_to_effective_reading(
            con, observation_id=final_observation, owner_id=owner_id,
            service=service, reading_date=final_date,
            status="observed" if final_ok else "conflict",
        )
        if not final_ok:
            _create_reading_issue(
                con, case_id, document_id, "READING_CONFLICT", field_name,
                "La lectura confirmada contradice una lectura canónica existente",
            )
            pending_review = True

    return pending_review


def record_reading_observation(
    con: sqlite3.Connection,
    *,
    owner_id: int,
    service: str,
    reading_date: str,
    observed_value: float,
    source_path: str,
    document_id: int,
) -> int:
    """Registra el valor de origen antes de decidir su proyección canónica."""
    existing = con.execute(
        """SELECT id_observation FROM reading_observations
           WHERE id_propietario=? AND tipo=? AND fecha_lectura=?
             AND observed_value=? AND id_document=?
           ORDER BY id_observation DESC LIMIT 1""",
        (owner_id, service, reading_date, observed_value, document_id),
    ).fetchone()
    if existing is not None:
        return int(existing["id_observation"])
    cursor = con.execute(
        """INSERT INTO reading_observations
           (id_propietario,tipo,fecha_lectura,observed_value,source_path,id_document,status)
           VALUES (?,?,?,?,?,?,'observed')""",
        (owner_id, service, reading_date, observed_value, source_path, document_id),
    )
    return int(cursor.lastrowid)


def _link_observation_to_effective_reading(
    con: sqlite3.Connection,
    *,
    observation_id: int,
    owner_id: int,
    service: str,
    reading_date: str,
    status: str,
) -> None:
    effective = con.execute(
        """SELECT id_lectura FROM lecturas_vecino
           WHERE id_propietario=? AND tipo=? AND fecha_lectura=?""",
        (owner_id, service, reading_date),
    ).fetchone()
    con.execute(
        """UPDATE reading_observations
           SET status=?, effective_reading_id=? WHERE id_observation=?""",
        (status, effective["id_lectura"] if effective is not None else None, observation_id),
    )


def apply_zero_carry_forward(
    con: sqlite3.Connection,
    *,
    owner_id: int,
    service: str,
    reading_date: str,
    period_id: int,
    observation_id: int,
    source_path: str,
) -> bool:
    """Conserva la última lectura válida al recibir un cero, sin ocultarlo."""
    current_reference = con.execute(
        """SELECT id_lectura FROM lecturas_vecino
           WHERE id_propietario=? AND tipo=? AND fecha_lectura=?
             AND valor_acumulado<>0
             AND (
                 estado='real'
                 OR (
                     estado='estimado'
                     AND metodo_estimacion='counter_reset_carry_forward'
                     AND trim(COALESCE(approved_by,''))<>''
                     AND trim(COALESCE(approved_at,''))<>''
                 )
             )""",
        (owner_id, service, reading_date),
    ).fetchone()
    if current_reference is not None:
        con.execute(
            "INSERT OR IGNORE INTO reading_periods(id_lectura,id_periodo) VALUES (?,?)",
            (current_reference["id_lectura"], period_id),
        )
        con.execute(
            """UPDATE reading_observations
               SET status='carried_forward',effective_reading_id=?
               WHERE id_observation=?""",
            (current_reference["id_lectura"], observation_id),
        )
        return True
    confirmed_zero = con.execute(
        """SELECT id_lectura FROM lecturas_vecino
           WHERE id_propietario=? AND tipo=? AND fecha_lectura=?
             AND valor_acumulado=0 AND estado='real'""",
        (owner_id, service, reading_date),
    ).fetchone()
    if confirmed_zero is not None:
        con.execute(
            "INSERT OR IGNORE INTO reading_periods(id_lectura,id_periodo) VALUES (?,?)",
            (confirmed_zero["id_lectura"], period_id),
        )
        con.execute(
            """UPDATE reading_observations
               SET status='observed',effective_reading_id=?,previous_reading_id=NULL
               WHERE id_observation=?""",
            (confirmed_zero["id_lectura"], observation_id),
        )
        return True
    previous = con.execute(
        """SELECT id_lectura,valor_acumulado FROM lecturas_vecino
           WHERE id_propietario=? AND tipo=? AND fecha_lectura < ?
             AND (estado='real' OR (estado='estimado'
                  AND trim(COALESCE(approved_by,''))<>''
                  AND trim(COALESCE(approved_at,''))<>''))
           ORDER BY fecha_lectura DESC,id_lectura DESC LIMIT 1""",
        (owner_id, service, reading_date),
    ).fetchone()
    if previous is None:
        con.execute(
            "UPDATE reading_observations SET status='review_required' WHERE id_observation=?",
            (observation_id,),
        )
        return False
    current = con.execute(
        """SELECT id_lectura,valor_acumulado,estado,metodo_estimacion
           FROM lecturas_vecino WHERE id_propietario=? AND tipo=? AND fecha_lectura=?""",
        (owner_id, service, reading_date),
    ).fetchone()
    if current is not None and (
        abs(float(current["valor_acumulado"]) - float(previous["valor_acumulado"])) > 0.01
        or current["estado"] == "real"
    ):
        con.execute(
            "UPDATE reading_observations SET status='conflict' WHERE id_observation=?",
            (observation_id,),
        )
        return False
    if current is None:
        cursor = con.execute(
            """INSERT INTO lecturas_vecino
               (id_propietario,id_periodo,tipo,fecha_lectura,valor_acumulado,estado,
                metodo_estimacion,fuente,notas)
               VALUES (?,?,?,?,?,'estimado','carry_forward_zero',?,?)""",
            (
                owner_id, period_id, service, reading_date, previous["valor_acumulado"],
                source_path, f"Lectura 0 recibida; se conserva la lectura válida de {reading_date}.",
            ),
        )
        effective_id = int(cursor.lastrowid)
    else:
        effective_id = int(current["id_lectura"])
    con.execute(
        "INSERT OR IGNORE INTO reading_periods(id_lectura,id_periodo) VALUES (?,?)",
        (effective_id, period_id),
    )
    con.execute(
        """UPDATE reading_observations
           SET status='carried_forward',effective_reading_id=?,previous_reading_id=?
           WHERE id_observation=?""",
        (effective_id, previous["id_lectura"], observation_id),
    )
    return True


def _confirmed_date(value: object) -> str:
    if not isinstance(value, str):
        raise ValueError("La fecha de una lectura confirmada es obligatoria")
    normalized = value.strip()
    try:
        return datetime.strptime(normalized, "%Y-%m-%d").date().isoformat()
    except ValueError:
        raise ValueError("La fecha de una lectura confirmada no es válida") from None


def _confirmed_number(value: object) -> float:
    if isinstance(value, bool) or value is None:
        raise ValueError("El valor de una lectura confirmada es obligatorio")
    try:
        number = float(str(value).strip().replace(",", "."))
    except ValueError:
        raise ValueError("El valor de una lectura confirmada debe ser numérico") from None
    if not math.isfinite(number):
        raise ValueError("El valor de una lectura confirmada debe ser finito")
    return number


def _insert_confirmed_reading(
    con: sqlite3.Connection,
    owner_id: int,
    period_id: int,
    service: str,
    reading_date: str,
    value: float,
    state: str,
    source_path: str,
    notes: str | None = None,
) -> bool:
    existing = con.execute(
        """SELECT id_lectura,valor_acumulado,estado,approved_by,approved_at FROM lecturas_vecino
           WHERE id_propietario=? AND tipo=? AND fecha_lectura=?""",
        (owner_id, service, reading_date),
    ).fetchone()
    if existing is not None:
        if abs(existing["valor_acumulado"] - value) > 0.01:
            return False
        if state == "real" and (existing['estado'] == 'contador_averiado' or
                (existing['estado'] == 'estimado' and not (existing['approved_by'] and existing['approved_at']))):
            return False
        con.execute("INSERT OR IGNORE INTO reading_periods(id_lectura,id_periodo) VALUES (?,?)",
                    (existing['id_lectura'], period_id))
        return True
    cursor = con.execute(
        """INSERT INTO lecturas_vecino
           (id_propietario,id_periodo,tipo,fecha_lectura,valor_acumulado,estado,fuente,notas)
           VALUES (?,?,?,?,?,?,?,?)""",
        (owner_id, period_id, service, reading_date, value, state, source_path, notes),
    )
    con.execute("INSERT OR IGNORE INTO reading_periods(id_lectura,id_periodo) VALUES (?,?)",
                (cursor.lastrowid, period_id))
    return True


def _counter_reset_status(
    con: sqlite3.Connection, document_id: int, field_name: str,
    initial_date: str, final_date: str,
) -> str | None:
    row = con.execute(
        """SELECT i.status FROM review_issues i
           LEFT JOIN counter_reset_targets t ON t.id_issue=i.id_issue
           LEFT JOIN lecturas_vecino a ON a.id_lectura=t.initial_reading_id
           LEFT JOIN lecturas_vecino b ON b.id_lectura=t.final_reading_id
           WHERE i.id_document=? AND i.code='COUNTER_RESET' AND i.field_name IN (?,?)
             AND (t.id_issue IS NULL OR (a.fecha_lectura=? AND b.fecha_lectura=?))
           ORDER BY i.id_issue DESC LIMIT 1""",
        (document_id, field_name, f"{field_name}|{initial_date}/{final_date}", initial_date, final_date),
    ).fetchone()
    return row["status"] if row is not None else None


def _create_reading_issue(
    con: sqlite3.Connection,
    case_id: int,
    document_id: int,
    code: str,
    field_name: str,
    message: str,
    detected_value: str | None = None,
):
    return document_review.create_review_issue(
        con, case_id, document_id, code=code, field_name=field_name,
        message=message, detected_value=detected_value,
    )

def _normalizar_vivienda(codigo: str) -> str:
    """
    Normaliza el código de vivienda para comparación robusta.
    'SS22 BAJO IZDA' → 'SS22BAJOIZDA'
    'SS22 1 IZDA' → 'SS221IZDA'
    """
    # Quitar espacios, puntos, º
    return re.sub(r'[\s\.º°]', '', codigo.upper())


def _insertar_o_detectar_conflicto(con, id_prop, id_periodo, tipo, fecha, valor, vivienda):
    """
    Inserta una lectura. Si ya existe una para (propietario, tipo, fecha) —
    algo habitual porque el periodo final de un informe Metrigest coincide
    con el periodo inicial del siguiente — comprueba que el valor coincida.
    Si coincide (o casi: contadores no dan decimales exactos), es un
    duplicado normal y se ignora. Si NO coincide, son dos informes oficiales
    de Metrigest contradiciéndose sobre la misma lectura del mismo contador
    el mismo día — no se puede saber automáticamente cuál es la buena, así
    que se reporta como conflicto en vez de quedarse callado con el primero
    que llegó (que es justo el bug que causó el -40.433 de la 644).

    Devuelve (insertada: bool, conflicto: dict|None).
    """
    try:
        con.execute("""
            INSERT INTO lecturas_vecino
                (id_propietario, id_periodo, tipo, fecha_lectura,
                 valor_acumulado, estado, fuente)
            VALUES (?,?,?,?,?,?,?)
        """, (id_prop, id_periodo, tipo, fecha, valor, "real", "metrigest_pdf"))
        return True, None
    except sqlite3.IntegrityError:
        existente = con.execute("""
            SELECT valor_acumulado FROM lecturas_vecino
            WHERE id_propietario=? AND tipo=? AND fecha_lectura=?
        """, (id_prop, tipo, fecha)).fetchone()
        valor_existente = existente["valor_acumulado"] if existente else None
        if valor_existente is not None and abs(valor_existente - valor) > 0.01:
            return False, {
                "vivienda": vivienda, "tipo": tipo, "fecha": fecha,
                "valor_en_bd": valor_existente, "valor_en_este_informe": valor,
            }
        return False, None


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
