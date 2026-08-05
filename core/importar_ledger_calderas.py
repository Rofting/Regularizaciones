"""
importar_ledger_calderas.py
============================
Parser + clasificador para los extractos contables internos de Meditrade
de mantenimiento/reparación de Sala de Calderas ("MTO CALDERAS ..." /
"REP CALDERAS ..."), y utilidades para llevarlos a la BD.

QUÉ SON ESTOS PDFS
-------------------
NO son facturas de un proveedor externo. Son extractos del propio libro
contable de Meditrade, uno por "cuenta de concepto" (p.ej. 6220301 - Mto.
Sala de Calderas, 6220303 - Reparación Instalación Calefacción y ACS,
6220306 - Reparación Sala Calderas y ACS), con una línea por cargo:

    Fecha  Número  Documento  Comentario                          Debe  Haber  Saldo
    31/10/2023 439 23047146   Mto. Sala Calderas Octubre 2023   371,60  0,00  371,60
    ...
    30/09/2023 1797 ASIENTO DE REGULARIZACIÓN                    0,00 4.399,12  0,00
    Suma total....                                             4.399,12 4.399,12  0,00

Detalles reales encontrados al inspeccionar los PDFs históricos de la 644
que hay que tener en cuenta:
  - El nº de "Documento" a veces falta (asientos manuales).
  - El comentario a veces se parte en dos líneas en el PDF (la continuación
    no empieza por fecha ni tiene importes) — hay que fusionarla con la
    línea anterior.
  - Puede haber importes NEGATIVOS (abonos de duplicados, p.ej. "Abono
    duplicidad enero 2023 -371,60").
  - La última línea real siempre es "ASIENTO DE REGULARIZACIÓN" (Debe=0,00,
    Haber=total) — es el asiento que cierra la cuenta a 0, no es un gasto.
  - La fila "Suma total...." es el total del extracto — se usa para
    verificar la suma de las líneas, no se inserta como gasto.
  - UN MISMO ejercicio (año fiscal) puede repartirse entre VARIOS PDFs con
    distinto código de concepto (p.ej. 2022-2023 tiene reparaciones en
    6220303 y en 6220306 a la vez) — hay que sumarlos todos al bloque
    "MANTENIMIENTO CORRECTIVO (REPARACIONES)" de ese ejercicio.

CLASIFICACIÓN AMORTIZADO (5 años) vs GASTO DIRECTO (1 año)
-------------------------------------------------------------
En el histórico 2020-2024 (38 líneas, ya verificadas contra el Excel de
referencia real línea a línea) el criterio real del gestor fue, a grandes
rasgos:
    - Si el comentario es un aviso/incidencia puntual sin sustitución de
      material ("aviso", "fallo", "rearmar", "abrir depósitos", "apertura
      depósitos...limpiar") -> gasto directo de ese año (1).
    - Si implica sustituir/instalar/colocar material -> amortizado a 5 años.

Esta heurística, contrastada contra las 38 líneas históricas ya validadas,
acierta 35/38 (92%). Las 3 excepciones son casos donde el comentario
empieza describiendo un "fallo/aviso" pero termina mencionando una pieza
sustituida (p.ej. "fallo Antivibratorio, Válvula contadores..." o "aviso
fallo ACS y cambio válvula contador") — el gestor los amortizó a 5 años
pese a la palabra "fallo". NO es un criterio 100% mecánico, así que
`revisar_pdf_calderas()` NUNCA escribe en la BD por sí sola: solo
devuelve la clasificación PROPUESTA para que un humano la confirme o
corrija antes de llamar a `insertar_lineas()`.
"""
from __future__ import annotations

import re
import sqlite3
from datetime import datetime
from typing import Optional


# ---------------------------------------------------------------------------
# PARSEO DEL TEXTO DEL PDF
# ---------------------------------------------------------------------------

RE_CONCEPTO = re.compile(r"^\s*(?P<codigo>\d{6,7})\s*-\s*(?P<nombre>.+?)\s*$", re.MULTILINE)
RE_LINEA = re.compile(
    r"^(?P<fecha>\d{2}/\d{2}/\d{4})\s+"
    r"(?P<numero>\d+)\s+"
    r"(?:(?P<documento>\d{6,})\s+)?"
    r"(?P<resto>.+)$"
)
RE_IMPORTES = re.compile(
    r"^(?P<comentario>.*?)\s+"
    r"(?P<debe>-?[\d.]+,\d{2})\s+"
    r"(?P<haber>-?[\d.]+,\d{2})\s+"
    r"(?P<saldo>-?[\d.]+,\d{2})\s*$"
)


def _num_es(txt: str) -> float:
    """'1.234,56' / '-371,60' -> float."""
    return float(txt.replace(".", "").replace(",", "."))


def parsear_ledger_meditrade(texto: str) -> dict:
    """
    Parsea el texto extraído de un PDF de extracto interno de Meditrade
    (MTO CALDERAS / REP CALDERAS y similares).

    Devuelve:
        {
          'concepto_codigo': '6220301',
          'concepto_nombre': 'Mto. Sala de Calderas',
          'lineas': [
             {'fecha': '2023-10-31', 'numero': '439', 'documento': '23047146',
              'comentario': 'Mto. Sala Calderas Octubre 2023', 'importe': 371.60},
             ...
          ],
          'suma_total': 4562.88,   # de la fila "Suma total...."
          'suma_lineas': 4562.88,  # suma real de 'lineas' (para verificar)
        }
    """
    m_concepto = RE_CONCEPTO.search(texto)
    concepto_codigo = m_concepto.group("codigo") if m_concepto else None
    concepto_nombre = m_concepto.group("nombre") if m_concepto else None

    lineas = []
    suma_total = None
    pendiente = None  # última línea de datos, por si la siguiente es continuación

    for cruda in texto.splitlines():
        linea = cruda.strip()
        if not linea:
            continue

        if linea.upper().startswith("SUMA TOTAL"):
            m = RE_IMPORTES.match(linea.replace("Suma total....", "").strip() + " x")
            # La fila "Suma total...." trae 3 importes pegados sin comentario;
            # nos basta el primero (Debe).
            nums = re.findall(r"-?[\d.]+,\d{2}", linea)
            if nums:
                suma_total = _num_es(nums[0])
            pendiente = None
            continue

        m_lin = RE_LINEA.match(linea)
        if m_lin:
            resto = m_lin.group("resto")
            m_imp = RE_IMPORTES.match(resto)
            if not m_imp:
                # Línea con fecha pero sin importes reconocibles: se ignora
                pendiente = None
                continue

            comentario = m_imp.group("comentario").strip()
            debe = _num_es(m_imp.group("debe"))

            if "ASIENTO DE REGULARIZACI" in comentario.upper():
                pendiente = None
                continue

            fecha_iso = datetime.strptime(m_lin.group("fecha"), "%d/%m/%Y").strftime("%Y-%m-%d")
            item = {
                "fecha": fecha_iso,
                "numero": m_lin.group("numero"),
                "documento": m_lin.group("documento"),
                "comentario": comentario,
                "importe": debe,
            }
            lineas.append(item)
            pendiente = item
        else:
            # No empieza por fecha: o es cabecera/ruido, o es la continuación
            # del comentario de la línea anterior (comentario partido en dos
            # líneas por el PDF, p.ej. "Serpentinas" o "Contador Agua Fría...").
            if pendiente is not None and not linea.upper().startswith(("FECHA", "MEDITRADE", "PS ", "EMAIL", "644 -")):
                pendiente["comentario"] = f"{pendiente['comentario']} {linea}".strip()

    return {
        "concepto_codigo": concepto_codigo,
        "concepto_nombre": concepto_nombre,
        "lineas": lineas,
        "suma_total": suma_total,
        "suma_lineas": round(sum(l["importe"] for l in lineas), 2),
    }


# ---------------------------------------------------------------------------
# CLASIFICACIÓN AMORTIZADO (5 AÑOS) vs GASTO DIRECTO (1 AÑO)
# ---------------------------------------------------------------------------

_PALABRAS_GASTO_DIRECTO = ("aviso", "avisos", "fallo", "rearmar")


def clasificar_amortizacion(comentario: str) -> int:
    """
    Heurística validada contra las 38 reparaciones históricas 2020-2024 de
    la 644 (acierta 35/38 = 92%): un aviso/incidencia puntual sin sustituir
    material se computa como gasto directo del año (1); todo lo demás
    (sustituir, instalar, colocar, cambiar...) se amortiza a 5 años. Falla
    en avisos que terminan mencionando una pieza sustituida — ver docstring
    del módulo.

    Devuelve 1 o 5. Es solo una PROPUESTA — revisar_pdf_calderas() la usa
    para presentar la clasificación a un humano, no para insertar sola.
    """
    c = comentario.lower()
    if any(p in c for p in _PALABRAS_GASTO_DIRECTO):
        return 1
    if "abrir dep" in c or ("apertura" in c and "dep" in c):
        return 1
    return 5


def clasificar_servicio(comentario: str) -> str:
    """
    Heurística de a qué servicio afecta la reparación, a partir de qué
    palabra aparece en el comentario. 'AMBOS' reparte 50/50 en el motor de
    reparto (ver motor_reparto.py) — es el valor por defecto cuando el
    comentario no distingue ACS de Calefacción o menciona los dos.
    """
    c = comentario.upper()
    tiene_acs = "ACS" in c
    tiene_calef = "CALEF" in c
    if tiene_acs and not tiene_calef:
        return "ACS"
    if tiene_calef and not tiene_acs:
        return "CALEFACCION"
    return "AMBOS"


# ---------------------------------------------------------------------------
# REVISIÓN (dry-run) — no escribe nada en la BD
# ---------------------------------------------------------------------------

def revisar_pdf_calderas(ruta_pdf: str) -> dict:
    """
    Extrae el texto de un PDF de ledger de calderas y devuelve la
    clasificación PROPUESTA de cada línea (tipo MANTENIMIENTO o REPARACION,
    y para reparaciones también años de amortización + servicio afectado
    propuestos), lista para que un humano la confirme antes de insertar.

    No escribe nada en la BD.
    """
    import pdfplumber

    with pdfplumber.open(ruta_pdf) as pdf:
        texto = "\n".join(p.extract_text() or "" for p in pdf.pages)

    datos = parsear_ledger_meditrade(texto)
    es_mantenimiento = bool(datos["concepto_nombre"]) and "MTO" in datos["concepto_nombre"].upper() \
        or (datos["concepto_nombre"] and "MANTENIMIENTO" in datos["concepto_nombre"].upper())
    es_reparacion = datos["concepto_nombre"] and "REPARACI" in datos["concepto_nombre"].upper()

    propuesta = []
    for l in datos["lineas"]:
        fila = dict(l)
        if es_reparacion:
            fila["tipo_gasto"] = "REPARACION"
            fila["años_amortizacion"] = clasificar_amortizacion(l["comentario"])
            fila["servicio_afectado"] = clasificar_servicio(l["comentario"])
        else:
            fila["tipo_gasto"] = "MANTENIMIENTO"
            fila["años_amortizacion"] = 1
            fila["servicio_afectado"] = "AMBOS"
        propuesta.append(fila)

    ok_suma = (datos["suma_total"] is None) or (abs(datos["suma_total"] - datos["suma_lineas"]) < 0.01)

    return {
        "archivo": ruta_pdf,
        "concepto_codigo": datos["concepto_codigo"],
        "concepto_nombre": datos["concepto_nombre"],
        "es_mantenimiento": bool(es_mantenimiento),
        "es_reparacion": bool(es_reparacion),
        "lineas": propuesta,
        "suma_total_pdf": datos["suma_total"],
        "suma_lineas": datos["suma_lineas"],
        "cuadra": ok_suma,
    }


def imprimir_revision(revision: dict) -> None:
    """Imprime una tabla legible de la revisión para confirmar a mano."""
    print(f"\n=== {revision['archivo']} ===")
    print(f"Concepto: {revision['concepto_codigo']} - {revision['concepto_nombre']}")
    estado = "OK" if revision["cuadra"] else "NO CUADRA"
    print(f"Suma PDF: {revision['suma_total_pdf']}  |  Suma líneas: {revision['suma_lineas']}  [{estado}]")
    for l in revision["lineas"]:
        extra = ""
        if l["tipo_gasto"] == "REPARACION":
            extra = f"  -> {l['años_amortizacion']}a  serv={l['servicio_afectado']}"
        print(f"  {l['fecha']}  {l['importe']:>10.2f} €  {l['comentario']}{extra}")


# ---------------------------------------------------------------------------
# INSERCIÓN (solo tras revisión humana — nunca automática)
# ---------------------------------------------------------------------------

def insertar_lineas(ruta_bd: str, id_comunidad: int, revision: dict,
                     id_periodo: Optional[int] = None,
                     origen: str = "importado_ledger_calderas") -> dict:
    """
    Inserta en la BD las líneas de una revisión YA CONFIRMADA (por un
    humano, tras revisar `imprimir_revision`). Es idempotente por
    `origen` + comunidad + concepto: si ya se insertó este mismo PDF, no
    duplica.

    - MANTENIMIENTO -> tabla `facturas` (requiere id_periodo).
    - REPARACION     -> tabla `gastos_extra` (amortización propia).
    """
    con = sqlite3.connect(ruta_bd)
    cur = con.cursor()

    marca = f"{origen}:{revision['concepto_codigo']}"
    ya = cur.execute(
        "SELECT COUNT(*) FROM facturas WHERE id_comunidad=? AND archivo_origen=?",
        (id_comunidad, marca)
    ).fetchone()[0]
    ya += cur.execute(
        "SELECT COUNT(*) FROM gastos_extra WHERE id_comunidad=? AND notas=?",
        (id_comunidad, marca)
    ).fetchone()[0]
    if ya:
        con.close()
        return {"ok": False, "error": f"Ya hay {ya} filas insertadas con origen '{marca}'"}

    insertadas = 0
    for l in revision["lineas"]:
        if l["tipo_gasto"] == "MANTENIMIENTO":
            if id_periodo is None:
                con.close()
                return {"ok": False, "error": "MANTENIMIENTO requiere id_periodo"}
            cur.execute("""INSERT INTO facturas
                (id_comunidad, id_periodo, tipo_suministro, proveedor, fecha_factura,
                 fecha_inicio, fecha_fin, importe_total, notas, archivo_origen)
                VALUES (?,?,?,?,?,?,?,?,?,?)""",
                (id_comunidad, id_periodo, "MANTENIMIENTO",
                 f"Meditrade — {revision['concepto_nombre']} (ficha interna)",
                 l["fecha"], l["fecha"], l["fecha"], l["importe"],
                 revision["concepto_nombre"], marca))
        else:
            cur.execute("""INSERT INTO gastos_extra
                (id_comunidad, descripcion, fecha, importe_total, años_amortizacion,
                 tipo_gasto, servicio_afectado, notas)
                VALUES (?,?,?,?,?,?,?,?)""",
                (id_comunidad, l["comentario"], l["fecha"], l["importe"],
                 l["años_amortizacion"], "REPARACION", l["servicio_afectado"], marca))
        insertadas += 1

    con.commit()
    con.close()
    return {"ok": True, "insertadas": insertadas}


if __name__ == "__main__":
    import sys
    if len(sys.argv) < 2:
        print("Uso: python importar_ledger_calderas.py <ruta_pdf> [<ruta_pdf> ...]")
        sys.exit(1)
    for ruta in sys.argv[1:]:
        imprimir_revision(revisar_pdf_calderas(ruta))
