"""Cuotas cobradas a los propietarios por ACS y calefacción.

El estudio compara dos cosas: lo que ha costado el servicio y lo que se ha
cobrado por él. El coste sale de las facturas; lo cobrado, de las cuotas que
gira la comunidad, que no aparecen en ningún documento del expediente. Este
módulo guarda ese libro y lo resume para el análisis.

Hay dos clases de apunte, las mismas que el despacho lleva a mano:

- ``fija``: la cuota mensual por vivienda, un apunte por mes del período.
- ``variable``: la liquidación por consumo de cada lectura de contadores.
"""

from __future__ import annotations

import sqlite3
from dataclasses import dataclass
from datetime import date
from decimal import Decimal


SERVICIOS = ("ACS", "CALEFACCION")


@dataclass(frozen=True)
class ResumenCuotas:
    """Lo cobrado en el período, separado como lo separa el análisis."""

    fija: Decimal
    variable: Decimal
    apuntes: int

    @property
    def total(self) -> Decimal:
        return self.fija + self.variable


def _importe(valor: object, etiqueta: str) -> Decimal:
    try:
        importe = Decimal(str(valor).strip().replace(",", "."))
    except (ArithmeticError, ValueError, AttributeError):
        raise ValueError(f"{etiqueta} debe ser un número") from None
    if not importe.is_finite():
        raise ValueError(f"{etiqueta} debe ser un número")
    return importe


def _fecha(valor: str, etiqueta: str) -> str:
    try:
        return date.fromisoformat(str(valor)[:10]).isoformat()
    except ValueError:
        raise ValueError(f"{etiqueta} no es una fecha válida (aaaa-mm-dd)") from None


def meses_del_periodo(fecha_inicio: str, fecha_fin: str) -> list[str]:
    """Primer día de cada mes que toca el período, de inicio a fin."""
    inicio, fin = date.fromisoformat(fecha_inicio[:10]), date.fromisoformat(fecha_fin[:10])
    if fin < inicio:
        raise ValueError("El período no forma un intervalo válido")
    meses, año, mes = [], inicio.year, inicio.month
    while (año, mes) <= (fin.year, fin.month):
        meses.append(date(año, mes, 1).isoformat())
        año, mes = (año + 1, 1) if mes == 12 else (año, mes + 1)
    return meses


def registrar_cuota(
    connection: sqlite3.Connection,
    *,
    community_id: int,
    period_id: int,
    servicio: str,
    concepto: str,
    fecha: str,
    importe: object,
    consumo: object | None = None,
    notas: str | None = None,
) -> int:
    """Anota una cuota, sustituyendo la del mismo servicio, concepto y fecha."""
    if servicio not in SERVICIOS:
        raise ValueError("El servicio debe ser ACS o CALEFACCION")
    if concepto not in ("fija", "variable"):
        raise ValueError("El concepto debe ser 'fija' o 'variable'")
    valor = _importe(importe, "El importe de la cuota")
    if valor < 0:
        raise ValueError("El importe de la cuota no puede ser negativo")
    consumo_valor = None if consumo in (None, "") else float(_importe(consumo, "El consumo"))
    cursor = connection.execute(
        """INSERT INTO cuotas_servicio
           (id_comunidad,id_periodo,servicio,concepto,fecha,importe,consumo,notas)
           VALUES (?,?,?,?,?,?,?,?)
           ON CONFLICT(id_comunidad,id_periodo,servicio,concepto,fecha) DO UPDATE SET
               importe=excluded.importe, consumo=excluded.consumo, notas=excluded.notas""",
        (
            community_id, period_id, servicio, concepto, _fecha(fecha, "La fecha de la cuota"),
            float(valor), consumo_valor, (notas or "").strip() or None,
        ),
    )
    return int(cursor.lastrowid)


def generar_cuotas_mensuales(
    connection: sqlite3.Connection,
    *,
    community_id: int,
    period_id: int,
    servicio: str,
    tramos: list[tuple[str, object]],
    notas: str | None = None,
) -> int:
    """Crea una cuota fija por mes a partir de los tramos de importe.

    ``tramos`` son pares (mes desde el que aplica, importe mensual), como los
    lleva el despacho: 400 € desde agosto y 300 € desde enero. Cada mes del
    período recibe el importe del último tramo que ya haya empezado.
    """
    periodo = connection.execute(
        "SELECT fecha_inicio,fecha_fin FROM periodos WHERE id_periodo=?", (period_id,)
    ).fetchone()
    if periodo is None:
        raise LookupError("El período no existe")
    if not tramos:
        raise ValueError("Indica al menos un importe mensual")

    ordenados = sorted(
        (( _fecha(desde, "El mes del tramo"), _importe(importe, "El importe mensual"))
         for desde, importe in tramos),
        key=lambda tramo: tramo[0],
    )
    creadas = 0
    for mes in meses_del_periodo(periodo["fecha_inicio"], periodo["fecha_fin"]):
        vigentes = [importe for desde, importe in ordenados if desde[:7] <= mes[:7]]
        if not vigentes:
            continue
        registrar_cuota(
            connection, community_id=community_id, period_id=period_id,
            servicio=servicio, concepto="fija", fecha=mes, importe=vigentes[-1],
            notas=notas,
        )
        creadas += 1
    return creadas


def resumen(
    connection: sqlite3.Connection, *, community_id: int, period_id: int, servicio: str,
) -> ResumenCuotas:
    """Totales cobrados del período: fijo, variable y número de apuntes."""
    fila = connection.execute(
        """SELECT COALESCE(SUM(CASE WHEN concepto='fija' THEN importe END),0) AS fija,
                  COALESCE(SUM(CASE WHEN concepto='variable' THEN importe END),0) AS variable,
                  COUNT(*) AS apuntes
           FROM cuotas_servicio
           WHERE id_comunidad=? AND id_periodo=? AND servicio=?""",
        (community_id, period_id, servicio),
    ).fetchone()
    return ResumenCuotas(
        fija=Decimal(str(fila["fija"])), variable=Decimal(str(fila["variable"])),
        apuntes=int(fila["apuntes"]),
    )


def listar(
    connection: sqlite3.Connection, *, community_id: int, period_id: int, servicio: str,
) -> list[sqlite3.Row]:
    """Apuntes del período en orden cronológico, como en la hoja de cobros."""
    return connection.execute(
        """SELECT id_cuota,concepto,fecha,importe,consumo,fecha_inicio,fecha_fin,notas
           FROM cuotas_servicio
           WHERE id_comunidad=? AND id_periodo=? AND servicio=?
           ORDER BY fecha,concepto,id_cuota""",
        (community_id, period_id, servicio),
    ).fetchall()


def borrar(connection: sqlite3.Connection, id_cuota: int) -> None:
    connection.execute("DELETE FROM cuotas_servicio WHERE id_cuota=?", (id_cuota,))
