"""Crea una instalación de prueba con comunidades inventadas para ensayar el flujo.

Uso (desde la carpeta del programa, con el entorno instalado)::

    .venv\\Scripts\\python.exe scripts\\crear_comunidades_prueba.py
    .venv\\Scripts\\python.exe scripts\\crear_comunidades_prueba.py --destino D:\\pruebas --reemplazar

Todo se crea en una carpeta propia (por defecto
``%LOCALAPPDATA%\\Regularizaciones-prueba``), con su propia base de datos.
La base de datos real del despacho no se abre ni se modifica. Para abrir la
aplicación sobre esa carpeta está ``Regularizaciones_prueba.bat``.

Comunidades creadas (todos los nombres, CIF, CUPS y correos son inventados;
los correos usan el dominio reservado ``example.invalid``):

* **901 · Lista para calcular**: facturas, lecturas, cuotas cobradas y gastos
  fijos ya confirmados. Sirve para pulsar «Generar Excel», «Calcular reparto»
  y «Generar cartas» sin preparar nada.
* **902 · Con incidencias**: los mismos datos con los problemas habituales de
  los documentos reales para practicar su resolución: una factura del
  ejercicio anterior, otra sin importe, una lectura final que falta, un
  contador que baja y un local sin contador.
* **903 · Bandeja de fuentes**: expediente vacío y una carpeta ``fuentes_903``
  con PDF de factura, un listado de propietarios y un informe de lecturas
  para probar «Añadir fuentes» de principio a fin.
"""

from __future__ import annotations

import argparse
import csv
import json
import os
import shutil
import sys
from contextlib import redirect_stdout
from dataclasses import dataclass
from datetime import date, timedelta
from io import StringIO
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
CORE = PROJECT_ROOT / "core"

PERIOD_START = date(2025, 8, 1)
PERIOD_END = date(2026, 7, 31)
RESOURCES = ("config", "plantillas")
MARKER = ".comunidades_prueba"


@dataclass(frozen=True)
class Dwelling:
    code: str
    name: str
    coefficient: str
    initial: int | None
    final: int | None
    kind: str = "vivienda"


def _dwellings(code: str) -> list[Dwelling]:
    people = (
        "Ana Prueba Uno", "Bruno Prueba Dos", "Carla Prueba Tres", "Diego Prueba Cuatro",
        "Elena Prueba Cinco", "Fermín Prueba Seis", "Gloria Prueba Siete", "Hugo Prueba Ocho",
    )
    result = []
    consumption = (14, 22, 9, 31, 18, 0, 26, 12)
    for index, person in enumerate(people):
        floor, door = divmod(index, 2)
        initial = 100 + 37 * index
        result.append(Dwelling(
            f"P{code[-1]}-{floor + 1}º{'AB'[door]}", person,
            f"{11 + (index % 3)},{25 * (index % 4):02d}",
            initial, initial + consumption[index],
        ))
    result.append(Dwelling(f"P{code[-1]}-LOC A", "Comercio de Prueba S.L.", "6,50", None, None, "local"))
    result.append(Dwelling(f"P{code[-1]}-GAR.", "Garajes de Prueba", "5,25", None, None, "garaje"))
    return result


def _email(name: str) -> str:
    import unicodedata
    plain = unicodedata.normalize("NFKD", name).encode("ascii", "ignore").decode().lower()
    return ".".join(plain.split()[:2]) + "@example.invalid"


def _months():
    current = date(PERIOD_START.year, PERIOD_START.month, 1)
    while current <= PERIOD_END:
        following = date(current.year + current.month // 12, current.month % 12 + 1, 1)
        yield current, following - timedelta(days=1)
        current = following


def _invoice_months():
    """Meses facturados dentro del ejercicio.

    El despacho cuenta la factura por su fecha de emisión: la de julio se emite
    en agosto. Para que las doce facturas queden dentro del ejercicio, el
    consumo facturado empieza un mes antes.
    """
    for start, _end in _months():
        previous_end = start - timedelta(days=1)
        yield date(previous_end.year, previous_end.month, 1), previous_end


def _write_pdf(path: Path, lines: list[str]) -> Path:
    """PDF de texto (no escaneado) con las líneas dadas, una página."""
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    from matplotlib.backends.backend_pdf import PdfPages

    path.parent.mkdir(parents=True, exist_ok=True)
    with matplotlib.rc_context({"pdf.fonttype": 42}), PdfPages(path) as pdf:
        figure = plt.figure(figsize=(8.27, 11.69))
        for index, line in enumerate(lines):
            figure.text(0.07, 0.95 - 0.028 * index, line, fontsize=10, family="DejaVu Sans")
        pdf.savefig(figure)
        plt.close(figure)
    return path


def _euros(value: float) -> str:
    return f"{value:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def _es(day: date) -> str:
    return day.strftime("%d/%m/%Y")


# ----------------------------------------------------------------------------
# Facturas inventadas. Los importes siguen una estacionalidad creíble: el gas
# sube en invierno (calefacción) y la luz y el agua apenas cambian.
# ----------------------------------------------------------------------------

def _gas_invoices():
    winter = {11: 2.6, 12: 4.1, 1: 4.4, 2: 3.6, 3: 2.2}
    for number, (start, end) in enumerate(_invoice_months(), start=1):
        kwh = round(3800 * winter.get(start.month, 1.0))
        variable = round(kwh * 0.0784, 2)
        fixed = 41.30
        total = round((fixed + variable) * 1.21, 2)
        fixed_tax, variable_tax = round(fixed * 1.21, 2), round(total - fixed * 1.21, 2)
        yield dict(
            tipo_suministro="GAS", num_factura=f"FE2639{number:010d}",
            fecha_factura=(end + timedelta(days=5)).isoformat(),
            fecha_inicio=start.isoformat(), fecha_fin=end.isoformat(),
            consumo_kwh=str(kwh), termino_fijo=f"{fixed_tax:.2f}",
            termino_variable=f"{variable_tax:.2f}", importe_total=f"{total:.2f}",
        )


def _electricity_invoices():
    for number, (start, end) in enumerate(_invoice_months(), start=1):
        kwh = round(980 + 60 * (start.month in (12, 1, 2)), 3)
        fixed, variable = 20.82, round(kwh * 0.192, 2)
        yield dict(
            tipo_suministro="ELECTRICIDAD", num_factura=f"P26CON{number:09d}",
            fecha_factura=(end + timedelta(days=4)).isoformat(),
            fecha_inicio=start.isoformat(), fecha_fin=end.isoformat(),
            consumo_kwh=str(kwh), termino_fijo=f"{fixed:.2f}",
            termino_variable=f"{variable:.2f}", importe_total=f"{fixed + variable:.2f}",
        )


def _water_invoices():
    months = list(_invoice_months())
    for number, index in enumerate(range(0, 12, 2), start=1):
        start, end = months[index][0], months[index + 1][1]
        m3 = 70 + 9 * number
        fixed, variable = 98.40, round(m3 * 1.61, 2)
        yield dict(
            tipo_suministro="AGUA", num_factura=f"AB{number:05d}2600",
            fecha_factura=(end + timedelta(days=20)).isoformat(),
            fecha_inicio=start.isoformat(), fecha_fin=end.isoformat(),
            consumo_m3=str(m3), termino_fijo=f"{fixed:.2f}",
            termino_variable=f"{variable:.2f}", importe_total=f"{fixed + variable:.2f}",
        )


# ----------------------------------------------------------------------------
# Fuentes para «Añadir fuentes» (903): formatos que el catálogo reconoce.
# ----------------------------------------------------------------------------

def _naturgy_pdf(path: Path, invoice: dict, cups: str) -> Path:
    start, end = (date.fromisoformat(invoice[k]) for k in ("fecha_inicio", "fecha_fin"))
    return _write_pdf(path, [
        "Naturgy Clientes S.A.U. - Mercado Libre (DOCUMENTO DE PRUEBA)",
        "Hola, aquí tienes tu factura de gas.",
        f"Nº factura: {invoice['num_factura']}",
        f"Fecha de emisión: {start.strftime('%d.%m.%Y') if False else _es(date.fromisoformat(invoice['fecha_factura']))}",
        f"Periodo gas: del {_es(start)} al {_es(end)}",
        f"CUPS: {cups}",
        f"Consumo gas {int(float(invoice['consumo_kwh'])):,} kWh".replace(",", "."),
        f"Término fijo {_euros(float(invoice['termino_fijo']))} €",
        f"Término variable {_euros(float(invoice['termino_variable']))} €",
        f"Total a pagar {_euros(float(invoice['importe_total']))} €",
    ])


def _endesa_pdf(path: Path, invoice: dict, cups: str) -> Path:
    start, end = (date.fromisoformat(invoice[k]) for k in ("fecha_inicio", "fecha_fin"))
    days = (end - start).days
    return _write_pdf(path, [
        "Endesa Energía S.A.U. CIF A81948077 (DOCUMENTO DE PRUEBA)",
        "DATOS DE LA FACTURA",
        f"Nº factura: {invoice['num_factura']}",
        f"Fecha emisión factura: {_es(date.fromisoformat(invoice['fecha_factura']))}",
        f"Periodo de facturación: del {_es(start)} a {_es(end)} ({days} días)",
        f"CUPS: {cups}",
        f"Consumo Total {_euros(float(invoice['consumo_kwh']))} kWh",
        f"Potencia {_euros(float(invoice['termino_fijo']))} €",
        f"Energía {_euros(float(invoice['termino_variable']))} €",
        f"Total {_euros(float(invoice['importe_total']))} €",
    ])


def _maintenance_pdf(path: Path, number: str, day: date, amount: float | None,
                     concept: str = "Mantenimiento preventivo mensual de la sala de calderas") -> Path:
    lines = [
        "Factura",
        f"Número # {number}",
        f"Fecha {_es(day)}",
        "MANTENIMIENTOS (DOCUMENTO DE PRUEBA)",
        "INSTALACIONES ZARAGOZA S.L.",
        concept,
        "BASE IMPONIBLE   IMPUESTO   TOTAL IMPUESTO   TOTAL",
    ]
    if amount is None:
        lines.append("(importe ilegible en la copia escaneada)")
    else:
        base = round(amount / 1.21, 2)
        lines += [
            f"{_euros(base)}€   IVA 21%   {_euros(round(amount - base, 2))}€   {_euros(amount)}€",
            f"{number} - {_euros(amount)}€ Vencimiento {_es(day + timedelta(days=30))}",
        ]
    return _write_pdf(path, lines)


# Proveedor de agua inventado: se declara en la capa del despacho de la
# instalación de prueba (config/proveedores_despacho.json), igual que el
# despacho añadiría uno propio.
TEST_WATER_PROVIDER = {
    "AGUAS_DE_PRUEBA": {
        "nombre_display": "Aguas de Prueba S.A. (proveedor inventado)",
        "tipo_suministro": "AGUA",
        "unidad_consumo": "m3",
        "firmas_identificacion": ["AGUAS\\s+DE\\s+PRUEBA"],
        "regex": {
            "num_factura": "Factura\\s+n[ºo°]\\s*(?P<valor>[A-Z0-9]+)",
            "fecha_factura": "Fecha\\s+de\\s+emisi[oó]n:\\s*(?P<valor>\\d{2}/\\d{2}/\\d{4})",
            "fecha_inicio": "Periodo:\\s*del\\s*(?P<valor>\\d{2}/\\d{2}/\\d{4})",
            "fecha_fin": "Periodo:\\s*del\\s*\\d{2}/\\d{2}/\\d{4}\\s*al\\s*(?P<valor>\\d{2}/\\d{2}/\\d{4})",
            "consumo_m3": "Consumo:\\s*(?P<valor>\\d+)\\s*m3",
            "termino_fijo": "Cuota\\s+fija:\\s*(?P<valor>[\\d.,]+)\\s*€",
            "termino_variable": "Cuota\\s+variable:\\s*(?P<valor>[\\d.,]+)\\s*€",
            "importe_total": "TOTAL\\s+A\\s+PAGAR:\\s*(?P<valor>[\\d.,]+)\\s*€",
        },
    }
}


def _water_pdf(path: Path, invoice: dict, provider: str = "AGUAS DE PRUEBA S.A.") -> Path:
    start, end = (date.fromisoformat(invoice[k]) for k in ("fecha_inicio", "fecha_fin"))
    return _write_pdf(path, [
        f"{provider} (DOCUMENTO DE PRUEBA)",
        f"Factura nº {invoice['num_factura']}",
        f"Fecha de emisión: {_es(date.fromisoformat(invoice['fecha_factura']))}",
        "Suministro de agua",
        f"Periodo: del {_es(start)} al {_es(end)}",
        f"Consumo: {invoice['consumo_m3']} m3",
        f"Cuota fija: {_euros(float(invoice['termino_fijo']))} €",
        f"Cuota variable: {_euros(float(invoice['termino_variable']))} €",
        f"TOTAL A PAGAR: {_euros(float(invoice['importe_total']))} €",
    ])


def _meter_service_pdf(path: Path, number: str, day: date, amount: float) -> Path:
    month = ("ENE", "FEB", "MAR", "ABR", "MAY", "JUN", "JUL", "AGO", "SEP", "OCT", "NOV", "DIC")[day.month - 1]
    base = round(amount / 1.21, 2)
    return _write_pdf(path, [
        "FACTURA   GOMEZ GROUP METERING, S.L. (DOCUMENTO DE PRUEBA)",
        "Nº FACTURA F.FACTURA F. VCTO. RAZÓN SOCIAL",
        f"{number} {_es(day)} {_es(day + timedelta(days=4))}",
        "COMUNIDAD DE PROPIETARIOS DE PRUEBA 903",
        "SERVICIO DE LECTURA, FACTURACIÓN Y MANTENIMIENTO",
        f"DE CONTADORES. LECTURAS FACTURADAS: {day.year} {month}",
        f"BASE IMPONIBLE ... {_euros(base)} €",
        f"TOTAL FACTURA ... {_euros(amount)} €",
    ])


# ----------------------------------------------------------------------------
# Construcción de la instalación de prueba
# ----------------------------------------------------------------------------

def _prepare_destination(destination: Path, replace: bool) -> None:
    if destination.exists() and any(destination.iterdir()):
        if not replace:
            raise SystemExit(
                f"La carpeta {destination} ya existe y no está vacía. "
                "Usa --reemplazar para volver a crearla."
            )
        if not (destination / MARKER).is_file():
            raise SystemExit(
                f"{destination} no es una carpeta creada por este script; no se borra."
            )
        shutil.rmtree(destination)
    destination.mkdir(parents=True, exist_ok=True)
    for name in RESOURCES:
        shutil.copytree(PROJECT_ROOT / name, destination / name, dirs_exist_ok=True)
    # Sin perfiles ni plantillas privadas: cada comunidad de prueba parte del
    # modelo común, como una comunidad nueva del despacho.
    shutil.rmtree(destination / "config" / "excel_profiles", ignore_errors=True)
    shutil.rmtree(destination / "plantillas" / "comunidades", ignore_errors=True)
    for private in ("proveedores_despacho.json", "letter_identities.json"):
        (destination / "config" / private).unlink(missing_ok=True)
    (destination / "config" / "proveedores_despacho.json").write_text(
        json.dumps({"_descripcion": "Capa de la instalación de prueba.", "proveedores": TEST_WATER_PROVIDER},
                   ensure_ascii=False, indent=2) + "\n", encoding="utf-8",
    )
    (destination / MARKER).write_text(
        "Instalación de prueba creada por scripts/crear_comunidades_prueba.py\n", encoding="utf-8",
    )


def _service_fees(connection, community_id: int, period_id: int) -> None:
    import cuotas_servicio
    cuotas_servicio.generar_cuotas_mensuales(
        connection, community_id=community_id, period_id=period_id, servicio="ACS",
        tramos=[("2025-08", "220"), ("2026-01", "200")],
    )
    for day, amount, consumption in (
        # Las liquidaciones suman los 132 m³ que marcan los contadores.
        ("2025-11-05", "2400", "40"), ("2026-03-09", "3100", "52"), ("2026-07-10", "2300", "40"),
    ):
        cuotas_servicio.registrar_cuota(
            connection, community_id=community_id, period_id=period_id, servicio="ACS",
            concepto="variable", fecha=day, importe=amount, consumo=consumption,
        )
    connection.commit()


def _fixed_costs(connection, community_id: int, period_id: int) -> None:
    import fixed_costs
    # Sólo se confirma que no hay lecturas de calefacción ni placas solares. El
    # mantenimiento de calderas y las lecturas de ACS salen de las facturas.
    fixed_costs.save_fixed_costs(connection, community_id, period_id, {
        "fixed_cost_meter_reading_heating": "",
        "fixed_cost_solar_maintenance": "",
    })


def _create_case(connection, code: str, name: str):
    import expedient_service
    import gestor_bd

    community_id = gestor_bd.obtener_o_crear_comunidad(connection, code, name)
    connection.execute(
        "UPDATE comunidades SET cif=?, direccion=? WHERE id_comunidad=?",
        (f"H99{code}000", f"Calle de Prueba {code}, Zaragoza", community_id),
    )
    connection.commit()
    case = expedient_service.create_case(
        connection, community_id, name="2025-2026", start_date=PERIOD_START, end_date=PERIOD_END,
    )
    return community_id, case


def _write_sources(folder: Path, code: str, *, defects: bool) -> list[Path]:
    """Documentos tal como llegan al despacho; con ``defects`` los problemas típicos."""
    from openpyxl import Workbook

    folder.mkdir(parents=True, exist_ok=True)
    dwellings = _dwellings(code)
    with (folder / f"{code} Listado de propietarios.csv").open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.writer(handle, delimiter=";")
        writer.writerow(["Fdenominacion", "Nombre", "Coeficiente", "Email"])
        for item in dwellings:
            writer.writerow([item.code, item.name, item.coefficient, _email(item.name)])
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Lecturas"
    sheet.append([f"C.P. DE PRUEBA {code} - LECTURAS DE CONSUMO (DOCUMENTO DE PRUEBA)"])
    sheet.append([])
    sheet.append(["Cod.", "Propiedad", "Nombre", "8/2025", "7/2026", "CONS. ACS"])
    for index, item in enumerate(dwellings, start=1):
        initial, final = item.initial or 0, item.final or 0
        if defects and index == 4:
            final = initial - 40          # contador que baja: estimación a aprobar
        if defects and index == 6:
            sheet.append([index, item.code, item.name, initial, None, None])  # falta la final
            continue
        sheet.append([index, item.code, item.name, initial, final, final - initial])
    workbook.save(folder / f"{code} Lecturas consumo 08 2025 a 07 2026.xlsx")
    for invoice in _gas_invoices():
        _naturgy_pdf(folder / f"{code} NATURGY {invoice['num_factura']}.pdf", invoice,
                     f"ES00999{code}GA0000000001")
    for invoice in _electricity_invoices():
        _endesa_pdf(folder / f"{code} ENDESA {invoice['num_factura']}.pdf", invoice,
                    f"ES00999{code}EL0000000001")
    for invoice in _water_invoices():
        _water_pdf(folder / f"{code} AGUA {invoice['num_factura']}.pdf", invoice)
    for index, (start, _end) in enumerate(_months(), start=1):
        _maintenance_pdf(folder / f"{code} MANTENIMIENTO F26{index:05d}.pdf",
                         f"F26{index:05d}", start + timedelta(days=24), 123.35)
        _meter_service_pdf(folder / f"{code} LECTURAS CONTADORES LF26{index:06d}.pdf",
                           f"LF26{index:06d}", start + timedelta(days=1), 72.60)
    if defects:
        previous = dict(next(_gas_invoices()), num_factura="FE2539000000099",
                        fecha_factura="2025-07-04", fecha_inicio="2025-06-01", fecha_fin="2025-06-30")
        _naturgy_pdf(folder / f"{code} NATURGY {previous['num_factura']} (ejercicio anterior).pdf",
                     previous, f"ES00999{code}GA0000000001")
        _maintenance_pdf(folder / f"{code} REPARACION F26REP01.pdf", "F26REP01", date(2026, 2, 12), None,
                         concept="Sustitución de válvula en la sala de calderas")
        unknown = dict(next(_water_invoices()), num_factura="XX0001")
        _water_pdf(folder / f"{code} AGUA proveedor nuevo.pdf", unknown, provider="CANAL DESCONOCIDO S.L.")
        duplicate = folder / f"{code} ENDESA copia reenviada.pdf"
        shutil.copy2(next(folder.glob(f"{code} ENDESA P26CON*.pdf")), duplicate)
    return sorted(path for path in folder.iterdir() if path.is_file())


def _intake(home: Path, case_id: int, code: str, paths: list[Path]) -> None:
    """Lo mismo que hace «Añadir fuentes» en la aplicación."""
    import case_ingestion
    import community_discovery
    import gestor_bd
    import source_batch
    from app_paths import ApplicationPaths

    database = ApplicationPaths.resolve().database
    lookup = gestor_bd.conectar(str(database))
    try:
        accepted, _foreign = community_discovery.partition_sources_for_community(
            paths, code, connection=lookup)
        lookup.commit()
    finally:
        lookup.close()
    for item in source_batch.analyse_batch(accepted, community_code=code, database_path=database):
        if item.analysis is None:
            continue
        connection = gestor_bd.conectar(str(database))
        try:
            case_ingestion.add_analysed_document_to_case(
                connection, case_id, source_path=str(item.path),
                archive_root=ApplicationPaths.resolve().sources, analysis=item.analysis,
            )
        finally:
            connection.close()
    connection = gestor_bd.conectar(str(database))
    try:
        case_ingestion.finalize_case_source_intake(connection, case_id)
    finally:
        connection.close()


def _confirm_everything(home: Path, case_id: int) -> None:
    """Confirma lo que el análisis deja pendiente de un vistazo humano (901)."""
    import case_ingestion
    import document_review
    import gestor_bd
    from app_paths import ApplicationPaths

    connection = gestor_bd.conectar(str(ApplicationPaths.resolve().database))
    try:
        for (document_id,) in connection.execute(
            """SELECT id_document FROM source_documents
               WHERE id_case=? AND status NOT IN ('validated','not_applicable')""",
            (case_id,),
        ).fetchall():
            for issue in document_review.list_open_issues(connection, case_id):
                if issue.id_document == document_id and issue.code == "READING_FORMAT_REVIEW":
                    document_review.resolve_issue(
                        connection, issue.id_issue, value="confirmado",
                        reason="Formato revisado (instalación de prueba)", resolved_by="prueba",
                    )
            case_ingestion.confirm_source_candidates(
                connection, case_id, document_id, confirmed_by="prueba",
            )
        document_review.validate_case_ready(connection, case_id, project_root=home)
    finally:
        connection.close()


def build(destination: Path, *, replace: bool = False) -> dict:
    destination = destination.resolve()
    _prepare_destination(destination, replace)
    os.environ["REGULARIZACIONES_HOME"] = str(destination)
    if str(CORE) not in sys.path:
        sys.path.insert(0, str(CORE))
    import expedient_service
    import gestor_bd
    import office_settings
    from app_paths import ApplicationPaths

    paths = ApplicationPaths.resolve()
    paths.prepare()
    with redirect_stdout(StringIO()):
        gestor_bd.crear_bd(str(paths.database))
    connection = gestor_bd.conectar(str(paths.database))
    try:
        office_settings.save_office_settings(connection, office_settings.OfficeSettings(
            name="PRUEBA · datos inventados", city="Zaragoza", address="Paseo de Prueba 1, Zaragoza",
            email="despacho@example.invalid", signature="El administrador (prueba)",
            fiscal_start_month=8,
        ))
        _community_901, case_901 = _create_case(connection, "901", "CP PRUEBA 901 · LISTA PARA CALCULAR")
        community_902, case_902 = _create_case(connection, "902", "CP PRUEBA 902 · CON INCIDENCIAS")
        _community_903, case_903 = _create_case(connection, "903", "CP PRUEBA 903 · BANDEJA DE FUENTES")
        for case in (case_901, case_902, case_903):
            period_id = expedient_service.link_case_to_period(connection, case.id_case)
            connection.commit()
            community = expedient_service.get_case(connection, case.id_case).community_id
            _service_fees(connection, community, period_id)
            _fixed_costs(connection, community, period_id)
    finally:
        connection.close()
    # 901: documentos limpios añadidos y confirmados; listo para calcular.
    sources_901 = destination / "documentos_recibidos_901"
    _intake(destination, case_901.id_case, "901", _write_sources(sources_901, "901", defects=False))
    shutil.rmtree(sources_901, ignore_errors=True)
    _confirm_everything(destination, case_901.id_case)
    # 902: sus documentos ya han pasado por «Añadir fuentes» y esperan revisión.
    sources_902 = destination / "documentos_recibidos_902"
    _intake(destination, case_902.id_case, "902", _write_sources(sources_902, "902", defects=True))
    shutil.rmtree(sources_902, ignore_errors=True)
    # 903: los documentos se dejan en una carpeta para añadirlos desde la aplicación.
    folder = destination / "fuentes_903"
    _write_sources(folder, "903", defects=False)
    connection = gestor_bd.conectar(str(paths.database))
    try:
        ready = expedient_service.get_case(connection, case_901.id_case)
        issues = expedient_service.get_case(connection, case_902.id_case)
        tray = expedient_service.get_case(connection, case_903.id_case)
        open_issues = connection.execute(
            "SELECT COUNT(*) FROM review_issues WHERE id_case=? AND status='open'", (issues.id_case,),
        ).fetchone()[0]
    finally:
        connection.close()
    return {
        "destino": destination,
        "901": ready.status,
        "902": f"{issues.status} ({open_issues} incidencias)",
        "903": f"{tray.status}; fuentes en {folder}",
    }


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    default = Path(os.environ.get("LOCALAPPDATA") or Path.home()) / "Regularizaciones-prueba"
    parser.add_argument("--destino", type=Path, default=default,
                        help=f"carpeta de la instalación de prueba (por defecto {default})")
    parser.add_argument("--reemplazar", action="store_true",
                        help="borra y vuelve a crear una instalación de prueba anterior")
    args = parser.parse_args(argv)
    result = build(args.destino, replace=args.reemplazar)
    print(f"Instalación de prueba creada en {result['destino']}")
    print(f"  901 · lista para calcular: {result['901']}")
    print(f"  902 · con incidencias:     {result['902']}")
    print(f"  903 · bandeja de fuentes:  {result['903']}")
    print("Ábrela con Regularizaciones_prueba.bat")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
