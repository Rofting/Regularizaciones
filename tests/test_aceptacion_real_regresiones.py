"""Regresiones encontradas al pasar por la aplicación el expediente real de una comunidad.

Todas usan libros y listados sintéticos que reproducen la forma de los
documentos reales que bloqueaban el flujo:

* El maestro de un año anterior trae la «Fecha fra.» arrastrada del ejercicio
  previo. Corregirla en la incidencia debe importar la factura (antes sólo se
  guardaba el valor y el Excel se bloqueaba por «faltan facturas»).
* Excluir una fila del maestro no debe invalidar el libro entero.
* Las filas de previsión sin importes no son facturas.
* Un maestro con las columnas desplazadas respecto del perfil.
* El desglose fijo/variable sin impuestos debe sumar el total en el Excel.
* Locales y garajes sin contador no bloquean por falta de lecturas.
* Una lectura que falta se completa desde la propia incidencia.
"""

import csv
import sqlite3
import sys
import tempfile
import unittest
from contextlib import redirect_stdout
from datetime import date
from io import StringIO
from pathlib import Path

from openpyxl import Workbook, load_workbook

PROJECT_ROOT = Path(__file__).resolve().parents[1]
CORE_DIR = PROJECT_ROOT / "core"
if str(CORE_DIR) not in sys.path:
    sys.path.insert(0, str(CORE_DIR))

import document_review
import excel_bootstrap_importer
import excel_profiles
import expedient_service
import gestor_bd
import owner_lists
from excel_bootstrap_importer import import_companion_sources, import_master_excel
from excel_export_service import invoice_components_with_taxes
from tests.test_excel_bootstrap_importer import _make_master, _make_readings


class _MasterCase(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory(ignore_cleanup_errors=True)
        self.root = Path(self.directory.name).resolve()
        self.database_path = self.root / "gestion.db"
        with redirect_stdout(StringIO()):
            gestor_bd.crear_bd(str(self.database_path))
        self.connection = sqlite3.connect(self.database_path)
        self.connection.row_factory = sqlite3.Row
        self.connection.execute("PRAGMA foreign_keys = ON")
        self.community_id = gestor_bd.obtener_o_crear_comunidad(
            self.connection, "658", "Comunidad sintética 658"
        )
        self.connection.commit()
        self.case = expedient_service.create_case(
            self.connection, self.community_id, name="2025-2026",
            start_date=date(2025, 9, 1), end_date=date(2026, 8, 31),
        )
        self.profile = excel_profiles.load_profile("658_acs_v1", PROJECT_ROOT)

    def tearDown(self):
        self.connection.close()
        self.directory.cleanup()

    def _import(self, workbook_path: Path):
        return import_master_excel(
            self.connection, id_case=self.case.id_case,
            workbook_path=workbook_path, profile=self.profile, actor="Prueba",
        )

    def _issue(self, field_name: str):
        return self.connection.execute(
            "SELECT * FROM review_issues WHERE field_name=? AND status='open'",
            (field_name,),
        ).fetchone()

    def _gas_invoices(self) -> int:
        return self.connection.execute(
            "SELECT COUNT(*) FROM facturas WHERE tipo_suministro='GAS'"
        ).fetchone()[0]


class MasterCorrectionLoopTest(_MasterCase):
    def test_corrected_stale_invoice_date_imports_the_invoice(self):
        workbook_path = _make_master(self.root / "fecha_arrastrada.xlsx")
        workbook = load_workbook(workbook_path)
        workbook["GAS"]["B10"] = date(2024, 10, 5)  # la del ejercicio anterior
        workbook.save(workbook_path)
        self._import(workbook_path)
        issue = self._issue("GAS.B10.invoice_date")
        self.assertEqual("INVOICE_OUTSIDE_PERIOD", issue["code"])
        self.assertEqual(0, self._gas_invoices())

        document_review.resolve_issue(
            self.connection, issue["id_issue"], value="05/10/2025",
            reason="Fecha arrastrada del año anterior", resolved_by="Prueba",
        )

        self.assertEqual(1, self._gas_invoices())
        self.assertEqual("2025-10-05", self.connection.execute(
            "SELECT fecha_factura FROM facturas WHERE tipo_suministro='GAS'"
        ).fetchone()[0])
        self.assertEqual("validated", self.connection.execute(
            "SELECT status FROM source_documents WHERE id_document=?", (issue["id_document"],)
        ).fetchone()[0])

    def test_master_cell_corrections_are_validated_as_dates(self):
        workbook_path = _make_master(self.root / "fecha_texto.xlsx")
        workbook = load_workbook(workbook_path)
        workbook["GAS"]["B10"] = "?"
        workbook.save(workbook_path)
        self._import(workbook_path)
        issue = self._issue("GAS.B10.invoice_date")
        with self.assertRaises(ValueError):
            document_review.resolve_issue(
                self.connection, issue["id_issue"], value="pendiente",
                reason="Prueba", resolved_by="Prueba",
            )

    def test_excluding_one_row_keeps_the_rest_of_the_master_valid(self):
        workbook_path = _make_master(self.root / "prevision.xlsx")
        workbook = load_workbook(workbook_path)
        workbook["GAS"]["B10"] = "sin recibir aún"
        workbook.save(workbook_path)
        self._import(workbook_path)
        issue = self._issue("GAS.B10.invoice_date")

        document_review.exclude_master_invoice_row(
            self.connection, issue["id_issue"], reason="Previsión", excluded_by="Prueba",
        )

        status = self.connection.execute(
            "SELECT status FROM source_documents WHERE id_document=?", (issue["id_document"],)
        ).fetchone()[0]
        self.assertEqual("validated", status)
        self.assertEqual(0, self._gas_invoices())
        # Las demás hojas del mismo maestro siguen importadas.
        self.assertEqual(1, self.connection.execute(
            "SELECT COUNT(*) FROM facturas WHERE tipo_suministro='ELECTRICIDAD'"
        ).fetchone()[0])

    def test_dismissing_an_outside_period_row_does_not_discard_the_master(self):
        workbook_path = _make_master(self.root / "otro_ejercicio.xlsx")
        workbook = load_workbook(workbook_path)
        workbook["GAS"]["B10"] = date(2027, 1, 1)
        workbook.save(workbook_path)
        self._import(workbook_path)
        issue = self._issue("GAS.B10.invoice_date")

        document_review.dismiss_invoice_outside_period(
            self.connection, issue["id_issue"], reason="Otro ejercicio", dismissed_by="Prueba",
        )

        self.assertEqual("validated", self.connection.execute(
            "SELECT status FROM source_documents WHERE id_document=?", (issue["id_document"],)
        ).fetchone()[0])
        self.assertEqual(1, self.connection.execute(
            "SELECT COUNT(*) FROM facturas WHERE tipo_suministro='ELECTRICIDAD' AND id_periodo IS NOT NULL"
        ).fetchone()[0])

    def test_forecast_row_without_amounts_is_not_an_invoice(self):
        workbook_path = _make_master(self.root / "fila_preparada.xlsx")
        workbook = load_workbook(workbook_path)
        gas = workbook["GAS"]
        gas["B11"] = date(2026, 7, 28)  # fecha preparada, sin importes
        gas["L11"] = 0
        gas["B12"] = "Suma 2025-2026"
        workbook.save(workbook_path)

        self._import(workbook_path)

        self.assertIsNone(self.connection.execute(
            "SELECT 1 FROM review_issues WHERE field_name LIKE 'GAS.%11.%'"
        ).fetchone())


class ShiftedLayoutTest(unittest.TestCase):
    def test_columns_are_realigned_when_the_master_has_one_column_less(self):
        profile = excel_profiles.load_profile("658_acs_v1", PROJECT_ROOT)
        layout = excel_bootstrap_importer._layout_invoice_columns(profile, "GAS")
        workbook = Workbook()
        sheet = workbook.active
        headers = {
            "B": "Fecha fra.", "C": "Días", "D": "Fecha final", "E": "Lectura final",
            "F": "Fecha inicial", "G": "Lectura inicial", "H": "Consumo (m3)",
            "I": "Consumo (kWh)", "J": "Término fijo (€ IVA incl.)",
            "K": "Término variable (€ IVA incl.)", "L": "Importe TOTAL (€ IVA incluido)",
            "M": "Fecha pago", "N": "Nota", "O": "Precio m3 (IVA incl.)", "X": "% Tér. Fijo", "Y": "% Tér. Variable",
        }
        for column, label in headers.items():
            sheet[f"{column}8"] = label

        _start, _end, columns = excel_bootstrap_importer._realign_layout_with_headers(sheet, layout)

        self.assertEqual(
            {"date": 2, "end": 4, "start": 6, "fixed": 10, "variable": 11, "total": 12, "notes": 14},
            {key: columns[key] for key in ("date", "end", "start", "fixed", "variable", "total", "notes")},
        )
        self.assertEqual((9,), tuple(columns["consumption"]))

    def test_matching_layout_is_left_untouched(self):
        profile = excel_profiles.load_profile("658_acs_v1", PROJECT_ROOT)
        layout = excel_bootstrap_importer._layout_invoice_columns(profile, "GAS")
        workbook = Workbook()
        sheet = workbook.active  # sin rótulos: se conserva el perfil
        self.assertEqual(layout, excel_bootstrap_importer._realign_layout_with_headers(sheet, layout))


class InvoiceComponentsTest(unittest.TestCase):
    def test_taxes_are_shared_so_fixed_plus_variable_equals_the_total(self):
        fixed, variable = invoice_components_with_taxes(15.42, 151.14, 214.04)
        self.assertAlmostEqual(214.04, fixed + variable, places=2)
        self.assertAlmostEqual(19.82, fixed, places=2)

    def test_components_that_already_add_up_are_kept(self):
        self.assertEqual((20.0, 80.0), invoice_components_with_taxes(20, 80, 100))

    def test_extra_decimals_are_reconciled_to_the_cent(self):
        fixed, variable = invoice_components_with_taxes(20.8241, 193.2128, 214.0369)
        self.assertEqual(21404, round((fixed + variable) * 100))

    def test_credit_notes_keep_their_sign(self):
        fixed, variable = invoice_components_with_taxes(-180.65, -1086.07, -1532.73)
        self.assertLess(fixed, 0)
        self.assertAlmostEqual(-1532.73, fixed + variable, places=2)

    def test_without_breakdown_nothing_is_invented(self):
        self.assertEqual((0.0, 0.0), invoice_components_with_taxes(None, None, 123.35))


class OwnerListTest(unittest.TestCase):
    def test_email_on_the_next_line_under_the_property_column_is_kept(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "listado.xlsx"
            workbook = Workbook()
            sheet = workbook.active
            sheet.append(["Código", None, "Propiedad", None, "Nombre", None, None, "NIF", None, None, "Coeficiente"])
            sheet.append(["000001", None, "PA2-1ºA", None, "PERSONA UNO", None, None, None, None, None, "1,46"])
            sheet.append([None, None, "e-mail : uno@example.invalid", None, None])
            sheet.append(["000002", None, "PA2-1ºB", None, "PERSONA DOS", None, None, None, None, None, "1,52"])
            workbook.save(path)

            rows = owner_lists.read_owner_rows(path)

        self.assertEqual(["PA2-1ºA", "PA2-1ºB"], [row.vivienda for row in rows])
        self.assertEqual("uno@example.invalid", rows[0].email)
        self.assertEqual("", rows[1].email)


class CompanionReadingsTest(_MasterCase):
    def _owners(self, path: Path, rows) -> Path:
        with path.open("w", encoding="utf-8", newline="") as handle:
            writer = csv.writer(handle, delimiter=";")
            writer.writerow(["Fdenominacion", "Nombre", "Coeficiente", "Email"])
            writer.writerows(rows)
        return path

    def test_premises_without_meter_do_not_block_the_case(self):
        owners = self._owners(self.root / "propietarios.csv", [
            ("P1-A", "Persona Uno", "50,00", ""),
            ("P1-B", "Persona Dos", "50,00", ""),
            ("LOC A", "Comercio", "3,00", ""),
            ("GAR.", "Garajes", "4,00", ""),
        ])
        readings = _make_readings(self.root / "lecturas.xlsx")
        self._import(_make_master(self.root / "maestro.xlsx"))

        import_companion_sources(
            self.connection, id_case=self.case.id_case, owner_list_path=owners,
            readings_path=readings, profile=self.profile, actor="Prueba",
        )

        codes = {
            row["field_name"] for row in self.connection.execute(
                "SELECT field_name FROM review_issues WHERE code='MISSING_READING_RANGE'"
            )
        }
        self.assertFalse(any("LOC" in code or "GAR" in code for code in codes), codes)

    def test_missing_reading_is_completed_from_the_issue(self):
        owners = self._owners(self.root / "propietarios.csv", [
            ("P1-A", "Persona Uno", "50,00", ""),
            ("P1-B", "Persona Dos", "50,00", ""),
        ])
        readings = _make_readings(self.root / "lecturas.xlsx", include_second=False)
        self._import(_make_master(self.root / "maestro.xlsx"))
        import_companion_sources(
            self.connection, id_case=self.case.id_case, owner_list_path=owners,
            readings_path=readings, profile=self.profile, actor="Prueba",
        )
        issue = self._issue("reading.P1-B.ACS")
        self.assertIsNotNone(issue)

        with self.assertRaises(ValueError):
            document_review.resolve_issue(
                self.connection, issue["id_issue"], value="210",
                reason="Lectura comprobada", resolved_by="Prueba",
            )
        document_review.resolve_issue(
            self.connection, issue["id_issue"], value="200;214",
            reason="Lecturas comprobadas en el contador", resolved_by="Prueba",
        )

        values = [
            (row["fecha_lectura"], row["valor_acumulado"], row["estado"], row["approved_by"])
            for row in self.connection.execute(
                """SELECT l.fecha_lectura,l.valor_acumulado,l.estado,l.approved_by
                   FROM lecturas_vecino l JOIN propietarios p USING(id_propietario)
                   WHERE p.codigo_vivienda='P1-B' ORDER BY l.fecha_lectura"""
            )
        ]
        self.assertEqual([
            ("2025-09-01", 200.0, "estimado", "Prueba"),
            ("2026-08-31", 214.0, "estimado", "Prueba"),
        ], values)
        self.assertIsNone(self._issue("reading.P1-B.ACS"))


class TotalEnergiesBreakdownTest(unittest.TestCase):
    """El desglose fijo/variable de TotalEnergies suma el total y no bloquea el Excel."""

    TEXT = """Factura gas
Importe factura gas: {total} €
CDAD. PROP. PRUEBA
Nº Factura: FGAS2600099999
FECHA FACTURA: 28 de abril de 2026
Periodo de facturación: De 31/03/2026 al 23/04/2026
Término Fijo {fixed} €
Término Variable {variable} €
Cuota GTS y Tasa CNMC {gts} €
IEH {ieh} €
Alquiler {rental} €
Otros {tolls} €
Impuesto Aplicado {iva} €
TOTAL IMPORTE FACTURA {total} €
Base imponible 1.100,00 €
Impuesto IVA 10,00 % sobre 1.100,00 €
{iva} €
TOTAL FACTURA {total} €
"""

    def _extract(self, sign=""):
        import json
        import lector_pdf
        config = json.loads((PROJECT_ROOT / "config" / "proveedores.json").read_text(
            encoding="utf-8"))["proveedores"]["TOTALENERGIES_GAS"]
        text = self.TEXT.format(
            total=f"{sign}1.210,00", iva=f"{sign}110,00", fixed=f"{sign}100,00",
            variable=f"{sign}500,00", gts=f"{sign}10,00", ieh=f"{sign}40,00",
            rental=f"{sign}10,00", tolls=f"{sign}300,00",
        )
        return lector_pdf.extraer_datos_totalenergies(text, config)

    def test_fixed_includes_tolls_gts_and_rental_and_sums_to_the_total(self):
        data = self._extract()
        # (100 + 300 + 10 + 10) = 420 antes de IVA; con el 10 % de IVA, 462.
        self.assertEqual(462.00, data["termino_fijo"])
        self.assertEqual(748.00, data["termino_variable"])
        self.assertEqual(1210.00, round(data["termino_fijo"] + data["termino_variable"], 2))

    def test_tax_only_shown_as_impuesto_aplicado(self):
        import json
        import lector_pdf
        config = json.loads((PROJECT_ROOT / "config" / "proveedores.json").read_text(
            encoding="utf-8"))["proveedores"]["TOTALENERGIES_GAS"]
        text = (
            "Nº Factura: FGAS2600099998\nFECHA FACTURA: 28 de abril de 2026\n"
            "Periodo de facturación: De 31/03/2026 al 23/04/2026\nFactura resumen\n"
            "Término Fijo 180,65 €\nTérmino Variable 4968,48 €\nCuota GTS y Tasa CNMC 18,90 €\n"
            "IEH 91,46 € Término Fijo\nAlquiler 7,43 € Término Variable\nOtros 1086,07 € Cuota GTS\n"
            "Impuesto Aplicado 635,30 €\nTOTAL IMPORTE FACTURA 6.988,29 €\n"
        )
        data = lector_pdf.extraer_datos_totalenergies(text, config)
        self.assertEqual(1422.36, data["termino_fijo"])
        self.assertEqual(5565.93, data["termino_variable"])

    def test_credit_note_keeps_the_same_proportion_with_negative_signs(self):
        data = self._extract(sign="-")
        self.assertEqual(-462.00, data["termino_fijo"])
        self.assertEqual(-748.00, data["termino_variable"])




class InvoiceTotalAmongLinesTest(unittest.TestCase):
    """El total de la factura no se confunde con el de una línea de detalle."""

    TEXT = (
        "Factura\nNúmero # F26260669\nFecha 07/05/2026\n"
        "CONCEPTO DESCRIPCIÓN UNIDADES PRECIO SUBTOTAL IVA TOTAL\n"
        'Rosca Doble latón 1/2" 2 1,40€ 2,80€ 21% 3,39€\n'
        'Válvula Tuller MAriposa 1/2" 2 10,69€ 21,38€ 21% 25,87€\n'
        "Mano de Obra Técnico (horas) 5 23,00€ 115,00€ 21% 139,15€\n"
        "BASE IMPONIBLE IMPUESTO TOTAL IMPUESTO TOTAL\n"
        "148,71€ IVA 21% 31,23€ 179,94€\n148,71€ 179,94€\n179,94€\n"
        "F26260669 - 179,94€ Vencimiento 06/06/2026 Pág. 1 de 1\n"
    )

    def test_largest_reconciling_total_wins(self):
        import invoice_extractors
        fields, diagnostics = {}, []
        invoice_extractors._extract_amounts(self.TEXT, "periodic_maintenance", fields, diagnostics)
        self.assertEqual("179.94", fields["importe_total"].value)


if __name__ == "__main__":
    unittest.main()
