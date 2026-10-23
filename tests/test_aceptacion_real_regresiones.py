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


if __name__ == "__main__":
    unittest.main()
