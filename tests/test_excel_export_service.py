import json
import sqlite3
import sys
import tempfile
import unittest
from contextlib import redirect_stdout
from datetime import date
from io import StringIO
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill


PROJECT_ROOT = Path(__file__).resolve().parents[1]
CORE_DIR = PROJECT_ROOT / "core"
if str(CORE_DIR) not in sys.path:
    sys.path.insert(0, str(CORE_DIR))

import excel_generator
import gestor_bd
from excel_export_service import ExportBlockedError, generate_official_excel
from excel_validation import WorkbookValidationError, validate_workbook
from office_recalculation import (
    DeterministicRecalculator,
    LibreOfficeRecalculator,
    RecalculationError,
)


PROFILE_DATA = {
    "key": "658_acs_v1",
    "version": "1",
    "community_code": "658",
    "template_relative_path": "plantillas/comunidades/658/658_acs_v1.xlsx",
    "active_modules": ["GAS", "ELECTRICIDAD", "AGUA", "OTROS_GASTOS", "ACS"],
    "required_sheets": [
        "DATOS", "GAS", "ELECTRICIDAD", "AGUA", "OTROS GASTOS",
        "LECTURAS ACS M3", "ANALISIS",
    ],
    "required_formula_cells": [["ANALISIS", "H23"]],
    "concepts": [
        {
            "key": "acs_fixed", "allocation_method": "equal",
            "actual_source": "period_parameters.acs_fixed_actual",
            "billed_source": "period_parameters.acs_fixed_billed", "required": True,
        },
        {
            "key": "acs_variable", "allocation_method": "consumption",
            "actual_source": "period_parameters.acs_variable_actual",
            "billed_source": "period_parameters.acs_variable_billed", "required": True,
        },
        {
            "key": "extraordinary_expense", "allocation_method": "coefficient",
            "actual_source": "period_parameters.extraordinary_expense_actual",
            "billed_source": None, "required": False,
        },
    ],
    "workbook_layout": {
        "metadata_cells": {
            "community_name": ["DATOS", "A1"],
            "period_label": ["DATOS", "A3"],
            "owner_count": ["DATOS", "D4"],
        },
        "tables": {
            "GAS": {
                "sheet": "GAS", "start_row": 10, "end_row": 20,
                "columns": {
                    "invoice_date": "B", "days": "C", "start_date": "D",
                    "end_date": "F", "consumption": "I", "fixed": "J",
                    "variable": "K", "total": "L", "provider": "N",
                },
            },
            "ELECTRICIDAD": {
                "sheet": "ELECTRICIDAD", "start_row": 10, "end_row": 20,
                "columns": {
                    "invoice_date": "B", "start_date": "C", "end_date": "D",
                    "consumption": "E", "fixed": "F", "variable": "G",
                    "total": "H", "provider": "J",
                },
            },
            "AGUA": {
                "sheet": "AGUA", "start_row": 14, "end_row": 24,
                "columns": {
                    "invoice_date": "B", "start_date": "D", "end_date": "F",
                    "consumption": "H", "total": "T", "variable": "V",
                    "fixed": "W", "provider": "X",
                },
            },
            "OTROS_GASTOS": {
                "sheet": "OTROS GASTOS", "start_row": 13, "end_row": 23,
                "columns": {"date": "D", "description": "E", "amount": "F"},
            },
            "ACS": {
                "sheet": "LECTURAS ACS M3", "start_row": 7, "end_row": 17,
                "columns": {
                    "period": "A", "initial_date": "C", "initial": "D",
                    "final_date": "E", "final": "F", "consumption": "G",
                },
            },
        },
        "parameter_cells": {
            "acs_fixed_actual": ["ANALISIS", "H64"],
            "acs_variable_actual": ["ANALISIS", "J64"],
            "acs_fixed_billed": ["ANALISIS", "H66"],
            "acs_variable_billed": ["ANALISIS", "J66"],
        },
        "total_checks": {
            "invoice_total:GAS": ["GAS", "L10:L20"],
            "invoice_component:GAS:fixed": ["GAS", "J10:J20"],
            "invoice_component:GAS:variable": ["GAS", "K10:K20"],
            "invoice_total:ELECTRICIDAD": ["ELECTRICIDAD", "H10:H20"],
            "invoice_component:ELECTRICIDAD:fixed": ["ELECTRICIDAD", "F10:F20"],
            "invoice_component:ELECTRICIDAD:variable": ["ELECTRICIDAD", "G10:G20"],
            "invoice_total:AGUA": ["AGUA", "T14:T24"],
            "invoice_component:AGUA:fixed": ["AGUA", "W14:W24"],
            "invoice_component:AGUA:variable": ["AGUA", "V14:V24"],
            "parameter:extraordinary_expense_actual": ["OTROS GASTOS", "F13:F23"],
            "reading_total:ACS": ["LECTURAS ACS M3", "G7:G17"],
            "parameter:acs_fixed_actual": ["ANALISIS", "H64"],
            "parameter:acs_variable_actual": ["ANALISIS", "J64"],
            "parameter:acs_fixed_billed": ["ANALISIS", "H66"],
            "parameter:acs_variable_billed": ["ANALISIS", "J66"],
        },
    },
}


def _make_template(path: Path) -> Path:
    workbook = Workbook()
    workbook.active.title = "DATOS"
    for sheet_name in PROFILE_DATA["required_sheets"][1:]:
        workbook.create_sheet(sheet_name)
    for sheet in workbook.worksheets:
        sheet.print_area = "A1:X40"
    workbook["ANALISIS"]["H23"] = "=SUM(H1:H22)"
    workbook["ANALISIS"]["B2"] = "NO TOCAR"
    workbook["ANALISIS"]["B2"].fill = PatternFill("solid", fgColor="D9EAF7")
    workbook["GAS"]["L11"] = 9999
    workbook["GAS"]["L12"] = "=SUM(L10:L11)"
    workbook["GAS"]["P5"] = "CELDA AJENA"
    workbook.save(path)
    with ZipFile(path, "a", ZIP_DEFLATED) as archive:
        archive.writestr("customXml/item1.xml", b"<audit>preservar</audit>")
    return path


class ExcelExportServiceTest(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory(ignore_cleanup_errors=True)
        self.root = Path(self.directory.name)
        self.project_root = self.root / "proyecto"
        config_dir = self.project_root / "config" / "excel_profiles"
        template_dir = self.project_root / "plantillas" / "comunidades" / "658"
        config_dir.mkdir(parents=True)
        template_dir.mkdir(parents=True)
        (config_dir / "658_acs_v1.json").write_text(
            json.dumps(PROFILE_DATA), encoding="utf-8"
        )
        self.template = _make_template(template_dir / "658_acs_v1.xlsx")
        self.output_root = self.root / "salidas"
        self.database_path = self.root / "gestion.db"
        with redirect_stdout(StringIO()):
            gestor_bd.crear_bd(str(self.database_path))
        self.connection = sqlite3.connect(self.database_path)
        self.connection.row_factory = sqlite3.Row
        self.connection.execute("PRAGMA foreign_keys = ON")
        self.community_id = gestor_bd.obtener_o_crear_comunidad(
            self.connection, "658", "Comunidad sintética 658"
        )
        case_cursor = self.connection.execute(
            """INSERT INTO regularization_cases
               (id_comunidad,nombre,fecha_inicio,fecha_fin,estado)
               VALUES (?,?,?,?,?)""",
            (self.community_id, "2025-2026", "2025-09-01", "2026-08-31",
             "ready_for_calculation"),
        )
        self.case_id = int(case_cursor.lastrowid)
        period_cursor = self.connection.execute(
            """INSERT INTO periodos
               (id_comunidad,nombre,fecha_inicio,fecha_fin,estado)
               VALUES (?,?,?,?,?)""",
            (self.community_id, "2025-2026", "2025-09-01", "2026-08-31", "abierto"),
        )
        self.period_id = int(period_cursor.lastrowid)
        self.connection.execute(
            "UPDATE regularization_cases SET id_periodo=? WHERE id_case=?",
            (self.period_id, self.case_id),
        )
        self._insert_normalized_inputs()
        self.connection.commit()

    def tearDown(self):
        self.connection.close()
        self.directory.cleanup()

    @property
    def official_output(self) -> Path:
        return self.output_root / "Excels_Maestros" / "Comunidad_658.xlsx"

    def _insert_normalized_inputs(self):
        invoices = [
            ("GAS", "Gas sintético", "G-1", "2025-10-05", "2025-09-01",
             "2025-09-30", 30, 110.0, "kWh", 20.0, 80.0, 100.0),
            ("ELECTRICIDAD", "Luz sintética", "E-1", "2025-10-08", "2025-09-01",
             "2025-09-30", 30, 250.0, "kWh", 25.0, 75.0, 100.0),
            ("AGUA", "Agua sintética", "A-1", "2025-10-09", "2025-09-01",
             "2025-09-30", 30, 15.0, "m3", 20.0, 40.0, 60.0),
        ]
        for row in invoices:
            cursor = self.connection.execute(
                """INSERT INTO facturas
                   (id_comunidad,id_periodo,tipo_suministro,proveedor,
                    cups_o_referencia,num_factura,fecha_factura,fecha_inicio,
                    fecha_fin,dias_facturados,consumo_total,unidad_consumo,
                    termino_fijo,termino_variable,importe_total)
                   VALUES (?,?,?,?,'SINTETICO',?,?,?,?,?,?,?,?,?,?)""",
                (self.community_id, self.period_id, *row),
            )
            for key, amount in (("fixed", row[-3]), ("variable", row[-2]), ("total", row[-1])):
                self.connection.execute(
                    """INSERT INTO invoice_components
                       (id_factura,component_key,amount,unit,source_sheet,source_cell)
                       VALUES (?,?,?,?,?,?)""",
                    (cursor.lastrowid, key, amount, "EUR", row[0], "A1"),
                )
        parameters = {
            "acs_fixed_actual": 78.0,
            "acs_variable_actual": 182.0,
            "acs_fixed_billed": 112.0,
            "acs_variable_billed": 168.0,
            "extraordinary_expense_actual": 120.0,
        }
        for key, amount in parameters.items():
            self.connection.execute(
                """INSERT INTO period_parameters
                   (id_comunidad,id_periodo,parameter_key,numeric_value,unit,
                    source_sheet,source_cell) VALUES (?,?,?,?,?,?,?)""",
                (self.community_id, self.period_id, key, amount, "EUR", "ANALISIS", "A1"),
            )
        for code, name, coefficient, first, last in (
            ("P1-A", "Persona Uno", 50.0, 100.0, 125.0),
            ("P1-B", "Persona Dos", 50.0, 200.0, 215.0),
        ):
            owner = self.connection.execute(
                """INSERT INTO propietarios
                   (id_comunidad,codigo_vivienda,nombre_propietario,coeficiente,activo)
                   VALUES (?,?,?,?,1)""",
                (self.community_id, code, name, coefficient),
            )
            self.connection.executemany(
                """INSERT INTO lecturas_vecino
                   (id_propietario,id_periodo,tipo,fecha_lectura,valor_acumulado,
                    estado,fuente) VALUES (?,?, 'ACS', ?, ?, 'real', 'sintetico')""",
                [
                    (owner.lastrowid, self.period_id, "2025-09-01", first),
                    (owner.lastrowid, self.period_id, "2026-08-31", last),
                ],
            )

    def _expected_totals(self):
        return {
            "invoice_total:GAS": 10_000,
            "invoice_component:GAS:fixed": 2_000,
            "invoice_component:GAS:variable": 8_000,
            "invoice_total:ELECTRICIDAD": 10_000,
            "invoice_component:ELECTRICIDAD:fixed": 2_500,
            "invoice_component:ELECTRICIDAD:variable": 7_500,
            "invoice_total:AGUA": 6_000,
            "invoice_component:AGUA:fixed": 2_000,
            "invoice_component:AGUA:variable": 4_000,
            "parameter:extraordinary_expense_actual": 12_000,
            "reading_total:ACS": 4_000,
            "parameter:acs_fixed_actual": 7_800,
            "parameter:acs_variable_actual": 18_200,
            "parameter:acs_fixed_billed": 11_200,
            "parameter:acs_variable_billed": 16_800,
        }

    def test_export_writes_typed_data_preserves_template_and_reports_progress(self):
        stages = []
        result = generate_official_excel(
            self.connection,
            id_case=self.case_id,
            project_root=self.project_root,
            output_root=self.output_root,
            progress=lambda stage, payload: stages.append(stage),
            recalculator=DeterministicRecalculator(),
        )

        self.assertEqual(self.official_output, result.output_path)
        self.assertIsNone(result.backup_path)
        self.assertEqual("validated", result.stage)
        self.assertEqual(
            "validated",
            self.connection.execute(
                "SELECT status FROM excel_export_runs WHERE id_export_run=?",
                (result.id_export_run,),
            ).fetchone()[0],
        )
        self.assertEqual(
            ["validate_case", "prepare_template", "write_gas",
             "write_electricidad", "write_agua", "write_otros_gastos",
             "write_acs", "recalculate", "reconcile", "publish"],
            stages,
        )
        workbook = load_workbook(result.output_path, data_only=False)
        self.assertIsInstance(workbook["GAS"]["B10"].value, date)
        self.assertIsInstance(workbook["GAS"]["L10"].value, (int, float))
        self.assertIsNone(workbook["GAS"]["L11"].value)
        self.assertEqual("=SUM(L10:L11)", workbook["GAS"]["L12"].value)
        self.assertEqual("=SUM(H1:H22)", workbook["ANALISIS"]["H23"].value)
        self.assertEqual(40, workbook["LECTURAS ACS M3"]["G7"].value)
        self.assertEqual("NO TOCAR", workbook["ANALISIS"]["B2"].value)
        self.assertEqual("00D9EAF7", workbook["ANALISIS"]["B2"].fill.fgColor.rgb)
        self.assertEqual("'ANALISIS'!$A$1:$X$40", str(workbook["ANALISIS"].print_area))
        workbook.close()
        with ZipFile(result.output_path) as archive:
            self.assertEqual(b"<audit>preservar</audit>", archive.read("customXml/item1.xml"))

    def test_open_issue_blocks_export_without_replacing_predecessor(self):
        self.official_output.parent.mkdir(parents=True)
        self.official_output.write_bytes(b"SALIDA ANTERIOR")
        document = self.connection.execute(
            """INSERT INTO source_documents
               (id_case,original_name,archived_path,sha256,document_kind,status)
               VALUES (?,?,?,?,?,'under_review')""",
            (self.case_id, "dato.xlsx", "archivo/dato.xlsx", "a" * 64, "excel_master"),
        )
        self.connection.execute(
            """INSERT INTO review_issues
               (id_case,id_document,code,field_name,message,status)
               VALUES (?,?,?,?,?,'open')""",
            (self.case_id, document.lastrowid, "MISSING_DATA", "AGUA.total", "Falta dato"),
        )
        self.connection.commit()

        with self.assertRaisesRegex(ExportBlockedError, "incidencia"):
            generate_official_excel(
                self.connection, id_case=self.case_id,
                project_root=self.project_root, output_root=self.output_root,
                recalculator=DeterministicRecalculator(),
            )

        self.assertEqual(b"SALIDA ANTERIOR", self.official_output.read_bytes())

    def test_mismatched_invoice_components_block_official_output(self):
        self.connection.execute(
            """UPDATE invoice_components SET amount=10
               WHERE component_key='fixed' AND id_factura=(
                   SELECT id_factura FROM facturas WHERE tipo_suministro='AGUA'
               )"""
        )
        self.connection.commit()

        with self.assertRaisesRegex(ExportBlockedError, "componentes"):
            generate_official_excel(
                self.connection, id_case=self.case_id,
                project_root=self.project_root, output_root=self.output_root,
                recalculator=DeterministicRecalculator(),
            )

        self.assertFalse(self.official_output.exists())

    def test_formula_error_fails_auditable_run_without_replacing_predecessor(self):
        self.official_output.parent.mkdir(parents=True)
        self.official_output.write_bytes(b"SALIDA VALIDADA PREVIA")

        def corrupt_formula(path: Path):
            workbook = load_workbook(path)
            workbook["ANALISIS"]["H23"] = "#VALUE!"
            workbook.save(path)

        with self.assertRaisesRegex(WorkbookValidationError, "fórmula"):
            generate_official_excel(
                self.connection, id_case=self.case_id,
                project_root=self.project_root, output_root=self.output_root,
                recalculator=DeterministicRecalculator(after_recalculate=corrupt_formula),
            )

        self.assertEqual(b"SALIDA VALIDADA PREVIA", self.official_output.read_bytes())
        run = self.connection.execute(
            """SELECT status,error_message,diagnostic_path FROM excel_export_runs
               ORDER BY id_export_run DESC LIMIT 1"""
        ).fetchone()
        self.assertEqual("failed", run["status"])
        self.assertIn("fórmula", run["error_message"])
        self.assertTrue(Path(run["diagnostic_path"]).is_file())

    def test_second_success_backs_up_exact_preceding_workbook(self):
        first = generate_official_excel(
            self.connection, id_case=self.case_id,
            project_root=self.project_root, output_root=self.output_root,
            recalculator=DeterministicRecalculator(),
        )
        preceding = first.output_path.read_bytes()
        self.connection.execute(
            "UPDATE period_parameters SET numeric_value=79 WHERE parameter_key='acs_fixed_actual'"
        )
        self.connection.commit()

        second = generate_official_excel(
            self.connection, id_case=self.case_id,
            project_root=self.project_root, output_root=self.output_root,
            recalculator=DeterministicRecalculator(),
        )

        self.assertIsNotNone(second.backup_path)
        self.assertEqual(preceding, second.backup_path.read_bytes())
        self.assertNotEqual(preceding, second.output_path.read_bytes())

    def test_optional_extraordinary_expense_does_not_block_export(self):
        self.connection.execute(
            "DELETE FROM period_parameters WHERE parameter_key='extraordinary_expense_actual'"
        )
        self.connection.commit()

        result = generate_official_excel(
            self.connection,
            id_case=self.case_id,
            project_root=self.project_root,
            output_root=self.output_root,
            recalculator=DeterministicRecalculator(),
        )

        self.assertEqual("validated", result.stage)
        workbook = load_workbook(result.output_path, data_only=False)
        self.assertIsNone(workbook["OTROS GASTOS"]["F13"].value)
        workbook.close()

    def test_validator_rejects_wrong_declared_total(self):
        result = generate_official_excel(
            self.connection, id_case=self.case_id,
            project_root=self.project_root, output_root=self.output_root,
            recalculator=DeterministicRecalculator(),
        )
        workbook = load_workbook(result.output_path)
        workbook["ANALISIS"]["H64"] = 999
        workbook.save(result.output_path)

        from excel_profiles import load_profile
        with self.assertRaisesRegex(WorkbookValidationError, "acs_fixed_actual"):
            validate_workbook(
                result.output_path,
                load_profile("658_acs_v1", self.project_root),
                self._expected_totals(),
            )

    def test_missing_libreoffice_has_an_understandable_error(self):
        recalculator = LibreOfficeRecalculator(
            executable=self.root / "programa-inexistente" / "soffice.exe"
        )
        with self.assertRaisesRegex(RecalculationError, "LibreOffice"):
            recalculator.recalculate(self.template, self.root / "recalculo")

    def test_legacy_entry_point_accepts_an_eligible_case(self):
        self.connection.close()
        result = excel_generator.regenerar_excel_comunidad(
            "658", ruta_bd=str(self.database_path),
            ruta_excels=str(self.output_root / "Excels_Maestros"),
            id_case=self.case_id, project_root=self.project_root,
            recalculator=DeterministicRecalculator(),
        )
        self.connection = sqlite3.connect(self.database_path)
        self.connection.row_factory = sqlite3.Row

        self.assertTrue(result["ok"])
        self.assertEqual(str(self.official_output), result["archivo"])
        self.assertEqual([], result["errores"])


if __name__ == "__main__":
    unittest.main()
