import json
import sqlite3
import sys
import tempfile
import unittest
from contextlib import redirect_stdout
from datetime import date
from io import StringIO
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
CORE_DIR = PROJECT_ROOT / "core"
if str(CORE_DIR) not in sys.path:
    sys.path.insert(0, str(CORE_DIR))

import excel_profiles
import expedient_service
import gestor_bd


class ExcelProfileTest(unittest.TestCase):
    def test_658_profile_declares_only_its_active_modules(self):
        profile = excel_profiles.load_profile("658_acs_v1", PROJECT_ROOT)

        self.assertEqual(
            ("GAS", "ELECTRICIDAD", "AGUA", "OTROS_GASTOS", "ACS"),
            profile.active_modules,
        )
        self.assertNotIn("CALEFACCION", profile.active_modules)
        self.assertIn(("ANALISIS", "H23"), profile.required_formula_cells)

    def test_658_agua_invoice_table_ends_at_row_20(self):
        profile = excel_profiles.load_profile("658_acs_v1", PROJECT_ROOT)

        self.assertEqual(20, profile.workbook_layout["tables"]["AGUA"]["end_row"])

    def test_profile_rejects_unsafe_path_duplicate_concepts_and_unknown_method(self):
        valid = {
            "key": "test_v1",
            "version": "1",
            "community_code": "TEST",
            "template_relative_path": "plantillas/comunidades/test.xlsx",
            "active_modules": ["ACS"],
            "required_sheets": ["LECTURAS ACS M3", "ANALISIS"],
            "required_formula_cells": [["ANALISIS", "H23"]],
            "concepts": [{
                "key": "acs_variable",
                "allocation_method": "consumption",
                "actual_source": "period_parameters.acs_variable_actual",
                "billed_source": "period_parameters.acs_variable_billed",
                "required": True,
            }],
        }

        invalid_profiles = []
        unsafe = dict(valid, template_relative_path="../privado.xlsx")
        invalid_profiles.append(unsafe)
        duplicate = dict(valid, concepts=[valid["concepts"][0], valid["concepts"][0]])
        invalid_profiles.append(duplicate)
        unknown_method = json.loads(json.dumps(valid))
        unknown_method["concepts"][0]["allocation_method"] = "azar"
        invalid_profiles.append(unknown_method)

        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            profile_dir = root / "config" / "excel_profiles"
            profile_dir.mkdir(parents=True)
            for index, payload in enumerate(invalid_profiles):
                key = f"invalid_{index}"
                payload["key"] = key
                (profile_dir / f"{key}.json").write_text(
                    json.dumps(payload), encoding="utf-8"
                )
                with self.subTest(index=index):
                    with self.assertRaises(ValueError):
                        excel_profiles.load_profile(key, root)

    def test_profile_rejects_missing_fields_and_required_module_sheets(self):
        incomplete = {
            "key": "incomplete",
            "version": "1",
            "community_code": "TEST",
            "template_relative_path": "plantillas/test.xlsx",
            "active_modules": ["ACS"],
            "required_sheets": ["ANALISIS"],
            "required_formula_cells": [["ANALISIS", "H23"]],
        }
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            profile_dir = root / "config" / "excel_profiles"
            profile_dir.mkdir(parents=True)
            (profile_dir / "incomplete.json").write_text(
                json.dumps(incomplete), encoding="utf-8"
            )
            with self.assertRaises(ValueError):
                excel_profiles.load_profile("incomplete", root)

            incomplete["concepts"] = []
            (profile_dir / "incomplete.json").write_text(
                json.dumps(incomplete), encoding="utf-8"
            )
            with self.assertRaisesRegex(ValueError, "LECTURAS ACS M3"):
                excel_profiles.load_profile("incomplete", root)

    def test_profile_rejects_posix_windows_root_and_drive_template_paths(self):
        valid = {
            "key": "unsafe",
            "version": "1",
            "community_code": "TEST",
            "template_relative_path": "plantillas/test.xlsx",
            "active_modules": [],
            "required_sheets": [],
            "required_formula_cells": [],
            "concepts": [],
        }
        unsafe_paths = (
            "/plantillas/test.xlsx",
            "\\plantillas\\test.xlsx",
            "C:\\plantillas\\test.xlsx",
        )
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            profile_dir = root / "config" / "excel_profiles"
            profile_dir.mkdir(parents=True)
            for index, unsafe_path in enumerate(unsafe_paths):
                key = f"unsafe_{index}"
                payload = dict(
                    valid,
                    key=key,
                    template_relative_path=unsafe_path,
                )
                (profile_dir / f"{key}.json").write_text(
                    json.dumps(payload), encoding="utf-8"
                )
                with self.subTest(path=unsafe_path):
                    with self.assertRaisesRegex(ValueError, "ruta relativa"):
                        excel_profiles.load_profile(key, root)

    def test_profile_rejects_formula_text_spaces_and_ranges(self):
        valid = {
            "key": "formula",
            "version": "1",
            "community_code": "TEST",
            "template_relative_path": "plantillas/test.xlsx",
            "active_modules": [],
            "required_sheets": ["ANALISIS"],
            "required_formula_cells": [["ANALISIS", "H23"]],
            "concepts": [],
        }
        invalid_references = ("TOTAL", "H 23", "H23:H24")
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            profile_dir = root / "config" / "excel_profiles"
            profile_dir.mkdir(parents=True)
            for index, reference in enumerate(invalid_references):
                key = f"formula_{index}"
                payload = dict(
                    valid,
                    key=key,
                    required_formula_cells=[["ANALISIS", reference]],
                )
                (profile_dir / f"{key}.json").write_text(
                    json.dumps(payload), encoding="utf-8"
                )
                with self.subTest(reference=reference):
                    with self.assertRaisesRegex(ValueError, "A1"):
                        excel_profiles.load_profile(key, root)


class CasePeriodLinkTest(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory(ignore_cleanup_errors=True)
        self.database_path = Path(self.directory.name) / "gestion.db"
        with redirect_stdout(StringIO()):
            gestor_bd.crear_bd(str(self.database_path))
        self.connection = sqlite3.connect(self.database_path)
        self.connection.row_factory = sqlite3.Row
        self.connection.execute("PRAGMA foreign_keys = ON")
        self.community_id = gestor_bd.obtener_o_crear_comunidad(
            self.connection, "SINTETICA", "Comunidad sintética"
        )

    def tearDown(self):
        self.connection.close()
        self.directory.cleanup()

    def test_link_case_reuses_exact_period_dates_and_exposes_period_id(self):
        cursor = self.connection.execute(
            """INSERT INTO periodos
               (id_comunidad, nombre, fecha_inicio, fecha_fin, estado)
               VALUES (?, ?, ?, ?, 'abierto')""",
            (self.community_id, "Nombre anterior", "2025-09-01", "2026-08-31"),
        )
        existing_period_id = cursor.lastrowid
        self.connection.commit()
        case = expedient_service.create_case(
            self.connection,
            self.community_id,
            name="2025-2026",
            start_date=date(2025, 9, 1),
            end_date=date(2026, 8, 31),
        )

        linked_id = expedient_service.link_case_to_period(
            self.connection, case.id_case
        )

        self.assertEqual(existing_period_id, linked_id)
        self.assertEqual(
            existing_period_id,
            expedient_service.get_case(self.connection, case.id_case).period_id,
        )
        self.assertEqual(
            1, self.connection.execute("SELECT COUNT(*) FROM periodos").fetchone()[0]
        )

    def test_link_case_rejects_same_period_name_with_incompatible_dates(self):
        self.connection.execute(
            """INSERT INTO periodos
               (id_comunidad, nombre, fecha_inicio, fecha_fin, estado)
               VALUES (?, ?, ?, ?, 'abierto')""",
            (self.community_id, "2025-2026", "2025-01-01", "2025-12-31"),
        )
        self.connection.commit()
        case = expedient_service.create_case(
            self.connection,
            self.community_id,
            name="2025-2026",
            start_date=date(2025, 9, 1),
            end_date=date(2026, 8, 31),
        )

        with self.assertRaisesRegex(ValueError, "fechas incompatibles"):
            expedient_service.link_case_to_period(self.connection, case.id_case)

        self.assertIsNone(
            expedient_service.get_case(self.connection, case.id_case).period_id
        )


if __name__ == "__main__":
    unittest.main()
