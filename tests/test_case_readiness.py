"""Requisitos estructurales compartidos por todas las etapas del expediente."""

from __future__ import annotations

import hashlib
import json
import sys
import tempfile
import unittest
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
CORE_DIR = PROJECT_ROOT / "core"
if str(CORE_DIR) not in sys.path:
    sys.path.insert(0, str(CORE_DIR))

import gestor_bd
from tests.helpers import temporary_database


class CaseReadinessTest(unittest.TestCase):
    def setUp(self):
        self.database = temporary_database()
        self.connection, self.database_path = self.database.__enter__()
        gestor_bd.crear_bd(str(self.database_path))
        self.community_id = gestor_bd.obtener_o_crear_comunidad(
            self.connection, "901", "Comunidad de preparación"
        )
        self.period_id = self.connection.execute(
            """INSERT INTO periodos(id_comunidad,nombre,fecha_inicio,fecha_fin)
               VALUES (?, '2025-2026', '2025-09-01', '2026-08-31')""",
            (self.community_id,),
        ).lastrowid
        self.case_id = self.connection.execute(
            """INSERT INTO regularization_cases
               (id_comunidad,nombre,fecha_inicio,fecha_fin,estado,id_periodo)
               VALUES (?, '2025-2026', '2025-09-01', '2026-08-31',
                       'ready_for_calculation', ?)""",
            (self.community_id, self.period_id),
        ).lastrowid
        self.connection.execute(
            """INSERT INTO source_documents
               (id_case,original_name,archived_path,sha256,document_kind,status,
                eligibility_status)
               VALUES (?, 'fuente.pdf', 'fuente.pdf', ?, 'invoice', 'validated',
                       'eligible')""",
            (self.case_id, "a" * 64),
        )
        self.directory = tempfile.TemporaryDirectory(ignore_cleanup_errors=True)
        self.project_root = Path(self.directory.name)
        self.addCleanup(self.directory.cleanup)

    def tearDown(self):
        self.database.__exit__(None, None, None)

    def _register_profile(self, method: str, *, key: str = "test_profile") -> None:
        concept_key = {
            "equal": "acs_fixed",
            "coefficient": "extraordinary_expense",
            "consumption": "acs_variable",
        }[method]
        concepts = [{
            "key": concept_key,
            "allocation_method": method,
            "actual_source": "period_parameters.test_actual",
            "billed_source": "period_parameters.test_billed",
            "required": True,
        }]
        payload = {
            "key": key,
            "version": "1",
            "community_code": "901",
            "template_relative_path": "plantillas/comunidades/901/modelo.xlsx",
            "active_modules": [],
            "required_sheets": [],
            "required_formula_cells": [],
            "concepts": concepts,
        }
        profile_path = self.project_root / "config" / "excel_profiles" / f"{key}.json"
        profile_path.parent.mkdir(parents=True, exist_ok=True)
        profile_path.write_text(json.dumps(payload), encoding="utf-8")
        template = self.project_root / payload["template_relative_path"]
        template.parent.mkdir(parents=True, exist_ok=True)
        template.write_bytes(b"modelo")
        digest = hashlib.sha256(profile_path.read_bytes()).hexdigest()
        self.connection.execute(
            """INSERT INTO excel_template_profiles
               (id_comunidad,profile_key,profile_version,template_relative_path,
                template_sha256,profile_sha256,status)
               VALUES (?,?,?,?,?,?, 'active')""",
            (
                self.community_id,
                key,
                "1",
                payload["template_relative_path"],
                hashlib.sha256(template.read_bytes()).hexdigest(),
                digest,
            ),
        )
        for parameter in ("test_actual", "test_billed"):
            self.connection.execute(
                """INSERT INTO period_parameters
                   (id_comunidad,id_periodo,parameter_key,numeric_value)
                   VALUES (?,?,?,10)""",
                (self.community_id, self.period_id, parameter),
            )
        self.connection.commit()

    def _add_owner(self, *, code: str = "A", coefficient: float = 1.0) -> int:
        owner_id = self.connection.execute(
            """INSERT INTO propietarios
               (id_comunidad,codigo_vivienda,nombre_propietario,coeficiente,
                tipo_unidad,activo)
               VALUES (?, ?, 'Propietario de prueba', ?, 'vivienda', 1)""",
            (self.community_id, code, coefficient),
        ).lastrowid
        self.connection.commit()
        return int(owner_id)

    def test_case_without_owners_is_not_excel_ready_even_without_open_issues(self):
        self._register_profile("equal")
        from case_readiness import evaluate_case_readiness

        report = evaluate_case_readiness(
            self.connection, self.case_id, self.project_root
        )

        self.assertIn(
            "MISSING_OWNERS", tuple(item.code for item in report.for_stage("excel"))
        )
        self.assertFalse(report.excel_ready)

    def test_coefficient_concept_requires_a_positive_coefficient_for_every_owner(self):
        self._register_profile("coefficient")
        self._add_owner(code="A", coefficient=1.0)
        self._add_owner(code="B", coefficient=0.0)
        from case_readiness import evaluate_case_readiness

        report = evaluate_case_readiness(
            self.connection, self.case_id, self.project_root
        )

        blocker = next(
            item for item in report.for_stage("distribution")
            if item.code == "MISSING_COEFFICIENTS"
        )
        self.assertEqual(1, blocker.count)
        self.assertEqual("review_owners", blocker.action)

    def test_equal_concept_does_not_require_meter_readings(self):
        self._register_profile("equal")
        self._add_owner()
        from case_readiness import evaluate_case_readiness

        report = evaluate_case_readiness(
            self.connection, self.case_id, self.project_root
        )

        self.assertNotIn(
            "MISSING_READINGS",
            tuple(item.code for item in report.for_stage("distribution")),
        )
        self.assertTrue(report.excel_ready)

    def test_single_configured_profile_is_usable_before_first_registration(self):
        self._register_profile("equal")
        self.connection.execute(
            "DELETE FROM excel_template_profiles WHERE id_comunidad=?",
            (self.community_id,),
        )
        self.connection.commit()
        self._add_owner()
        from case_readiness import evaluate_case_readiness

        report = evaluate_case_readiness(
            self.connection, self.case_id, self.project_root
        )

        self.assertNotIn(
            "MISSING_PROFILE",
            tuple(item.code for item in report.for_stage("excel")),
        )
        self.assertTrue(report.excel_ready)

    def test_consumption_concept_requires_effective_start_and_end_readings(self):
        self._register_profile("consumption")
        self._add_owner()
        from case_readiness import evaluate_case_readiness

        report = evaluate_case_readiness(
            self.connection, self.case_id, self.project_root
        )

        blocker = next(
            item for item in report.for_stage("distribution")
            if item.code == "MISSING_READINGS"
        )
        self.assertEqual(1, blocker.count)
        self.assertIn("1 propiedad", blocker.message)


if __name__ == "__main__":
    unittest.main()
