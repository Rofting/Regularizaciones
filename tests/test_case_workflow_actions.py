"""Acciones del flujo guiado, comprobables sin arrancar Tk."""

from __future__ import annotations

import sys
import unittest
from pathlib import Path
from unittest.mock import patch


PROJECT_ROOT = Path(__file__).resolve().parents[1]
CORE_DIR = PROJECT_ROOT / "core"
if str(CORE_DIR) not in sys.path:
    sys.path.insert(0, str(CORE_DIR))

import gestor_bd
from tests.helpers import temporary_database


class CaseWorkflowActionsTest(unittest.TestCase):
    def setUp(self):
        self.database = temporary_database()
        self.connection, self.database_path = self.database.__enter__()
        gestor_bd.crear_bd(str(self.database_path))
        self.community_id = gestor_bd.obtener_o_crear_comunidad(
            self.connection, "658", "Comunidad portátil"
        )
        self.other_community_id = gestor_bd.obtener_o_crear_comunidad(
            self.connection, "999", "Otra comunidad"
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
        self.other_case_id = self.connection.execute(
            """INSERT INTO regularization_cases
               (id_comunidad,nombre,fecha_inicio,fecha_fin,estado)
               VALUES (?, 'otro', '2025-09-01', '2026-08-31', 'draft')""",
            (self.other_community_id,),
        ).lastrowid
        self.connection.execute(
            """INSERT INTO excel_template_profiles
               (id_comunidad,profile_key,profile_version,template_relative_path,
                template_sha256,profile_sha256,status)
               VALUES (?, '658_acs_v1', '1',
                       'plantillas/comunidades/658/658_acs_v1.xlsx',
                       'template', 'profile', 'active')""",
            (self.community_id,),
        )
        self.connection.commit()

    def tearDown(self):
        self.database.__exit__(None, None, None)

    def test_rejects_an_expedient_from_another_active_community(self):
        from case_workflow_actions import WorkflowBlockedError, run_generate_excel

        with self.assertRaisesRegex(WorkflowBlockedError, "otra comunidad"):
            run_generate_excel(
                self.database_path,
                id_case=self.other_case_id,
                active_community_id=self.community_id,
                project_root=PROJECT_ROOT,
                output_root=PROJECT_ROOT / "salidas-prueba",
            )

    def test_resolves_the_active_profile_from_the_community_not_a_fixed_code(self):
        from case_workflow_actions import resolve_case_profile

        profile = resolve_case_profile(
            self.connection,
            id_case=self.case_id,
            active_community_id=self.community_id,
            project_root=PROJECT_ROOT,
        )

        self.assertEqual("658_acs_v1", profile.key)
        self.assertEqual("658", profile.community_code)

    def test_excel_distribution_and_letters_forward_consistent_progress_events(self):
        from case_workflow_actions import (
            run_calculate_distribution,
            run_generate_excel,
            run_generate_letters,
        )

        events = []
        with patch("case_workflow_actions.generate_official_excel", return_value="excel") as export:
            result = run_generate_excel(
                self.database_path, id_case=self.case_id,
                active_community_id=self.community_id, project_root=PROJECT_ROOT,
                output_root=PROJECT_ROOT / "salidas-prueba",
                progress=lambda stage, payload: events.append((stage, payload)),
            )
        self.assertEqual("excel", result)
        self.assertEqual(self.case_id, export.call_args.kwargs["id_case"])
        self.assertEqual(["validate_case", "generar_excel"], [event[0] for event in events])

        events.clear()
        with patch("case_workflow_actions.calculate_case_distribution", return_value="reparto") as distribution:
            result = run_calculate_distribution(
                self.database_path, id_case=self.case_id,
                active_community_id=self.community_id, project_root=PROJECT_ROOT,
                progress=lambda stage, payload: events.append((stage, payload)),
            )
        self.assertEqual("reparto", result)
        self.assertEqual(self.case_id, distribution.call_args.kwargs["id_case"])
        self.assertEqual(["validate_case", "calcular_reparto"], [event[0] for event in events])

        events.clear()
        from case_letter_service import LetterBatchResult
        with patch(
            "case_workflow_actions.generate_case_letters",
            return_value=LetterBatchResult(1, PROJECT_ROOT / "salidas-prueba", 1, ("pendiente",)),
        ) as letters:
            result = run_generate_letters(
                self.database_path, id_case=self.case_id,
                active_community_id=self.community_id, project_root=PROJECT_ROOT,
                selected_concepts=("acs_fixed",),
                progress=lambda stage, payload: events.append((stage, payload)),
            )
        self.assertEqual(1, result.id_letter_run)
        self.assertEqual(("acs_fixed",), letters.call_args.kwargs["selected_concepts"])
        self.assertEqual(["validate_case", "generar_cartas"], [event[0] for event in events])

    def test_successful_distribution_marks_the_case_reconciled_before_letters(self):
        from case_workflow_actions import run_calculate_distribution

        with patch("case_workflow_actions.calculate_case_distribution", return_value="reparto"):
            run_calculate_distribution(
                self.database_path, id_case=self.case_id,
                active_community_id=self.community_id, project_root=PROJECT_ROOT,
            )

        self.assertEqual(
            "reconciled",
            self.connection.execute(
                "SELECT estado FROM regularization_cases WHERE id_case=?", (self.case_id,)
            ).fetchone()[0],
        )

    def test_complete_letter_batch_marks_the_case_as_delivered(self):
        from case_letter_service import LetterBatchResult
        from case_workflow_actions import run_generate_letters

        self.connection.execute(
            "UPDATE regularization_cases SET estado='reconciled' WHERE id_case=?",
            (self.case_id,),
        )
        self.connection.commit()
        batch = LetterBatchResult(1, PROJECT_ROOT / "salidas-prueba", 2, ())
        with patch("case_workflow_actions.generate_case_letters", return_value=batch):
            run_generate_letters(
                self.database_path, id_case=self.case_id,
                active_community_id=self.community_id, project_root=PROJECT_ROOT,
                selected_concepts=("acs_fixed",),
            )

        self.assertEqual(
            "deliveries_generated",
            self.connection.execute(
                "SELECT estado FROM regularization_cases WHERE id_case=?", (self.case_id,)
            ).fetchone()[0],
        )

    def test_bootstrap_requires_both_optional_companion_sources_together(self):
        from case_workflow_actions import WorkflowBlockedError, run_bootstrap_import

        with self.assertRaisesRegex(WorkflowBlockedError, "ambas fuentes complementarias"):
            run_bootstrap_import(
                self.database_path, id_case=self.case_id,
                active_community_id=self.community_id, project_root=PROJECT_ROOT,
                master_path=PROJECT_ROOT / "README.md",
                owner_list_path=PROJECT_ROOT / "README.md",
                readings_path=None,
            )

    def test_clean_bootstrap_marks_the_case_ready_for_excel_generation(self):
        from excel_bootstrap_importer import BootstrapImportResult
        from case_workflow_actions import run_bootstrap_import

        imported = BootstrapImportResult(1, self.period_id, 2, 0, 0, 0)
        with patch(
            "case_workflow_actions.excel_bootstrap_importer.import_master_excel",
            return_value=imported,
        ), patch("case_workflow_actions.document_review.validate_case_ready") as ready:
            result, companions = run_bootstrap_import(
                self.database_path, id_case=self.case_id,
                active_community_id=self.community_id, project_root=PROJECT_ROOT,
                master_path=PROJECT_ROOT / "README.md",
            )

        self.assertEqual(imported, result)
        self.assertIsNone(companions)
        self.assertEqual(self.case_id, ready.call_args.args[1])

    def test_letter_selector_exposes_only_concepts_active_for_the_case_profile(self):
        from case_workflow_actions import available_case_letter_concepts

        self.connection.executemany(
            """UPDATE regularization_concepts
               SET label=?,display_order=?,active=1 WHERE concept_key=?""",
            [
                ("Cuota fija de ACS", 10, "acs_fixed"),
                ("Cuota fija de calefacción", 20, "heating_fixed"),
            ],
        )
        owner_id = self.connection.execute(
            """INSERT INTO propietarios
               (id_comunidad,codigo_vivienda,nombre_propietario,coeficiente)
               VALUES (?, 'A', 'Persona de prueba', 1)""",
            (self.community_id,),
        ).lastrowid
        self.connection.execute(
            """INSERT INTO owner_concept_results
               (id_propietario,id_periodo,concept_key,billed_cents,actual_cents,difference_cents)
               VALUES (?,?,?,?,?,?)""",
            (owner_id, self.period_id, "acs_fixed", 100, 100, 0),
        )
        self.connection.commit()

        concepts = available_case_letter_concepts(
            self.database_path, id_case=self.case_id,
            active_community_id=self.community_id, project_root=PROJECT_ROOT,
        )

        self.assertEqual((("acs_fixed", "Cuota fija de ACS"),), concepts)


if __name__ == "__main__":
    unittest.main()
