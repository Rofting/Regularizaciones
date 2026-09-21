"""Acciones del flujo guiado, comprobables sin arrancar Tk."""

from __future__ import annotations

import json
import hashlib
import sys
import tempfile
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
                       'template', ?, 'active')""",
            (
                self.community_id,
                hashlib.sha256(
                    (PROJECT_ROOT / "config" / "excel_profiles" / "658_acs_v1.json").read_bytes()
                ).hexdigest(),
            ),
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

    def test_rejects_profile_bytes_changed_without_version_bump(self):
        from case_workflow_actions import WorkflowBlockedError, resolve_case_profile

        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            profile_path = root / "config" / "excel_profiles" / "658_acs_v1.json"
            profile_path.parent.mkdir(parents=True)
            profile_path.write_bytes(
                (PROJECT_ROOT / "config" / "excel_profiles" / "658_acs_v1.json").read_bytes()
            )
            registered_hash = hashlib.sha256(profile_path.read_bytes()).hexdigest()
            self.connection.execute(
                "UPDATE excel_template_profiles SET profile_sha256=? WHERE id_comunidad=?",
                (registered_hash, self.community_id),
            )
            self.connection.commit()
            profile_path.write_bytes(profile_path.read_bytes() + b"\n")

            with self.assertRaisesRegex(WorkflowBlockedError, "huella"):
                resolve_case_profile(
                    self.connection,
                    id_case=self.case_id,
                    active_community_id=self.community_id,
                    project_root=root,
                )

    def test_resolves_single_json_profile_before_first_export_registration(self):
        from case_workflow_actions import resolve_case_profile

        self.connection.execute("DELETE FROM excel_template_profiles WHERE id_comunidad=?", (self.community_id,))
        self.connection.commit()

        profile = resolve_case_profile(
            self.connection,
            id_case=self.case_id,
            active_community_id=self.community_id,
            project_root=PROJECT_ROOT,
        )

        self.assertEqual("658_acs_v1", profile.key)
        self.assertTrue(profile.workbook_layout)

    def test_blocks_bootstrap_when_no_registered_or_configured_profile_exists(self):
        from case_workflow_actions import WorkflowBlockedError, resolve_case_profile

        self.connection.execute("DELETE FROM excel_template_profiles WHERE id_comunidad=?", (self.community_id,))
        self.connection.commit()
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            (root / "config" / "excel_profiles").mkdir(parents=True)
            with self.assertRaisesRegex(WorkflowBlockedError, "configuración"):
                resolve_case_profile(
                    self.connection, id_case=self.case_id,
                    active_community_id=self.community_id, project_root=root,
                )

    def test_blocks_bootstrap_when_multiple_json_profiles_match_the_community(self):
        from case_workflow_actions import WorkflowBlockedError, resolve_case_profile

        self.connection.execute("DELETE FROM excel_template_profiles WHERE id_comunidad=?", (self.community_id,))
        self.connection.commit()
        reference = json.loads(
            (PROJECT_ROOT / "config" / "excel_profiles" / "658_acs_v1.json").read_text(encoding="utf-8")
        )
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            profile_directory = root / "config" / "excel_profiles"
            profile_directory.mkdir(parents=True)
            for key in ("portable_a", "portable_b"):
                candidate = {**reference, "key": key}
                (profile_directory / f"{key}.json").write_text(
                    json.dumps(candidate), encoding="utf-8"
                )
            with self.assertRaisesRegex(WorkflowBlockedError, "varios perfiles"):
                resolve_case_profile(
                    self.connection, id_case=self.case_id,
                    active_community_id=self.community_id, project_root=root,
                )

    def test_generate_excel_provisions_a_missing_template_instead_of_blocking(self):
        """Una comunidad sin plantilla no detiene el expediente: se le crea.

        El modelo de estudio es común a todas las comunidades, así que exigir
        que alguien eligiera antes un Excel maestro dejaba parados expedientes
        con todas sus fuentes ya validadas.
        """
        from case_workflow_actions import run_generate_excel

        self.connection.execute(
            "DELETE FROM excel_template_profiles WHERE id_comunidad=?",
            (self.community_id,),
        )
        self.connection.commit()

        with patch("case_workflow_actions.plantilla_comunidad.asegurar_plantilla") as provision:
            provision.return_value = None
            with patch("case_workflow_actions.generate_official_excel", return_value="excel") as export:
                result = run_generate_excel(
                    self.database_path,
                    id_case=self.case_id,
                    active_community_id=self.community_id,
                    project_root=PROJECT_ROOT,
                    output_root=PROJECT_ROOT / "salidas-prueba",
                )

        self.assertEqual("excel", result)
        export.assert_called_once()
        provision.assert_called_once()
        self.assertEqual(self.community_id, provision.call_args.kwargs["community_id"])

    def test_generate_excel_blocks_when_the_canonical_model_is_missing(self):
        """Sin modelo del que copiar sí hay que parar, y decir por qué."""
        from case_workflow_actions import WorkflowBlockedError, run_generate_excel
        from plantilla_comunidad import PlantillaNoDisponible

        self.connection.execute(
            "DELETE FROM excel_template_profiles WHERE id_comunidad=?",
            (self.community_id,),
        )
        self.connection.commit()

        with patch("case_workflow_actions.plantilla_comunidad.asegurar_plantilla") as provision:
            provision.side_effect = PlantillaNoDisponible("Falta el modelo canónico")
            with patch("case_workflow_actions.generate_official_excel") as export:
                with self.assertRaisesRegex(WorkflowBlockedError, "modelo canónico"):
                    run_generate_excel(
                        self.database_path,
                        id_case=self.case_id,
                        active_community_id=self.community_id,
                        project_root=PROJECT_ROOT,
                        output_root=PROJECT_ROOT / "salidas-prueba",
                    )

        export.assert_not_called()

    def test_successful_excel_generation_advances_the_case_to_calculated(self):
        from case_workflow_actions import run_generate_excel

        with patch("case_workflow_actions.generate_official_excel", return_value="excel"):
            result = run_generate_excel(
                self.database_path,
                id_case=self.case_id,
                active_community_id=self.community_id,
                project_root=PROJECT_ROOT,
                output_root=PROJECT_ROOT / "salidas-prueba",
            )

        status = self.connection.execute(
            "SELECT estado FROM regularization_cases WHERE id_case=?",
            (self.case_id,),
        ).fetchone()[0]
        self.assertEqual("excel", result)
        self.assertEqual("calculated", status)

    def test_missing_invoice_date_returns_case_to_actionable_review(self):
        from case_workflow_actions import WorkflowBlockedError, run_generate_excel

        document_id = self.connection.execute(
            """INSERT INTO source_documents
               (id_case,original_name,archived_path,sha256,document_kind,status,
                classification_confidence,eligibility_status)
               VALUES (?, 'factura.pdf', 'archivo/factura.pdf', ?, 'invoice',
                       'validated', 'high', 'eligible')""",
            (self.case_id, "a" * 64),
        ).lastrowid
        invoice_id = self.connection.execute(
            """INSERT INTO facturas
               (id_comunidad,id_periodo,tipo_suministro,proveedor,fecha_factura,
                fecha_inicio,fecha_fin,importe_total,archivo_origen)
               VALUES (?, ?, 'GAS', 'Proveedor', NULL,
                       '2025-09-01', '2025-09-30', 100, 'archivo/factura.pdf')""",
            (self.community_id, self.period_id),
        ).lastrowid
        self.connection.execute(
            """INSERT INTO invoice_components
               (id_factura,component_key,amount,unit)
               VALUES (?, 'total', 100, 'EUR')""",
            (invoice_id,),
        )
        self.connection.execute(
            """INSERT INTO archivos_procesados
               (nombre_archivo,id_factura,resultado)
               VALUES (?, ?, 'ok')""",
            (f"source_document:{document_id}", invoice_id),
        )
        self.connection.commit()

        with patch("case_workflow_actions.plantilla_comunidad.asegurar_plantilla", return_value=None), patch(
            "case_workflow_actions.generate_official_excel"
        ) as export:
            with self.assertRaisesRegex(WorkflowBlockedError, "fecha de factura"):
                run_generate_excel(
                    self.database_path,
                    id_case=self.case_id,
                    active_community_id=self.community_id,
                    project_root=PROJECT_ROOT,
                    output_root=PROJECT_ROOT / "salidas-prueba",
                )

        export.assert_not_called()
        issue = self.connection.execute(
            """SELECT field_name,status FROM review_issues
               WHERE id_case=? AND id_document=?""",
            (self.case_id, document_id),
        ).fetchone()
        self.assertEqual(("fecha_factura", "open"), tuple(issue))
        status = self.connection.execute(
            "SELECT estado FROM regularization_cases WHERE id_case=?",
            (self.case_id,),
        ).fetchone()[0]
        self.assertEqual("under_review", status)

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

    def test_completed_case_can_regenerate_excel_and_invalidates_later_stage(self):
        from case_workflow_actions import run_generate_excel

        self.connection.execute(
            "UPDATE regularization_cases SET estado='deliveries_generated' WHERE id_case=?",
            (self.case_id,),
        )
        self.connection.commit()
        with patch("case_workflow_actions.generate_official_excel", return_value="nuevo excel"):
            result = run_generate_excel(
                self.database_path, id_case=self.case_id,
                active_community_id=self.community_id, project_root=PROJECT_ROOT,
                output_root=PROJECT_ROOT / "salidas-prueba",
            )

        self.assertEqual("nuevo excel", result)
        self.assertEqual("calculated", self.connection.execute(
            "SELECT estado FROM regularization_cases WHERE id_case=?", (self.case_id,),
        ).fetchone()[0])

    def test_completed_case_can_recalculate_and_then_repeat_letters(self):
        from case_letter_service import LetterBatchResult
        from case_workflow_actions import run_calculate_distribution, run_generate_letters

        self.connection.execute(
            "UPDATE regularization_cases SET estado='deliveries_generated' WHERE id_case=?",
            (self.case_id,),
        )
        self.connection.commit()
        with patch("case_workflow_actions.calculate_case_distribution", return_value="nuevo reparto"):
            self.assertEqual("nuevo reparto", run_calculate_distribution(
                self.database_path, id_case=self.case_id,
                active_community_id=self.community_id, project_root=PROJECT_ROOT,
            ))
        self.assertEqual("reconciled", self.connection.execute(
            "SELECT estado FROM regularization_cases WHERE id_case=?", (self.case_id,),
        ).fetchone()[0])

        batch = LetterBatchResult(9, PROJECT_ROOT / "salidas-prueba", 2, ())
        with patch("case_workflow_actions.generate_case_letters", return_value=batch):
            run_generate_letters(
                self.database_path, id_case=self.case_id,
                active_community_id=self.community_id, project_root=PROJECT_ROOT,
                selected_concepts=("acs_fixed",),
            )
        self.assertEqual("deliveries_generated", self.connection.execute(
            "SELECT estado FROM regularization_cases WHERE id_case=?", (self.case_id,),
        ).fetchone()[0])

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
