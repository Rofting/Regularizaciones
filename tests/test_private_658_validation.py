import os
import re
import subprocess
import sys
import tempfile
import unittest
import hashlib
import shutil
from types import SimpleNamespace
from pathlib import Path
from unittest.mock import patch


PROJECT_ROOT = Path(__file__).resolve().parents[1]
CORE_DIR = PROJECT_ROOT / "core"
if str(CORE_DIR) not in sys.path:
    sys.path.insert(0, str(CORE_DIR))


class Private658ValidationGuardsTest(unittest.TestCase):
    def test_missing_explicit_source_variables_are_blocked_before_any_work(self):
        from private_658_validation import PrivateValidationBlockedError, validation_environment

        with self.assertRaisesRegex(PrivateValidationBlockedError, "REGULARIZACION_658_MASTER"):
            validation_environment({})

    def test_validation_root_cannot_be_project_or_parent_directory(self):
        from private_658_validation import PrivateValidationBlockedError, prepare_validation_run

        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as directory:
            root = Path(directory)
            with self.assertRaisesRegex(PrivateValidationBlockedError, "segura"):
                prepare_validation_run(root, project_root=root)
            with self.assertRaisesRegex(PrivateValidationBlockedError, "segura"):
                prepare_validation_run(root.parent, project_root=root)

    def test_validation_root_cannot_be_drive_home_or_its_broad_parent(self):
        from private_658_validation import PrivateValidationBlockedError, prepare_validation_run

        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as directory:
            root = Path(directory)
            home = root / "operator"
            home.mkdir()
            project = root / "separate-project"
            profile = project / "config" / "excel_profiles" / "658_acs_v1.json"
            profile.parent.mkdir(parents=True)
            profile.write_text("{}", encoding="utf-8")
            (project / "plantillas").mkdir()
            (project / "plantillas" / "Plantilla_Cartas.docx").write_bytes(b"public-template")
            with patch("private_658_validation.Path.home", return_value=home):
                for broad_root in (Path(root.anchor), home, home.parent):
                    with self.subTest(validation_root=broad_root):
                        with self.assertRaisesRegex(PrivateValidationBlockedError, "segura"):
                            prepare_validation_run(broad_root, project_root=project)
                self.assertEqual([], list(home.glob("validacion_*")))

    def test_safe_empty_root_creates_one_new_run_and_copies_only_public_assets(self):
        from private_658_validation import prepare_validation_run

        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as directory:
            root = Path(directory)
            project = root / "project"
            profile = project / "config" / "excel_profiles" / "658_acs_v1.json"
            profile.parent.mkdir(parents=True)
            profile.write_text("{}", encoding="utf-8")
            (project / "config" / "office-private.json").write_text("no copiar", encoding="utf-8")
            (project / "config" / "letter_identities.json").write_text("no copiar", encoding="utf-8")
            (project / "plantillas").mkdir()
            (project / "plantillas" / "Plantilla_Cartas.docx").write_bytes(b"public-template")
            destination = root / "private-validation"

            run = prepare_validation_run(destination, project_root=project)

            self.assertTrue(run.is_dir())
            self.assertTrue((run / "config" / "excel_profiles" / "658_acs_v1.json").is_file())
            self.assertFalse((run / "config" / "office-private.json").exists())
            self.assertFalse((run / "config" / "letter_identities.json").exists())
            self.assertTrue((run / "plantillas" / "Plantilla_Cartas.docx").is_file())
            self.assertFalse((run / "gestion.db").exists())
            self.assertFalse((run / "expedientes").exists())
            self.assertEqual([], list(destination.glob("gestion.db")))

    def test_sanitized_report_does_not_disclose_source_paths_or_owner_failures(self):
        from private_658_validation import render_sanitized_report

        report = render_sanitized_report(
            status="blocked_open_issues",
            run_root=Path("C:/private/validation/run"),
            issue_counts={"MISSING_READING_RANGE": 2},
            generated_letters=0,
            source_paths=(Path("C:/secret/owner list.xlsx"),),
            failures=("Nombre privado: error",),
        )

        self.assertIn("blocked_open_issues", report)
        self.assertIn("MISSING_READING_RANGE: 2", report)
        self.assertNotIn("secret", report)
        self.assertNotIn("Nombre privado", report)

    def test_resume_rejects_project_and_non_validation_directories(self):
        from private_658_validation import PrivateValidationBlockedError, resume_private_658_validation

        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as directory:
            root = Path(directory)
            project = root / "project"
            project.mkdir()
            with self.assertRaisesRegex(PrivateValidationBlockedError, "segura"):
                resume_private_658_validation(project, project_root=project)
            arbitrary = root / "carpeta_arbitraria"
            arbitrary.mkdir()
            with self.assertRaisesRegex(PrivateValidationBlockedError, "ejecución"):
                resume_private_658_validation(arbitrary, project_root=project)

    def test_resume_existing_clean_run_reuses_database_and_does_not_duplicate_outputs(self):
        import gestor_bd
        from case_letter_service import LetterBatchResult
        from private_658_validation import resume_private_658_validation

        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as directory:
            root = Path(directory)
            project = root / "project"
            project.mkdir()
            run = root / "private" / "validacion_sintetica"
            profile_source = PROJECT_ROOT / "config" / "excel_profiles" / "658_acs_v1.json"
            profile_path = run / "config" / "excel_profiles" / "658_acs_v1.json"
            profile_path.parent.mkdir(parents=True)
            shutil.copy2(profile_source, profile_path)
            (run / "plantillas").mkdir()
            shutil.copy2(
                PROJECT_ROOT / "plantillas" / "Plantilla_Cartas.docx",
                run / "plantillas" / "Plantilla_Cartas.docx",
            )
            database_path = run / "data" / "gestion.db"
            gestor_bd.crear_bd(str(database_path))
            connection = gestor_bd.conectar(str(database_path))
            community_id = gestor_bd.obtener_o_crear_comunidad(
                connection, "658", "Comunidad sintética"
            )
            period_id = connection.execute(
                """INSERT INTO periodos(id_comunidad,nombre,fecha_inicio,fecha_fin)
                   VALUES (?, '2025-2026', '2025-09-01', '2026-08-31')""",
                (community_id,),
            ).lastrowid
            case_id = connection.execute(
                """INSERT INTO regularization_cases
                   (id_comunidad,nombre,fecha_inicio,fecha_fin,estado,id_periodo)
                   VALUES (?, 'sintético', '2025-09-01', '2026-08-31',
                           'under_review', ?)""",
                (community_id, period_id),
            ).lastrowid
            owner_id = connection.execute(
                """INSERT INTO propietarios
                   (id_comunidad,codigo_vivienda,nombre_propietario,coeficiente,activo)
                   VALUES (?, 'A', 'Persona sintética', 1, 1)""",
                (community_id,),
            ).lastrowid
            connection.commit()
            connection.close()

            def fake_export(database_path, **kwargs):
                output = run / "salidas" / "Excels_Maestros" / "Comunidad_658.xlsx"
                output.parent.mkdir(parents=True, exist_ok=True)
                output.write_bytes(b"excel-sintetico")
                db = gestor_bd.conectar(str(database_path))
                profile_id = db.execute(
                    """INSERT INTO excel_template_profiles
                       (id_comunidad,profile_key,profile_version,template_relative_path,
                        template_sha256,profile_sha256,status)
                       VALUES (?, '658_acs_v1', '1',
                               'plantillas/comunidades/658/658_acs_v1.xlsx',
                               'template', ?, 'active')""",
                    (community_id, hashlib.sha256(profile_path.read_bytes()).hexdigest()),
                ).lastrowid
                export_id = db.execute(
                    """INSERT INTO excel_export_runs
                       (id_case,id_periodo,id_template_profile,input_sha256,
                        template_sha256,output_path,status)
                       VALUES (?,?,?,?,?,?,'validated')""",
                    (case_id, period_id, profile_id, "input", "template", str(output)),
                ).lastrowid
                db.commit()
                db.close()
                return SimpleNamespace(id_export_run=export_id, output_path=output)

            def fake_distribution(database_path, **kwargs):
                db = gestor_bd.conectar(str(database_path))
                db.execute(
                    "UPDATE regularization_cases SET estado='reconciled' WHERE id_case=?",
                    (case_id,),
                )
                db.commit()
                db.close()
                return SimpleNamespace(owner_result_count=1)

            def fake_letters(database_path, **kwargs):
                batch = run / "salidas" / "cartas" / "658" / "2025-2026" / "lote_1"
                batch.mkdir(parents=True, exist_ok=False)
                letter = batch / "CARTA_A.docx"
                letter.write_bytes(b"carta-sintetica")
                db = gestor_bd.conectar(str(database_path))
                export_id = db.execute(
                    "SELECT id_export_run FROM excel_export_runs WHERE id_case=?",
                    (case_id,),
                ).fetchone()[0]
                run_id = db.execute(
                    """INSERT INTO letter_generation_runs
                       (id_case,id_periodo,id_export_run,input_sha256,template_sha256,
                        output_path,status)
                       VALUES (?,?,?,?,?,?,'completed')""",
                    (case_id, period_id, export_id, "letters", "word", str(batch)),
                ).lastrowid
                db.execute(
                    """INSERT INTO generated_letters
                       (id_letter_run,id_propietario,input_sha256,template_sha256,
                        output_path,status)
                       VALUES (?,?,?,?,?,'generated')""",
                    (run_id, owner_id, "letters", "word", str(letter)),
                )
                db.execute(
                    "UPDATE regularization_cases SET estado='deliveries_generated' WHERE id_case=?",
                    (case_id,),
                )
                db.commit()
                db.close()
                return LetterBatchResult(run_id, batch, 1, ())

            def fake_render(document, output_directory):
                output_directory.mkdir(parents=True, exist_ok=True)
                (output_directory / f"{document.stem}.pdf").write_bytes(b"pdf-sintetico")

            with patch("private_658_validation.case_workflow_actions.run_generate_excel", side_effect=fake_export), patch(
                "private_658_validation.case_workflow_actions.run_calculate_distribution",
                side_effect=fake_distribution,
            ), patch(
                "private_658_validation.case_workflow_actions.available_case_letter_concepts",
                return_value=(("acs_fixed", "Cuota"),),
            ), patch(
                "private_658_validation.case_workflow_actions.run_generate_letters",
                side_effect=fake_letters,
            ), patch("private_658_validation._render_one_page", side_effect=fake_render):
                first = resume_private_658_validation(run, project_root=project)
                second = resume_private_658_validation(run, project_root=project)

            self.assertEqual("completed", first.status)
            self.assertEqual("completed", second.status)
            self.assertEqual(1, second.generated_letters)
            self.assertEqual(1, len(list((run / "salidas" / "Excels_Maestros").glob("*.xlsx"))))
            self.assertEqual(1, len(list((run / "salidas" / "cartas").rglob("lote_*"))))
            self.assertEqual(1, len(list((run / "salidas" / "cartas").rglob("*.docx"))))

    def test_resume_with_open_issues_stays_blocked_without_outputs(self):
        import gestor_bd
        from private_658_validation import resume_private_658_validation

        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as directory:
            root = Path(directory)
            project = root / "project"
            project.mkdir()
            run = root / "private" / "validacion_bloqueada"
            profile_path = run / "config" / "excel_profiles" / "658_acs_v1.json"
            profile_path.parent.mkdir(parents=True)
            shutil.copy2(
                PROJECT_ROOT / "config" / "excel_profiles" / "658_acs_v1.json",
                profile_path,
            )
            (run / "plantillas").mkdir()
            shutil.copy2(
                PROJECT_ROOT / "plantillas" / "Plantilla_Cartas.docx",
                run / "plantillas" / "Plantilla_Cartas.docx",
            )
            database_path = run / "data" / "gestion.db"
            gestor_bd.crear_bd(str(database_path))
            db = gestor_bd.conectar(str(database_path))
            community_id = gestor_bd.obtener_o_crear_comunidad(db, "658", "Sintética")
            case_id = db.execute(
                """INSERT INTO regularization_cases
                   (id_comunidad,nombre,fecha_inicio,fecha_fin,estado)
                   VALUES (?, 'sintético', '2025-09-01', '2026-08-31', 'under_review')""",
                (community_id,),
            ).lastrowid
            document_id = db.execute(
                """INSERT INTO source_documents
                   (id_case,original_name,archived_path,sha256,document_kind,status)
                   VALUES (?, 'fuente.xlsx', 'archivo.xlsx', 'sha', 'excel_master_bootstrap',
                           'under_review')""",
                (case_id,),
            ).lastrowid
            db.execute(
                """INSERT INTO review_issues
                   (id_case,id_document,code,field_name,message,status)
                   VALUES (?,?, 'MISSING_REQUIRED_FIELD', 'campo', 'Falta dato', 'open')""",
                (case_id, document_id),
            )
            db.commit()
            db.close()

            result = resume_private_658_validation(run, project_root=project)

            self.assertEqual("blocked_open_issues", result.status)
            self.assertEqual(2, result.exit_code)
            self.assertFalse((run / "salidas").exists())

    def test_tracked_documents_do_not_contain_known_private_identifiers_or_local_paths(self):
        tracked = subprocess.check_output(
            ["git", "ls-files", "*.md", "*.txt"],
            cwd=PROJECT_ROOT,
            text=True,
            encoding="utf-8",
        ).splitlines()
        contents = "\n".join(
            (PROJECT_ROOT / relative).read_text(encoding="utf-8", errors="replace")
            for relative in tracked
        ).lower()
        for private_fragment in (
            "658 estudio",
            "estudio acs-cal",
            "monasterio de poblet",
            "pintor aguayo",
            "reina fabiola",
            "658 regularizacion 2025 2026",
            "comunidad_644_liquidado",
        ):
            with self.subTest(private_fragment=private_fragment):
                self.assertNotIn(private_fragment, contents)
        self.assertIsNone(
            re.search(r"[a-z]:\\users\\[^\\\s]+", contents),
            "La documentación versionada contiene una ruta local de usuario",
        )


if __name__ == "__main__":
    unittest.main()
