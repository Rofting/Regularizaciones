import os
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path


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

        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            with self.assertRaisesRegex(PrivateValidationBlockedError, "segura"):
                prepare_validation_run(root, project_root=root)
            with self.assertRaisesRegex(PrivateValidationBlockedError, "segura"):
                prepare_validation_run(root.parent, project_root=root)

    def test_safe_empty_root_creates_one_new_run_and_copies_only_public_assets(self):
        from private_658_validation import prepare_validation_run

        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            project = root / "project"
            (project / "config").mkdir(parents=True)
            (project / "config" / "sample.json").write_text("{}", encoding="utf-8")
            (project / "plantillas").mkdir()
            (project / "plantillas" / "Plantilla_Cartas.docx").write_bytes(b"public-template")
            destination = root / "private-validation"

            run = prepare_validation_run(destination, project_root=project)

            self.assertTrue(run.is_dir())
            self.assertTrue((run / "config" / "sample.json").is_file())
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

    def test_repository_does_not_track_private_source_names(self):
        tracked = subprocess.check_output(
            ["git", "ls-files"], cwd=PROJECT_ROOT, text=True, encoding="utf-8"
        ).lower()
        for private_fragment in ("658 estudio", "estudio acs-cal", "reina fabiola"):
            with self.subTest(private_fragment=private_fragment):
                self.assertNotIn(private_fragment, tracked)


if __name__ == "__main__":
    unittest.main()
