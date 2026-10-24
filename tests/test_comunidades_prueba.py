"""La instalación de prueba se crea entera y en el estado que promete."""

import sqlite3
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
SCRIPT = PROJECT_ROOT / "scripts" / "crear_comunidades_prueba.py"


class CommunityTestInstallationTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.directory = tempfile.TemporaryDirectory(ignore_cleanup_errors=True)
        cls.home = Path(cls.directory.name).resolve() / "prueba"
        # Proceso aparte: el script fija REGULARIZACIONES_HOME y no debe
        # alterar el resto de la suite.
        cls.result = subprocess.run(
            [sys.executable, str(SCRIPT), "--destino", str(cls.home)],
            capture_output=True, text=True, encoding="utf-8", timeout=600,
            env={**__import__("os").environ, "PYTHONUTF8": "1"},
        )

    @classmethod
    def tearDownClass(cls):
        cls.directory.cleanup()

    def _connection(self):
        connection = sqlite3.connect(self.home / "data" / "gestion.db")
        connection.row_factory = sqlite3.Row
        return connection

    def test_script_finishes(self):
        self.assertEqual(0, self.result.returncode, self.result.stderr[-3000:])

    def test_each_community_is_in_its_promised_state(self):
        with self._connection() as connection:
            states = dict(connection.execute(
                """SELECT c.codigo, r.estado FROM regularization_cases r
                   JOIN comunidades c USING(id_comunidad)"""
            ).fetchall())
            issues = {
                row["code"] for row in connection.execute(
                    """SELECT i.code FROM review_issues i JOIN regularization_cases r USING(id_case)
                       JOIN comunidades c USING(id_comunidad) WHERE c.codigo='902' AND i.status='open'"""
                )
            }
        self.assertEqual(
            {"901": "ready_for_calculation", "902": "under_review", "903": "draft"}, states,
        )
        self.assertEqual({"INVOICE_OUTSIDE_PERIOD", "MISSING_REQUIRED_FIELD", "PROVIDER_UNKNOWN"}, issues)

    def test_tray_folder_has_every_kind_of_source(self):
        files = sorted(path.name for path in (self.home / "fuentes_903").iterdir())
        self.assertTrue(any("Listado de propietarios" in name for name in files))
        self.assertTrue(any(name.endswith(".xlsx") for name in files))
        for provider in ("NATURGY", "ENDESA", "AGUA", "MANTENIMIENTO", "LECTURAS CONTADORES"):
            self.assertEqual(12 if provider != "AGUA" else 6,
                             sum(provider in name for name in files), provider)

    def test_only_invented_contact_data(self):
        with self._connection() as connection:
            emails = [row[0] for row in connection.execute("SELECT email FROM propietarios")]
        self.assertTrue(emails)
        self.assertTrue(all(email.endswith("@example.invalid") for email in emails))

    def test_refuses_to_delete_a_folder_it_did_not_create(self):
        with tempfile.TemporaryDirectory() as other:
            (Path(other) / "algo.txt").write_text("datos", encoding="utf-8")
            result = subprocess.run(
                [sys.executable, str(SCRIPT), "--destino", other, "--reemplazar"],
                capture_output=True, text=True, encoding="utf-8", timeout=120,
            )
            self.assertNotEqual(0, result.returncode)
            self.assertTrue((Path(other) / "algo.txt").is_file())


if __name__ == "__main__":
    unittest.main()
