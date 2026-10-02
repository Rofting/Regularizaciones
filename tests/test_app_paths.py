import sqlite3
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "core"))

from app_paths import ApplicationPaths, PUBLIC_RESOURCES


class ApplicationPathsTest(unittest.TestCase):
    def test_packaged_default_and_portable_override_support_accents(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            app = root / "Programa" / "Regularizaciones.exe"
            paths = ApplicationPaths.resolve({}, executable=app,
                local_app_data=root / "Datos de Álvaro", resources=ROOT, frozen=True)
            self.assertEqual(root / "Datos de Álvaro" / "Regularizaciones" / "data" / "gestion.db",
                             paths.database)
            portable = ApplicationPaths.resolve(
                {"REGULARIZACIONES_HOME": str(root / "Mi despacho áé")},
                executable=app, resources=ROOT, frozen=True)
            portable.prepare()
            self.assertTrue((portable.home / "plantillas" / "Plantilla_Cartas.docx").is_file())
            self.assertTrue((portable.home / "config" / "proveedores.json").is_file())
            self.assertTrue(portable.sources.is_dir())
            self.assertTrue(portable.backups.is_dir())
            self.assertTrue(portable.logs.is_dir())
            self.assertTrue(portable.outputs.is_dir())
            for private in ("data/gestion.db", "config/ui_prefs.json",
                            "config/proveedores_despacho.json"):
                self.assertNotIn(private, PUBLIC_RESOURCES)
                self.assertFalse((portable.home / private).exists())

    def test_legacy_database_is_copied_with_integrity_and_not_overwritten(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            legacy = root / "Programa anterior"
            (legacy / "data").mkdir(parents=True)
            old = legacy / "data" / "gestion.db"
            with sqlite3.connect(old) as con:
                con.execute("CREATE TABLE nota (texto TEXT)")
                con.execute("INSERT INTO nota VALUES ('conservado')")
            paths = ApplicationPaths(ROOT, root / "Nuevo despacho", legacy / "Regularizaciones.exe")
            paths.prepare()
            with sqlite3.connect(paths.database) as con:
                self.assertEqual("conservado", con.execute("SELECT texto FROM nota").fetchone()[0])
                con.execute("INSERT INTO nota VALUES ('nuevo')")
            paths.prepare()
            with sqlite3.connect(paths.database) as con:
                self.assertEqual(2, con.execute("SELECT COUNT(*) FROM nota").fetchone()[0])

    def test_unwritable_home_fails_before_creating_window(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            blocked = root / "archivo"
            blocked.write_text("no es un directorio", encoding="utf-8")
            paths = ApplicationPaths(ROOT, blocked, root / "Regularizaciones.exe")
            with self.assertRaisesRegex(RuntimeError, "No se puede escribir"):
                paths.prepare()


if __name__ == "__main__":
    unittest.main()
