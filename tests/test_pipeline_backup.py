"""El flujo desatendido protege la base antes de tocar archivos o datos."""

import sqlite3
import sys
import tempfile
import unittest
from contextlib import closing
from pathlib import Path
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "core"))
import database_backup
import pipeline


class PipelineBackupTest(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.addCleanup(self.directory.cleanup)
        self.root = Path(self.directory.name)
        self.database = self.root / "gestion.db"
        with closing(sqlite3.connect(self.database)) as connection:
            connection.execute("CREATE TABLE marker(value TEXT)")
            connection.execute("INSERT INTO marker VALUES ('antes del lote')")
            connection.commit()
        self.extra = self.root / "extra"
        self.extra.mkdir()
        (self.extra / "factura.pdf").write_bytes(b"documento pendiente")
        self.routes = {name: self.root / name for name in (
            "entrada", "procesados", "cuarentena", "excels", "cartas", "logs",
        )}
        for folder in self.routes.values():
            folder.mkdir()
        self.routes["bd"] = self.database
        self.routes["plantilla_cartas"] = self.root / "modelo.docx"

    def test_backup_failure_stops_before_copying_or_ingesting(self):
        with patch.object(pipeline, "cargar_rutas", return_value=self.routes), \
             patch.object(pipeline.database_backup, "backup_database",
                          side_effect=database_backup.DatabaseBackupError("disco lleno")), \
             patch.object(pipeline, "ingerir_entrada") as ingest:
            result = pipeline.procesar_todo(carpeta_extra=str(self.extra), log_callback=lambda *_: None)

        self.assertFalse(result["ok"])
        self.assertIn("disco lleno", result["errores"][0])
        self.assertEqual([], list(self.routes["entrada"].iterdir()))
        ingest.assert_not_called()

    def test_successful_backup_records_state_before_copying(self):
        with patch.object(pipeline, "cargar_rutas", return_value=self.routes), \
             patch.object(pipeline, "ingerir_entrada", return_value=({}, {})):
            result = pipeline.procesar_todo(carpeta_extra=str(self.extra), log_callback=lambda *_: None)

        self.assertTrue(result["ok"])
        self.assertTrue((self.routes["entrada"] / "factura.pdf").exists())
        copies = list((self.root / "backups").glob("*.db"))
        self.assertEqual(1, len(copies))
        database_backup.verify_backup(copies[0])
        with closing(sqlite3.connect(copies[0])) as connection:
            self.assertEqual(("antes del lote",), connection.execute(
                "SELECT value FROM marker"
            ).fetchone())


if __name__ == "__main__":
    unittest.main()
