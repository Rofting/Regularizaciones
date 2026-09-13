import sqlite3
import sys
import tempfile
import unittest
from contextlib import closing
from pathlib import Path
from unittest.mock import patch


PROJECT_ROOT = Path(__file__).resolve().parents[1]
CORE_DIR = PROJECT_ROOT / "core"
if str(CORE_DIR) not in sys.path:
    sys.path.insert(0, str(CORE_DIR))

import database_reset


class DatabaseResetTest(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.root = Path(self.directory.name)
        self.database = self.root / "gestion.db"
        self.backups = self.root / "backups"
        self._write_marker(self.database, "old")

    def tearDown(self):
        self.directory.cleanup()

    @staticmethod
    def _write_marker(path, marker):
        with closing(sqlite3.connect(path)) as connection:
            connection.execute("CREATE TABLE marker (value TEXT NOT NULL)")
            connection.execute("INSERT INTO marker(value) VALUES (?)", (marker,))
            connection.commit()

    @staticmethod
    def read_marker(path):
        with closing(sqlite3.connect(path)) as connection:
            return connection.execute("SELECT value FROM marker").fetchone()[0]

    def initialise(self, path):
        self._write_marker(path, "new")

    @staticmethod
    def broken_initialise(path):
        raise RuntimeError("initialisation failed")

    def test_reset_copies_and_verifies_old_database_before_replacing_it(self):
        result = database_reset.reset_database(
            self.database, backup_root=self.backups, initialise=self.initialise
        )

        self.assertTrue(result.backup_path.exists())
        self.assertEqual("old", self.read_marker(result.backup_path))
        self.assertEqual("new", self.read_marker(self.database))
        self.assertEqual(self.database, result.new_path)

    def test_reset_keeps_existing_database_when_backup_verification_fails(self):
        with patch.object(
            database_reset,
            "_verify_database",
            side_effect=database_reset.DatabaseResetError("integrity check failed"),
        ):
            with self.assertRaises(database_reset.DatabaseResetError):
                database_reset.reset_database(
                    self.database, backup_root=self.backups, initialise=self.initialise
                )

        self.assertEqual("old", self.read_marker(self.database))

    def test_reset_keeps_existing_database_when_new_initialisation_fails(self):
        with self.assertRaises(database_reset.DatabaseResetError):
            database_reset.reset_database(
                self.database,
                backup_root=self.backups,
                initialise=self.broken_initialise,
            )

        self.assertEqual("old", self.read_marker(self.database))

    def test_reset_generates_unique_backup_names(self):
        first = database_reset.reset_database(
            self.database, backup_root=self.backups, initialise=self.initialise
        )
        second = database_reset.reset_database(
            self.database, backup_root=self.backups, initialise=self.initialise
        )

        self.assertNotEqual(first.backup_path, second.backup_path)
        self.assertTrue(first.backup_path.exists())
        self.assertTrue(second.backup_path.exists())


if __name__ == "__main__":
    unittest.main()
