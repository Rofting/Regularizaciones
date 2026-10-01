import json
import shutil
import sqlite3
import sys
import tempfile
import unittest
from contextlib import closing
from pathlib import Path
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "core"))
import database_backup
import db_migrations


class DatabaseBackupTest(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.addCleanup(self.directory.cleanup)
        self.root = Path(self.directory.name).resolve()
        self.database = self.root / "gestion.db"
        with closing(sqlite3.connect(self.database)) as con:
            con.execute("CREATE TABLE marker(value TEXT)")
            con.execute("INSERT INTO marker VALUES ('original')")
            con.commit()

    def test_wal_backup_can_be_restored_without_sidecar_files(self):
        with closing(sqlite3.connect(self.database)) as con:
            con.execute("PRAGMA journal_mode=WAL")
            con.execute("PRAGMA wal_autocheckpoint=0")
            con.execute("INSERT INTO marker VALUES ('confirmado en WAL')")
            con.commit()
            self.assertTrue(Path(str(self.database) + "-wal").is_file())
            result = database_backup.backup_database(self.database, reason="before_excel")
            restored = self.root / "restaurada.db"
            shutil.copyfile(result.backup_path, restored)
            database_backup.verify_backup(restored)
            with closing(sqlite3.connect(restored)) as recovered:
                self.assertEqual([('original',), ('confirmado en WAL',)],
                                 recovered.execute("SELECT value FROM marker").fetchall())
        record = json.loads(result.backup_path.with_suffix('.db.json').read_text())
        self.assertEqual("before_excel", record['reason'])

    def test_retention_preserves_manual_corrupt_and_other_database_backups(self):
        (self.root / "backup_settings.json").write_text('{"max_backups": 3}')
        first = database_backup.backup_database(self.database, reason="startup")
        second = database_backup.backup_database(self.database, reason="startup")
        first.backup_path.write_bytes(b"copia alterada que no debe borrarse")
        manual = self.root / "backups" / "manual.db"
        shutil.copyfile(self.database, manual)
        other = self.root / "otra.db"
        shutil.copyfile(self.database, other)
        other_copy = database_backup.backup_database(other, reason="startup")
        for _ in range(3):
            last = database_backup.backup_database(self.database, reason="before_excel")
        self.assertFalse(second.backup_path.exists())
        self.assertTrue(first.backup_path.exists())
        self.assertTrue(manual.exists())
        self.assertTrue(other_copy.backup_path.exists())
        self.assertEqual(1, len(last.removed_paths))
        self.assertTrue(last.cleanup_warnings)

    def test_failed_verification_preserves_source_and_previous_backup(self):
        previous = database_backup.backup_database(self.database, reason="startup")
        original = self.database.read_bytes()
        with patch('database_backup.verify_backup', side_effect=database_backup.DatabaseBackupError('fallo')):
            with self.assertRaises(database_backup.DatabaseBackupError):
                database_backup.backup_database(self.database, reason="before_excel")
        self.assertEqual(original, self.database.read_bytes())
        self.assertEqual([previous.backup_path], list((self.root / 'backups').glob('*.db')))

    def test_invalid_retention_stops_before_creating_or_deleting_copies(self):
        for settings in ('{"max_backups": 0}', '{"max_backups": true}', '[]', '{broken'):
            with self.subTest(settings=settings):
                (self.root / 'backup_settings.json').write_text(settings)
                with self.assertRaises(database_backup.DatabaseBackupError):
                    database_backup.backup_database(self.database, reason="startup")
                self.assertFalse((self.root / 'backups').exists())

    def test_memory_connection_does_not_create_files(self):
        with closing(sqlite3.connect(':memory:')) as con:
            self.assertIsNone(database_backup.backup_connection(con, reason="before_migration"))
        self.assertFalse((self.root / 'backups').exists())

    def test_cleanup_failure_keeps_verified_new_copy(self):
        (self.root / 'backup_settings.json').write_text('{"max_backups": 3}')
        for _ in range(3):
            database_backup.backup_database(self.database, reason="startup")
        with patch('database_backup.Path.unlink', side_effect=PermissionError('no se puede limpiar')):
            result = database_backup.backup_database(self.database, reason="startup")
        self.assertTrue(result.backup_path.is_file())
        self.assertTrue(result.cleanup_warnings)

    def test_migration_backup_contains_original_schema_and_data(self):
        with closing(sqlite3.connect(self.database)) as con:
            con.execute('CREATE TABLE schema_migrations(version INTEGER PRIMARY KEY)')
            con.execute('INSERT INTO schema_migrations VALUES (1)')
            con.commit()
            migrations = {1: lambda c: None, 2: lambda c: c.execute('ALTER TABLE marker ADD COLUMN extra TEXT')}
            with patch.dict(db_migrations.MIGRATIONS, migrations, clear=True):
                db_migrations.migrate(con)
                db_migrations.migrate(con)
            self.assertEqual(['value', 'extra'], [r[1] for r in con.execute('PRAGMA table_info(marker)')])
        copies = list((self.root / 'backups').glob('*.db'))
        self.assertEqual(1, len(copies))
        with closing(sqlite3.connect(copies[0])) as con:
            self.assertEqual(['value'], [r[1] for r in con.execute('PRAGMA table_info(marker)')])
            self.assertEqual('original', con.execute('SELECT value FROM marker').fetchone()[0])

    def test_backup_failure_prevents_migration(self):
        with closing(sqlite3.connect(self.database)) as con:
            migration = unittest.mock.Mock()
            with patch.dict(db_migrations.MIGRATIONS, {1: migration}, clear=True), \
                 patch('database_backup.backup_connection', side_effect=database_backup.DatabaseBackupError('fallo')):
                with self.assertRaises(database_backup.DatabaseBackupError):
                    db_migrations.migrate(con)
            migration.assert_not_called()
            self.assertEqual('original', con.execute('SELECT value FROM marker').fetchone()[0])

    def test_empty_new_database_does_not_need_a_migration_backup(self):
        empty = self.root / 'nueva.db'
        with closing(sqlite3.connect(empty)) as con:
            con.execute('CREATE TABLE empty_table(value TEXT)')
            with patch.dict(db_migrations.MIGRATIONS, {1: lambda c: None}, clear=True):
                db_migrations.migrate(con)
        self.assertFalse((self.root / 'backups').exists())


if __name__ == '__main__':
    unittest.main()
