import sqlite3
import sys
import tempfile
import unittest
from contextlib import closing, redirect_stdout
from io import StringIO
from pathlib import Path
from unittest.mock import patch


PROJECT_ROOT = Path(__file__).resolve().parents[1]
CORE_DIR = PROJECT_ROOT / "core"
if str(CORE_DIR) not in sys.path:
    sys.path.insert(0, str(CORE_DIR))

import gestor_bd


EXPECTED_TABLES = {
    "schema_migrations",
    "import_batches",
    "source_values",
    "regularization_concepts",
    "owner_concept_results",
    "reconciliations",
    "community_letter_settings",
    "invoice_components",
    "period_parameters",
    "excel_template_profiles",
    "excel_export_runs",
    "letter_generation_runs",
    "generated_letters",
}


class DatabaseMigrationTest(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory(ignore_cleanup_errors=True)
        self.database_path = Path(self.directory.name) / "gestion.db"

    def tearDown(self):
        self.directory.cleanup()

    def _connect(self):
        connection = sqlite3.connect(self.database_path)
        connection.row_factory = sqlite3.Row
        connection.execute("PRAGMA foreign_keys = ON")
        return connection

    def test_crear_bd_applies_version_one_to_empty_database(self):
        with redirect_stdout(StringIO()):
            gestor_bd.crear_bd(str(self.database_path))

        with closing(self._connect()) as connection:
            tables = {
                row[0]
                for row in connection.execute(
                    "SELECT name FROM sqlite_master WHERE type='table'"
                )
            }
            columns = {
                row[1]
                for row in connection.execute("PRAGMA table_info(propietarios)")
            }
            versions = connection.execute(
                "SELECT version FROM schema_migrations ORDER BY version"
            ).fetchall()
            concepts = connection.execute(
                "SELECT concept_key FROM regularization_concepts ORDER BY display_order"
            ).fetchall()

        self.assertTrue(EXPECTED_TABLES.issubset(tables))
        self.assertIn("email", columns)
        self.assertEqual([1, 2, 3], [row[0] for row in versions])
        self.assertEqual(
            [
                "acs_fixed",
                "acs_variable",
                "heating_fixed",
                "heating_variable",
                "extraordinary_expense",
                "adjustment",
                "credit",
                "other",
            ],
            [row[0] for row in concepts],
        )

    def test_migration_is_idempotent_for_current_database(self):
        with redirect_stdout(StringIO()):
            gestor_bd.crear_bd(str(self.database_path))
            gestor_bd.crear_bd(str(self.database_path))

        with closing(self._connect()) as connection:
            versions = connection.execute(
                "SELECT version FROM schema_migrations ORDER BY version"
            ).fetchall()
            concept_count = connection.execute(
                "SELECT COUNT(*) FROM regularization_concepts"
            ).fetchone()[0]

        self.assertEqual([1, 2, 3], [row[0] for row in versions])
        self.assertEqual(8, concept_count)

    def test_migration_two_creates_case_and_review_tables(self):
        with closing(self._connect()) as connection:
            for statement in gestor_bd.TABLAS:
                connection.execute(statement)
            connection.commit()
            version = gestor_bd.aplicar_migraciones(connection)
            tables = {
                row[0]
                for row in connection.execute(
                    "SELECT name FROM sqlite_master WHERE type='table'"
                )
            }
            columns = {
                row[1]
                for row in connection.execute("PRAGMA table_info(review_issues)")
            }
            indexes = {
                row[0]
                for row in connection.execute(
                    "SELECT name FROM sqlite_master WHERE type='index'"
                )
            }

        self.assertEqual(version, 3)
        self.assertTrue({
            "regularization_cases", "source_documents", "extraction_candidates",
            "review_issues", "manual_corrections",
        }.issubset(tables))
        self.assertTrue({"id_case", "id_document", "field_name", "status"}.issubset(columns))
        self.assertTrue({
            "idx_cases_community",
            "idx_documents_case",
            "idx_issues_case_open",
        }.issubset(indexes))

    def test_migration_two_is_idempotent(self):
        with closing(self._connect()) as connection:
            for statement in gestor_bd.TABLAS:
                connection.execute(statement)
            connection.commit()
            self.assertEqual(3, gestor_bd.aplicar_migraciones(connection))
            self.assertEqual(3, gestor_bd.aplicar_migraciones(connection))
            versions = connection.execute(
                "SELECT version FROM schema_migrations ORDER BY version"
            ).fetchall()

        self.assertEqual([1, 2, 3], [row[0] for row in versions])

    def test_migration_three_links_cases_and_creates_export_audit_tables(self):
        with closing(self._connect()) as connection:
            for statement in gestor_bd.TABLAS:
                connection.execute(statement)
            connection.commit()

            version = gestor_bd.aplicar_migraciones(connection)
            tables = {
                row[0]
                for row in connection.execute(
                    "SELECT name FROM sqlite_master WHERE type='table'"
                )
            }
            case_columns = {
                row[1]
                for row in connection.execute(
                    "PRAGMA table_info(regularization_cases)"
                )
            }
            indexes = {
                row[0]
                for row in connection.execute(
                    "SELECT name FROM sqlite_master WHERE type='index'"
                )
            }

        self.assertEqual(3, version)
        self.assertTrue({
            "invoice_components",
            "period_parameters",
            "excel_template_profiles",
            "excel_export_runs",
            "letter_generation_runs",
            "generated_letters",
        }.issubset(tables))
        self.assertIn("id_periodo", case_columns)
        self.assertIn("idx_case_period", indexes)

    def test_failed_migration_rolls_back_every_statement(self):
        with closing(self._connect()) as connection:
            for statement in gestor_bd.TABLAS:
                connection.execute(statement)
            connection.commit()

            def broken_migration(target):
                target.execute("CREATE TABLE partial_change (id INTEGER)")
                target.execute("THIS IS NOT VALID SQL")

            with patch("db_migrations.MIGRATIONS", {1: broken_migration}):
                with self.assertRaises(sqlite3.OperationalError):
                    gestor_bd.aplicar_migraciones(connection)

            partial = connection.execute(
                "SELECT name FROM sqlite_master WHERE name='partial_change'"
            ).fetchone()
            version_table = connection.execute(
                "SELECT name FROM sqlite_master WHERE name='schema_migrations'"
            ).fetchone()

        self.assertIsNone(partial)
        self.assertIsNone(version_table)


if __name__ == "__main__":
    unittest.main()
