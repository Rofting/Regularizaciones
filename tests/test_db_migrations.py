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
    "case_import_batches",
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
        self.assertEqual([1, 2, 3, 4, 5, 6, 7, 8, 9, 10], [row[0] for row in versions])
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

        self.assertEqual([1, 2, 3, 4, 5, 6, 7, 8, 9, 10], [row[0] for row in versions])
        self.assertEqual(8, concept_count)

    def test_migration_nine_creates_an_immutable_case_history(self):
        with redirect_stdout(StringIO()):
            gestor_bd.crear_bd(str(self.database_path))

        with closing(self._connect()) as connection:
            tables = {
                row[0] for row in connection.execute(
                    "SELECT name FROM sqlite_master WHERE type='table'"
                )
            }
            columns = {
                row[1] for row in connection.execute(
                    "PRAGMA table_info(case_history_events)"
                )
            }
            triggers = {
                row[0] for row in connection.execute(
                    "SELECT name FROM sqlite_master WHERE type='trigger'"
                )
            }

        self.assertIn("case_history_events", tables)
        self.assertTrue({"id_case", "id_document", "event_type", "details_json", "created_at"}.issubset(columns))
        self.assertTrue({
            "history_source_registered",
            "history_source_status_changed",
            "history_manual_correction",
            "history_case_status_changed",
        }.issubset(triggers))

    def test_migration_ten_keeps_closed_issue_history_but_rejects_duplicate_open_issue(self):
        with redirect_stdout(StringIO()):
            gestor_bd.crear_bd(str(self.database_path))

        with closing(self._connect()) as connection:
            connection.execute("INSERT INTO comunidades(codigo,nombre) VALUES ('M10','Migración diez')")
            connection.execute(
                """INSERT INTO regularization_cases
                   (id_comunidad,nombre,fecha_inicio,fecha_fin,estado)
                   VALUES (1,'Caso','2026-01-01','2026-01-31','under_review')"""
            )
            connection.execute(
                """INSERT INTO source_documents
                   (id_case,original_name,archived_path,sha256,document_kind,status)
                   VALUES (1,'lecturas.xls','archivo/lecturas.xls','m10-source','reading','under_review')"""
            )
            values = (1, 1, "MISSING_REQUIRED_FIELD", "fecha_fin", "Falta fecha final")
            connection.execute(
                """INSERT INTO review_issues
                   (id_case,id_document,code,field_name,message,status,origin)
                   VALUES (?,?,?,?,?,'resolved','automatic')""",
                values,
            )
            connection.execute(
                """INSERT INTO review_issues
                   (id_case,id_document,code,field_name,message,status,origin)
                   VALUES (?,?,?,?,?,'resolved','automatic')""",
                values,
            )
            connection.execute(
                """INSERT INTO review_issues
                   (id_case,id_document,code,field_name,message,status,origin)
                   VALUES (?,?,?,?,?,'open','automatic')""",
                values,
            )
            with self.assertRaises(sqlite3.IntegrityError):
                connection.execute(
                    """INSERT INTO review_issues
                       (id_case,id_document,code,field_name,message,status,origin)
                       VALUES (?,?,?,?,?,'open','automatic')""",
                    values,
                )

    def test_migration_ten_creates_reading_observations_with_source_and_effective_links(self):
        with redirect_stdout(StringIO()):
            gestor_bd.crear_bd(str(self.database_path))

        with closing(self._connect()) as connection:
            columns = {
                row[1] for row in connection.execute("PRAGMA table_info(reading_observations)")
            }
            self.assertTrue({
                "id_propietario", "tipo", "fecha_lectura", "observed_value",
                "source_path", "id_document", "status", "effective_reading_id",
            }.issubset(columns))

    def test_migration_nine_backfills_source_history_for_existing_periods(self):
        import db_migrations

        with closing(self._connect()) as connection:
            for statement in gestor_bd.TABLAS:
                connection.execute(statement)
            connection.execute(
                "CREATE TABLE schema_migrations(version INTEGER PRIMARY KEY, applied_at TEXT)"
            )
            connection.commit()
            for version in range(1, 9):
                db_migrations.MIGRATIONS[version](connection)
            connection.execute("INSERT INTO comunidades(codigo,nombre) VALUES ('HIS','Histórica')")
            connection.execute(
                "INSERT INTO periodos(id_comunidad,nombre,fecha_inicio,fecha_fin) VALUES (1,'2024','2024-01-01','2024-12-31')"
            )
            connection.execute(
                """INSERT INTO regularization_cases(id_comunidad,id_periodo,nombre,fecha_inicio,fecha_fin,estado)
                   VALUES (1,1,'2024','2024-01-01','2024-12-31','ready_for_calculation')"""
            )
            connection.execute(
                """INSERT INTO source_documents(id_case,original_name,archived_path,sha256,document_kind,status)
                   VALUES (1,'factura.pdf','archivo/factura.pdf','hash-historico','invoice','validated')"""
            )
            connection.commit()

            db_migrations.MIGRATIONS[9](connection)
            events = connection.execute(
                "SELECT event_type, details_json FROM case_history_events WHERE id_case=1"
            ).fetchall()

        self.assertEqual(1, len(events))
        self.assertEqual("source_registered", events[0][0])
        self.assertIn("factura.pdf", events[0][1])

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
            document_columns = {
                row[1]
                for row in connection.execute("PRAGMA table_info(source_documents)")
            }
            candidate_columns = {
                row[1]
                for row in connection.execute("PRAGMA table_info(extraction_candidates)")
            }
            issue_columns = {
                row[1]
                for row in connection.execute("PRAGMA table_info(review_issues)")
            }
            indexes = {
                row[0]
                for row in connection.execute(
                    "SELECT name FROM sqlite_master WHERE type='index'"
                )
            }

        self.assertEqual(version, 10)
        self.assertTrue({
            "regularization_cases", "source_documents", "extraction_candidates",
            "review_issues", "manual_corrections",
        }.issubset(tables))
        self.assertTrue({"id_case", "id_document", "field_name", "status"}.issubset(columns))
        self.assertIn("classification_confidence", document_columns)
        self.assertIn("source_context", candidate_columns)
        self.assertIn("origin", issue_columns)
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
            self.assertEqual(10, gestor_bd.aplicar_migraciones(connection))
            self.assertEqual(10, gestor_bd.aplicar_migraciones(connection))
            versions = connection.execute(
                "SELECT version FROM schema_migrations ORDER BY version"
            ).fetchall()

        self.assertEqual([1, 2, 3, 4, 5, 6, 7, 8, 9, 10], [row[0] for row in versions])

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

        self.assertEqual(10, version)
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

    def test_migration_five_records_explicit_estimation_approval(self):
        with closing(self._connect()) as connection:
            for statement in gestor_bd.TABLAS:
                connection.execute(statement)
            connection.commit()
            gestor_bd.aplicar_migraciones(connection)
            columns = {
                row[1] for row in connection.execute("PRAGMA table_info(lecturas_vecino)")
            }

        self.assertTrue({"approved_by", "approved_at"}.issubset(columns))

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
