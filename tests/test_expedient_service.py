import sqlite3
import sys
import tempfile
import unittest
from contextlib import redirect_stdout
from datetime import date
from io import StringIO
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
CORE_DIR = PROJECT_ROOT / "core"
if str(CORE_DIR) not in sys.path:
    sys.path.insert(0, str(CORE_DIR))

import gestor_bd
import expedient_service


class ExpedientServiceTest(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory(ignore_cleanup_errors=True)
        self.database_path = Path(self.directory.name) / "gestion.db"
        with redirect_stdout(StringIO()):
            gestor_bd.crear_bd(str(self.database_path))
        self.connection = sqlite3.connect(self.database_path)
        self.connection.row_factory = sqlite3.Row
        self.connection.execute("PRAGMA foreign_keys = ON")
        self.community_id = gestor_bd.obtener_o_crear_comunidad(
            self.connection, "SINTETICA", "Comunidad sintética"
        )

    def tearDown(self):
        self.connection.close()
        self.directory.cleanup()

    def _new_connection(self, timeout=5.0):
        connection = sqlite3.connect(self.database_path, timeout=timeout)
        connection.row_factory = sqlite3.Row
        connection.execute("PRAGMA foreign_keys = ON")
        return connection

    def test_create_case_preserves_any_exact_date_range(self):
        created = expedient_service.create_case(
            self.connection, self.community_id, name="Invierno parcial",
            start_date=date(2026, 1, 15), end_date=date(2026, 3, 14),
        )

        self.assertEqual(created.start_date, date(2026, 1, 15))
        self.assertEqual(created.end_date, date(2026, 3, 14))
        self.assertEqual(created.status, "draft")

    def test_create_case_rejects_end_before_start(self):
        with self.assertRaisesRegex(ValueError, "fin posterior"):
            expedient_service.create_case(
                self.connection, self.community_id, name="incorrecto",
                start_date=date(2026, 5, 1), end_date=date(2026, 4, 30),
            )

    def test_register_source_document_archives_once_and_starts_gathering(self):
        case = expedient_service.create_case(
            self.connection, self.community_id, name="Fuentes",
            start_date=date(2026, 1, 1), end_date=date(2026, 1, 31),
        )
        source_path = Path(self.directory.name) / "origen.pdf"
        source_path.write_bytes(b"%PDF-1.4 prueba")
        archive_root = Path(self.directory.name) / "expedientes"

        first, first_created = expedient_service.register_source_document(
            self.connection, case.id_case, source_path=source_path,
            archive_root=archive_root, document_kind="invoice",
        )
        second, second_created = expedient_service.register_source_document(
            self.connection, case.id_case, source_path=source_path,
            archive_root=archive_root, document_kind="invoice",
        )

        archived_files = list((archive_root / str(case.id_case) / "fuentes").iterdir())
        self.assertTrue(first_created)
        self.assertFalse(second_created)
        self.assertEqual(first.id_document, second.id_document)
        self.assertEqual(1, len(archived_files))
        self.assertEqual("gathering_sources", expedient_service.get_case(
            self.connection, case.id_case
        ).status)

    def test_register_source_document_rejects_outer_transaction_without_side_effects(self):
        case = expedient_service.create_case(
            self.connection, self.community_id, name="Transacción externa",
            start_date=date(2026, 3, 1), end_date=date(2026, 3, 31),
        )
        source_path = Path(self.directory.name) / "externa.pdf"
        source_path.write_bytes(b"%PDF-1.4 externa")
        archive_root = Path(self.directory.name) / "expedientes"

        self.connection.execute("BEGIN")
        try:
            with self.assertRaisesRegex(RuntimeError, "transacción externa"):
                expedient_service.register_source_document(
                    self.connection, case.id_case, source_path=source_path,
                    archive_root=archive_root, document_kind="invoice",
                )
        finally:
            self.connection.rollback()

        self.assertFalse((archive_root / str(case.id_case) / "fuentes").exists())
        self.assertEqual(0, self.connection.execute(
            "SELECT COUNT(*) FROM source_documents WHERE id_case = ?", (case.id_case,)
        ).fetchone()[0])

    def test_register_source_document_serializes_before_deduplication(self):
        case = expedient_service.create_case(
            self.connection, self.community_id, name="Concurrencia",
            start_date=date(2026, 4, 1), end_date=date(2026, 4, 30),
        )
        source_path = Path(self.directory.name) / "compartida.pdf"
        source_path.write_bytes(b"%PDF-1.4 compartida")
        archive_root = Path(self.directory.name) / "expedientes"
        first, first_created = expedient_service.register_source_document(
            self.connection, case.id_case, source_path=source_path,
            archive_root=archive_root, document_kind="invoice",
        )
        archived_bytes = first.archived_path.read_bytes()
        blocker = self._new_connection()
        second = self._new_connection(timeout=0.0)
        blocker.execute("BEGIN IMMEDIATE")
        try:
            with self.assertRaises(sqlite3.OperationalError):
                expedient_service.register_source_document(
                    second, case.id_case, source_path=source_path,
                    archive_root=archive_root, document_kind="invoice",
                )
        finally:
            blocker.commit()
        try:
            duplicate, duplicate_created = expedient_service.register_source_document(
                second, case.id_case, source_path=source_path,
                archive_root=archive_root, document_kind="invoice",
            )
        finally:
            blocker.close()
            second.close()

        self.assertTrue(first_created)
        self.assertFalse(duplicate_created)
        self.assertEqual(first.id_document, duplicate.id_document)
        self.assertEqual(archived_bytes, first.archived_path.read_bytes())
        self.assertEqual(1, len(list(first.archived_path.parent.iterdir())))

    def test_set_case_status_rejects_skipping_gathering_sources(self):
        case = expedient_service.create_case(
            self.connection, self.community_id, name="Estados",
            start_date=date(2026, 2, 1), end_date=date(2026, 2, 28),
        )

        with self.assertRaisesRegex(ValueError, "gathering_sources"):
            expedient_service.set_case_status(
                self.connection, case.id_case, "ready_for_calculation"
            )

    def test_list_cases_orders_by_start_date_then_id_and_unknown_case_fails(self):
        earliest = expedient_service.create_case(
            self.connection, self.community_id, name="Primero",
            start_date=date(2026, 1, 1), end_date=date(2026, 1, 31),
        )
        same_date = expedient_service.create_case(
            self.connection, self.community_id, name="Segundo",
            start_date=date(2026, 2, 1), end_date=date(2026, 2, 14),
        )
        latest = expedient_service.create_case(
            self.connection, self.community_id, name="Tercero",
            start_date=date(2026, 2, 1), end_date=date(2026, 2, 28),
        )

        self.assertEqual(
            [latest.id_case, same_date.id_case, earliest.id_case],
            [case.id_case for case in expedient_service.list_cases(
                self.connection, self.community_id
            )],
        )
        with self.assertRaisesRegex(LookupError, "^Expediente no encontrado$"):
            expedient_service.get_case(self.connection, 999999)


if __name__ == "__main__":
    unittest.main()
