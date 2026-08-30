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

import case_ingestion
import document_review
import expedient_service
import gestor_bd


class ExpedientFlowTest(unittest.TestCase):
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
        self.case = expedient_service.create_case(
            self.connection,
            self.community_id,
            name="Expediente sintético",
            start_date=date(2026, 1, 1),
            end_date=date(2026, 1, 31),
        )
        self.source_path = Path(self.directory.name) / "factura.pdf"
        self.source_path.write_bytes(b"%PDF-1.4 factura sintetica")
        self.archive_root = Path(self.directory.name) / "expedientes"

    def tearDown(self):
        self.connection.close()
        self.directory.cleanup()

    def _add_invoice_and_resolve_start_date(self):
        result = case_ingestion.add_document_to_case(
            self.connection,
            self.case.id_case,
            source_path=self.source_path,
            archive_root=self.archive_root,
            document_kind="invoice",
            candidates={"importe_total": "123.45"},
            required_fields=("importe_total", "fecha_inicio"),
        )
        issue = document_review.list_open_issues(
            self.connection, self.case.id_case
        )[0]
        document_review.resolve_issue(
            self.connection,
            issue.id_issue,
            value="2026-01-01",
            reason="Confirmado en documento",
        )
        return result

    def test_invoice_review_flow_reaches_ready_for_calculation(self):
        result = self._add_invoice_and_resolve_start_date()

        ready_case = document_review.validate_case_ready(
            self.connection, self.case.id_case
        )

        self.assertTrue(result.created)
        self.assertEqual(1, result.open_issue_count)
        self.assertEqual(
            1,
            case_ingestion.count_case_documents(
                self.connection, self.case.id_case
            ),
        )
        self.assertEqual("ready_for_calculation", ready_case.status)

    def test_readding_resolved_source_preserves_manual_value_and_correction(self):
        self._add_invoice_and_resolve_start_date()

        repeated = case_ingestion.add_document_to_case(
            self.connection,
            self.case.id_case,
            source_path=self.source_path,
            archive_root=self.archive_root,
            document_kind="invoice",
            candidates={"importe_total": "999.99", "fecha_inicio": "otro"},
            required_fields=("importe_total", "fecha_inicio"),
        )

        source_count = self.connection.execute(
            "SELECT COUNT(*) FROM source_documents WHERE id_case = ?",
            (self.case.id_case,),
        ).fetchone()[0]
        manual_start_date = self.connection.execute(
            """SELECT value FROM extraction_candidates
               WHERE id_document = ? AND field_name = 'fecha_inicio'""",
            (repeated.document.id_document,),
        ).fetchone()[0]
        correction_count = self.connection.execute(
            "SELECT COUNT(*) FROM manual_corrections"
        ).fetchone()[0]

        self.assertFalse(repeated.created)
        self.assertEqual(1, source_count)
        self.assertEqual("2026-01-01", manual_start_date)
        self.assertEqual(0, repeated.open_issue_count)
        self.assertEqual(1, correction_count)


if __name__ == "__main__":
    unittest.main()
