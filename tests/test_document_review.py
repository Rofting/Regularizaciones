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

import document_review
import expedient_service
import gestor_bd


class DocumentReviewTest(unittest.TestCase):
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
            self.connection, self.community_id, name="Revisión sintética",
            start_date=date(2026, 1, 1), end_date=date(2026, 1, 31),
        )
        source_path = Path(self.directory.name) / "origen.pdf"
        source_path.write_bytes(b"%PDF-1.4 prueba")
        self.document, _ = expedient_service.register_source_document(
            self.connection, self.case.id_case, source_path=source_path,
            archive_root=Path(self.directory.name) / "expedientes",
            document_kind="invoice",
        )

    def tearDown(self):
        self.connection.close()
        self.directory.cleanup()

    def _create_missing_date_issue(self):
        document_review.record_candidates(
            self.connection, self.document.id_document,
            {"importe_total": "123.45", "fecha_inicio": None}, source="pdf",
        )
        return document_review.create_missing_field_issues(
            self.connection, self.case.id_case, self.document.id_document,
            required_fields=("importe_total", "fecha_inicio"),
        )[0]

    def test_missing_required_candidate_creates_open_issue(self):
        document_review.record_candidates(
            self.connection, self.document.id_document,
            {"importe_total": "123.45", "fecha_inicio": None}, source="pdf",
        )

        issues = document_review.create_missing_field_issues(
            self.connection, self.case.id_case, self.document.id_document,
            required_fields=("importe_total", "fecha_inicio"),
        )

        self.assertEqual([(item.field_name, item.status) for item in issues],
                         [("fecha_inicio", "open")])
        self.assertEqual(issues[0].archived_path, self.document.archived_path)

    def test_resolving_issue_persists_manual_correction_and_validates_candidate(self):
        issue = self._create_missing_date_issue()

        resolved = document_review.resolve_issue(
            self.connection, issue.id_issue, value="2026-01-01",
            reason="Confirmado en la primera página",
        )

        correction = self.connection.execute(
            """SELECT original_value, corrected_value, reason, resolved_by
               FROM manual_corrections WHERE id_issue = ?""",
            (issue.id_issue,),
        ).fetchone()
        candidate = self.connection.execute(
            """SELECT value, source, validation_status FROM extraction_candidates
               WHERE id_document = ? AND field_name = 'fecha_inicio'""",
            (self.document.id_document,),
        ).fetchone()
        issue_row = self.connection.execute(
            "SELECT status, resolved_at FROM review_issues WHERE id_issue = ?",
            (issue.id_issue,),
        ).fetchone()

        self.assertEqual((None, "2026-01-01", "Confirmado en la primera página", "usuario_local"),
                         tuple(correction))
        self.assertEqual(("2026-01-01", "manual", "validated"), tuple(candidate))
        self.assertEqual("resolved", resolved.status)
        self.assertEqual("resolved", issue_row["status"])
        self.assertIsNotNone(issue_row["resolved_at"])

    def test_resolving_issue_normalizes_responsible_person_before_auditing(self):
        issue = self._create_missing_date_issue()

        document_review.resolve_issue(
            self.connection,
            issue.id_issue,
            value="2026-01-01",
            reason="Confirmado en documento",
            resolved_by="  gestora principal  ",
        )

        responsible = self.connection.execute(
            "SELECT resolved_by FROM manual_corrections WHERE id_issue = ?",
            (issue.id_issue,),
        ).fetchone()[0]
        self.assertEqual("gestora principal", responsible)

    def test_resolving_issue_rejects_blank_responsible_before_auditing(self):
        issue = self._create_missing_date_issue()

        with self.assertRaisesRegex(ValueError, "responsable"):
            document_review.resolve_issue(
                self.connection,
                issue.id_issue,
                value="2026-01-01",
                reason="Confirmado en documento",
                resolved_by="   ",
            )

        correction_count = self.connection.execute(
            "SELECT COUNT(*) FROM manual_corrections WHERE id_issue = ?",
            (issue.id_issue,),
        ).fetchone()[0]
        issue_status = self.connection.execute(
            "SELECT status FROM review_issues WHERE id_issue = ?",
            (issue.id_issue,),
        ).fetchone()[0]
        self.assertEqual(0, correction_count)
        self.assertEqual("open", issue_status)

    def test_readiness_requires_resolving_issues_then_validates_documents(self):
        issue = self._create_missing_date_issue()

        with self.assertRaisesRegex(ValueError, "1 incidencia abierta"):
            document_review.validate_case_ready(self.connection, self.case.id_case)

        document_review.resolve_issue(
            self.connection, issue.id_issue, value="2026-01-01",
            reason="Confirmado en la primera página",
        )
        ready = document_review.validate_case_ready(self.connection, self.case.id_case)
        document_status = self.connection.execute(
            "SELECT status FROM source_documents WHERE id_document = ?",
            (self.document.id_document,),
        ).fetchone()["status"]

        self.assertEqual("ready_for_calculation", ready.status)
        self.assertEqual("validated", document_status)

    def test_repeating_missing_review_reuses_the_open_issue(self):
        first = self._create_missing_date_issue()

        repeated = document_review.create_missing_field_issues(
            self.connection, self.case.id_case, self.document.id_document,
            required_fields=("fecha_inicio",),
        )

        self.assertEqual([first.id_issue], [item.id_issue for item in repeated])
        self.assertEqual(1, self.connection.execute(
            """SELECT COUNT(*) FROM review_issues
               WHERE id_document = ? AND field_name = 'fecha_inicio' AND status = 'open'""",
            (self.document.id_document,),
        ).fetchone()[0])


if __name__ == "__main__":
    unittest.main()
