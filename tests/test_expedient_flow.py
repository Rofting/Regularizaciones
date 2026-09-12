import sqlite3
import sys
import tempfile
import unittest
from contextlib import redirect_stdout
from datetime import date
from io import StringIO
from pathlib import Path
from unittest.mock import patch


PROJECT_ROOT = Path(__file__).resolve().parents[1]
CORE_DIR = PROJECT_ROOT / "core"
if str(CORE_DIR) not in sys.path:
    sys.path.insert(0, str(CORE_DIR))

import case_ingestion
import document_review
import expedient_service
import gestor_bd
from source_analysis import SourceAnalysis, SourceLocator


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
        self.reading_file = Path(self.directory.name) / "lecturas.csv"
        self.reading_file.write_text("contador;lectura\nA;12\n", encoding="utf-8")
        self.unknown_file = Path(self.directory.name) / "desconocido.dat"
        self.unknown_file.write_bytes(b"contenido no clasificable")
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

    def open_issue_code(self):
        return self.connection.execute(
            """SELECT code FROM review_issues
               WHERE id_case = ? AND status = 'open' ORDER BY id_issue""",
            (self.case.id_case,),
        ).fetchone()[0]

    def candidate_value(self, document_id, field_name):
        return self.connection.execute(
            """SELECT value FROM extraction_candidates
               WHERE id_document = ? AND field_name = ?""",
            (document_id, field_name),
        ).fetchone()[0]

    def add_invoice_with_manual_total(self, total):
        result = case_ingestion.add_analysed_document_to_case(
            self.connection,
            self.case.id_case,
            source_path=self.source_path,
            archive_root=self.archive_root,
            analysis=SourceAnalysis.invoice(
                {
                    "fecha_inicio": "2026-01-01",
                    "fecha_fin": "2026-01-31",
                    "importe_total": None,
                }
            ),
        )
        issue = next(
            issue for issue in document_review.list_open_issues(self.connection, self.case.id_case)
            if issue.field_name == "importe_total"
        )
        document_review.resolve_issue(
            self.connection, issue.id_issue, value=total,
            reason="Confirmado manualmente",
        )
        return result.document

    def changed_analyser(self, _path):
        return SourceAnalysis.invoice({
            "fecha_inicio": "2026-01-01",
            "fecha_fin": "2026-01-31",
            "importe_total": "20.00",
        })

    def test_reading_document_creates_no_invoice_missing_field_issues(self):
        result = case_ingestion.add_analysed_document_to_case(
            self.connection,
            self.case.id_case,
            source_path=self.reading_file,
            archive_root=self.archive_root,
            analysis=SourceAnalysis.reading(),
        )

        self.assertEqual(0, result.open_issue_count)

    def test_unknown_document_creates_one_classification_issue(self):
        result = case_ingestion.add_analysed_document_to_case(
            self.connection,
            self.case.id_case,
            source_path=self.unknown_file,
            archive_root=self.archive_root,
            analysis=SourceAnalysis.unknown(),
        )
        repeated = case_ingestion.add_analysed_document_to_case(
            self.connection,
            self.case.id_case,
            source_path=self.unknown_file,
            archive_root=self.archive_root,
            analysis=SourceAnalysis.unknown(),
        )

        self.assertEqual(1, result.open_issue_count)
        self.assertEqual(1, repeated.open_issue_count)
        self.assertEqual("DOCUMENT_CLASSIFICATION_REQUIRED", self.open_issue_code())
        self.assertEqual(1, self.connection.execute(
            """SELECT COUNT(*) FROM review_issues
               WHERE id_document = ? AND code = 'DOCUMENT_CLASSIFICATION_REQUIRED'
                 AND status = 'open'""",
            (result.document.id_document,),
        ).fetchone()[0])

    def test_analysed_document_persists_classification_and_compact_context(self):
        result = case_ingestion.add_analysed_document_to_case(
            self.connection,
            self.case.id_case,
            source_path=self.source_path,
            archive_root=self.archive_root,
            analysis=SourceAnalysis.invoice(
                {"importe_total": "123.45"},
                locator=SourceLocator(page=2, fragment="TOTAL"),
            ),
        )

        document = self.connection.execute(
            """SELECT document_kind, classification_confidence FROM source_documents
               WHERE id_document = ?""",
            (result.document.id_document,),
        ).fetchone()
        candidate = self.connection.execute(
            """SELECT value, validation_status, source_context FROM extraction_candidates
               WHERE id_document = ? AND field_name = 'importe_total'""",
            (result.document.id_document,),
        ).fetchone()

        self.assertEqual(("invoice", "high"), tuple(document))
        self.assertEqual(("123.45", "candidate", '{"page":2,"fragment":"TOTAL"}'), tuple(candidate))

    def test_reanalysis_keeps_manually_confirmed_candidate(self):
        document = self.add_invoice_with_manual_total("10.00")

        case_ingestion.reanalyze_case_documents(
            self.connection, self.case.id_case, analyser=self.changed_analyser,
        )

        self.assertEqual("10.00", self.candidate_value(document.id_document, "importe_total"))

    def test_reanalysis_removes_only_open_automatic_issues(self):
        result = case_ingestion.add_analysed_document_to_case(
            self.connection,
            self.case.id_case,
            source_path=self.source_path,
            archive_root=self.archive_root,
            analysis=SourceAnalysis.invoice(
                {
                    "fecha_inicio": "2026-01-01",
                    "fecha_fin": "2026-01-31",
                    "importe_total": None,
                }
            ),
        )
        manual = document_review.create_review_issue(
            self.connection,
            self.case.id_case,
            result.document.id_document,
            code="MANUAL_REVIEW",
            field_name="importe_total",
            message="Revisión solicitada por la gestora",
        )

        case_ingestion.reanalyze_case_documents(
            self.connection,
            self.case.id_case,
            analyser=lambda _path: SourceAnalysis.invoice({
                "fecha_inicio": "2026-01-01",
                "fecha_fin": "2026-01-31",
                "importe_total": "123.45",
            }),
        )

        open_issues = document_review.list_open_issues(self.connection, self.case.id_case)
        self.assertEqual([manual.id_issue], [issue.id_issue for issue in open_issues])

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

    def test_new_incomplete_source_returns_ready_case_to_review(self):
        self._add_invoice_and_resolve_start_date()
        document_review.validate_case_ready(self.connection, self.case.id_case)
        second_source = Path(self.directory.name) / "factura-adicional.pdf"
        second_source.write_bytes(b"%PDF-1.4 factura adicional")

        result = case_ingestion.add_document_to_case(
            self.connection,
            self.case.id_case,
            source_path=second_source,
            archive_root=self.archive_root,
            document_kind="invoice",
            candidates={"fecha_inicio": "2026-01-15"},
            required_fields=("fecha_inicio", "importe_total"),
        )

        issues = document_review.list_open_issues(
            self.connection, self.case.id_case
        )
        self.assertTrue(result.created)
        self.assertEqual(
            [("importe_total", "open")],
            [(issue.field_name, issue.status) for issue in issues],
        )
        self.assertEqual(
            "under_review",
            expedient_service.get_case(self.connection, self.case.id_case).status,
        )

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

    def test_retry_after_archiving_failure_completes_review_without_duplicates(self):
        candidates = {"fecha_inicio": "2026-01-01", "importe_total": None}
        with patch(
            "document_review.record_candidates",
            side_effect=RuntimeError("fallo inyectado antes de candidatos"),
        ):
            with self.assertRaisesRegex(RuntimeError, "fallo inyectado"):
                case_ingestion.add_document_to_case(
                    self.connection,
                    self.case.id_case,
                    source_path=self.source_path,
                    archive_root=self.archive_root,
                    document_kind="invoice",
                    candidates=candidates,
                    required_fields=("fecha_inicio", "importe_total"),
                )

        self.assertEqual(
            1,
            case_ingestion.count_case_documents(
                self.connection, self.case.id_case
            ),
        )
        self.assertEqual(
            b"%PDF-1.4 factura sintetica",
            next((self.archive_root / str(self.case.id_case) / "fuentes").iterdir()).read_bytes(),
        )

        recovered = case_ingestion.add_document_to_case(
            self.connection,
            self.case.id_case,
            source_path=self.source_path,
            archive_root=self.archive_root,
            document_kind="invoice",
            candidates=candidates,
            required_fields=("fecha_inicio", "importe_total"),
        )

        stored_candidates = self.connection.execute(
            """SELECT field_name, value FROM extraction_candidates
               WHERE id_document = ? ORDER BY field_name""",
            (recovered.document.id_document,),
        ).fetchall()
        issues = document_review.list_open_issues(
            self.connection, self.case.id_case
        )
        self.assertFalse(recovered.created)
        self.assertEqual(
            [("fecha_inicio", "2026-01-01"), ("importe_total", None)],
            [tuple(row) for row in stored_candidates],
        )
        self.assertEqual(["importe_total"], [issue.field_name for issue in issues])
        self.assertEqual(1, recovered.open_issue_count)
        self.assertEqual(
            1,
            case_ingestion.count_case_documents(
                self.connection, self.case.id_case
            ),
        )
        self.assertEqual(
            1,
            self.connection.execute(
                "SELECT COUNT(*) FROM review_issues WHERE id_case = ?",
                (self.case.id_case,),
            ).fetchone()[0],
        )
        self.assertEqual(
            "under_review",
            expedient_service.get_case(self.connection, self.case.id_case).status,
        )

    def test_case_ownership_rejects_another_community_and_accepts_its_own(self):
        other_community_id = gestor_bd.obtener_o_crear_comunidad(
            self.connection, "OTRA", "Otra comunidad sintética"
        )

        with self.assertRaisesRegex(LookupError, "pertenece a otra comunidad"):
            case_ingestion.assert_case_belongs_to_community(
                self.connection,
                self.case.id_case,
                other_community_id,
            )

        selected_case = case_ingestion.assert_case_belongs_to_community(
            self.connection,
            self.case.id_case,
            self.community_id,
        )
        self.assertEqual(self.case.id_case, selected_case.id_case)
        self.assertEqual(self.community_id, selected_case.community_id)


if __name__ == "__main__":
    unittest.main()
