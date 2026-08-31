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
import expedient_ui
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

    def _create_counter_reset_issue(self, *, property_code="A", initial=100, final=5):
        period_id = self.connection.execute(
            """INSERT INTO periodos
               (id_comunidad,nombre,fecha_inicio,fecha_fin,estado)
               VALUES (?,'enero','2026-01-01','2026-01-31','abierto')""",
            (self.community_id,),
        ).lastrowid
        self.connection.execute(
            "UPDATE regularization_cases SET id_periodo=? WHERE id_case=?",
            (period_id, self.case.id_case),
        )
        owner_id = self.connection.execute(
            """INSERT INTO propietarios
               (id_comunidad,codigo_vivienda,nombre_propietario,coeficiente)
               VALUES (?,?,?,1)""",
            (self.community_id, property_code, f"Propietario {property_code}"),
        ).lastrowid
        self.connection.executemany(
            """INSERT INTO lecturas_vecino
               (id_propietario,id_periodo,tipo,fecha_lectura,valor_acumulado,estado,fuente)
               VALUES (?,?,'ACS',?,?,?, 'origen')""",
            (
                (owner_id, period_id, "2026-01-01", initial, "real"),
                (owner_id, period_id, "2026-01-31", final, "contador_averiado"),
            ),
        )
        self.connection.execute(
            "UPDATE propietarios SET estado_contador_acs='averiado' WHERE id_propietario=?",
            (owner_id,),
        )
        self.connection.commit()
        return document_review.create_review_issue(
            self.connection,
            self.case.id_case,
            self.document.id_document,
            code="COUNTER_RESET",
            field_name=f"reading.{property_code}.ACS",
            message="El contador disminuye y requiere una estimación aprobada",
            detected_value=f"{initial} -> {final}",
        ), owner_id, period_id

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

    def test_approving_counter_reset_estimate_replaces_only_canonical_final_reading(self):
        issue, owner_id, period_id = self._create_counter_reset_issue()
        other_owner = self.connection.execute(
            """INSERT INTO propietarios
               (id_comunidad,codigo_vivienda,nombre_propietario,coeficiente)
               VALUES (?,'B','Propietario B',1)""",
            (self.community_id,),
        ).lastrowid
        self.connection.executemany(
            """INSERT INTO lecturas_vecino
               (id_propietario,id_periodo,tipo,fecha_lectura,valor_acumulado,estado)
               VALUES (?,?,'ACS',?,?, 'real')""",
            (
                (other_owner, period_id, "2026-01-01", 20),
                (other_owner, period_id, "2026-01-31", 30),
            ),
        )
        self.connection.commit()

        resolved = document_review.approve_counter_reset_estimate(
            self.connection,
            issue.id_issue,
            consumption="12.5",
            reason="Sustitución documentada por mantenimiento",
            approved_by="gestora",
        )

        reading = self.connection.execute(
            """SELECT valor_acumulado,estado,metodo_estimacion,approved_by,
                      approved_at,notas,fuente
               FROM lecturas_vecino
               WHERE id_propietario=? AND fecha_lectura='2026-01-31'""",
            (owner_id,),
        ).fetchone()
        correction = self.connection.execute(
            """SELECT original_value,corrected_value,reason,resolved_by
               FROM manual_corrections WHERE id_issue=?""",
            (issue.id_issue,),
        ).fetchone()
        candidate = self.connection.execute(
            """SELECT value,source,validation_status FROM extraction_candidates
               WHERE id_document=? AND field_name=?""",
            (self.document.id_document, issue.field_name),
        ).fetchone()
        other_final = self.connection.execute(
            """SELECT valor_acumulado,estado FROM lecturas_vecino
               WHERE id_propietario=? AND fecha_lectura='2026-01-31'""",
            (other_owner,),
        ).fetchone()

        self.assertEqual("resolved", resolved.status)
        self.assertEqual((112.5, "estimado", "counter_reset_manual", "gestora"), tuple(reading)[:4])
        self.assertIsNotNone(reading["approved_at"])
        self.assertIn("valor original=5", reading["notas"])
        self.assertEqual("origen", reading["fuente"])
        self.assertEqual(("5", "112.5", "Sustitución documentada por mantenimiento", "gestora"), tuple(correction))
        self.assertEqual(("112.5", "manual_counter_reset", "validated"), tuple(candidate))
        self.assertEqual((30.0, "real"), tuple(other_final))
        self.assertEqual("ok", self.connection.execute(
            "SELECT estado_contador_acs FROM propietarios WHERE id_propietario=?", (owner_id,)
        ).fetchone()[0])

    def test_counter_reset_estimate_rejects_non_positive_or_non_numeric_consumption(self):
        issue, owner_id, _period_id = self._create_counter_reset_issue()

        for consumption in ("0", "-1", "no es un número"):
            with self.subTest(consumption=consumption):
                with self.assertRaisesRegex(ValueError, "positivo"):
                    document_review.approve_counter_reset_estimate(
                        self.connection, issue.id_issue, consumption=consumption,
                        reason="Soporte comprobado", approved_by="gestora",
                    )

        reading = self.connection.execute(
            """SELECT valor_acumulado,estado,approved_by FROM lecturas_vecino
               WHERE id_propietario=? AND fecha_lectura='2026-01-31'""",
            (owner_id,),
        ).fetchone()
        self.assertEqual((5.0, "contador_averiado", None), tuple(reading))
        self.assertEqual(0, self.connection.execute(
            "SELECT COUNT(*) FROM manual_corrections WHERE id_issue=?", (issue.id_issue,)
        ).fetchone()[0])

    def test_counter_reset_estimate_rejects_a_property_from_another_community(self):
        self._create_counter_reset_issue()
        other_community = gestor_bd.obtener_o_crear_comunidad(
            self.connection, "OTRA", "Otra comunidad"
        )
        self.connection.execute(
            """INSERT INTO propietarios
               (id_comunidad,codigo_vivienda,nombre_propietario,coeficiente)
               VALUES (?, 'Z', 'Propietario ajeno', 1)""",
            (other_community,),
        )
        self.connection.commit()
        issue = document_review.create_review_issue(
            self.connection, self.case.id_case, self.document.id_document,
            code="COUNTER_RESET", field_name="reading.Z.ACS",
            message="Contador reiniciado", detected_value="10 -> 5",
        )

        with self.assertRaisesRegex(LookupError, "propietario del expediente"):
            document_review.approve_counter_reset_estimate(
                self.connection, issue.id_issue, consumption="12",
                reason="Soporte comprobado", approved_by="gestora",
            )

        self.assertEqual("open", self.connection.execute(
            "SELECT status FROM review_issues WHERE id_issue=?", (issue.id_issue,)
        ).fetchone()[0])
        self.assertEqual(0, self.connection.execute(
            "SELECT COUNT(*) FROM manual_corrections WHERE id_issue=?", (issue.id_issue,)
        ).fetchone()[0])

    def test_counter_reset_estimate_rolls_back_audit_when_canonical_update_fails(self):
        issue, owner_id, _period_id = self._create_counter_reset_issue()
        self.connection.execute(
            """CREATE TRIGGER reject_estimated_reading BEFORE UPDATE ON lecturas_vecino
               WHEN NEW.estado = 'estimado'
               BEGIN SELECT RAISE(ABORT, 'fallo de lectura inyectado'); END"""
        )
        self.connection.commit()

        with self.assertRaisesRegex(sqlite3.IntegrityError, "fallo de lectura inyectado"):
            document_review.approve_counter_reset_estimate(
                self.connection, issue.id_issue, consumption="12",
                reason="Soporte comprobado", approved_by="gestora",
            )

        self.assertEqual(0, self.connection.execute(
            "SELECT COUNT(*) FROM manual_corrections WHERE id_issue=?", (issue.id_issue,)
        ).fetchone()[0])
        self.assertEqual("open", self.connection.execute(
            "SELECT status FROM review_issues WHERE id_issue=?", (issue.id_issue,)
        ).fetchone()[0])
        self.assertEqual((5.0, "contador_averiado"), tuple(self.connection.execute(
            """SELECT valor_acumulado,estado FROM lecturas_vecino
               WHERE id_propietario=? AND fecha_lectura='2026-01-31'""", (owner_id,)
        ).fetchone()))

    def test_dismisses_only_an_invoice_outside_period_without_creating_invoice(self):
        issue = document_review.create_review_issue(
            self.connection, self.case.id_case, self.document.id_document,
            code="INVOICE_OUTSIDE_PERIOD", field_name="GAS.B10.invoice_date",
            message="La fecha de la factura queda fuera del expediente",
            detected_value="2025-12-31",
        )

        dismissed = document_review.dismiss_invoice_outside_period(
            self.connection, issue.id_issue,
            reason="Corresponde al ejercicio anterior", dismissed_by="gestora",
        )

        self.assertEqual("dismissed", dismissed.status)
        self.assertEqual(0, self.connection.execute("SELECT COUNT(*) FROM facturas").fetchone()[0])
        correction = self.connection.execute(
            "SELECT original_value,corrected_value,reason,resolved_by FROM manual_corrections WHERE id_issue=?",
            (issue.id_issue,),
        ).fetchone()
        self.assertEqual(
            ("2025-12-31", "no_corresponde_al_periodo", "Corresponde al ejercicio anterior", "gestora"),
            tuple(correction),
        )

    def test_issue_route_selects_specialized_flow_without_creating_a_window(self):
        counter, _owner_id, _period_id = self._create_counter_reset_issue()
        invoice = document_review.create_review_issue(
            self.connection, self.case.id_case, self.document.id_document,
            code="INVOICE_OUTSIDE_PERIOD", field_name="GAS.B10.invoice_date",
            message="Fuera de período", detected_value="2025-12-31",
        )
        generic = document_review.create_review_issue(
            self.connection, self.case.id_case, self.document.id_document,
            code="MISSING_REQUIRED_FIELD", field_name="importe_total",
            message="Falta el importe", detected_value=None,
        )

        self.assertEqual("counter_reset_estimate", expedient_ui.resolution_route_for_issue(counter))
        self.assertEqual("dismiss_invoice_outside_period", expedient_ui.resolution_route_for_issue(invoice))
        self.assertEqual("generic_correction", expedient_ui.resolution_route_for_issue(generic))


if __name__ == "__main__":
    unittest.main()
