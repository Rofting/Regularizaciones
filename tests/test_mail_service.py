import sqlite3
import sys
import unittest
from email.parser import BytesParser
from email.policy import default
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
for folder in (PROJECT_ROOT / "core", PROJECT_ROOT / "tests"):
    if str(folder) not in sys.path:
        sys.path.insert(0, str(folder))

from test_case_letter_service import CaseLetterServiceTest
from case_letter_service import generate_case_letters
from mail_service import exclude_mail_delivery, generate_eml_drafts, prepare_mail_run, send_mail_run
from mail_transport import FakeMailTransport
from office_settings import OfficeSettings, save_office_settings


class MailServiceTest(unittest.TestCase):
    def setUp(self):
        self.fixture = CaseLetterServiceTest("test_case_letters_use_active_concepts_and_record_each_owner")
        self.fixture.setUp()
        self.addCleanup(self.fixture.tearDown)
        f = self.fixture
        save_office_settings(f.connection, OfficeSettings(name="Despacho de prueba", email="despacho@example.org"))
        f.connection.execute("UPDATE propietarios SET email='ana@example.org' WHERE id_propietario=?", (f.owner_one,))
        f.connection.commit()
        self.letters = generate_case_letters(f.database_path, id_case=f.case_id, project_root=f.root)
        f.connection.execute("UPDATE regularization_cases SET estado='deliveries_generated' WHERE id_case=?", (f.case_id,))
        f.connection.commit()

    def _prepare(self):
        f = self.fixture
        result = prepare_mail_run(f.connection, f.case_id, f.root / "salidas" / "correo")
        return generate_eml_drafts(f.connection, result.id_mail_run)

    def test_eml_records_missing_email_and_attaches_corresponding_pdf(self):
        result = self._prepare()
        self.assertEqual((1, 1, 0), (result.draft_count, result.skipped_count, result.duplicate_count))
        eml = next(result.output_path.glob("*.eml"))
        message = BytesParser(policy=default).parsebytes(eml.read_bytes())
        self.assertEqual("ana@example.org", message["To"])
        attachment = list(message.iter_attachments())[0]
        self.assertEqual("application/pdf", attachment.get_content_type())
        self.assertEqual(
            (self.letters.output_path / "CARTA_A-1_ANA_VECINA.pdf").read_bytes(),
            attachment.get_payload(decode=True),
        )
        rows = self.fixture.connection.execute(
            "SELECT status FROM mail_deliveries WHERE id_mail_run=? ORDER BY id_propietario",
            (result.id_mail_run,),
        ).fetchall()
        self.assertEqual(["draft", "skipped"], [row[0] for row in rows])
        self.assertEqual(result.id_mail_run, self._prepare().id_mail_run)

    def test_duplicate_recipient_is_audited_and_swapped_pdf_requires_regeneration(self):
        f = self.fixture
        f.connection.execute("UPDATE propietarios SET email='ana@example.org' WHERE id_propietario=?", (f.owner_two,))
        f.connection.commit()
        first = self._prepare()
        self.assertEqual(1, first.duplicate_count)
        self.assertEqual(1, len(list(first.output_path.glob("*.eml"))))
        # Otro PDF válido puede pertenecer a un vecino distinto: exige regenerar.
        source = self.letters.output_path / "CARTA_B-2_BRUNO_VECINO.pdf"
        destination = self.letters.output_path / "CARTA_A-1_ANA_VECINA.pdf"
        destination.write_bytes(source.read_bytes())
        before = f.connection.execute('SELECT COUNT(*) FROM mail_runs').fetchone()[0]
        with self.assertRaisesRegex(ValueError, 'Regenera las cartas'):
            self._prepare()
        self.assertEqual(before, f.connection.execute('SELECT COUNT(*) FROM mail_runs').fetchone()[0])

    def test_letters_without_a_generation_hash_cannot_prepare_or_send_email(self):
        f = self.fixture
        run = self._prepare()
        f.connection.execute('UPDATE generated_letters SET pdf_sha256=NULL WHERE id_letter_run=?', (self.letters.id_letter_run,))
        f.connection.commit()
        with self.assertRaisesRegex(ValueError, 'Regenera las cartas'):
            self._prepare()
        transport = FakeMailTransport()
        self.assertEqual((0, 1), send_mail_run(f.connection, run.id_mail_run, transport, confirmed_by='gestor'))
        self.assertEqual([], transport.messages)

    def test_confirmation_failure_and_retry_never_resends_success(self):
        f = self.fixture
        f.connection.execute("UPDATE propietarios SET email='bruno@example.org' WHERE id_propietario=?", (f.owner_two,))
        f.connection.commit()
        result = self._prepare()
        transport = FakeMailTransport()
        with self.assertRaisesRegex(ValueError, "Confirma"):
            send_mail_run(f.connection, result.id_mail_run, transport, confirmed_by="")

        class FailSecond:
            def __init__(self):
                self.count = 0

            def send(self, message):
                self.count += 1
                if self.count == 2:
                    raise OSError("fallo simulado")

        self.assertEqual((1, 1), send_mail_run(f.connection, result.id_mail_run,
                                               FailSecond(), confirmed_by="gestor"))
        self.assertEqual((1, 0), send_mail_run(f.connection, result.id_mail_run,
                                               transport, confirmed_by="gestor"))
        self.assertEqual(1, len(transport.messages))
        self.assertEqual("bruno@example.org", transport.messages[0]["To"])
        self.assertEqual((0, 0), send_mail_run(f.connection, result.id_mail_run,
                                               transport, confirmed_by="gestor"))

    def test_smtp_errors_are_explained_without_server_details(self):
        import smtplib
        result = self._prepare()
        f = self.fixture

        class Refuses:
            def __init__(self, error):
                self.error = error

            def send(self, message):
                raise self.error

        cases = (
            (smtplib.SMTPAuthenticationError(535, b"5.7.3 secreto-del-servidor"), "usuario o la contraseña"),
            (TimeoutError("timed out"), "No se pudo conectar"),
            (smtplib.SMTPRecipientsRefused({"ana@example.org": (550, b"no existe")}), "rechazó la dirección"),
        )
        for error, expected in cases:
            with self.subTest(error=type(error).__name__):
                send_mail_run(f.connection, result.id_mail_run, Refuses(error), confirmed_by="gestor")
                message = f.connection.execute(
                    "SELECT error_message FROM mail_deliveries WHERE id_mail_run=? AND status='failed'",
                    (result.id_mail_run,),
                ).fetchone()[0]
                self.assertIn(expected, message)
                self.assertNotIn("secreto", message)

    def test_run_is_sent_when_every_sendable_letter_went_out(self):
        result = self._prepare()  # Bruno no tiene correo: queda «skipped»
        f = self.fixture
        self.assertEqual((1, 0), send_mail_run(f.connection, result.id_mail_run,
                                               FakeMailTransport(), confirmed_by="gestor"))
        status = f.connection.execute("SELECT status FROM mail_runs WHERE id_mail_run=?",
                                      (result.id_mail_run,)).fetchone()[0]
        self.assertEqual("sent", status)

    def test_excluded_owner_has_no_sendable_draft(self):
        result = self._prepare()
        f = self.fixture
        exclude_mail_delivery(f.connection, result.id_mail_run, f.owner_one,
                              reason="Solicitó entrega en papel")
        self.assertEqual([], list(result.output_path.glob("*.eml")))
        transport = FakeMailTransport()
        self.assertEqual((0, 0), send_mail_run(f.connection, result.id_mail_run,
                                               transport, confirmed_by="gestor"))
        self.assertEqual([], transport.messages)


if __name__ == "__main__":
    unittest.main()
