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

    def test_duplicate_recipient_and_changed_attachment_get_separate_audit(self):
        f = self.fixture
        f.connection.execute("UPDATE propietarios SET email='ana@example.org' WHERE id_propietario=?", (f.owner_two,))
        f.connection.commit()
        first = self._prepare()
        self.assertEqual(1, first.duplicate_count)
        self.assertEqual(1, len(list(first.output_path.glob("*.eml"))))
        # Sustituir por otro PDF válido cambia la huella y obliga a preparar otro lote.
        source = self.letters.output_path / "CARTA_B-2_BRUNO_VECINO.pdf"
        destination = self.letters.output_path / "CARTA_A-1_ANA_VECINA.pdf"
        destination.write_bytes(source.read_bytes())
        second = self._prepare()
        self.assertNotEqual(first.id_mail_run, second.id_mail_run)

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
