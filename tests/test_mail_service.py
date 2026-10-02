"""Mejora 16: correo de cartas. Nunca se envía correo real en las pruebas."""

import email
import email.policy
import sqlite3
import sys
import tempfile
import unittest
from contextlib import redirect_stdout
from datetime import date
from io import StringIO
from pathlib import Path
from unittest import mock

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "core"))

import expedient_service
import gestor_bd
import mail_service
import mail_transport
from mail_transport import FakeMailTransport, MailTransportError


class MailServiceTest(unittest.TestCase):
    def setUp(self):
        directory = tempfile.TemporaryDirectory(ignore_cleanup_errors=True)
        self.addCleanup(directory.cleanup)
        self.root = Path(directory.name)
        database = self.root / "g.db"
        with redirect_stdout(StringIO()):
            gestor_bd.crear_bd(str(database))
        self.connection = sqlite3.connect(database)
        self.connection.row_factory = sqlite3.Row
        self.addCleanup(self.connection.close)
        community = gestor_bd.obtener_o_crear_comunidad(self.connection, "901", "CP Ensayo")
        self.case = expedient_service.create_case(
            self.connection, community, name="Año 2026", start_date=date(2026, 1, 1), end_date=date(2026, 12, 31))
        owners = (("1A", "Ana Uno", "ana@ejemplo.es"), ("1B", "Benito Dos", ""),
                  ("2A", "Carla Tres", "correo-roto"), ("2B", "Dani Cuatro", "COMPARTIDO@Ejemplo.ES"),
                  ("3A", "Eva Cinco", "compartido@ejemplo.es"))
        self.connection.execute("PRAGMA foreign_keys=OFF")
        letters = self.root / "cartas"
        letters.mkdir()
        run = self.connection.execute(
            "INSERT INTO letter_generation_runs(id_case,id_periodo,id_export_run,input_sha256,template_sha256,"
            "output_path,status) VALUES (?,1,1,'i','t',?,'completed')", (self.case.id_case, str(letters))).lastrowid
        self.owner_ids = {}
        for code, name, address in owners:
            owner = self.connection.execute(
                "INSERT INTO propietarios(id_comunidad,codigo_vivienda,nombre_propietario,coeficiente,email) "
                "VALUES (?,?,?,20,?)", (community, code, name, address)).lastrowid
            self.owner_ids[code] = owner
            path = letters / f"{code}.docx"
            path.write_bytes(b"PK carta " + code.encode())
            self.connection.execute(
                "INSERT INTO generated_letters(id_letter_run,id_propietario,input_sha256,template_sha256,output_path,status) "
                "VALUES (?,?,'i','t',?,'generated')", (run, owner, str(path)))
        self.connection.commit()
        self.letters = letters

    def prepare(self, **kwargs):
        return mail_service.prepare_mail_run(self.connection, self.case.id_case, created_by="Prueba", **kwargs)

    def statuses(self, run_id):
        summary = mail_service.mail_run_summary(self.connection, run_id)
        return {row.dwelling: (row.status, row.reason) for row in summary.deliveries}

    def test_plan_flags_missing_invalid_and_shared_addresses(self):
        plan = mail_service.plan_mail(self.connection, self.case.id_case)
        by_dwelling = {item.dwelling: item for item in plan.candidates}
        self.assertEqual("ana@ejemplo.es", by_dwelling["1A"].email)
        self.assertEqual(mail_service.NO_EMAIL, by_dwelling["1B"].problem)
        self.assertEqual(mail_service.INVALID_EMAIL, by_dwelling["2A"].problem)
        self.assertEqual("COMPARTIDO@ejemplo.es", by_dwelling["2B"].email)
        self.assertEqual({"2B", "3A"}, {item.dwelling for item in plan.duplicates})
        self.assertEqual(3, len(plan.sendable))

    def test_drafts_are_mime_messages_with_the_letter_attached(self):
        run_id = self.prepare(excluded_owner_ids=[self.owner_ids["3A"]])
        summary = mail_service.generate_eml_drafts(self.connection, run_id, self.root / "salidas", sender="despacho@ejemplo.es")
        self.assertEqual(2, summary.count("draft"))
        statuses = self.statuses(run_id)
        self.assertEqual(("skipped", mail_service.NO_EMAIL), statuses["1B"])
        self.assertEqual(("skipped", mail_service.EXCLUDED), statuses["3A"])
        draft = next(row for row in summary.deliveries if row.dwelling == "1A")
        message = email.message_from_bytes(Path(draft.eml_path).read_bytes(), policy=email.policy.default)
        self.assertEqual("ana@ejemplo.es", message["To"])
        self.assertEqual("1", message["X-Unsent"])
        self.assertIn("Año 2026", message["Subject"])
        attachments = [part for part in message.walk() if part.get_filename()]
        self.assertEqual(["1A.docx"], [part.get_filename() for part in attachments])
        self.assertEqual(b"PK carta 1A", attachments[0].get_payload(decode=True))
        body = next(part for part in message.walk() if part.get_content_type() == "text/plain")
        self.assertIn("Ana Uno", body.get_content())

    def test_identical_draft_is_reused_and_a_changed_letter_gets_a_new_one(self):
        first = self.prepare()
        mail_service.generate_eml_drafts(self.connection, first, self.root / "salidas")
        second = self.prepare()
        summary = mail_service.generate_eml_drafts(self.connection, second, self.root / "salidas")
        reused = {row.dwelling: row for row in summary.deliveries}["1A"]
        self.assertEqual("Borrador idéntico ya generado", reused.reason)
        (self.letters / "1A.docx").write_bytes(b"PK carta corregida")
        third = self.prepare()
        summary = mail_service.generate_eml_drafts(self.connection, third, self.root / "salidas")
        fresh = {row.dwelling: row for row in summary.deliveries}["1A"]
        self.assertIsNone(fresh.reason)
        self.assertIn("lote_" + str(third), fresh.eml_path)

    def test_sending_needs_confirmation(self):
        run_id = self.prepare()
        with self.assertRaises(PermissionError):
            mail_service.send_mail_run(self.connection, run_id, FakeMailTransport(), confirmed_by=" ", sender="d@e.es")

    def test_failure_is_recorded_and_retry_does_not_resend(self):
        run_id = self.prepare()
        transport = FakeMailTransport(fail_for={"COMPARTIDO@ejemplo.es"})
        summary = mail_service.send_mail_run(self.connection, run_id, transport, confirmed_by="Prueba", sender="d@e.es")
        self.assertEqual("incomplete", summary.status)
        statuses = self.statuses(run_id)
        self.assertEqual("sent", statuses["1A"][0])
        self.assertEqual("sent", statuses["3A"][0])
        self.assertEqual(("failed", "550 Buzón inexistente"), statuses["2B"])
        self.assertEqual(2, len(transport.sent))

        retry = FakeMailTransport()
        summary = mail_service.send_mail_run(self.connection, run_id, retry, confirmed_by="Prueba", sender="d@e.es")
        self.assertEqual("completed", summary.status)
        self.assertEqual(["COMPARTIDO@ejemplo.es"], [str(m["To"]) for m in retry.sent])

        # Un lote nuevo con el mismo contenido no reenvía nada.
        again = self.prepare()
        third = FakeMailTransport()
        summary = mail_service.send_mail_run(self.connection, again, third, confirmed_by="Prueba", sender="d@e.es")
        self.assertEqual([], third.sent)
        self.assertTrue(all(row.status == "skipped" for row in summary.deliveries))
        self.assertIn("no se reenvía", self.statuses(again)["1A"][1])

    def test_letter_changed_after_preparing_is_not_sent(self):
        run_id = self.prepare()
        (self.letters / "1A.docx").write_bytes(b"PK otra carta")
        transport = FakeMailTransport()
        mail_service.send_mail_run(self.connection, run_id, transport, confirmed_by="Prueba", sender="d@e.es")
        status, reason = self.statuses(run_id)["1A"]
        self.assertEqual("failed", status)
        self.assertIn("cambió", reason)
        self.assertNotIn("ana@ejemplo.es", [str(m["To"]) for m in transport.sent])

    def test_connection_error_does_not_leave_the_run_sending(self):
        class Broken(FakeMailTransport):
            def __enter__(self):
                raise MailTransportError("No se pudo conectar con el servidor de correo")
        run_id = self.prepare()
        summary = mail_service.send_mail_run(self.connection, run_id, Broken(), confirmed_by="Prueba", sender="d@e.es")
        self.assertEqual("incomplete", summary.status)
        self.assertEqual(3, summary.count("failed"))

    def test_unknown_placeholder_is_rejected(self):
        with self.assertRaisesRegex(ValueError, "importe"):
            self.prepare(subject_template="Carta {importe}")

    def test_without_letters_there_is_nothing_to_prepare(self):
        self.connection.execute("DELETE FROM generated_letters")
        self.connection.execute("DELETE FROM letter_generation_runs")
        self.connection.commit()
        with self.assertRaisesRegex(ValueError, "Genera primero"):
            self.prepare()


class MailDialogTest(unittest.TestCase):
    setUp = MailServiceTest.setUp

    def test_dialog_asks_for_letters_first(self):
        import expedient_ui
        self.connection.execute("DELETE FROM generated_letters")
        self.connection.execute("DELETE FROM letter_generation_runs")
        self.connection.commit()
        app = mock.Mock(ruta_bd_expedientes=self.connection.execute("PRAGMA database_list").fetchone()[2])
        with mock.patch.object(expedient_ui, "messagebox") as messages, \
                mock.patch.object(expedient_ui, "_dialog") as dialog:
            expedient_ui.open_mail_dialog(app, self.case.id_case)
        self.assertIn("Genera primero", messages.showinfo.call_args.args[1])
        dialog.assert_not_called()


class TransportTest(unittest.TestCase):
    def test_password_comes_from_the_keyring_and_never_from_settings(self):
        keyring = mock.Mock()
        keyring.get_password.return_value = "secreta"
        with mock.patch.dict(sys.modules, {"keyring": keyring}):
            self.assertEqual("secreta", mail_transport.keyring_password("usuario@ejemplo.es"))
            mail_transport.store_password("usuario@ejemplo.es", "nueva")
        keyring.get_password.assert_called_with(mail_transport.KEYRING_SERVICE, "usuario@ejemplo.es")
        keyring.set_password.assert_called_with(mail_transport.KEYRING_SERVICE, "usuario@ejemplo.es", "nueva")
        self.assertNotIn("password", mail_transport.SmtpSettings.__dataclass_fields__)

    def test_smtp_uses_starttls_and_hides_the_password_in_errors(self):
        settings = mail_transport.SmtpSettings(host="smtp.ejemplo.es", port=587, username="u@ejemplo.es")
        server = mock.Mock()
        server.login.side_effect = mail_transport.smtplib.SMTPAuthenticationError(535, b"bad secreta")
        with mock.patch.object(mail_transport.smtplib, "SMTP", return_value=server) as smtp:
            transport = mail_transport.SmtpTransport(settings, password_provider=lambda _u: "secreta")
            with self.assertRaises(MailTransportError) as raised:
                transport.__enter__()
        smtp.assert_called_once_with("smtp.ejemplo.es", 587, timeout=30)
        server.starttls.assert_called_once()
        self.assertNotIn("secreta", str(raised.exception))

    def test_settings_round_trip_without_password(self):
        directory = tempfile.TemporaryDirectory(ignore_cleanup_errors=True)
        self.addCleanup(directory.cleanup)
        database = Path(directory.name) / "g.db"
        with redirect_stdout(StringIO()):
            gestor_bd.crear_bd(str(database))
        connection = sqlite3.connect(database)
        self.addCleanup(connection.close)
        settings = mail_transport.SmtpSettings(host="smtp.ejemplo.es", port=465, username="u@ejemplo.es", security="ssl")
        mail_transport.save_smtp_settings(connection, settings)
        self.assertEqual(settings, mail_transport.load_smtp_settings(connection))
        stored = connection.execute("SELECT settings_json FROM mail_settings").fetchone()[0]
        self.assertNotIn("secreta", stored)
        with self.assertRaises(ValueError):
            mail_transport.save_smtp_settings(connection, mail_transport.SmtpSettings(host="", username="x"))


if __name__ == "__main__":
    unittest.main()
