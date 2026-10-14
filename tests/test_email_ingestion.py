import hashlib
import json
import sqlite3
import struct
import sys
import tempfile
import unittest
from contextlib import redirect_stdout
from datetime import date
from email.message import EmailMessage
from io import StringIO
from pathlib import Path
from unittest.mock import Mock, patch

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "core"))
from email_sources import read_email_sources, safe_attachment_name
import email_ingestion
import case_ingestion
import expedient_service
import expedient_ui
import gestor_bd
import db_migrations
from source_analysis import SourceAnalysis
from tests.test_source_preview import write_pdf


def write_eml(path, attachments=()):
    message = EmailMessage()
    message["Subject"] = "Facturas de la comunidad — enero"
    message["From"] = "Proveedor <proveedor@example.test>"
    message.set_content("Adjunto los documentos. TOTAL 999999 no es una factura.")
    for name, data in attachments:
        message.add_attachment(data, maintype="application", subtype="octet-stream", filename=name)
    path.write_bytes(message.as_bytes())
    return path


def write_msg(path, attachments):
    """MSG OLE/MAPI sintético leído por extract-msg real, sin mocks."""
    from extract_msg.ole_writer import OleWriter
    writer = OleWriter()
    prop = lambda tag, value: struct.pack("<IIII", tag, 0, value, 0)
    writer.addEntry("__properties_version1.0", b"\0" * 32 + prop(0x340D0003, 0x40000))
    writer.addEntry("__substg1.0_001A001F", "IPM.Note".encode("utf-16-le"))
    writer.addEntry("__substg1.0_0037001F", "Facturas de prueba".encode("utf-16-le"))
    writer.addEntry("__nameid_version1.0", storage=True)
    for stream in ("00020102", "00030102", "00040102"):
        writer.addEntry(f"__nameid_version1.0/__substg1.0_{stream}", b"")
    for index, (name, data) in enumerate(attachments):
        folder = f"__attach_version1.0_#{index:08X}"
        writer.addEntry(folder, storage=True)
        writer.addEntry(folder + "/__properties_version1.0", b"\0" * 8 + prop(0x37050003, 1))
        writer.addEntry(folder + "/__substg1.0_3707001F", name.encode("utf-16-le"))
        writer.addEntry(folder + "/__substg1.0_37010102", data)
    writer.write(str(path))
    return path


class EmailSourceReadingTest(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.addCleanup(self.directory.cleanup)
        self.root = Path(self.directory.name)

    def test_eml_decodes_names_and_keeps_only_supported_attachments(self):
        path = write_eml(self.root / "correo.eml", (
            ("Lecturas Álvarez.CSV", b"A;10\n"), ("logo.png", b"imagen"),
            ("vacio.pdf", b""), ("../../CON.pdf", b"pdf")))
        before = path.read_bytes()
        result = read_email_sources(path)
        self.assertEqual(["Lecturas Álvarez.CSV", "_CON.pdf"], [a.filename for a in result.attachments])
        self.assertIn("enero", result.subject)
        self.assertIn("proveedor@example.test", result.sender)
        self.assertEqual(2, len(result.notes))
        self.assertEqual(before, path.read_bytes())

    def test_msg_extracts_real_ole_mapi_attachment_and_closes_original(self):
        path = write_msg(self.root / "correo.msg", (("lecturas.csv", b"vivienda;consumo\nA;10\n"),))
        result = read_email_sources(path)
        self.assertEqual("Facturas de prueba", result.subject)
        self.assertEqual(b"vivienda;consumo\nA;10\n", result.attachments[0].data)
        path.unlink()  # También debe funcionar en Windows: handle OLE cerrado.

    def test_sizes_and_attachment_count_are_bounded(self):
        path = write_eml(self.root / "correo.eml", (("a.csv", b"123456"), ("b.csv", b"123")))
        with patch("email_sources.MAX_ATTACHMENT_BYTES", 4):
            result = read_email_sources(path)
        self.assertEqual(["b.csv"], [a.filename for a in result.attachments])
        self.assertIn("tamaño", result.notes[0])
        with patch("email_sources.MAX_ATTACHMENTS", 1):
            self.assertEqual(1, len(read_email_sources(path).attachments))
        with patch("email_sources.MAX_EMAIL_BYTES", 10):
            with self.assertRaisesRegex(ValueError, "50 MB"):
                read_email_sources(path)

    def test_folder_selection_includes_mail_only_for_the_supported_intake(self):
        path = write_eml(self.root / "correo.eml")
        (self.root / "~$temporal.msg").write_bytes(b"temp")
        self.assertEqual([], expedient_ui.source_files_in_folder(self.root))
        self.assertEqual([path], expedient_ui.source_files_in_folder(self.root, include_emails=True))
        for name in (r"C:\despacho\factura.pdf", "../../factura.pdf"):
            self.assertEqual("factura.pdf", safe_attachment_name(name))


class EmailIngestionTest(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.addCleanup(self.directory.cleanup)
        self.root = Path(self.directory.name).resolve()
        self.database = self.root / "gestion.db"
        with redirect_stdout(StringIO()):
            gestor_bd.crear_bd(str(self.database))
        self.connection = gestor_bd.conectar(str(self.database))
        self.addCleanup(self.connection.close)
        community = gestor_bd.obtener_o_crear_comunidad(self.connection, "658", "Prueba")
        other = gestor_bd.obtener_o_crear_comunidad(self.connection, "701", "Otra comunidad")
        self.case = expedient_service.create_case(self.connection, community, name="2026",
            start_date=date(2026, 1, 1), end_date=date(2026, 12, 31))
        self.archive = self.root / "archive"
        self.analysis = SourceAnalysis.invoice({"tipo_suministro": "GAS",
            "fecha_inicio": "2026-01-01", "fecha_fin": "2026-01-31", "importe_total": "123.00"})
        self.analyser = Mock(return_value=self.analysis)
        self.pdf = self.root / "original.pdf"
        write_pdf(self.pdf, [["Factura gas", "Total 123,00"]])

    def ingest(self, path, analyser=None):
        return email_ingestion.import_email_to_case(self.connection, self.case.id_case,
            source_path=path, archive_root=self.archive, analyser=analyser or self.analyser)

    def test_eml_archives_original_applies_invoice_and_links_attachment(self):
        path = write_eml(self.root / "correo.eml", (("factura.pdf", self.pdf.read_bytes()),))
        result = self.ingest(path)
        self.assertFalse(result.errors)
        self.assertTrue(result.results[0].created)
        self.assertEqual(1, self.connection.execute("SELECT COUNT(*) FROM facturas").fetchone()[0])
        email = self.connection.execute("SELECT * FROM source_documents WHERE id_document=?", (result.email_document_id,)).fetchone()
        self.assertEqual("not_applicable", email["status"])
        self.assertEqual(path.read_bytes(), Path(email["archived_path"]).read_bytes())
        attachment = self.connection.execute("SELECT * FROM email_attachments").fetchone()
        self.assertEqual(result.results[0].document.id_document, attachment["id_document"])
        self.assertEqual(hashlib.sha256(self.pdf.read_bytes()).hexdigest(), attachment["sha256"])
        self.assertEqual(self.pdf.read_bytes(), Path(attachment["archived_path"]).read_bytes())

    def test_repeated_mail_and_other_mail_and_loose_document_do_not_duplicate(self):
        loose = case_ingestion.add_analysed_document_to_case(self.connection, self.case.id_case,
            source_path=self.pdf, archive_root=self.archive, analysis=self.analysis)
        first = write_eml(self.root / "primero.eml", (("renombrada.pdf", self.pdf.read_bytes()),))
        second = write_eml(self.root / "segundo.eml", (("otra.pdf", self.pdf.read_bytes()),))
        second.write_bytes(second.read_bytes().replace(b"enero", b"febrero"))
        for path in (first, first, second):
            result = self.ingest(path)
            self.assertFalse(result.results[0].created)
            self.assertEqual(loose.document.id_document, result.results[0].document.id_document)
        self.analyser.assert_not_called()
        self.assertEqual(1, self.connection.execute("SELECT COUNT(*) FROM facturas").fetchone()[0])
        self.assertEqual(2, self.connection.execute("SELECT COUNT(*) FROM email_attachments").fetchone()[0])

    def test_same_attachment_names_keep_distinct_bytes_and_bad_one_does_not_stop_rest(self):
        path = write_eml(self.root / "correo.eml", (("factura.pdf", b"malformado"),
            ("factura.pdf", self.pdf.read_bytes())))
        def analyse(source):
            if source.read_bytes() == b"malformado":
                raise ValueError("PDF ilegible")
            return self.analysis
        result = self.ingest(path, analyse)
        self.assertEqual(1, len(result.errors))
        self.assertEqual(1, len(result.results))
        rows = self.connection.execute("SELECT * FROM email_attachments ORDER BY attachment_index").fetchall()
        self.assertEqual(["error", "imported"], [r["outcome"] for r in rows])
        self.assertNotEqual(rows[0]["archived_path"], rows[1]["archived_path"])

    def test_foreign_attachment_is_archived_without_entering_calculation(self):
        path = write_eml(self.root / "correo.eml", (("701_factura.pdf", self.pdf.read_bytes()),))
        result = self.ingest(path)
        self.assertEqual(1, result.foreign_count)
        self.assertEqual((), result.results)
        self.analyser.assert_not_called()
        row = self.connection.execute("SELECT * FROM email_attachments").fetchone()
        self.assertEqual("foreign_community", row["outcome"])
        self.assertIsNone(row["id_document"])
        self.assertTrue(Path(row["archived_path"]).is_file())

    def test_empty_and_broken_mail_are_reported_and_originals_remain_archived(self):
        empty = self.ingest(write_eml(self.root / "vacio.eml"))
        broken_path = self.root / "roto.msg"
        broken_path.write_bytes(b"no es un contenedor Outlook")
        broken = self.ingest(broken_path)
        self.assertIn("no contiene", empty.notes[0])
        self.assertEqual(1, len(broken.errors))
        self.assertEqual(2, self.connection.execute("SELECT COUNT(*) FROM email_imports").fetchone()[0])
        self.assertEqual(0, self.connection.execute("SELECT COUNT(*) FROM review_issues").fetchone()[0])

    def test_modified_attachment_or_original_is_not_silently_reused(self):
        path = write_eml(self.root / "correo.eml", (("factura.pdf", self.pdf.read_bytes()),))
        imported = self.ingest(path)
        row = self.connection.execute("SELECT archived_path FROM email_attachments").fetchone()
        Path(row[0]).write_bytes(b"sustituido")
        repeated = self.ingest(path)
        self.assertIn("ha cambiado", repeated.errors[0][1])
        original = self.connection.execute("SELECT archived_path FROM source_documents WHERE id_document=?",
                                           (imported.email_document_id,)).fetchone()[0]
        Path(original).write_bytes(b"correo sustituido")
        self.assertIn("ha cambiado", self.ingest(path).errors[0][1])

    def test_reanalysis_reads_attachment_and_never_uses_email_body_as_invoice(self):
        result = self.ingest(write_eml(self.root / "correo.eml", (("factura.pdf", self.pdf.read_bytes()),)))
        analyser = Mock(return_value=self.analysis)
        updated = case_ingestion.reanalyze_case_documents(self.connection, self.case.id_case,
            analyser=analyser, archive_root=self.archive)
        self.assertEqual(1, analyser.call_count)
        self.assertEqual(".pdf", analyser.call_args.args[0].suffix)
        self.assertEqual(1, len(updated))
        self.assertEqual("not_applicable", self.connection.execute("SELECT status FROM source_documents WHERE id_document=?",
                         (result.email_document_id,)).fetchone()[0])

    def test_modified_source_copy_is_rejected_and_provenance_link_survives(self):
        path = write_eml(self.root / "correo.eml", (("factura.pdf", self.pdf.read_bytes()),))
        result = self.ingest(path)
        child = result.results[0].document
        child.archived_path.write_bytes(b"otra factura")
        repeated = self.ingest(path)
        self.assertIn("ha cambiado", repeated.errors[0][1])
        self.assertEqual(child.id_document, self.connection.execute("SELECT id_document FROM email_attachments").fetchone()[0])

    def test_msg_flows_through_real_reader_and_same_ingestion(self):
        result = self.ingest(write_msg(self.root / "correo.msg", (("factura.pdf", self.pdf.read_bytes()),)))
        self.assertFalse(result.errors)
        self.assertEqual(1, len(result.results))
        self.assertEqual("Facturas de prueba", self.connection.execute("SELECT subject FROM email_imports").fetchone()[0])

    def test_repeated_migration_preserves_email_and_attachment_links(self):
        result = self.ingest(write_eml(self.root / "correo.eml", (("factura.pdf", self.pdf.read_bytes()),)))
        db_migrations.MIGRATIONS[23](self.connection)
        db_migrations.MIGRATIONS[23](self.connection)
        row = self.connection.execute("SELECT id_email_document,id_document FROM email_attachments").fetchone()
        self.assertEqual((result.email_document_id, result.results[0].document.id_document), tuple(row))
