import sqlite3
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock


PROJECT_ROOT = Path(__file__).resolve().parents[1]
CORE_DIR = PROJECT_ROOT / "core"
if str(CORE_DIR) not in sys.path:
    sys.path.insert(0, str(CORE_DIR))

from document_text_service import TextExtraction, get_document_text
import lector_pdf


class DocumentTextServiceTest(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.path = Path(self.temp_dir.name) / "factura.pdf"
        self.path.write_bytes(b"synthetic-pdf")
        self.connection = sqlite3.connect(":memory:")
        self.connection.execute("""
            CREATE TABLE document_text_cache (
                sha256 TEXT NOT NULL,
                extractor_version TEXT NOT NULL,
                text_content TEXT NOT NULL,
                method TEXT NOT NULL,
                pages_json TEXT NOT NULL,
                diagnostics_json TEXT NOT NULL,
                duration_ms INTEGER NOT NULL,
                created_at TEXT NOT NULL DEFAULT (datetime('now')),
                PRIMARY KEY (sha256, extractor_version)
            )
        """)

    def tearDown(self):
        self.connection.close()
        self.temp_dir.cleanup()

    def test_same_sha_and_version_uses_cached_text(self):
        calls = []

        def extractor(path, max_pages):
            calls.append(path)
            return TextExtraction("texto factura", "pdf_text", (1,), {}, 3, False)

        first = get_document_text(self.connection, self.path, extractor=extractor)
        second = get_document_text(self.connection, self.path, extractor=extractor)

        self.assertEqual("texto factura", second.text)
        self.assertEqual(1, len(calls))
        self.assertFalse(first.from_cache)
        self.assertTrue(second.from_cache)

    def test_new_extractor_version_invalidates_cache(self):
        calls = []

        def extractor(path, max_pages):
            calls.append(path)
            return TextExtraction("texto", "pdf_text", (1,), {}, 1, False)

        get_document_text(
            self.connection, self.path, extractor_version="v1", extractor=extractor
        )
        get_document_text(
            self.connection, self.path, extractor_version="v2", extractor=extractor
        )
        self.assertEqual(2, len(calls))

    def test_timeout_returns_recoverable_diagnostic(self):
        def timed_out(_path, _max_pages):
            raise TimeoutError("OCR excedió 45 segundos")

        result = get_document_text(self.connection, self.path, extractor=timed_out)
        self.assertEqual("timeout", result.method)
        self.assertEqual("OCR_TIMEOUT", result.diagnostics["code"])

    def test_cache_key_changes_when_file_content_changes(self):
        calls = []

        def extractor(path, max_pages):
            calls.append(path.read_bytes())
            return TextExtraction("texto", "pdf_text", (1,), {}, 1, False)

        get_document_text(self.connection, self.path, extractor=extractor)
        self.path.write_bytes(b"changed-pdf")
        get_document_text(self.connection, self.path, extractor=extractor)

        self.assertEqual(2, len(calls))

    def test_pdf_reader_reuses_supplied_text_without_reopening_document(self):
        config = {
            "tipo_suministro": "MANTENIMIENTO",
            "regex": {},
            "cups_comunidades": {},
        }
        with (
            mock.patch.object(lector_pdf, "extraer_texto") as text_reader,
            mock.patch.object(lector_pdf, "extraer_texto_ocr_con_diagnostico") as ocr,
            mock.patch.object(
                lector_pdf, "identificar_proveedor", return_value=("SYNTHETIC", config)
            ),
            mock.patch.object(
                lector_pdf,
                "extraer_datos_factura",
                return_value={"importe_total": 121.0},
            ),
        ):
            result = lector_pdf.procesar_archivo(
                str(self.path),
                "658",
                proveedores={"proveedores": {}},
                extracted_text="FACTURA F-1 TOTAL 121,00 EUR",
            )

        self.assertTrue(result["ok"])
        text_reader.assert_not_called()
        ocr.assert_not_called()


if __name__ == "__main__":
    unittest.main()
