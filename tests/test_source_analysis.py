import sys
import unittest
from pathlib import Path
from unittest import mock


PROJECT_ROOT = Path(__file__).resolve().parents[1]
CORE_DIR = PROJECT_ROOT / "core"
if str(CORE_DIR) not in sys.path:
    sys.path.insert(0, str(CORE_DIR))

import source_analysis


class SourceAnalysisTest(unittest.TestCase):
    def test_invoice_result_requires_only_missing_invoice_fields(self):
        result = source_analysis.analyse_pdf(
            Path("invoice.pdf"),
            pdf_processor=lambda *_: {"ok": True, "tipo": "FACTURA", "datos": {
                "fecha_inicio": "2026-01-01", "fecha_fin": "2026-01-31", "importe_total": 42.5,
            }},
        )
        self.assertEqual("invoice", result.kind)
        self.assertEqual(("fecha_inicio", "fecha_fin", "importe_total"), result.required_fields)
        self.assertEqual("42.5", result.candidates["importe_total"])

    def test_reading_result_never_requires_invoice_fields(self):
        result = source_analysis.analyse_pdf(
            Path("meters.pdf"),
            pdf_processor=lambda *_: {"ok": True, "tipo": "LECTURA_METRIGEST", "datos": {"tipo": "ACS"}},
        )
        self.assertEqual("reading", result.kind)
        self.assertEqual((), result.required_fields)

    def test_unknown_tabular_file_creates_one_classification_review(self):
        result = source_analysis.classify_headers(("Referencia", "Observacion"), suffix=".csv")
        self.assertEqual("unknown", result.kind)
        self.assertTrue(result.review_message.startswith("No se ha podido identificar"))

    def test_property_and_dwelling_headers_classify_owner_lists(self):
        for header in ("Propiedad", "Vivienda"):
            with self.subTest(header=header):
                result = source_analysis.classify_headers((header, "Nombre"), suffix=".csv")
                self.assertEqual("owners", result.kind)

    @mock.patch("lector_pdf.procesar_archivo")
    def test_default_pdf_analysis_uses_project_provider_configuration(self, processor):
        processor.return_value = {
            "ok": True,
            "tipo": "FACTURA",
            "datos": {"fecha_inicio": "2026-01-01", "fecha_fin": "2026-01-31", "importe_total": 42.5},
        }

        source_analysis.analyse_pdf(Path("invoice.pdf"), community_code="658")

        processor.assert_called_once_with(
            "invoice.pdf", "658",
            ruta_proveedores=str(PROJECT_ROOT / "config" / "proveedores.json"),
        )


if __name__ == "__main__":
    unittest.main()
