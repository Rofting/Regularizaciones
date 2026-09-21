import sys
import unittest
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
CORE_DIR = PROJECT_ROOT / "core"
if str(CORE_DIR) not in sys.path:
    sys.path.insert(0, str(CORE_DIR))

from document_classifier import classify_document


class DocumentClassifierTest(unittest.TestCase):
    def test_invoice_requires_invoice_structure_not_only_an_amount(self):
        result = classify_document(
            "FACTURA Nº F-102 Fecha 12/03/2026 Base imponible 100,00 "
            "IVA 21,00 Total 121,00 €",
            "F-102.pdf",
        )
        self.assertEqual(("invoice", "high"), (result.kind, result.confidence))

        amount_only = classify_document("TOTAL 121,00 €", "documento.pdf")
        self.assertNotEqual("invoice", amount_only.kind)

    def test_quote_wins_over_generic_total(self):
        result = classify_document(
            "PRESUPUESTO Nº P-8 Fecha 12/03/2026 Total 121,00 € Validez 30 días",
            "presupuesto.pdf",
        )
        self.assertEqual("quote", result.kind)

    def test_delivery_note_and_bank_receipt_never_become_invoices(self):
        fixtures = (
            ("ALBARÁN Entrega de gasóleo 900 litros Total 1.000 €", "delivery_note"),
            (
                "JUSTIFICANTE DE TRANSFERENCIA IBAN ES00 Ordenante Comunidad "
                "Importe 121,00 €",
                "bank_receipt",
            ),
        )
        for text, expected in fixtures:
            with self.subTest(expected=expected):
                self.assertEqual(expected, classify_document(text, "documento.pdf").kind)

    def test_meter_table_is_reading(self):
        result = classify_document(
            "Propiedad Lectura anterior Lectura actual Consumo ACS PA2-1A 144 166 22",
            "Lecturas consumo 07 2025 a 07 2026.pdf",
        )
        self.assertEqual("reading", result.kind)

    def test_credit_note_owners_and_report_have_explicit_kinds(self):
        fixtures = (
            (
                "FACTURA RECTIFICATIVA R-18 Abono factura F-9 Total -121,00 €",
                "credit_note",
            ),
            (
                "LISTADO DE PROPIETARIOS Vivienda Propietario Email Coeficiente",
                "owners",
            ),
            (
                "INFORME TÉCNICO Estado de la instalación y recomendaciones",
                "report",
            ),
        )
        for text, expected in fixtures:
            with self.subTest(expected=expected):
                self.assertEqual(expected, classify_document(text, "documento.pdf").kind)

    def test_generic_operational_document_is_other_and_blank_is_unknown(self):
        self.assertEqual(
            "other",
            classify_document("PARTE DE TRABAJO Visita preventiva", "parte.pdf").kind,
        )
        self.assertEqual("unknown", classify_document("", "scan_001.pdf").kind)


if __name__ == "__main__":
    unittest.main()
