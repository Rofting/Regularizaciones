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

    def test_invoice_structure_wins_when_utility_invoice_contains_meter_readings(self):
        result = classify_document(
            "FACTURA Nº 00568008 Fecha de factura 27/10/2025 "
            "Lectura anterior 120 Lectura actual 164 Consumo 44 m3 "
            "Base imponible 159,21 IVA 15,93 Total a pagar 175,14 €",
            "00568008.pdf",
        )

        self.assertEqual(("invoice", "high"), (result.kind, result.confidence))

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


class DocumentClassifierPassingMentionsTest(unittest.TestCase):
    INVOICE = (
        "FACTURA Nº F-9 Fecha factura 01/04/2026 Base imponible 100,00 IVA 21,00 "
        "Total 121,00 € Vencimiento 15/04/2026 "
    )

    def test_invoice_paid_by_direct_debit_stays_an_invoice(self):
        result = classify_document(self.INVOICE + "x " * 400 + "Forma de pago: adeudo SEPA", "f9.pdf")
        self.assertEqual("invoice", result.kind)

    def test_invoice_referencing_a_quote_or_delivery_note_stays_an_invoice(self):
        filler = "x " * 400
        text = self.INVOICE + filler + "según presupuesto 123 y albarán 45"
        self.assertEqual("invoice", classify_document(text, "f9.pdf").kind)

    def test_word_abono_alone_does_not_make_a_credit_note(self):
        result = classify_document(self.INVOICE + "Cuota de abono mensual", "f9.pdf")
        self.assertEqual("invoice", result.kind)
        credit = classify_document("FACTURA DE ABONO A-1 IVA Total -121,00", "a1.pdf")
        self.assertEqual("credit_note", credit.kind)

    def test_header_keyword_still_wins_for_a_real_quote_with_totals(self):
        text = "PRESUPUESTO P-8 Factura proforma Base imponible 100 IVA 21 Total 121"
        self.assertEqual("quote", classify_document(text, "p8.pdf").kind)


if __name__ == "__main__":
    unittest.main()
