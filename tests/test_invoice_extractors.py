import sys
import unittest
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
CORE_DIR = PROJECT_ROOT / "core"
if str(CORE_DIR) not in sys.path:
    sys.path.insert(0, str(CORE_DIR))

from invoice_extractors import extract_invoice_fields
from provider_registry import ProviderProfile


def profile(extractor_family: str) -> ProviderProfile:
    return ProviderProfile(
        key="SYNTHETIC",
        display_name="Proveedor sintético",
        tax_ids=(),
        aliases=("PROVEEDOR SINTETICO",),
        document_types=("invoice",),
        service_family="MANTENIMIENTO",
        extractor_family=extractor_family,
        required_signatures=(),
        excluded_signatures=(),
        legacy={},
    )


class InvoiceExtractorsTest(unittest.TestCase):
    def test_standard_invoice_returns_field_level_evidence(self):
        result = extract_invoice_fields(
            profile("standard_spanish_invoice"),
            "Factura F-9 Fecha 01/04/2026 Periodo 01/03/2026 a 31/03/2026 "
            "Base imponible 100,00 € IVA 21,00 € Total 121,00 €",
        )
        self.assertEqual("2026-03-01", result.fields["fecha_inicio"].value)
        self.assertEqual("2026-03-31", result.fields["fecha_fin"].value)
        self.assertEqual("121.00", result.fields["importe_total"].value)
        self.assertEqual("high", result.fields["importe_total"].confidence)

    def test_total_that_does_not_reconcile_is_not_high_confidence(self):
        result = extract_invoice_fields(
            profile("standard_spanish_invoice"),
            "Factura F-9 Base imponible 100,00 € IVA 21,00 € Total 999,00 €",
        )
        self.assertNotEqual("high", result.fields["importe_total"].confidence)
        self.assertIn("total_not_reconciled", result.diagnostics)

    def test_unlabelled_numbers_are_never_dates_or_totals(self):
        result = extract_invoice_fields(
            profile("standard_spanish_invoice"), "1231 1233 131231"
        )
        self.assertNotIn("fecha_inicio", result.fields)
        self.assertNotIn("importe_total", result.fields)

    def test_electricity_family_extracts_labelled_consumption(self):
        result = extract_invoice_fields(
            profile("electricity"),
            "Factura F-10 Periodo 01/04/2026 a 30/04/2026 "
            "Consumo total 1.234 kWh Total factura 250,00 €",
        )
        self.assertEqual("1234", result.fields["consumo_kwh"].value)
        self.assertEqual("MANTENIMIENTO", result.fields["tipo_suministro"].value)

    def test_all_supported_families_are_registered(self):
        families = (
            "standard_spanish_invoice",
            "electricity",
            "gas_fuel",
            "periodic_maintenance",
            "elevators",
            "boilers_hvac",
            "metering_management",
            "water_public_fees",
        )
        for family in families:
            with self.subTest(family=family):
                result = extract_invoice_fields(
                    profile(family), "Factura F-1 Total factura 10,00 €"
                )
                self.assertEqual("10.00", result.fields["importe_total"].value)


class InvoiceExtractorsRobustnessTest(unittest.TestCase):
    def extract(self, text, family="standard_spanish_invoice"):
        return extract_invoice_fields(profile(family), text)

    def test_summary_table_with_values_on_the_next_line(self):
        result = self.extract(
            "Base imponible % IVA Cuota IVA Total factura\n100,00 21 21,00 121,00\n"
        )
        self.assertEqual("100.00", result.fields["base_imponible"].value)
        self.assertEqual("21.00", result.fields["iva"].value)
        self.assertEqual("121.00", result.fields["importe_total"].value)
        self.assertEqual("high", result.fields["importe_total"].confidence)

    def test_reconciliation_skips_subtotals_and_tax_base_on_the_iva_line(self):
        result = self.extract(
            "Subtotal 1.000,00 €\nIVA 21% s/ 1.000,00 210,00 €\nTOTAL FACTURA 1.210,00 €"
        )
        self.assertEqual("210.00", result.fields["iva"].value)
        self.assertEqual("1210.00", result.fields["importe_total"].value)
        self.assertEqual("high", result.fields["importe_total"].confidence)

    def test_total_iva_line_is_not_the_invoice_total(self):
        result = self.extract("Total IVA 21,00\nTotal a pagar 121,00 €")
        self.assertEqual("121.00", result.fields["importe_total"].value)

    def test_dates_written_in_words_and_with_two_digit_years(self):
        result = self.extract(
            "Fecha de emisión: 5 de abril de 2026\nPeriodo de facturación: del 1/3/26 al 31/3/26"
        )
        self.assertEqual("2026-04-05", result.fields["fecha_factura"].value)
        self.assertEqual("2026-03-01", result.fields["fecha_inicio"].value)
        self.assertEqual("2026-03-31", result.fields["fecha_fin"].value)

    def test_period_with_en_dash_and_between_form(self):
        dash = self.extract("Periodo de consumo 01/03/2026 – 31/03/2026")
        self.assertEqual("2026-03-31", dash.fields["fecha_fin"].value)
        between = self.extract("Consumo entre el 01/03/2026 y el 31/03/2026")
        self.assertEqual("2026-03-01", between.fields["fecha_inicio"].value)

    def test_inverted_or_impossible_period_is_discarded(self):
        inverted = self.extract("Periodo 31/03/2026 a 01/03/2026")
        self.assertNotIn("fecha_inicio", inverted.fields)
        self.assertIn("period_inverted", inverted.diagnostics)
        impossible = self.extract("Periodo 31/02/2026 a 01/03/2026")
        self.assertNotIn("fecha_inicio", impossible.fields)

    def test_invoice_number_needs_a_digit_and_is_never_a_date(self):
        self.assertNotIn("num_factura", self.extract("FACTURA ELECTRICIDAD comunidad").fields)
        self.assertNotIn("num_factura", self.extract("Factura 01/04/2026").fields)
        self.assertEqual(
            "FE-2026/123", self.extract("Número de factura: FE-2026/123").fields["num_factura"].value
        )
        self.assertEqual("A-77", self.extract("Factura nº A-77").fields["num_factura"].value)

    def test_credit_note_amounts_keep_their_sign(self):
        self.assertEqual("-121.00", self.extract("Total a pagar 121,00- €").fields["importe_total"].value)
        self.assertEqual("-121.00", self.extract("Total a pagar -121,00 €").fields["importe_total"].value)

    def test_english_formatted_amounts(self):
        self.assertEqual("1234.56", self.extract("Total factura 1,234.56 EUR").fields["importe_total"].value)

    def test_service_type_is_inferred_only_without_a_profile_and_when_unambiguous(self):
        generic = ProviderProfile("U", "U", (), (), ("invoice",), "", "standard_spanish_invoice", (), (), {})
        inferred = extract_invoice_fields(generic, "Factura 1 suministro de gas natural")
        self.assertEqual("GAS", inferred.fields["tipo_suministro"].value)
        self.assertEqual("medium", inferred.fields["tipo_suministro"].confidence)
        ambiguous = extract_invoice_fields(generic, "Factura 1 gas natural y electricidad")
        self.assertNotIn("tipo_suministro", ambiguous.fields)
        known = extract_invoice_fields(profile("standard_spanish_invoice"), "Factura 1 gas natural")
        self.assertEqual("MANTENIMIENTO", known.fields["tipo_suministro"].value)

    def test_cups_is_extracted_for_energy_families(self):
        result = self.extract(
            "Factura F-1 CUPS: ES 0031 3007 1835 4001 RR0F Consumo facturado 1.234 kWh", "electricity"
        )
        self.assertEqual("ES0031300718354001RR0F", result.fields["cups"].value)
        self.assertEqual("1234", result.fields["consumo_kwh"].value)

    def test_water_consumption_with_cubic_metres(self):
        result = self.extract("Consumo facturado: 45 m³ Total 30,00 €", "water_public_fees")
        self.assertEqual("45", result.fields["consumo_m3"].value)


if __name__ == "__main__":
    unittest.main()
