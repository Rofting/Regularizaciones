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


if __name__ == "__main__":
    unittest.main()
