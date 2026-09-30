import sys
import tempfile
import unittest
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
CORE_DIR = PROJECT_ROOT / "core"
if str(CORE_DIR) not in sys.path:
    sys.path.insert(0, str(CORE_DIR))

from provider_registry import (
    load_provider_registry,
    provider_registry_from_payload,
    resolve_provider,
)


def _profile(*, aliases, document_types=("invoice",)):
    return {
        "tax_ids": [],
        "aliases": list(aliases),
        "document_types": list(document_types),
        "service_family": "MANTENIMIENTO",
        "extractor_family": "standard_spanish_invoice",
        "required_signatures": [r"\bFACTURA\b"],
        "excluded_signatures": [r"\bPRESUPUESTO\b"],
    }


class ProviderRegistryTest(unittest.TestCase):
    def test_current_legacy_catalog_loads_through_schema_two_adapter(self):
        registry = load_provider_registry(PROJECT_ROOT / "config" / "proveedores.json")

        self.assertGreaterEqual(len(registry), 38)
        self.assertIn("ENDESA_LUZ_GENERAL", registry)

    def test_recurrent_global_profiles_require_compatible_invoice_structure(self):
        registry = load_provider_registry(PROJECT_ROOT / "config" / "proveedores.json")
        keys = (
            "TIERSAN", "ECHEMAN", "FENIE_ENERGIA", "CERRAJERA_MONCASI",
            "LIMPIEZAS_COTE", "SCHINDLER", "JPG_REPARACIONES_ELECTRICAS",
            "LIMPIEZAS_UTEBO", "ORONA", "MOEVE", "TRITERMIA", "TELESER",
            "ISS", "LABOIL", "ISTA", "ASCENSORS_SALES",
            "GAS_INSTALACIONES_MANTENIMIENTOS", "VILAHEXDOSS",
            "MARTINEZ_OTERO", "LIMPIEZAS_MOREDA", "JARDINERIA_JESUS_GRACIA",
            "LIMPIEZAS_MARCEN",
        )
        for key in keys:
            profile = registry[key]
            alias = profile.aliases[0]
            with self.subTest(provider_key=key):
                match = resolve_provider(
                    registry,
                    f"{alias} FACTURA F-1 Base imponible 10,00 IVA 2,10 Total 12,10 EUR",
                    "factura.pdf",
                    "invoice",
                )
                self.assertIsNotNone(match)
                self.assertEqual(key, match.provider_key)
                self.assertIsNone(
                    resolve_provider(
                        registry,
                        f"{alias} PRESUPUESTO P-1 Total 12,10 EUR",
                        "presupuesto.pdf",
                        "quote",
                    )
                )

    def test_tax_id_beats_customer_and_bank_names(self):
        registry = provider_registry_from_payload({"proveedores": {
            "EMISOR": {
                "tax_ids": ["B12345678"],
                "aliases": ["Emisor Real"],
                "document_types": ["invoice"],
                "service_family": "MANTENIMIENTO",
                "extractor_family": "standard_spanish_invoice",
            },
            "BANCO": {
                "tax_ids": ["A87654321"],
                "aliases": ["Banco Ejemplo"],
                "document_types": ["bank_receipt"],
                "service_family": "PAGO",
                "extractor_family": "standard_spanish_invoice",
            },
        }})
        match = resolve_provider(
            registry,
            "Cliente Comunidad CIF H00000000 Banco Ejemplo Emisor Real "
            "CIF B12345678 FACTURA F-9",
            "factura.pdf",
            "invoice",
        )
        self.assertIsNotNone(match)
        self.assertEqual(("EMISOR", "high"), (match.provider_key, match.confidence))
        self.assertIn("B12345678", match.evidence)

    def test_equal_alias_evidence_is_ambiguous(self):
        registry = provider_registry_from_payload({"proveedores": {
            "UNO": _profile(aliases=["ACME"]),
            "DOS": _profile(aliases=["ACME"]),
        }})
        self.assertIsNone(
            resolve_provider(registry, "ACME FACTURA F-1", "f.pdf", "invoice")
        )

    def test_provider_document_type_must_match(self):
        registry = provider_registry_from_payload({"proveedores": {
            "UNO": _profile(aliases=["ACME"], document_types=["invoice"]),
        }})
        self.assertIsNone(
            resolve_provider(registry, "ACME PRESUPUESTO", "p.pdf", "quote")
        )

    def test_duplicate_tax_id_is_rejected(self):
        profile = {
            "tax_ids": ["B12345678"],
            "aliases": ["Proveedor"],
            "document_types": ["invoice"],
            "service_family": "MANTENIMIENTO",
            "extractor_family": "standard_spanish_invoice",
        }
        with self.assertRaisesRegex(ValueError, "B12345678"):
            provider_registry_from_payload({"proveedores": {
                "UNO": profile,
                "DOS": {**profile, "aliases": ["Otro"]},
            }})

    def test_legacy_profile_is_adapted_without_losing_regex(self):
        registry = provider_registry_from_payload({"proveedores": {
            "LEGACY": {
                "nombre_display": "Proveedor antiguo",
                "tipo_suministro": "GAS",
                "firmas_identificacion": ["Proveedor Antiguo"],
                "firmas_requeridas": [r"\bFACTURA\b"],
                "firma_exclusion": r"\bPRESUPUESTO\b",
                "regex": {"importe_total": "TOTAL (?P<valor>.+)"},
            },
        }})
        profile = registry["LEGACY"]
        self.assertEqual("Proveedor antiguo", profile.display_name)
        self.assertEqual("GAS", profile.service_family)
        self.assertEqual("TOTAL (?P<valor>.+)", profile.legacy["regex"]["importe_total"])
        self.assertEqual(
            "LEGACY",
            resolve_provider(
                registry,
                "Proveedor Antiguo FACTURA F-1",
                "factura.pdf",
                "invoice",
            ).provider_key,
        )


class ProviderResolutionRobustnessTest(unittest.TestCase):
    def registry(self, **profiles):
        return provider_registry_from_payload({"proveedores": profiles})

    def profile(self, **overrides):
        base = {
            "aliases": [], "tax_ids": [], "document_types": ["invoice"],
            "service_family": "MANTENIMIENTO",
        }
        base.update(overrides)
        return base

    def test_alias_is_found_when_the_pdf_glued_the_words_together(self):
        registry = self.registry(ACME=self.profile(aliases=["Ascensores Acme"]))
        match = resolve_provider(registry, "FACTURA de ASCENSORESACME S.L.", "f.pdf", "invoice")
        self.assertEqual("ACME", match.provider_key)

    def test_short_aliases_are_not_matched_inside_other_words(self):
        registry = self.registry(TU=self.profile(aliases=["Tu Luz"]))
        self.assertIsNone(resolve_provider(registry, "ESTU LUZ BRILLANTE", "f.pdf", "invoice"))

    def test_tax_id_with_separators_matches_but_not_inside_longer_tokens(self):
        registry = self.registry(ACME=self.profile(tax_ids=["B50123456"]))
        self.assertEqual(
            "ACME", resolve_provider(registry, "CIF: B-50.123.456", "f.pdf", "invoice").provider_key
        )
        self.assertIsNone(resolve_provider(registry, "Ref ZB501234567 pago", "f.pdf", "invoice"))

    def test_excluded_signature_vetoes_a_tax_id_match(self):
        registry = self.registry(
            ACME=self.profile(tax_ids=["B50123456"], excluded_signatures=["Energia XXI"])
        )
        self.assertIsNone(
            resolve_provider(registry, "Energía XXI CIF B50123456", "f.pdf", "invoice")
        )

    def test_two_tax_ids_are_disambiguated_by_the_issuer_signature(self):
        registry = self.registry(
            ISSUER=self.profile(tax_ids=["B50123456"], aliases=["Comercial Norte"]),
            DISTRIBUTOR=self.profile(tax_ids=["A28123456"], aliases=["Distribuidora Sur"]),
        )
        match = resolve_provider(
            registry, "Comercial Norte CIF B50123456 distribuidora CIF A28123456", "f.pdf", "invoice"
        )
        self.assertEqual("ISSUER", match.provider_key)


if __name__ == "__main__":
    unittest.main()
