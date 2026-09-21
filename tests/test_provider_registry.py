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

        self.assertGreaterEqual(len(registry), 16)
        self.assertIn("ENDESA_LUZ_GENERAL", registry)

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


if __name__ == "__main__":
    unittest.main()
