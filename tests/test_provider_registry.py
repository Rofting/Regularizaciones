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



class TaxIdTest(unittest.TestCase):
    def test_control_digit_validation(self):
        from provider_registry import valid_spanish_tax_id
        for value in ("A95758389", "B50000009", "12345678Z", "X1234567L", "Q2826000H", "H12345674"):
            with self.subTest(value=value):
                self.assertTrue(valid_spanish_tax_id(value))
        for value in ("A95758388", "B50000000", "12345678A", "Q28260008"):
            with self.subTest(value=value):
                self.assertFalse(valid_spanish_tax_id(value))

    def test_find_tax_ids_accepts_separators_and_skips_invalid_tokens(self):
        from provider_registry import find_tax_ids
        text = "CIF: A-95.758.389 · cliente H12345674 · ref B50000000 · IBAN ES12 3456"
        self.assertEqual(("A95758389", "H12345674"), find_tax_ids(text))

    def test_community_tax_id_with_separators_is_found_by_the_pdf_reader(self):
        from lector_pdf import extraer_cif_pdf
        self.assertEqual("H12345674", extraer_cif_pdf("Comunidad N.I.F. H-12.345.674"))
        self.assertIsNone(extraer_cif_pdf("Proveedor B50000009"))

    def test_catalog_lookup_by_display_name_or_alias(self):
        from provider_registry import profile_for_name
        registry = provider_registry_from_payload({"proveedores": {
            "ACME": {"display_name": "Acme Energía", "aliases": ["ACME ENERGIA SL"],
                     "document_types": ["invoice"]},
        }})
        for value in ("ACME", "acme energia", "Acme Energía SL"):
            with self.subTest(value=value):
                self.assertEqual("ACME", profile_for_name(registry, value).key)
        self.assertIsNone(profile_for_name(registry, "Otro"))


class OfficeCatalogLayerTest(unittest.TestCase):
    BASE = {"proveedores": {
        "ACME": {"aliases": ["ACME"], "document_types": ["invoice"], "tipo_suministro": "GAS",
                 "regex": {"importe_total": "a"}},
        "OTRO": {"aliases": ["OTRO"], "document_types": ["invoice"]},
    }}

    def test_office_layer_adds_completes_and_disables_providers(self):
        from provider_registry import merge_catalogs
        merged = merge_catalogs(self.BASE, {
            "desactivados": ["otro"],
            "proveedores": {
                "acme": {"regex": {"fecha_inicio": "b"}, "cups_comunidades": {"7": "ES1"}},
                "LOCAL": {"aliases": ["LOCAL SL"], "document_types": ["invoice"]},
            },
        })["proveedores"]
        self.assertEqual({"ACME", "LOCAL"}, set(merged))
        self.assertEqual({"importe_total": "a", "fecha_inicio": "b"}, merged["ACME"]["regex"])
        self.assertEqual("GAS", merged["ACME"]["tipo_suministro"])
        self.assertEqual({"7": "ES1"}, merged["ACME"]["cups_comunidades"])

    def test_registry_and_pdf_reader_load_the_office_layer_next_to_the_catalog(self):
        import json
        from lector_pdf import cargar_proveedores
        from provider_registry import load_provider_registry
        with tempfile.TemporaryDirectory() as directory:
            base = Path(directory) / "proveedores.json"
            base.write_text(json.dumps(self.BASE), encoding="utf-8")
            (Path(directory) / "proveedores_despacho.json").write_text(json.dumps(
                {"proveedores": {"LOCAL": {"aliases": ["LOCAL SL"], "document_types": ["invoice"]}}}
            ), encoding="utf-8")
            self.assertIn("LOCAL", load_provider_registry(base))
            self.assertIn("LOCAL", cargar_proveedores(str(base))["proveedores"])

    def test_product_catalog_carries_no_community_data(self):
        import json
        from provider_registry import OFFICE_ONLY_KEYS
        payload = json.loads((PROJECT_ROOT / "config" / "proveedores.json").read_text(encoding="utf-8"))
        for key, provider in payload["proveedores"].items():
            for field in OFFICE_ONLY_KEYS:
                with self.subTest(provider=key, field=field):
                    self.assertNotIn(field, provider)


class FuzzyProviderTest(unittest.TestCase):
    REGISTRY = {"proveedores": {
        "ACME": {"aliases": ["Ascensores Acme Levante"], "document_types": ["invoice"]},
        "BRILLO": {"aliases": ["Limpiezas Brillo Total"], "document_types": ["invoice"],
                   "excluded_signatures": ["PRESUPUESTO"]},
    }}

    def test_ocr_damaged_name_is_matched_with_medium_confidence(self):
        registry = provider_registry_from_payload(self.REGISTRY)
        match = resolve_provider(registry, "FACTURA ASCENS0RES ACME LEVANTF S.L.", "f.pdf", "invoice")
        self.assertEqual(("ACME", "medium", "provider:fuzzy:v1"),
                         (match.provider_key, match.confidence, match.rule_id))

    def test_unrelated_text_and_exclusions_never_match(self):
        registry = provider_registry_from_payload(self.REGISTRY)
        self.assertIsNone(resolve_provider(registry, "Factura de suministro general", "f.pdf", "invoice"))
        self.assertIsNone(resolve_provider(
            registry, "PRESUPUESTO LIMPIEZAS BRILL0 TOTAL", "f.pdf", "invoice"))

    def test_generic_texts_do_not_match_the_real_catalog(self):
        registry = load_provider_registry(PROJECT_ROOT / "config" / "proveedores.json")
        for text in ("Factura de limpieza de la comunidad total 121",
                     "Factura de agua caliente total 50", "Factura mantenimiento ascensor"):
            with self.subTest(text=text):
                self.assertIsNone(resolve_provider(registry, text, "x.pdf", "invoice"))


class OfficeKeywordsTest(unittest.TestCase):
    def test_office_keywords_extend_the_product_lists(self):
        import json
        import keywords
        with tempfile.TemporaryDirectory() as directory:
            (Path(directory) / "palabras_clave.json").write_text(json.dumps(
                {"servicios": {"GAS": ["gas natural"]}}), encoding="utf-8")
            (Path(directory) / "palabras_clave_despacho.json").write_text(json.dumps(
                {"servicios": {"GAS": ["Butano Ejemplo"], "PISCINA": ["piscina"]}}), encoding="utf-8")
            keywords.load_keywords.cache_clear()
            merged = keywords.load_keywords(directory)
        keywords.load_keywords.cache_clear()
        self.assertEqual(["gas natural", "Butano Ejemplo"], merged["servicios"]["GAS"])
        self.assertEqual(["piscina"], merged["servicios"]["PISCINA"])

if __name__ == "__main__":
    unittest.main()
