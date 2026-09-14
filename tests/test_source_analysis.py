import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock


PROJECT_ROOT = Path(__file__).resolve().parents[1]
CORE_DIR = PROJECT_ROOT / "core"
if str(CORE_DIR) not in sys.path:
    sys.path.insert(0, str(CORE_DIR))

import source_analysis
import lector_pdf


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

    def test_legacy_pdf_reader_finds_project_provider_configuration(self):
        providers = lector_pdf.cargar_proveedores()

        self.assertTrue(providers)

    @mock.patch("lector_pdf.procesar_archivo")
    def test_analyse_pdf_passes_a_preloaded_catalog_to_the_reader(self, processor):
        providers = {"proveedores": {}}
        processor.return_value = {"ok": True, "tipo": "FACTURA", "datos": {}}

        source_analysis.analyse_pdf(
            Path("invoice.pdf"),
            community_code="658", providers=providers,
        )

        processor.assert_called_once_with(
            "invoice.pdf", "658",
            ruta_proveedores=str(PROJECT_ROOT / "config" / "proveedores.json"),
            proveedores=providers,
        )

    def test_provider_catalog_is_loaded_once_per_path(self):
        lector_pdf.cargar_proveedores.cache_clear()
        config = str(PROJECT_ROOT / "config" / "proveedores.json")
        with mock.patch("lector_pdf.json.load", wraps=lector_pdf.json.load) as load:
            lector_pdf.cargar_proveedores(config)
            lector_pdf.cargar_proveedores(config)

        self.assertEqual(1, load.call_count)
        lector_pdf.cargar_proveedores.cache_clear()

    def test_unknown_provider_invoice_with_number_date_and_total_is_classified(self):
        result = source_analysis.analyse_pdf(
            Path("naturgy.pdf"),
            pdf_processor=lambda *_args, **_kwargs: {
                "ok": False,
                "motivo": "PROVEEDOR_NO_IDENTIFICADO",
                "detalle": "Ningún proveedor reconocido",
                "fragment": (
                    "Factura N.º FE26390022198715 Fecha de emisión: 08/06/2026 "
                    "Total a pagar 326,98 €"
                ),
            },
        )

        self.assertEqual("invoice", result.kind)
        self.assertEqual("medium", result.confidence)
        self.assertEqual("326,98", result.candidates["importe_total"])

    def test_unstructured_unknown_pdf_remains_classification_review(self):
        result = source_analysis.analyse_pdf(
            Path("nota.pdf"),
            pdf_processor=lambda *_args, **_kwargs: {
                "ok": False,
                "motivo": "PROVEEDOR_NO_IDENTIFICADO",
                "fragment": "Aviso interno",
            },
        )

        self.assertEqual("unknown", result.kind)
        self.assertIn("PROVEEDOR_NO_IDENTIFICADO", result.review_message)

    def test_catalogue_identifies_mantenimientos_zaragoza_invoice(self):
        providers = lector_pdf.cargar_proveedores(
            str(PROJECT_ROOT / "config" / "proveedores.json")
        )
        key, config = lector_pdf.identificar_proveedor(
            "INSTALACIONES ZARAGOZA S.L. MANTENIMIENTO DE SALAS DE CALDERAS",
            "F2524083.pdf", providers,
        )

        self.assertEqual("MANTENIMIENTOS_ZARAGOZA", key)
        self.assertEqual("MANTENIMIENTO", config["tipo_suministro"])

    def test_catalogue_identifies_naturgy_clientes_gas_invoice(self):
        providers = lector_pdf.cargar_proveedores(
            str(PROJECT_ROOT / "config" / "proveedores.json")
        )
        key, config = lector_pdf.identificar_proveedor(
            "Naturgy Clientes, S.A.U. Estás en mercado libre. "
            "Período gas: del 26/04/2026 al 29/05/2026",
            "naturgy.pdf", providers,
        )

        self.assertEqual("NATURGY_CLIENTES_GAS", key)
        self.assertEqual("GAS", config["tipo_suministro"])

    def test_catalogue_identifies_totalenergies_gas_invoice(self):
        providers = lector_pdf.cargar_proveedores(
            str(PROJECT_ROOT / "config" / "proveedores.json")
        )
        key, config = lector_pdf.identificar_proveedor(
            "Factura gas TotalEnergies Electricidad y Gas España, S.A.U. "
            "TOTAL IMPORTE FACTURA 6.988,29 €",
            "FGAS_2600026546.pdf", providers,
        )

        self.assertEqual("TOTALENERGIES_GAS", key)
        self.assertEqual("GAS", config["tipo_suministro"])

    def test_totalenergies_electricity_document_is_not_classified_as_gas(self):
        providers = lector_pdf.cargar_proveedores(
            str(PROJECT_ROOT / "config" / "proveedores.json")
        )
        key, config = lector_pdf.identificar_proveedor(
            "Factura electricidad TotalEnergies Electricidad y Gas España, S.A.U. "
            "TOTAL IMPORTE FACTURA 125,00 €",
            "electricidad.pdf", providers,
        )

        self.assertIsNone(key)
        self.assertIsNone(config)

    def test_totalenergies_profile_extracts_the_invoice_period_and_total(self):
        providers = lector_pdf.cargar_proveedores(
            str(PROJECT_ROOT / "config" / "proveedores.json")
        )
        config = providers["proveedores"]["TOTALENERGIES_GAS"]

        result = lector_pdf.extraer_datos_factura(
            "Nº Factura: FGAS2600026546 FECHA FACTURA: 28 de abril de 2026 "
            "Periodo de facturación: De 31/03/2026 al 23/04/2026 "
            "Total (kWh): 84.684,16 TOTAL IMPORTE FACTURA 6.988,29 €",
            config,
        )

        self.assertEqual("2026-03-31", result["fecha_inicio"])
        self.assertEqual("2026-04-23", result["fecha_fin"])
        self.assertEqual(84684.16, result["consumo_kwh"])
        self.assertEqual(6988.29, result["importe_total"])

    def test_catalogue_keeps_mizar_invoices_under_instalaciones_zaragoza(self):
        providers = lector_pdf.cargar_proveedores(
            str(PROJECT_ROOT / "config" / "proveedores.json")
        )
        key, config = lector_pdf.identificar_proveedor(
            "Factura Número # F260009 MANTENIMIENTOS "
            "INSTALACIONES ZARAGOZA\nS.L. B99091316 Aviso de fuga de ACS",
            "MIZAR_F260009.pdf", providers,
        )

        self.assertEqual("MANTENIMIENTOS_ZARAGOZA", key)
        self.assertEqual("MANTENIMIENTO", config["tipo_suministro"])

    def test_totalenergies_credit_note_is_accepted_as_a_negative_invoice(self):
        providers = lector_pdf.cargar_proveedores(
            str(PROJECT_ROOT / "config" / "proveedores.json")
        )
        text = (
            "Factura gas TotalEnergies Electricidad y Gas España, S.A.U. "
            "Nº Factura: ABOGAS2600002334 FECHA FACTURA: 8 de junio de 2026 "
            "Periodo de facturación: De 31/03/2026 al 23/04/2026 "
            "Total (kWh): 84.684,16 TOTAL IMPORTE FACTURA -6.988,29 €"
        )
        with tempfile.TemporaryDirectory() as temporary_directory:
            path = Path(temporary_directory) / "credit.pdf"
            path.touch()
            with mock.patch("lector_pdf.extraer_texto", return_value=text):
                result = lector_pdf.procesar_archivo(
                    str(path), "658", proveedores=providers,
                )

        self.assertTrue(result["ok"])
        self.assertEqual("FACTURA", result["tipo"])
        self.assertEqual(-6988.29, result["datos"]["importe_total"])

    def test_invoice_analysis_never_renames_the_archived_source_file(self):
        providers = lector_pdf.cargar_proveedores(
            str(PROJECT_ROOT / "config" / "proveedores.json")
        )
        text = (
            "Factura gas TotalEnergies Electricidad y Gas España, S.A.U. "
            "CIF/NIF: H50385863 Nº Factura: FGAS2600026546 "
            "FECHA FACTURA: 28 de abril de 2026 "
            "Periodo de facturación: De 31/03/2026 al 23/04/2026 "
            "Total (kWh): 84.684,16 TOTAL IMPORTE FACTURA 6.988,29 €"
        )
        with tempfile.TemporaryDirectory() as temporary_directory:
            path = Path(temporary_directory) / "archived-source.pdf"
            path.touch()
            with mock.patch("lector_pdf.extraer_texto", return_value=text):
                result = lector_pdf.procesar_archivo(
                    str(path), "658", proveedores=providers,
                )

            self.assertTrue(path.exists())
            self.assertFalse((path.parent / "658_archived-source.pdf").exists())

        self.assertTrue(result["ok"])


if __name__ == "__main__":
    unittest.main()
