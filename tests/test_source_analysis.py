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
            "Naturgy Clientes, S.A.U. Estás en mercado libre. Hola, aquí tienes tu factura de gas. "
            "Nº de factura: FE263900221987 Período gas: del 26/04/2026 al 29/05/2026 "
            "Total a pagar 326,98 €",
            "naturgy.pdf", providers,
        )

        self.assertEqual("NATURGY_CLIENTES_GAS", key)
        self.assertEqual("GAS", config["tipo_suministro"])

    def test_catalogue_identifies_naturgy_iberia_legacy_gas_invoice(self):
        providers = lector_pdf.cargar_proveedores(
            str(PROJECT_ROOT / "config" / "proveedores.json")
        )
        key, config = lector_pdf.identificar_proveedor(
            "Naturgy Iberia, S.A. - Mercado Libre Hola, aquí tienes tu factura de gas. "
            "Nº factura: FE25321495922757 "
            "Gas: del 25.07.2025 al 27.07.2025 Total a pagar 31,77 €",
            "naturgy-iberia.pdf", providers,
        )

        self.assertEqual("NATURGY_CLIENTES_GAS", key)
        self.assertEqual("GAS", config["tipo_suministro"])

    def test_catalogue_rejects_naturgy_iberia_without_gas_invoice_evidence(self):
        providers = lector_pdf.cargar_proveedores(
            str(PROJECT_ROOT / "config" / "proveedores.json")
        )
        key, config = lector_pdf.identificar_proveedor(
            "Naturgy Iberia, S.A. - Mercado Libre. Comunicado de mantenimiento.",
            "comunicado.pdf", providers,
        )

        self.assertIsNone(key)
        self.assertIsNone(config)

    def test_catalogue_rejects_naturgy_payment_notice_without_invoice_number(self):
        providers = lector_pdf.cargar_proveedores(
            str(PROJECT_ROOT / "config" / "proveedores.json")
        )
        key, config = lector_pdf.identificar_proveedor(
            "Naturgy Clientes, S.A.U. Aviso de pago pendiente de tu factura de gas. "
            "Período gas: del 25/06/2026 al 23/07/2026 Total a pagar 537,74 €",
            "aviso-naturgy.pdf", providers,
        )

        self.assertIsNone(key)
        self.assertIsNone(config)

    def test_catalogue_rejects_detailed_naturgy_payment_reminder(self):
        providers = lector_pdf.cargar_proveedores(
            str(PROJECT_ROOT / "config" / "proveedores.json")
        )
        key, config = lector_pdf.identificar_proveedor(
            "AVISO DE PAGO Naturgy Clientes, S.A.U. Hola, aquí tienes tu factura de gas. "
            "Nº de factura: FE263900221987 Período gas: del 26/04/2026 al 29/05/2026 "
            "Total a pagar 326,98 €. Esta factura está pendiente.",
            "aviso-detallado-naturgy.pdf", providers,
        )

        self.assertIsNone(key)
        self.assertIsNone(config)

    def test_naturgy_iberia_profile_extracts_dot_dates_consumption_and_total(self):
        providers = lector_pdf.cargar_proveedores(
            str(PROJECT_ROOT / "config" / "proveedores.json")
        )
        config = providers["proveedores"]["NATURGY_CLIENTES_GAS"]

        result = lector_pdf.extraer_datos_factura(
            "Nº factura: FE25321495922757 Fecha de emisión: 13.08.2025 "
            "Gas: del 25.07.2025 al 27.07.2025 Total a pagar 31,77 € "
            "Consumo kWh: 30 347kWh",
            config,
        )

        self.assertEqual("2025-07-25", result["fecha_inicio"])
        self.assertEqual("2025-07-27", result["fecha_fin"])
        self.assertEqual(347.0, result["consumo_kwh"])
        self.assertEqual(31.77, result["importe_total"])

    def test_naturgy_profile_ignores_chart_axis_before_explicit_total(self):
        providers = lector_pdf.cargar_proveedores(
            str(PROJECT_ROOT / "config" / "proveedores.json")
        )
        config = providers["proveedores"]["NATURGY_CLIENTES_GAS"]

        result = lector_pdf.extraer_datos_factura(
            "Naturgy Clientes, S.A.U. Período gas: del 25/06/2026 al 23/07/2026 "
            "Total a pagar\n4500\n537, 74 €\nTotal gas 444,41 €\n"
            "Total a pagar 537,74 €",
            config,
        )

        self.assertEqual(537.74, result["importe_total"])

    def test_naturgy_profile_interprets_dotted_kwh_as_thousands(self):
        providers = lector_pdf.cargar_proveedores(
            str(PROJECT_ROOT / "config" / "proveedores.json")
        )
        config = providers["proveedores"]["NATURGY_CLIENTES_GAS"]

        result = lector_pdf.extraer_datos_factura(
            "Naturgy Clientes, S.A.U. Período gas: del 25/06/2026 al 23/07/2026 "
            "Consumo gas 4.011 kWh Total a pagar 537,74 €",
            config,
        )

        self.assertEqual(4011.0, result["consumo_kwh"])

    def test_catalogue_identifies_endesa_electricity_without_a_community_cups(self):
        providers = lector_pdf.cargar_proveedores(
            str(PROJECT_ROOT / "config" / "proveedores.json")
        )
        key, config = lector_pdf.identificar_proveedor(
            "DATOSDELAFACTURA Nºfactura:P26CON005275625 "
            "Periododefacturación:del31/12/2025a31/01/2026 "
            "EndesaEnergía,S.A.Unipersonal. CIFA81948077. "
            "RESUMENDELAFACTURA Potencia 16,13€ Energía 177,54€ Total 248,75€",
            "factura-endesa.pdf", providers,
        )

        self.assertEqual("ENDESA_LUZ_GENERAL", key)
        self.assertEqual("ELECTRICIDAD", config["tipo_suministro"])

    def test_catalogue_does_not_classify_endesa_name_without_invoice_evidence(self):
        providers = lector_pdf.cargar_proveedores(
            str(PROJECT_ROOT / "config" / "proveedores.json")
        )
        key, config = lector_pdf.identificar_proveedor(
            "Aviso de mantenimiento de Endesa Energía, S.A. Unipersonal. CIF A81948077.",
            "aviso.pdf", providers,
        )

        self.assertIsNone(key)
        self.assertIsNone(config)

    def test_catalogue_does_not_classify_endesa_payment_notice_without_invoice_number(self):
        providers = lector_pdf.cargar_proveedores(
            str(PROJECT_ROOT / "config" / "proveedores.json")
        )
        key, config = lector_pdf.identificar_proveedor(
            "DATOS DE LA FACTURA RESUMEN DE LA FACTURA Endesa Energía, S.A. "
            "Unipersonal. CIF A81948077. Este aviso informa de un recibo pendiente.",
            "aviso-cobro.pdf", providers,
        )

        self.assertIsNone(key)
        self.assertIsNone(config)

    def test_general_endesa_profile_extracts_collapsed_electricity_consumption(self):
        providers = lector_pdf.cargar_proveedores(
            str(PROJECT_ROOT / "config" / "proveedores.json")
        )
        config = providers["proveedores"]["ENDESA_LUZ_GENERAL"]

        result = lector_pdf.extraer_datos_factura(
            "DATOSDELAFACTURA Nºfactura:P26CON005275625 "
            "Periododefacturación:del31/12/2025a31/01/2026(31días) "
            "EndesaEnergía,S.A.Unipersonal. CIFA81948077. "
            "ConsumoTotal 1.124,785 kWh Total 248,75€",
            config,
        )

        self.assertEqual(1124.785, result["consumo_kwh"])

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

    def test_catalogue_identifies_gomez_group_metering_service_invoice(self):
        providers = lector_pdf.cargar_proveedores(
            str(PROJECT_ROOT / "config" / "proveedores.json")
        )
        key, config = lector_pdf.identificar_proveedor(
            "FACTURA GOMEZ GROUP METERING, S.L.U. B80853237 F.FACTURA F. VCTO. "
            "Nº FACTURA LF26022852 "
            "MEDITRADE DEL EBRO SERVICIO DE LECTURA, FACTURACIÓN Y MANTENIMIENTO "
            "DE CONTADORES CONFORME CONTRATO EN VIGOR TOTAL FACTURA 125,24 €",
            "lectura-contadores.pdf", providers,
        )

        self.assertEqual("GOMEZ_GROUP_METERING", key)
        self.assertEqual("MANTENIMIENTO", config["tipo_suministro"])

    def test_gomez_group_metering_profile_extracts_invoice_date_and_total(self):
        providers = lector_pdf.cargar_proveedores(
            str(PROJECT_ROOT / "config" / "proveedores.json")
        )
        config = providers["proveedores"]["GOMEZ_GROUP_METERING"]

        result = lector_pdf.extraer_datos_factura(
            "F.FACTURA F. VCTO. Nº FACTURA RAZÓN SOCIAL LF26022852 02/03/2026 "
            "06/03/2026 GOMEZ GROUP METERING, S.L.U. TOTAL FACTURA ... 125,24 €",
            config,
        )

        self.assertEqual("LF26022852", result["num_factura"])
        self.assertEqual("2026-03-02", result["fecha_factura"])
        self.assertEqual(125.24, result["importe_total"])

    def test_gomez_group_metering_anchors_date_to_invoice_header(self):
        providers = lector_pdf.cargar_proveedores(
            str(PROJECT_ROOT / "config" / "proveedores.json")
        )
        config = providers["proveedores"]["GOMEZ_GROUP_METERING"]

        result = lector_pdf.extraer_datos_factura(
            "Fecha de prestación 01/02/2026 F.FACTURA F. VCTO. Nº FACTURA "
            "RAZÓN SOCIAL LF26022852 02/03/2026 06/03/2026 TOTAL FACTURA 125,24 €",
            config,
        )

        self.assertEqual("2026-03-02", result["fecha_factura"])

    def test_gomez_group_metering_infers_full_month_from_billed_readings_label(self):
        providers = lector_pdf.cargar_proveedores(
            str(PROJECT_ROOT / "config" / "proveedores.json")
        )
        config = providers["proveedores"]["GOMEZ_GROUP_METERING"]

        result = lector_pdf.extraer_datos_factura(
            "F.FACTURA F. VCTO. Nº FACTURA LF25040215 01/07/2025 05/07/2025 "
            "SERVICIO DE LECTURA, FACTURACIÓN Y MANTENIMIENTO DE CONTADORES. "
            "LECTURAS FACTURADAS: 2025 JUL TOTAL FACTURA 121,00 €",
            config,
        )

        self.assertEqual("2025-07-01", result["fecha_inicio"])
        self.assertEqual("2025-07-31", result["fecha_fin"])

    def test_maintenance_profile_uses_invoice_date_when_no_service_period_exists(self):
        providers = lector_pdf.cargar_proveedores(
            str(PROJECT_ROOT / "config" / "proveedores.json")
        )
        config = providers["proveedores"]["MANTENIMIENTOS_ZARAGOZA"]

        result = lector_pdf.extraer_datos_factura(
            "INSTALACIONES ZARAGOZA S.L. Factura Número F2523682 "
            "Fecha 04/07/2025 MANTENIMIENTO DE SALAS DE CALDERAS F2523682 - 100,00 €",
            config,
        )

        self.assertEqual("2025-07-04", result["fecha_inicio"])
        self.assertEqual("2025-07-04", result["fecha_fin"])

    def test_catalogue_rejects_gomez_payment_notice_without_service_evidence(self):
        providers = lector_pdf.cargar_proveedores(
            str(PROJECT_ROOT / "config" / "proveedores.json")
        )
        key, config = lector_pdf.identificar_proveedor(
            "GOMEZ GROUP METERING, S.L.U. FACTURA pendiente LF26022852 "
            "TOTAL FACTURA 125,24 €. Consulte el estado de su pago.",
            "aviso-gomez.pdf", providers,
        )

        self.assertIsNone(key)
        self.assertIsNone(config)

    def test_catalogue_rejects_detailed_gomez_payment_reminder(self):
        providers = lector_pdf.cargar_proveedores(
            str(PROJECT_ROOT / "config" / "proveedores.json")
        )
        key, config = lector_pdf.identificar_proveedor(
            "RECORDATORIO DE PAGO GOMEZ GROUP METERING, S.L.U. F.FACTURA F. VCTO. "
            "Nº FACTURA RAZÓN SOCIAL LF26022852 SERVICIO DE LECTURA, FACTURACIÓN "
            "Y MANTENIMIENTO DE CONTADORES CONFORME CONTRATO EN VIGOR TOTAL FACTURA 125,24 €",
            "aviso-detallado-gomez.pdf", providers,
        )

        self.assertIsNone(key)
        self.assertIsNone(config)

    def test_catalogue_rejects_gomez_service_text_without_structured_invoice_header(self):
        providers = lector_pdf.cargar_proveedores(
            str(PROJECT_ROOT / "config" / "proveedores.json")
        )
        key, config = lector_pdf.identificar_proveedor(
            "GOMEZ GROUP METERING, S.L.U. Nº FACTURA LF26022852 "
            "SERVICIO DE LECTURA, FACTURACIÓN Y MANTENIMIENTO DE CONTADORES "
            "TOTAL FACTURA 125,24 €",
            "gomez-sin-cabecera.pdf", providers,
        )

        self.assertIsNone(key)
        self.assertIsNone(config)

    def test_gomez_group_metering_profile_reads_legacy_total_before_label(self):
        providers = lector_pdf.cargar_proveedores(
            str(PROJECT_ROOT / "config" / "proveedores.json")
        )
        config = providers["proveedores"]["GOMEZ_GROUP_METERING"]

        result = lector_pdf.extraer_datos_factura(
            "FACTURA\nGOMEZ GROUP METERING, S.L.U.\nNº FACTURA\nRAZÓN SOCIAL\n"
            "LF25040215 01/07/2025 05/07/2025\nCONCEPTO CANTIDAD PRECIO IMPORTE\n"
            "100,00 €\nBASE IMPONIBLE ...\nI.V.A. 100,00 € (21%) 21,00 €\n"
            "Cargo IBAN: ES41\nFORMA DE PAGO:\n121,00 €\nTOTAL FACTURA ...\n",
            config,
        )

        self.assertEqual(121.0, result["importe_total"])

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
