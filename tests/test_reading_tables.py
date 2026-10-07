import json
import sqlite3
import sys
import tempfile
import unittest
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
CORE_DIR = PROJECT_ROOT / "core"
if str(CORE_DIR) not in sys.path:
    sys.path.insert(0, str(CORE_DIR))

import keywords
import source_analysis
from document_text_service import TextExtraction
from reading_tables import detect_company, parse_reading_document, parse_reading_rows, rows_from_text, tables_from_pdf

ISTA_TEXT = """ista Metering Services España
Informe de lecturas de agua caliente sanitaria
Periodo de lectura: del 01/09/2025 al 31/08/2026
Vivienda      Nº Contador     Lectura anterior    Lectura actual    Consumo
BL1 BAJO A    12345678        120,5               150,0             29,5
BL1 1º B      12345679        98                  110               12
TOTAL                                                               41,5
"""


class ReadingTablesTest(unittest.TestCase):
    def test_text_report_from_a_reading_company(self):
        table = parse_reading_document(text=ISTA_TEXT)
        self.assertEqual(("ACS",), table.services)
        self.assertEqual(("2025-09-01", "2026-08-31"), (table.start, table.end))
        self.assertEqual(("ISTA", "high"), (table.company, table.confidence))
        self.assertEqual(["BL1 BAJO A", "BL1 1º B"], [row["vivienda"] for row in table.rows])
        self.assertEqual((120.5, 150.0), (table.rows[0]["val_ant"], table.rows[0]["val_act"]))

    def test_combined_report_with_dated_columns_per_service(self):
        rows = [["Comunidad X"],
                ["Piso", "ACS 01/09/2025", "ACS 31/08/2026", "Calef. 01/09/2025", "Calef. 31/08/2026"],
                ["1A", "10", "20", "100", "300"]]
        table = parse_reading_rows(rows)
        self.assertEqual(("ACS", "CALEFACCION"), table.services)
        self.assertEqual([("ACS", 10.0, 20.0), ("CALEFACCION", 100.0, 300.0)],
                         [(row["tipo"], row["val_ant"], row["val_act"]) for row in table.rows])
        self.assertEqual("2026-08-31", table.rows[0]["fecha_act"])

    def test_split_header_with_ocr_errors_and_spanish_numbers(self):
        rows = [["Vivienda", "LECTURA", "LECTURA"], ["", "ANTERI0R", "ACTUAL"], ["2A", "1.234,5", "1.250,0"]]
        table = parse_reading_rows(rows, "calefaccion")
        self.assertEqual([(1234.5, 1250.0)], [(row["val_ant"], row["val_act"]) for row in table.rows])

    def test_month_headers_give_the_period_but_not_exact_row_dates(self):
        table = parse_reading_rows([["Propiedad", "Titular", "sep-25", "ago-26"], ["3C", "Ana", "7", "19"]], "ACS")
        self.assertEqual(("2025-09-01", "2026-08-31"), (table.start, table.end))
        self.assertNotIn("fecha_ant", table.rows[0])
        self.assertEqual("Ana", table.rows[0]["nombre"])

    def test_repeated_headers_and_summary_rows_are_skipped(self):
        rows = [["Vivienda", "Lectura anterior", "Lectura actual"], ["1A", "1", "2"],
                ["Vivienda", "Lectura anterior", "Lectura actual"], ["1B", "3", "5"],
                ["Totales", "4", "7"]]
        self.assertEqual(["1A", "1B"], [row["vivienda"] for row in parse_reading_rows(rows, "ACS").rows])

    def test_summary_consumption_is_checked_without_becoming_a_dwelling(self):
        rows = [["Vivienda", "Lectura anterior", "Lectura actual", "Consumo"],
                ["1A", "10", "15", "5"], ["1B", "20", "23", "3"],
                ["TOTAL", "", "", "9"]]
        table = parse_reading_rows(rows, "ACS")
        self.assertEqual(["1A", "1B"], [row["vivienda"] for row in table.rows])
        self.assertEqual("medium", table.confidence)
        self.assertIn("total_consumo_no_cuadra", table.diagnostics)
        from_text = "Vivienda  Lectura anterior  Lectura actual  Consumo\n1A  10  15  5\n1B  20  23  3"
        combined = parse_reading_document(text=from_text, table_rows=rows)
        self.assertIn("total_consumo_no_cuadra", combined.diagnostics)
        rows[-1][-1] = "8"
        self.assertEqual((), parse_reading_rows(rows, "ACS").diagnostics)

    def test_lined_and_unlined_pdf_tables_have_the_same_readings(self):
        import matplotlib.pyplot as plt
        from matplotlib.backends.backend_pdf import PdfPages

        data = [["Vivienda", "Lectura anterior", "Lectura actual", "Consumo"],
                ["1A", "10", "15", "5"], ["1B", "20", "23", "3"],
                ["TOTAL", "", "", "8"]]
        with tempfile.TemporaryDirectory() as directory:
            for lined in (True, False):
                path = Path(directory) / ("lineas.pdf" if lined else "sin_lineas.pdf")
                figure, axis = plt.subplots(figsize=(8, 4))
                axis.axis("off")
                grid = axis.table(cellText=data, loc="center", cellLoc="left",
                                  colWidths=[.2, .28, .25, .2])
                for cell in grid.get_celld().values():
                    cell.set_edgecolor("black" if lined else "white")
                    cell.set_linewidth(1 if lined else 0)
                with PdfPages(path) as pdf:
                    pdf.savefig(figure)
                plt.close(figure)
                table = parse_reading_document(
                    text="Informe de lecturas ACS, del 01/01/2026 al 31/01/2026",
                    table_rows=tables_from_pdf(path, text="Informe de lecturas ACS"),
                )
                self.assertEqual([("1A", 10.0, 15.0), ("1B", 20.0, 23.0)],
                                 [(row["vivienda"], row["val_ant"], row["val_act"])
                                  for row in table.rows])
                self.assertEqual("high", table.confidence)
                self.assertEqual(("2026-01-01", "2026-01-31"), (table.start, table.end))

    def test_consumption_that_does_not_match_lowers_confidence(self):
        rows = [["Vivienda", "Lectura anterior", "Lectura actual", "Consumo"], ["1A", "1", "2", "9"], ["1B", "1", "3", "7"]]
        table = parse_reading_rows(rows, "ACS")
        self.assertEqual("medium", table.confidence)

    def test_owner_lists_and_invoices_are_not_readings(self):
        self.assertIsNone(parse_reading_rows([["Propiedad", "Nombre", "Coeficiente"], ["1A", "Ana", "2,5"]]))
        self.assertIsNone(parse_reading_document(text="Factura F-1\\nBase imponible 100,00\\nTotal 121,00"))

    def test_rows_from_text_splits_trailing_numbers(self):
        self.assertEqual([["1º IZDA GARCIA", "12,5", "14"]], rows_from_text("1º IZDA GARCIA 12,5 14"))

    def test_company_detection_and_fuzzy_headers(self):
        self.assertEqual("TECHEM", detect_company("Informe Techem de repartidores"))
        self.assertIsNone(detect_company("ista metering y techem"))
        roles = keywords.section("cabeceras_lecturas")
        self.assertEqual("val_ant", keywords.best_role("LECTURA ANTERI0R", roles)[0])
        self.assertIsNone(keywords.best_role("Importe", roles))


class ReadingIntegrationTest(unittest.TestCase):
    def test_spreadsheet_with_company_headers_becomes_a_reading(self):
        from openpyxl import Workbook
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "lecturas.xlsx"
            book = Workbook()
            sheet = book.active
            sheet.title = "Agua caliente"
            for row in (["Informe de lecturas ACS"], ["Ubicación", "Nº serie", "L. anterior", "L. actual", "Consumo"],
                        ["1A", "A1", 10, 15, 5], ["1B", "A2", 20, 21, 1]):
                sheet.append(row)
            book.save(path)
            analysis = source_analysis.analyse_tabular(path)
        self.assertEqual(("reading", "ACS"), (analysis.kind, analysis.candidates["tipo"]))
        self.assertEqual(2, len(json.loads(analysis.candidates["vecinos"])))
        self.assertIn("fecha_inicio", analysis.required_fields)

    def test_pdf_report_from_unknown_company_is_read_through_the_pipeline(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "informe.pdf"
            path.write_bytes(b"pdf sintetico")
            connection = sqlite3.connect(":memory:")
            try:
                analysis = source_analysis.analyse_source(
                    path, community_code="1", connection=connection,
                    provider_registry={},
                    text_extractor=lambda *_: TextExtraction(ISTA_TEXT, "pdf_text", (1,), {}, 1, False),
                )
            finally:
                connection.close()
        self.assertEqual("reading", analysis.kind)
        self.assertEqual("ISTA", analysis.candidates["empresa_lecturas"])
        self.assertEqual("2026-08-31", analysis.candidates["fecha_fin"])
        self.assertEqual((), analysis.required_fields)


if __name__ == "__main__":
    unittest.main()
