"""Mejora 14: vista previa de la evidencia de una incidencia."""

import json
import sqlite3
import sys
import tempfile
import unittest
from contextlib import redirect_stdout
from io import StringIO
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
from openpyxl import Workbook

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "core"))

import gestor_bd
import source_preview as preview


def write_pdf(path, pages):
    with matplotlib.rc_context({"pdf.fonttype": 42}), PdfPages(path) as pdf:
        for lines in pages:
            figure = plt.figure(figsize=(8.27, 11.69))
            for index, line in enumerate(lines):
                figure.text(0.1, 0.9 - 0.05 * index, line, fontsize=12)
            pdf.savefig(figure)
            plt.close(figure)


def word(text, x, top):
    return {"text": text, "x0": x, "x1": x + 10 * len(text), "top": top, "bottom": top + 10}


class LocateTest(unittest.TestCase):
    def setUp(self):
        self.pages = [
            [word("Total", 0, 10), word("130", 60, 10), word("kWh", 100, 10)],
            [word("Importe", 0, 10), word("total", 80, 10), word("1.234,56", 140, 10), word("EUR", 240, 10),
             word("Consumo", 0, 30), word("130", 80, 30)],
        ]

    def test_value_inside_the_fragment_is_highlighted(self):
        match = preview.locate_in_pdf(self.pages, preview.Evidence(
            value="1234.56", fragment="Importe total 1.234,56 EUR"))
        self.assertEqual((2, "value", "1.234,56"), (match.page, match.matched, match.text))
        self.assertEqual(1, len(match.boxes))

    def test_value_found_elsewhere_is_not_taken_from_the_wrong_place(self):
        # «130» aparece en dos páginas: sin fragmento localizable no se elige ninguna.
        self.assertIsNone(preview.locate_in_pdf(self.pages, preview.Evidence(value="130")))
        match = preview.locate_in_pdf(self.pages, preview.Evidence(value="130", fragment="Consumo 130"))
        self.assertEqual((2, "value"), (match.page, match.matched))

    def test_fragment_without_value_highlights_the_fragment(self):
        match = preview.locate_in_pdf(self.pages, preview.Evidence(
            value="999", fragment="Importe total 1.234,56 EUR"))
        self.assertEqual((2, "fragment"), (match.page, match.matched))

    def test_text_that_is_not_in_the_pdf_gives_no_position(self):
        self.assertIsNone(preview.locate_in_pdf(self.pages, preview.Evidence(
            value="2026-01-31", fragment="Periodo de facturacion hasta 31/01/2026")))
        self.assertIsNone(preview.locate_in_pdf([[]], preview.Evidence(value="ES0000000000000000GA")))

    def test_distinctive_unique_value_is_found_without_fragment(self):
        match = preview.locate_in_pdf(self.pages, preview.Evidence(value="1234.56"))
        self.assertEqual((2, "value"), (match.page, match.matched))

    def test_printed_variants(self):
        self.assertIn("31/01/2026", preview.value_variants("2026-01-31"))
        self.assertIn("1.234,56", preview.value_variants("1234.56"))
        self.assertIn("50,00", preview.value_variants("50.0"))


class PreviewTest(unittest.TestCase):
    def setUp(self):
        directory = tempfile.TemporaryDirectory(ignore_cleanup_errors=True)
        self.addCleanup(directory.cleanup)
        self.root = Path(directory.name)

    def test_pdf_page_and_highlight(self):
        path = self.root / "factura.pdf"
        write_pdf(path, [["Portada"], ["Importe total 1.234,56 EUR", "CUPS ES0000000000000000GA"]])
        result = preview.build_preview(path, preview.Evidence(value="ES0000000000000000GA", fragment="CUPS ES0000000000000000GA"))
        self.assertEqual(("pdf", 2, 2), (result.kind, result.page_count, result.page))
        self.assertIn("página 2 de 2", result.headline)
        image = preview.render_pdf_page(path, result.page, result.match.boxes, resolution=40)
        self.assertGreater(image.size[0], 100)

    def test_unlocated_fragment_is_shown_as_text_on_the_declared_page(self):
        path = self.root / "escaneada.pdf"
        write_pdf(path, [["Uno"], ["Dos"]])
        result = preview.build_preview(path, preview.Evidence(fragment="Texto leído por OCR", page=2))
        self.assertIsNone(result.match)
        self.assertEqual(2, result.page)
        self.assertIn("No se ha localizado", result.headline)

    def test_sheet_window_marks_the_cell(self):
        path = self.root / "lecturas.xlsx"
        book = Workbook()
        for row in range(1, 41):
            book.active.append([f"V-{row}", row * 10])
        book.save(path)
        result = preview.build_preview(path, preview.Evidence(sheet="Sheet", cell="B30"))
        self.assertEqual("sheet", result.kind)
        self.assertEqual((30, 2), result.sheet.target)
        row = 30 - result.sheet.first_row
        self.assertEqual("300", result.sheet.rows[row][2 - result.sheet.first_column])
        self.assertIn("Celda B30", result.headline)


class EvidenceTest(unittest.TestCase):
    def test_collects_value_fragment_and_position(self):
        directory = tempfile.TemporaryDirectory(ignore_cleanup_errors=True)
        self.addCleanup(directory.cleanup)
        database = Path(directory.name) / "g.db"
        with redirect_stdout(StringIO()):
            gestor_bd.crear_bd(str(database))
        connection = sqlite3.connect(database)
        self.addCleanup(connection.close)
        connection.execute("PRAGMA foreign_keys=OFF")
        connection.execute(
            "INSERT INTO source_documents(id_document,id_case,original_name,archived_path,sha256,document_kind,status,source_context) "
            "VALUES (1,1,'a','a','x','invoice','registered',?)", (json.dumps({"page": 2, "fragment": "general"}),))
        connection.execute(
            "INSERT INTO source_field_evidence(id_document,field_name,value,confidence,source,locator_json,rule_id,extractor_version) "
            "VALUES (1,'importe_total','1234.56','high','text',?,'r','v')", (json.dumps({"fragment": "Importe total 1.234,56"}),))
        evidence = preview.evidence_for_issue(connection, 1, "importe_total")
        self.assertEqual(("1234.56", "Importe total 1.234,56", 2), (evidence.value, evidence.fragment, evidence.page))
        self.assertEqual("A7", preview.evidence_for_issue(connection, 1, "owner_list.row_7").cell)


if __name__ == "__main__":
    unittest.main()
