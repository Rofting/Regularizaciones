"""Mejora 2: preparar escaneos torcidos, tenues o con ruido antes del OCR."""

import sys
import tempfile
import unittest
from pathlib import Path

import numpy as np
from PIL import Image, ImageDraw, ImageFont

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "core"))

import scan_preprocessing as scans

LINES = (
    "FACTURA ELECTRICIDAD N. F-2026-0042",
    "Fecha de factura: 31/01/2026",
    "Periodo de facturacion: 01/01/2026 a 31/01/2026",
    "CIF emisor: B12345674",
    "Termino fijo 123,45 EUR",
    "Importe total 1.234,56 EUR",
)


def clean_page() -> Image.Image:
    page = Image.new("L", (1650, 1000), 255)
    draw = ImageDraw.Draw(page)
    font = ImageFont.load_default(size=34)
    for index, line in enumerate(LINES):
        draw.text((120, 120 + index * 120), line, fill=0, font=font)
    return page


def degraded_page(angle=4.5, noise=0.10) -> Image.Image:
    faint = 175 + np.asarray(clean_page(), dtype=float) * (55 / 255)
    page = Image.fromarray(faint.astype(np.uint8)).rotate(
        angle, expand=True, fillcolor=240, resample=Image.Resampling.BICUBIC)
    values = np.asarray(page).copy()
    rng = np.random.default_rng(1)
    dots = rng.random(values.shape) < noise
    values[dots] = rng.choice([40, 255], dots.sum())
    return Image.fromarray(values)


class AssessTest(unittest.TestCase):
    def test_clean_page_needs_nothing(self):
        prepared = scans.prepare(clean_page())
        self.assertEqual((), prepared.quality.problems)
        self.assertFalse(prepared.changed)

    def test_problems_are_measured(self):
        quality = scans.assess(degraded_page())
        self.assertEqual(("torcida", "ruido", "poco contraste"), quality.problems)
        self.assertAlmostEqual(-4.5, quality.skew_degrees, delta=0.5)

    def test_straightening_undoes_the_rotation(self):
        for angle in (-3.0, 2.0):
            with self.subTest(angle=angle):
                rotated = clean_page().rotate(angle, expand=True, fillcolor=255)
                prepared = scans.prepare(rotated)
                self.assertTrue(any(step.startswith("enderezado") for step in prepared.steps))
                self.assertAlmostEqual(0.0, scans.estimate_skew(prepared.image), delta=0.5)

    def test_the_original_image_is_not_modified(self):
        page = degraded_page()
        before = np.asarray(page).copy()
        scans.prepare(page)
        self.assertTrue(np.array_equal(before, np.asarray(page)))


class ChoiceTest(unittest.TestCase):
    def test_clean_page_is_read_once(self):
        calls = []
        choice = scans.best_ocr(clean_page(), lambda image: calls.append(image) or "Fecha 31/01/2026")
        self.assertEqual(1, len(calls))
        self.assertIsNone(choice.prepared_score)

    def test_prepared_version_is_used_only_when_it_reads_more(self):
        page = degraded_page()
        better = scans.best_ocr(page, lambda image: "" if image is page else "Fecha 31/01/2026 Total 12,50")
        self.assertTrue(better.used_prepared)
        self.assertIn("31/01/2026", better.text)
        self.assertIn("Imagen preparada", better.summary)

        same = scans.best_ocr(page, lambda image: "Fecha 31/01/2026")
        self.assertFalse(same.used_prepared)
        self.assertIn("el original se leía igual o mejor", same.summary)

    def test_preparation_failure_falls_back_to_the_original(self):
        choice = scans.best_ocr(object(), lambda _image: "texto")
        self.assertEqual(("texto", ()), (choice.text, choice.steps))

    def test_score_values_dates_amounts_and_tax_ids(self):
        self.assertGreater(scans.ocr_score("31/01/2026 1.234,56 B12345674"), scans.ocr_score("FACTURA ELECTRICIDAD"))


class RealOcrTest(unittest.TestCase):
    """Con RapidOCR y Poppler reales, como en la aplicación."""

    def test_degraded_scanned_pdf_is_read_after_preparation(self):
        import lector_pdf
        try:
            lector_pdf._rapidocr_engine()
        except Exception as error:  # pragma: no cover - depende de la instalación
            self.skipTest(f"RapidOCR no disponible: {error}")
        with tempfile.TemporaryDirectory() as directory:
            pdf = Path(directory) / "escaneo.pdf"
            degraded_page().convert("RGB").save(pdf, "PDF", resolution=200)
            result = lector_pdf.extraer_texto_ocr_con_diagnostico(str(pdf))
            clean_pdf = Path(directory) / "limpio.pdf"
            clean_page().convert("RGB").save(clean_pdf, "PDF", resolution=200)
            clean = lector_pdf.extraer_texto_ocr_con_diagnostico(str(clean_pdf))
        self.assertEqual("rapidocr", result.status)
        self.assertIn("31/01/2026", result.text)
        self.assertIn("Importe total 1.234", result.text)
        self.assertTrue(result.preprocessing)
        self.assertIn("Imagen preparada", result.detail)
        self.assertEqual((), clean.preprocessing)
        self.assertIn("31/01/2026", clean.text)


if __name__ == "__main__":
    unittest.main()
