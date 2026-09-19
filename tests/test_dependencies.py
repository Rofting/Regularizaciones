import importlib
import os
import sys
import tempfile
import unittest
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
CORE_DIR = PROJECT_ROOT / "core"
if str(CORE_DIR) not in sys.path:
    sys.path.insert(0, str(CORE_DIR))


class DependencySmokeTest(unittest.TestCase):
    def test_runtime_dependencies_and_current_modules_import(self):
        with tempfile.TemporaryDirectory() as cache_dir:
            os.environ["MPLCONFIGDIR"] = cache_dir
            for module_name in (
                "customtkinter",
                "openpyxl",
                "docx",
                "pdfplumber",
                "pdf2image",
                "pytesseract",
                "rapidocr",
                "onnxruntime",
                "matplotlib",
                "numpy",
                "xlrd",
                "gestor_bd",
                "motor_reparto",
                "carta_writer",
            ):
                with self.subTest(module=module_name):
                    self.assertIsNotNone(importlib.import_module(module_name))


if __name__ == "__main__":
    unittest.main()
