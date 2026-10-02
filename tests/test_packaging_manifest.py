import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "core"))

from app_paths import PUBLIC_RESOURCES


class PackagingManifestTest(unittest.TestCase):
    def test_only_public_resources_are_staged_for_the_executable(self):
        self.assertTrue((ROOT / "Regularizaciones.spec").is_file())
        self.assertTrue((ROOT / "scripts" / "build_windows.ps1").is_file())
        self.assertIn("plantillas/Plantilla_Cartas.docx", PUBLIC_RESOURCES)
        self.assertIn("config/proveedores.json", PUBLIC_RESOURCES)
        self.assertFalse(any(path.startswith(("data/", "salidas/", "fuentes/")) for path in PUBLIC_RESOURCES))
        for forbidden in ("gestion.db", "ui_prefs.json", "proveedores_despacho.json",
                          "letter_identities.json"):
            self.assertFalse(any(path.endswith(forbidden) for path in PUBLIC_RESOURCES))


if __name__ == "__main__":
    unittest.main()
