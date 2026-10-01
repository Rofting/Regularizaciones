import sys
import unittest
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
CORE_DIR = PROJECT_ROOT / "core"
if str(CORE_DIR) not in sys.path:
    sys.path.insert(0, str(CORE_DIR))

from dwelling_matching import canonical_dwelling, match_dwelling


class DwellingMatchingTest(unittest.TestCase):
    def test_spanish_floor_notations_share_a_canonical_form(self):
        for variants in (("1º B", "1B", "PISO 1 PUERTA B", "01-B"),
                         ("BAJO IZQ.", "0 IZDA", "Bajo izquierda", "BJ IZ"),
                         ("ÁTICO DCHA", "At. Dcha", "ATICO DERECHA"),
                         ("BL1-1ºB", "BL1 1 B", "Bloque 1 1B")):
            with self.subTest(variants=variants):
                self.assertEqual(1, len({canonical_dwelling(item) for item in variants}))

    def test_only_unambiguous_matches_are_returned(self):
        owners = {"BL1 1B": "a", "BL1 1A": "b", "BL2 1B": "c", "BL1 BAJO IZDA": "d"}
        self.assertEqual("a", match_dwelling("BL1 1º B", owners))
        self.assertEqual("d", match_dwelling("Bajo izquierda", owners))
        self.assertEqual("b", match_dwelling("1A", owners))  # único que termina en 1A
        self.assertIsNone(match_dwelling("1º B", owners))   # BL1 y BL2: ambiguo
        self.assertIsNone(match_dwelling("1C", owners))
        self.assertIsNone(match_dwelling("", owners))

    def test_neighbouring_doors_are_never_confused(self):
        self.assertIsNone(match_dwelling("1C", {"1A": 1, "1B": 2}))
        self.assertEqual(2, match_dwelling("1 b", {"1A": 1, "1B": 2}))


if __name__ == "__main__":
    unittest.main()
