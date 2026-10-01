import sys
import unittest
from pathlib import Path
from unittest import mock

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "core"))

import office_recalculation as recalculation


class LibreOfficeVersionTest(unittest.TestCase):
    def test_parses_linux_and_windows_output(self):
        self.assertEqual((24, 2, 7, 2), recalculation.parse_libreoffice_version(
            "LibreOffice 24.2.7.2 420(Build:2)"))
        self.assertEqual((26, 2, 6, 3), recalculation.parse_libreoffice_version(
            "LibreOffice 26.2.6.3 a1b2c3d4 (X86_64)"))
        self.assertIsNone(recalculation.parse_libreoffice_version("soffice: error"))

    def _check(self, version):
        recalculator = recalculation.LibreOfficeRecalculator()
        with mock.patch.object(recalculator, "installed_version", return_value=version):
            return recalculator.check_minimum_version()

    def test_accepts_the_verified_range(self):
        self.assertEqual((24, 2, 7, 2), self._check((24, 2, 7, 2)))
        self.assertEqual((26, 2, 6, 3), self._check((26, 2, 6, 3)))
        self.assertEqual((24, 2), self._check((24, 2)))

    def test_rejects_older_versions_with_a_clear_message(self):
        with self.assertRaisesRegex(recalculation.RecalculationError, r"7\.6\.4\.1.*24\.2"):
            self._check((7, 6, 4, 1))
        with self.assertRaises(recalculation.RecalculationError):
            self._check((24, 1, 9))


if __name__ == "__main__":
    unittest.main()
