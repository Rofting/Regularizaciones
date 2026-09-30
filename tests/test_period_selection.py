import sys
import unittest
from datetime import date
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
CORE_DIR = PROJECT_ROOT / "core"
if str(CORE_DIR) not in sys.path:
    sys.path.insert(0, str(CORE_DIR))

from period_selection import (
    ExistingCase,
    default_case_name,
    describe_range,
    parse_user_date,
    period_presets,
    validate_range,
)


class ParseUserDateTest(unittest.TestCase):
    def test_accepts_common_spanish_spellings(self):
        expected = date(2026, 8, 31)
        for text in ("31/08/2026", "31-8-26", "2026-08-31", "31.08.2026",
                     "31 agosto 2026", "31 de agosto de 2026", " 31/8/2026 "):
            with self.subTest(text=text):
                self.assertEqual(expected, parse_user_date(text))

    def test_day_and_month_use_the_reference_year(self):
        self.assertEqual(date(2025, 9, 1), parse_user_date("1/9", reference_year=2025))
        with self.assertRaises(ValueError):
            parse_user_date("1/9")

    def test_invalid_dates_raise_readable_errors(self):
        for text in ("", "31/02/2026", "mañana", "2026/13/01"):
            with self.subTest(text=text), self.assertRaises(ValueError):
                parse_user_date(text)


class PresetTest(unittest.TestCase):
    def test_continue_previous_case_with_the_same_length(self):
        previous = [ExistingCase("2024-2025", date(2024, 9, 1), date(2025, 8, 31))]
        first = period_presets(date(2026, 9, 30), previous)[0]
        self.assertEqual("continue", first.key)
        self.assertEqual((date(2025, 9, 1), date(2026, 8, 31)), (first.start, first.end))

    def test_continue_irregular_length(self):
        previous = [ExistingCase("Tramo", date(2026, 1, 10), date(2026, 1, 19))]
        first = period_presets(date(2026, 2, 1), previous)[0]
        self.assertEqual((date(2026, 1, 20), date(2026, 1, 29)), (first.start, first.end))

    def test_default_presets_cover_fiscal_year_calendar_year_and_month(self):
        presets = {item.key: item for item in period_presets(date(2026, 9, 30))}
        self.assertEqual((date(2025, 9, 1), date(2026, 8, 31)),
                         (presets["last_fiscal"].start, presets["last_fiscal"].end))
        self.assertEqual((date(2025, 1, 1), date(2025, 12, 31)),
                         (presets["last_year"].start, presets["last_year"].end))
        self.assertEqual((date(2026, 4, 1), date(2026, 6, 30)),
                         (presets["last_quarter"].start, presets["last_quarter"].end))
        self.assertEqual((date(2026, 8, 1), date(2026, 8, 31)),
                         (presets["last_month"].start, presets["last_month"].end))

    def test_presets_are_not_duplicated(self):
        previous = [ExistingCase("2024-2025", date(2024, 9, 1), date(2025, 8, 31))]
        presets = period_presets(date(2026, 9, 30), previous)
        self.assertEqual(len(presets), len({(item.start, item.end) for item in presets}))


class NamingAndValidationTest(unittest.TestCase):
    def test_default_names(self):
        self.assertEqual("2025-2026", default_case_name(date(2025, 9, 1), date(2026, 8, 31)))
        self.assertEqual("Año 2025", default_case_name(date(2025, 1, 1), date(2025, 12, 31)))
        self.assertEqual("Enero 2026", default_case_name(date(2026, 1, 1), date(2026, 1, 31)))
        self.assertEqual("Ene–Mar 2026", default_case_name(date(2026, 1, 1), date(2026, 3, 31)))
        self.assertEqual("05/01/2026 – 20/01/2026", default_case_name(date(2026, 1, 5), date(2026, 1, 20)))

    def test_describe_range_counts_months_and_days(self):
        self.assertIn("12 meses (365 días)", describe_range(date(2025, 9, 1), date(2026, 8, 31)))

    def test_validation_reports_inversion_and_overlap(self):
        errors, _ = validate_range(date(2026, 2, 1), date(2026, 1, 1))
        self.assertTrue(errors)
        existing = [ExistingCase("Anterior", date(2025, 9, 1), date(2026, 8, 31))]
        errors, warnings = validate_range(date(2026, 8, 1), date(2027, 7, 31), existing)
        self.assertEqual([], errors)
        self.assertTrue(any("Anterior" in item for item in warnings))
        _, contiguous = validate_range(date(2026, 9, 1), date(2027, 8, 31), existing)
        self.assertEqual([], contiguous)


if __name__ == "__main__":
    unittest.main()
