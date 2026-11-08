import sys
import unittest
from datetime import date
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
CORE_DIR = PROJECT_ROOT / "core"
if str(CORE_DIR) not in sys.path:
    sys.path.insert(0, str(CORE_DIR))

from source_eligibility import evaluate_invoice_eligibility


class SourceEligibilityTest(unittest.TestCase):
    def test_recognised_invoice_outside_active_modules_is_archived_not_applied(self):
        decision = evaluate_invoice_eligibility(
            community_code="658",
            case_start=date(2025, 9, 1),
            case_end=date(2026, 8, 31),
            active_modules=("GAS", "ELECTRICIDAD", "AGUA", "ACS"),
            document_kind="invoice",
            service_family="ASCENSORES",
            period_start=date(2026, 4, 1),
            period_end=date(2026, 4, 30),
            community_confidence="high",
        )
        self.assertEqual(
            ("not_applicable", "service_not_active"),
            (decision.status, decision.reason),
        )

    def test_invoice_for_another_period_requires_one_grouped_decision(self):
        decision = evaluate_invoice_eligibility(
            community_code="658",
            case_start=date(2025, 9, 1),
            case_end=date(2026, 8, 31),
            active_modules=("GAS",),
            document_kind="invoice",
            service_family="GAS",
            period_start=date(2024, 1, 1),
            period_end=date(2024, 1, 31),
            community_confidence="high",
        )
        self.assertEqual(
            ("review_required", "period_outside_case"),
            (decision.status, decision.reason),
        )

    def test_matching_invoice_is_eligible_for_its_concepts(self):
        decision = evaluate_invoice_eligibility(
            community_code="644",
            case_start=date(2025, 1, 1),
            case_end=date(2025, 12, 31),
            active_modules=("ACS", "CALEFACCION"),
            document_kind="invoice",
            service_family="CONTADORES",
            period_start=date(2025, 3, 1),
            period_end=date(2025, 3, 31),
            community_confidence="high",
        )
        self.assertEqual("eligible", decision.status)
        self.assertEqual(("ACS", "CALEFACCION"), decision.concept_keys)

    def _decide(self, issue_date, start=date(2025, 12, 1), end=date(2025, 12, 31)):
        return evaluate_invoice_eligibility(
            community_code="658",
            case_start=date(2026, 1, 1),
            case_end=date(2026, 12, 31),
            active_modules=("GAS",),
            document_kind="invoice",
            service_family="GAS",
            period_start=start,
            period_end=end,
            community_confidence="high",
            issue_date=issue_date,
        )

    def test_invoice_issued_in_the_exercise_counts_even_if_consumed_before(self):
        decision = self._decide(date(2026, 1, 8))
        self.assertEqual("eligible", decision.status)

    def test_invoice_issued_after_the_exercise_is_reviewed_even_if_consumed_inside(self):
        decision = self._decide(date(2027, 1, 8), start=date(2026, 12, 1), end=date(2026, 12, 31))
        self.assertEqual(("review_required", "period_outside_case"), (decision.status, decision.reason))

    def test_without_issue_date_the_billed_period_decides(self):
        self.assertEqual("review_required", self._decide(None).status)

    def test_uncertain_community_is_reviewed_before_fields(self):
        decision = evaluate_invoice_eligibility(
            community_code=None,
            case_start=date(2025, 1, 1),
            case_end=date(2025, 12, 31),
            active_modules=("GAS",),
            document_kind="invoice",
            service_family="GAS",
            period_start=date(2025, 3, 1),
            period_end=date(2025, 3, 31),
            community_confidence="low",
        )
        self.assertEqual(("review_required", "community_unknown"), (decision.status, decision.reason))


if __name__ == "__main__":
    unittest.main()
