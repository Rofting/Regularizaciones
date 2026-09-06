import sys
import unittest
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
CORE_DIR = PROJECT_ROOT / "core"
if str(CORE_DIR) not in sys.path:
    sys.path.insert(0, str(CORE_DIR))

import expedient_ui


class CommunityOnboardingRouteTest(unittest.TestCase):
    def test_onboarding_step_route_requires_identity_first(self):
        state = {"step": "summary", "identity_valid": False, "has_sources": True}

        self.assertEqual("identity", expedient_ui.onboarding_step_route(state))

    def test_onboarding_step_route_requires_sources_before_detection(self):
        state = {"step": "identity", "identity_valid": True, "has_sources": False}

        self.assertEqual("sources", expedient_ui.onboarding_step_route(state))

    def test_onboarding_step_route_requires_answers_before_summary(self):
        state = {
            "step": "detected",
            "identity_valid": True,
            "has_sources": True,
            "has_required_questions": True,
            "answers_complete": False,
        }

        self.assertEqual("confirmations", expedient_ui.onboarding_step_route(state))

    def test_onboarding_step_route_blocks_incomplete_summary(self):
        state = {
            "step": "summary",
            "identity_valid": True,
            "has_sources": True,
            "has_required_questions": True,
            "answers_complete": False,
        }

        self.assertEqual("confirmations", expedient_ui.onboarding_step_route(state))

    def test_onboarding_step_route_allows_complete_summary(self):
        state = {
            "step": "confirmations",
            "identity_valid": True,
            "has_sources": True,
            "has_required_questions": True,
            "answers_complete": True,
        }

        self.assertEqual("summary", expedient_ui.onboarding_step_route(state))


if __name__ == "__main__":
    unittest.main()
