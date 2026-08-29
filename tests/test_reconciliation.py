import sys
import unittest
from contextlib import redirect_stdout
from io import StringIO
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
CORE_DIR = PROJECT_ROOT / "core"
if str(CORE_DIR) not in sys.path:
    sys.path.insert(0, str(CORE_DIR))

import gestor_bd
from reconciliation import ReconciliationError, assert_period_reconciled, reconcile_period
from regularization_service import calculate_period
from tests.helpers import temporary_database
from tests.test_regularization_service import prepare_period


class ReconciliationTest(unittest.TestCase):
    def test_accepts_one_cent_and_blocks_two_cent_difference(self):
        with temporary_database() as (connection, path):
            with redirect_stdout(StringIO()):
                gestor_bd.crear_bd(str(path))
            period_id, batch_id, _ = prepare_period(connection)
            calculate_period(connection, period_id, batch_id)
            exact = reconcile_period(connection, period_id)
            self.assertEqual("cuadrado", exact.status)
            assert_period_reconciled(connection, period_id)

            result_id = connection.execute(
                "SELECT id_result FROM owner_concept_results WHERE concept_key='acs_fixed' LIMIT 1"
            ).fetchone()[0]
            connection.execute(
                "UPDATE owner_concept_results SET actual_cents=actual_cents+1 WHERE id_result=?",
                (result_id,),
            )
            one_cent = reconcile_period(connection, period_id)
            self.assertEqual("cuadrado", one_cent.status)

            connection.execute(
                "UPDATE owner_concept_results SET actual_cents=actual_cents+1 WHERE id_result=?",
                (result_id,),
            )
            two_cents = reconcile_period(connection, period_id)
            self.assertEqual("descuadrado", two_cents.status)
            with self.assertRaises(ReconciliationError):
                assert_period_reconciled(connection, period_id)


if __name__ == "__main__":
    unittest.main()
