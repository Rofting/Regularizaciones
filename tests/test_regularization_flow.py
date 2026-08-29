import sys
import tempfile
import unittest
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
CORE_DIR = PROJECT_ROOT / "core"
if str(CORE_DIR) not in sys.path:
    sys.path.insert(0, str(CORE_DIR))

import gestor_bd
from regularization_flow import run_regularization


class RegularizationFlowTest(unittest.TestCase):
    def test_rejects_missing_source_files(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            db = Path(td) / "flow.db"
            gestor_bd.crear_bd(str(db))
            con = gestor_bd.conectar(str(db))
            community = gestor_bd.obtener_o_crear_comunidad(con, "TEST", "Comunidad")
            with self.assertRaises(FileNotFoundError):
                run_regularization(con, community, "missing.xlsx", "owners.csv", "readings.xls")
            con.close()


if __name__ == "__main__":
    unittest.main()
