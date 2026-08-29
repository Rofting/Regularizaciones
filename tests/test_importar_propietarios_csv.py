import sys
import tempfile
import unittest
from contextlib import closing, redirect_stdout
from io import StringIO
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
CORE_DIR = PROJECT_ROOT / "core"
if str(CORE_DIR) not in sys.path:
    sys.path.insert(0, str(CORE_DIR))

import gestor_bd
from importar_propietarios_csv import import_owner_csv
from tests.helpers import make_owner_csv


class OwnerCsvImportTest(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory(ignore_cleanup_errors=True)
        self.database = Path(self.directory.name) / "gestion.db"
        self.csv_path = Path(self.directory.name) / "owners.csv"
        with redirect_stdout(StringIO()):
            gestor_bd.crear_bd(str(self.database))
        self.connection = gestor_bd.conectar(str(self.database))
        self.community_id = gestor_bd.obtener_o_crear_comunidad(
            self.connection, "TEST", "COMUNIDAD DE PRUEBA"
        )

    def tearDown(self):
        self.connection.close()
        self.directory.cleanup()

    def test_imports_and_updates_owner_by_property_code(self):
        make_owner_csv(
            self.csv_path,
            [
                {
                    "Codigo": "101",
                    "Nombre": "VECINO 01",
                    "Fdenominacion": " PA2-1ºA ",
                    "Coeficiente": "1,46",
                    "Email": "vecino01@example.test",
                }
            ],
        )
        first = import_owner_csv(
            self.connection, self.community_id, self.csv_path, expected_count=1
        )

        make_owner_csv(
            self.csv_path,
            [
                {
                    "Codigo": "101",
                    "Nombre": "VECINO ACTUALIZADO",
                    "Fdenominacion": "PA2-1ºA",
                    "Coeficiente": "1,500",
                    "Email": "malo; valido@example.test",
                }
            ],
        )
        second = import_owner_csv(
            self.connection, self.community_id, self.csv_path, expected_count=1
        )
        row = self.connection.execute(
            "SELECT * FROM propietarios WHERE id_comunidad=?",
            (self.community_id,),
        ).fetchone()

        self.assertEqual((1, 0, 0), (first.inserted, first.updated, first.skipped))
        self.assertEqual((0, 1, 0), (second.inserted, second.updated, second.skipped))
        self.assertEqual("PA2-1ºA", row["codigo_vivienda"])
        self.assertEqual("VECINO ACTUALIZADO", row["nombre_propietario"])
        self.assertAlmostEqual(1.5, row["coeficiente"])
        self.assertEqual("valido@example.test", row["email"])

    def test_expected_count_mismatch_rolls_back(self):
        make_owner_csv(
            self.csv_path,
            [
                {
                    "Codigo": "101",
                    "Nombre": "VECINO 01",
                    "Fdenominacion": "PA2-1ºA",
                    "Coeficiente": "1,46",
                    "Email": "vecino01@example.test",
                }
            ],
        )
        with self.assertRaisesRegex(ValueError, "esperaban 2"):
            import_owner_csv(
                self.connection, self.community_id, self.csv_path, expected_count=2
            )
        count = self.connection.execute("SELECT COUNT(*) FROM propietarios").fetchone()[0]
        self.assertEqual(0, count)


if __name__ == "__main__":
    unittest.main()
