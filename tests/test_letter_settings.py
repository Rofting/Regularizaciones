import sqlite3
import tempfile
import sys
import unittest
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
CORE_DIR = PROJECT_ROOT / "core"
if str(CORE_DIR) not in sys.path:
    sys.path.insert(0, str(CORE_DIR))

import gestor_bd
from letter_settings import available_concepts, load_selected_concepts, save_selected_concepts


class LetterSettingsTest(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory(ignore_cleanup_errors=True)
        self.database = str(Path(self.directory.name) / "settings.db")
        gestor_bd.crear_bd(self.database)
        self.connection = gestor_bd.conectar(self.database)
        self.community_id = gestor_bd.obtener_o_crear_comunidad(
            self.connection, "TEST", "Comunidad de prueba"
        )

    def tearDown(self):
        self.connection.close()
        self.directory.cleanup()

    def test_defaults_to_concepts_with_results_and_persists_selection(self):
        owner = self.connection.execute(
            "INSERT INTO propietarios(id_comunidad,codigo_vivienda,nombre_propietario) VALUES (?,?,?)",
            (self.community_id, "A-1", "Vecino"),
        ).lastrowid
        period = self.connection.execute(
            "INSERT INTO periodos(id_comunidad,nombre,fecha_inicio,fecha_fin) VALUES (?,?,?,?)",
            (self.community_id, "2024-2025", "2024-08-01", "2025-07-31"),
        ).lastrowid
        self.connection.execute(
            "INSERT INTO owner_concept_results(id_propietario,id_periodo,concept_key,billed_cents,actual_cents,difference_cents) VALUES (?,?,?,?,?,?)",
            (owner, period, "acs_variable", 100, 90, -10),
        )
        self.connection.commit()
        self.assertEqual(("acs_variable",), load_selected_concepts(self.connection, self.community_id))
        self.assertEqual(("acs_fixed", "acs_variable"), save_selected_concepts(
            self.connection, self.community_id, ["acs_variable", "acs_fixed"]
        ))
        self.assertEqual(("acs_fixed", "acs_variable"), load_selected_concepts(self.connection, self.community_id))

    def test_rejects_unknown_or_empty_selection(self):
        with self.assertRaisesRegex(ValueError, "no disponibles"):
            save_selected_concepts(self.connection, self.community_id, ["inventado"])
        with self.assertRaisesRegex(ValueError, "al menos un"):
            save_selected_concepts(self.connection, self.community_id, [])


if __name__ == "__main__":
    unittest.main()
