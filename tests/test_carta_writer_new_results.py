import sys
import tempfile
import unittest
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
CORE_DIR = PROJECT_ROOT / "core"
if str(CORE_DIR) not in sys.path:
    sys.path.insert(0, str(CORE_DIR))

import gestor_bd
from carta_writer import obtener_todos_los_vecinos_con_repartos


class CartaWriterNewResultsTest(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory(ignore_cleanup_errors=True)
        path = str(Path(self.directory.name) / "carta.db")
        gestor_bd.crear_bd(path)
        self.connection = gestor_bd.conectar(path)
        self.community = gestor_bd.obtener_o_crear_comunidad(self.connection, "TEST", "Comunidad")
        self.owner = self.connection.execute(
            "INSERT INTO propietarios(id_comunidad,codigo_vivienda,nombre_propietario) VALUES (?,?,?)",
            (self.community, "A-1", "Vecino"),
        ).lastrowid
        self.period = self.connection.execute(
            "INSERT INTO periodos(id_comunidad,nombre,fecha_inicio,fecha_fin) VALUES (?,?,?,?)",
            (self.community, "2024-2025", "2024-08-01", "2025-07-31"),
        ).lastrowid
        self.connection.executemany(
            "INSERT INTO lecturas_vecino(id_propietario,id_periodo,tipo,fecha_lectura,valor_acumulado) VALUES (?,?,?,?,?)",
            [(self.owner, self.period, "ACS", "2024-08-01", 100), (self.owner, self.period, "ACS", "2025-07-31", 112)],
        )
        self.connection.executemany(
            "INSERT INTO owner_concept_results(id_propietario,id_periodo,concept_key,consumption,billed_cents,actual_cents,difference_cents) VALUES (?,?,?,?,?,?,?)",
            [(self.owner, self.period, "acs_fixed", None, 1000, 900, -100), (self.owner, self.period, "acs_variable", 12, 2000, 2500, 500)],
        )
        self.connection.commit()

    def tearDown(self):
        self.connection.close()
        self.directory.cleanup()

    def test_reads_phase_one_results_when_legacy_repartos_are_empty(self):
        rows = obtener_todos_los_vecinos_con_repartos(self.connection, self.community, self.period, ("acs_variable",))
        self.assertEqual(1, len(rows))
        self.assertEqual(12.0, rows[0]["acs"]["consumo_real"])
        self.assertEqual(20.0, rows[0]["acs"]["importe_cobrado"])
        self.assertEqual(25.0, rows[0]["acs"]["importe_real"])


if __name__ == "__main__":
    unittest.main()
