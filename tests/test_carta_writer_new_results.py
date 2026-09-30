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



class CartaChartsTest(unittest.TestCase):
    def test_each_consumption_series_gets_its_own_chart(self):
        from docx import Document
        from carta_writer import generar_carta

        series = [
            {"title": "Agua caliente sanitaria", "owner_consumption": 72.4,
             "neighbor_consumptions": [20, 50, 72.4, 110], "history": [("2024-2025", 65)],
             "current_label": "2025-2026", "unit": "m³"},
            {"title": "Calefacción", "owner_consumption": 3120,
             "neighbor_consumptions": [900, 3120, 5000], "history": [],
             "current_label": "2025-2026", "unit": "kWh"},
        ]
        data = {
            "vecino": {"nombre": "Ana", "vivienda": "1A"},
            "periodo": {"nombre": "2025-2026", "fecha_inicio": "2025-09-01", "fecha_fin": "2026-08-31"},
            "total": {"cobrado": 0, "real": 0, "diferencia": 0},
            "conceptos": [{"label": "ACS", "importe_cobrado": 10, "importe_real": 12, "diferencia": 2}],
            "consumo_grafica": {"series": series},
        }
        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory) / "carta.docx"
            generar_carta(data, str(PROJECT_ROOT / "plantillas" / "Plantilla_Cartas.docx"), str(output))
            document = Document(str(output))
            self.assertEqual(2, len(document.inline_shapes))
            self.assertIn("COMPARATIVA DE CONSUMO", "\n".join(p.text for p in document.paragraphs))

if __name__ == "__main__":
    unittest.main()
