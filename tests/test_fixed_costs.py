import sys
import tempfile
import unittest
from contextlib import redirect_stdout
from decimal import Decimal
from io import StringIO
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
CORE_DIR = PROJECT_ROOT / "core"
if str(CORE_DIR) not in sys.path:
    sys.path.insert(0, str(CORE_DIR))

from openpyxl import Workbook, load_workbook

import fixed_costs
import gestor_bd


class FixedCostsTest(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory(ignore_cleanup_errors=True)
        with redirect_stdout(StringIO()):
            gestor_bd.crear_bd(str(Path(self.directory.name) / "g.db"))
        self.connection = gestor_bd.conectar(str(Path(self.directory.name) / "g.db"))
        self.community = gestor_bd.obtener_o_crear_comunidad(self.connection, "1", "C")
        self.period = self.connection.execute(
            "INSERT INTO periodos(id_comunidad,nombre,fecha_inicio,fecha_fin) VALUES (?,?,?,?)",
            (self.community, "2026", "2026-01-01", "2026-12-31"),
        ).lastrowid

    def tearDown(self):
        self.connection.close()
        self.directory.cleanup()

    def test_save_load_and_clear(self):
        fixed_costs.save_fixed_costs(self.connection, self.community, self.period, {
            "fixed_cost_boiler_maintenance": "1.095,50", "fixed_cost_meter_reading_acs": "82",
        })
        self.assertEqual(
            {"fixed_cost_boiler_maintenance": Decimal("1095.5"), "fixed_cost_meter_reading_acs": Decimal("82")},
            fixed_costs.load_fixed_costs(self.connection, self.community, self.period),
        )
        fixed_costs.save_fixed_costs(self.connection, self.community, self.period,
                                     {"fixed_cost_meter_reading_acs": ""})
        self.assertNotIn("fixed_cost_meter_reading_acs",
                         fixed_costs.load_fixed_costs(self.connection, self.community, self.period))
        with self.assertRaises(ValueError):
            fixed_costs.save_fixed_costs(self.connection, self.community, self.period,
                                         {"fixed_cost_boiler_maintenance": "-3"})

    def test_values_are_written_only_on_matching_rows(self):
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "OTROS GASTOS"
        sheet["A9"] = "MANTENIMIENTO SALA CALDERAS"
        sheet["A6"] = "Otra cosa distinta"
        written = fixed_costs.write_fixed_costs(
            workbook,
            {"fixed_cost_boiler_maintenance": Decimal("95.5"), "fixed_cost_meter_reading_acs": Decimal("80")},
            lambda cell, value: setattr(cell, "value", value),
        )
        self.assertEqual(["B9"], written)
        self.assertEqual(95.5, sheet["B9"].value)
        self.assertIsNone(sheet["B6"].value)

    def test_common_model_carries_no_community_costs(self):
        sheet = load_workbook(PROJECT_ROOT / "plantillas" / "modelo" / "modelo_acs_v1.xlsx")["OTROS GASTOS"]
        for item in fixed_costs.FIXED_COSTS:
            with self.subTest(cell=item.cell):
                self.assertIn(sheet[item.cell].value, (None, ""))
                self.assertTrue(str(sheet[f"A{sheet[item.cell].row}"].value).upper().startswith(item.sheet_label))
        self.assertEqual("=B6*12", sheet["C6"].value)


if __name__ == "__main__":
    unittest.main()
