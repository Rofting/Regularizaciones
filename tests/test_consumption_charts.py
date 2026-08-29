import sys
import tempfile
import unittest
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
CORE_DIR = PROJECT_ROOT / "core"
if str(CORE_DIR) not in sys.path:
    sys.path.insert(0, str(CORE_DIR))

from consumption_charts import consumption_band, consumption_bands, render_consumption_charts


class ConsumptionChartsTest(unittest.TestCase):
    def test_assigns_ten_unit_neighbor_bands(self):
        self.assertEqual(((20.0, 30.0), 2), consumption_band(27.4))
        self.assertEqual(((0.0, 10.0), 0), consumption_band(0))
        self.assertEqual(((0.0, 10.0), (10.0, 20.0), (20.0, 30.0)), consumption_bands([2, 11, 27]))

    def test_renders_two_panel_png(self):
        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory) / "consumo.png"
            result = render_consumption_charts(
                target, owner_consumption=27, neighbor_consumptions=[2, 6, 11, 27, 34],
                history=[("2024–25", 21), ("2025–26", 27)],
            )
            self.assertEqual(target, result)
            self.assertGreater(target.stat().st_size, 1000)


if __name__ == "__main__":
    unittest.main()
