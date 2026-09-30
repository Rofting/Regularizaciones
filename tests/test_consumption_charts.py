import sys
import tempfile
import unittest
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
CORE_DIR = PROJECT_ROOT / "core"
if str(CORE_DIR) not in sys.path:
    sys.path.insert(0, str(CORE_DIR))

from consumption_charts import (
    FIVE_BAND_LABELS,
    consumption_band,
    adaptive_step,
    consumption_bands,
    render_consumption_charts,
)


class ConsumptionChartsTest(unittest.TestCase):
    def test_assigns_ten_unit_neighbor_bands(self):
        self.assertEqual(((20.0, 30.0), 2), consumption_band(27.4))
        self.assertEqual(((0.0, 10.0), 0), consumption_band(0))
        self.assertEqual(
            ((0.0, 10.0), (10.0, 20.0), (20.0, 30.0),
             (30.0, 40.0), (40.0, float("inf"))),
            consumption_bands([2, 11, 27]),
        )
        self.assertEqual(("Muy bajo", "Bajo", "Medio", "Alto", "Muy alto"), FIVE_BAND_LABELS)

    def test_renders_unavailable_comparison_when_there_are_no_valid_neighbors(self):
        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory) / "sin-vecinos.png"
            result = render_consumption_charts(
                target, owner_consumption=12, neighbor_consumptions=[], history=[]
            )
            self.assertEqual(target, result)
            self.assertGreater(target.stat().st_size, 1000)

    def test_renders_two_panel_png(self):
        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory) / "consumo.png"
            result = render_consumption_charts(
                target, owner_consumption=27, neighbor_consumptions=[2, 6, 11, 27, 34],
                history=[("2024–25", 21), ("2025–26", 27)],
            )
            self.assertEqual(target, result)
            self.assertGreater(target.stat().st_size, 1000)


    def test_adaptive_step_spreads_high_consumptions_over_the_five_bands(self):
        values = [15 + index * 3.1 for index in range(40)]  # 15–136 m³
        step = adaptive_step(values)
        bands = consumption_bands(values, step=step)
        counts = [sum(low <= value < high for value in values) for low, high in bands]
        self.assertTrue(all(count > 0 for count in counts), counts)
        self.assertLess(max(counts), len(values) * 0.5)
        self.assertEqual(10.0, adaptive_step([]))
        self.assertEqual(1500.0, adaptive_step([800, 5800, 6000]))

    def test_renders_series_title_current_period_and_missing_owner(self):
        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory) / "calefaccion.png"
            result = render_consumption_charts(
                target, owner_consumption=None, neighbor_consumptions=[800, 2400, 5100],
                history=[("2024-2025", 2900)], unit="kWh", title="Calefacción",
                current_label="2025-2026",
            )
            self.assertEqual(target, result)
            self.assertGreater(target.stat().st_size, 1000)

if __name__ == "__main__":
    unittest.main()
