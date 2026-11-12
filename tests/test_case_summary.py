"""El resumen de la pantalla principal sólo cuenta lo que ya hay en la base."""

import sys
import tempfile
import unittest
from contextlib import redirect_stdout
from io import StringIO
from pathlib import Path

CORE_DIR = Path(__file__).resolve().parents[1] / "core"
if str(CORE_DIR) not in sys.path:
    sys.path.insert(0, str(CORE_DIR))

import case_summary
import cuotas_servicio
import gestor_bd


class CaseSummaryTest(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory(ignore_cleanup_errors=True)
        path = str(Path(self.directory.name) / "g.db")
        with redirect_stdout(StringIO()):
            gestor_bd.crear_bd(path)
        self.connection = gestor_bd.conectar(path)
        self.community = gestor_bd.obtener_o_crear_comunidad(self.connection, "1", "C")
        self.period = self.connection.execute(
            "INSERT INTO periodos(id_comunidad,nombre,fecha_inicio,fecha_fin) VALUES (?,?,?,?)",
            (self.community, "2026", "2026-01-01", "2026-12-31"),
        ).lastrowid
        self.case = self.connection.execute(
            """INSERT INTO regularization_cases(id_comunidad,id_periodo,nombre,fecha_inicio,fecha_fin,estado)
               VALUES (?,?,?,?,?,'draft')""",
            (self.community, self.period, "2026", "2026-01-01", "2026-12-31"),
        ).lastrowid
        self.connection.commit()

    def tearDown(self):
        self.connection.close()
        self.directory.cleanup()

    def test_formats_amounts_like_the_office(self):
        self.assertEqual("12.345,67 €", case_summary.euros(1234567))
        self.assertEqual("-0,05 €", case_summary.euros(-5))
        self.assertEqual("1.235 €", case_summary.euros(123450, decimals=False))
        self.assertEqual("358", case_summary.quantity(358.0))
        self.assertEqual("1.234,5", case_summary.quantity(1234.5))

    def test_unknown_case_has_no_summary(self):
        self.assertIsNone(case_summary.summarise_case(self.connection, 999))

    def test_counts_invoices_by_supply_and_charged_fees(self):
        for kind, amount in (("GAS", 600.10), ("GAS", -100.10), ("AGUA", 250), ("MANTENIMIENTO", 123.35)):
            self.connection.execute(
                "INSERT INTO facturas(id_comunidad,id_periodo,tipo_suministro,importe_total) VALUES (?,?,?,?)",
                (self.community, self.period, kind, amount),
            )
        cuotas_servicio.registrar_cuota(
            self.connection, community_id=self.community, period_id=self.period,
            servicio="ACS", concepto="fija", fecha="2026-01-01", importe="400")
        cuotas_servicio.registrar_cuota(
            self.connection, community_id=self.community, period_id=self.period,
            servicio="ACS", concepto="variable", fecha="2026-03-01", importe="612,50")
        self.connection.commit()
        summary = case_summary.summarise_case(self.connection, self.case)
        self.assertEqual(4, summary.invoice_count)
        self.assertEqual(87335, summary.invoice_cents)
        self.assertEqual(["GAS", "AGUA", "MANTENIMIENTO"], [item.key for item in summary.supplies])
        self.assertEqual((40000, 61250, 2), (
            summary.charged_fixed_cents, summary.charged_variable_cents, summary.charged_entries))
        self.assertFalse(summary.has_distribution)
        self.assertIsNone(summary.consumption)
        self.assertEqual(0, summary.letters)


if __name__ == "__main__":
    unittest.main()
