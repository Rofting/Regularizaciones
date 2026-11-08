"""Mantenimiento, lecturas y reparaciones se deducen de las facturas."""

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

import excel_export_service
import gestor_bd
import other_expenses


class AutomaticOtherExpensesTest(unittest.TestCase):
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
        self.case = {
            "id_comunidad": self.community, "id_periodo": self.period,
            "fecha_inicio": "2026-01-01", "fecha_fin": "2026-12-31",
        }

    def tearDown(self):
        self.connection.close()
        self.directory.cleanup()

    def _invoice(self, number, day, amount, provider="Mantenimientos Zaragoza", notes=None):
        self.connection.execute(
            """INSERT INTO facturas(id_comunidad,id_periodo,tipo_suministro,proveedor,num_factura,
                                    fecha_factura,importe_total,archivo_origen,notas)
               VALUES (?,?,?,?,?,?,?,?,?)""",
            (self.community, self.period, "MANTENIMIENTO", provider, number, day, amount,
             f"/archivo/{number}.pdf", notes),
        )
        self.connection.commit()

    def test_maintenance_and_meter_reading_become_monthly_averages(self):
        self._invoice("M1", "2026-02-10", 600)
        self._invoice("M2", "2026-08-10", 600)
        self._invoice("L1", "2026-03-02", 72.6, provider="Gómez Group Metering — Servicio de contadores")
        values = other_expenses.automatic_fixed_costs(
            self.connection, self.community, "2026-01-01", "2026-12-31")
        self.assertEqual(Decimal(100), values[other_expenses.MAINTENANCE_KEY])
        self.assertEqual(Decimal("72.60") / 12, values[other_expenses.METER_READING_KEY])
        # El modelo multiplica por doce: el anual es exactamente la suma de facturas.
        self.assertEqual(Decimal("72.60"), values[other_expenses.METER_READING_KEY] * 12)

    def test_only_invoices_issued_in_the_exercise_count(self):
        self._invoice("OLD", "2025-12-31", 500)
        self._invoice("IN", "2026-01-01", 120)
        self._invoice("NEXT", "2027-01-01", 900)
        values = other_expenses.automatic_fixed_costs(
            self.connection, self.community, "2026-01-01", "2026-12-31")
        self.assertEqual(Decimal(10), values[other_expenses.MAINTENANCE_KEY])

    def test_manual_fixed_costs_win_over_invoices(self):
        self._invoice("M1", "2026-02-10", 600)
        from fixed_costs import save_fixed_costs
        save_fixed_costs(self.connection, self.community, self.period,
                         {other_expenses.MAINTENANCE_KEY: "75"})
        effective = excel_export_service._effective_fixed_costs(self.connection, self.case)
        self.assertEqual(Decimal(75), effective[other_expenses.MAINTENANCE_KEY])

    def test_repair_invoice_is_not_counted_as_monthly_maintenance(self):
        self._invoice("M1", "2026-02-10", 600)
        self._invoice("R1", "2026-04-10", 1500, notes="Reparación de la bomba de recirculación")
        values = other_expenses.automatic_fixed_costs(
            self.connection, self.community, "2026-01-01", "2026-12-31")
        self.assertEqual(Decimal(50), values[other_expenses.MAINTENANCE_KEY])
        shares = other_expenses.repair_shares(
            self.connection, self.community, "2026-01-01", "2026-12-31")
        self.assertEqual([Decimal("300.00")], [share.share for share in shares])
        # Y la quinta parte sigue entrando en los cuatro ejercicios siguientes.
        later = other_expenses.repair_shares(
            self.connection, self.community, "2030-01-01", "2030-12-31")
        self.assertEqual([Decimal("300.00")], [share.share for share in later])
        gone = other_expenses.repair_shares(
            self.connection, self.community, "2031-01-01", "2031-12-31")
        self.assertEqual([], gone)

    def test_repair_is_spread_over_five_years(self):
        self.connection.execute(
            "INSERT INTO gastos_extra(id_comunidad,descripcion,fecha,importe_total,años_amortizacion)"
            " VALUES (?,?,?,?,?)",
            (self.community, "Bomba de recirculación", "2025-06-01", 1000, 5),
        )
        self.connection.execute(
            "INSERT INTO gastos_extra(id_comunidad,descripcion,fecha,importe_total,años_amortizacion)"
            " VALUES (?,?,?,?,?)",
            (self.community, "Antigua ya amortizada", "2020-06-01", 1000, 5),
        )
        self.connection.execute(
            "INSERT INTO gastos_extra(id_comunidad,descripcion,fecha,importe_total,años_amortizacion)"
            " VALUES (?,?,?,?,?)",
            (self.community, "Cargo directo", "2026-05-01", 80, 1),
        )
        self.connection.commit()
        rows = excel_export_service._other_expense_rows(self.connection, self.case)
        self.assertEqual(
            [("2025-06-01", 200.0), ("2026-05-01", 80.0)],
            [(row[0], row[2]) for row in rows],
        )
        self.assertIn("1/5", rows[0][1])


if __name__ == "__main__":
    unittest.main()
