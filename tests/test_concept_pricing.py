import sys
import unittest
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
CORE_DIR = PROJECT_ROOT / "core"
if str(CORE_DIR) not in sys.path:
    sys.path.insert(0, str(CORE_DIR))

import gestor_bd
from tests.helpers import temporary_database


class ConceptPricingTest(unittest.TestCase):
    def setUp(self):
        self.database = temporary_database()
        self.connection, path = self.database.__enter__()
        gestor_bd.crear_bd(str(path))
        self.community_id = gestor_bd.obtener_o_crear_comunidad(
            self.connection, "900", "Comunidad de precios"
        )
        self.period_id = self.connection.execute(
            """INSERT INTO periodos
               (id_comunidad,nombre,fecha_inicio,fecha_fin,estado)
               VALUES (?,'enero parcial','2026-01-16','2026-01-31','abierto')""",
            (self.community_id,),
        ).lastrowid
        self.case = {
            "id_comunidad": self.community_id,
            "id_periodo": self.period_id,
            "fecha_inicio": "2026-01-16",
            "fecha_fin": "2026-01-31",
        }

    def tearDown(self):
        self.database.__exit__(None, None, None)

    def _invoice(self, *, consumption=310, with_component=True):
        invoice_id = self.connection.execute(
            """INSERT INTO facturas
               (id_comunidad,id_periodo,tipo_suministro,fecha_inicio,fecha_fin,
                consumo_total,unidad_consumo,termino_variable,impuestos,iva,importe_total)
               VALUES (?,?,'GAS','2026-01-01','2026-01-31',?,'kWh',31,6,5,42)""",
            (self.community_id, self.period_id, consumption),
        ).lastrowid
        if with_component:
            self.connection.execute(
                """INSERT INTO invoice_components
                   (id_factura,component_key,amount,unit) VALUES (?,'variable',31,'EUR')""",
                (invoice_id,),
            )
        self.connection.commit()

    def test_derives_net_unit_price_and_prorates_overlapping_invoice_days(self):
        from concept_pricing import derive_unit_price

        self._invoice()
        result = derive_unit_price(
            self.connection, self.case, service="GAS", component="variable"
        )

        self.assertEqual(1600, result.total_net_cents)
        self.assertEqual(160.0, result.total_consumption)
        self.assertAlmostEqual(0.10, result.unit_price_euros, places=8)
        self.assertEqual(1, result.invoice_count)

    def test_zero_consumption_keeps_cost_but_has_no_unit_price(self):
        from concept_pricing import derive_unit_price

        self._invoice(consumption=0)
        result = derive_unit_price(
            self.connection, self.case, service="GAS", component="variable"
        )

        self.assertEqual(1600, result.total_net_cents)
        self.assertEqual(0.0, result.total_consumption)
        self.assertIsNone(result.unit_price_euros)

    def test_missing_invoice_component_is_not_silently_replaced_by_total(self):
        from concept_pricing import DerivedPricingError, derive_unit_price

        self._invoice(with_component=False)
        with self.assertRaisesRegex(DerivedPricingError, "componente variable"):
            derive_unit_price(
                self.connection, self.case, service="GAS", component="variable"
            )


if __name__ == "__main__":
    unittest.main()
