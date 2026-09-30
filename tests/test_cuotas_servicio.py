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

import cuotas_servicio
import gestor_bd


class CuotasServicioTest(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory(ignore_cleanup_errors=True)
        path = str(Path(self.directory.name) / "cuotas.db")
        with redirect_stdout(StringIO()):
            gestor_bd.crear_bd(path)
        self.connection = gestor_bd.conectar(path)
        self.community = gestor_bd.obtener_o_crear_comunidad(self.connection, "T", "Comunidad")
        self.period = self.connection.execute(
            "INSERT INTO periodos(id_comunidad,nombre,fecha_inicio,fecha_fin) VALUES (?,?,?,?)",
            (self.community, "2025-2026", "2025-09-01", "2026-08-31"),
        ).lastrowid

    def tearDown(self):
        self.connection.close()
        self.directory.cleanup()

    def summary(self, servicio="ACS"):
        return cuotas_servicio.resumen(
            self.connection, community_id=self.community, period_id=self.period, servicio=servicio,
        )

    def test_amounts_accept_spanish_and_english_formats(self):
        for text, expected in (("1.234,50", "1234.50"), ("1234,5", "1234.5"), ("1,234.50", "1234.50"),
                               ("400 €", "400"), ("1.200", "1200"), ("12.5", "12.5")):
            with self.subTest(text=text):
                self.assertEqual(Decimal(expected), cuotas_servicio._importe(text, "Importe"))

    def test_months_accept_several_spellings(self):
        for text in ("2026-01", "01/2026", "1/26", "enero 2026", "ene-26", "Enero de 2026"):
            with self.subTest(text=text):
                self.assertEqual("2026-01-01", cuotas_servicio.parse_mes(text))
        with self.assertRaises(ValueError):
            cuotas_servicio.parse_mes("13/2026")

    def test_monthly_fees_follow_tramos_and_regeneration_replaces_them(self):
        created = cuotas_servicio.generar_cuotas_mensuales(
            self.connection, community_id=self.community, period_id=self.period,
            servicio="ACS", tramos=[("09/2025", "400"), ("enero 2026", "300,00")],
        )
        self.assertEqual(12, created)
        self.assertEqual(Decimal("4000"), self.summary().fija)  # 4×400 + 8×300
        cuotas_servicio.generar_cuotas_mensuales(
            self.connection, community_id=self.community, period_id=self.period,
            servicio="ACS", tramos=[("2026-03", "100")],
        )
        self.assertEqual((Decimal("600"), 6), (self.summary().fija, self.summary().apuntes))

    def test_variable_fee_accepts_day_month_year_and_is_kept_on_regeneration(self):
        cuotas_servicio.registrar_cuota(
            self.connection, community_id=self.community, period_id=self.period,
            servicio="CALEFACCION", concepto="variable", fecha="31/12/2025", importe="1.250,75",
        )
        cuotas_servicio.generar_cuotas_mensuales(
            self.connection, community_id=self.community, period_id=self.period,
            servicio="CALEFACCION", tramos=[("2025-09", "50")],
        )
        summary = self.summary("CALEFACCION")
        self.assertEqual((Decimal("1250.75"), Decimal("600")), (summary.variable, summary.fija))
        self.assertEqual(Decimal("0"), self.summary("ACS").total)

    def test_dwellings_count_only_active_homes(self):
        self.connection.executemany(
            "INSERT INTO propietarios(id_comunidad,codigo_vivienda,nombre_propietario,activo,tipo_unidad) VALUES (?,?,?,?,?)",
            [(self.community, "1A", "A", 1, "vivienda"), (self.community, "1B", "B", 1, "vivienda"),
             (self.community, "L1", "Local", 1, "local"), (self.community, "2A", "Baja", 0, "vivienda")],
        )
        self.assertEqual(2, cuotas_servicio.viviendas_facturables(self.connection, self.community))


if __name__ == "__main__":
    unittest.main()
