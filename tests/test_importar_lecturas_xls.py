import sys
import tempfile
import unittest
from contextlib import redirect_stdout
from io import StringIO
from pathlib import Path

import xlwt


PROJECT_ROOT = Path(__file__).resolve().parents[1]
CORE_DIR = PROJECT_ROOT / "core"
if str(CORE_DIR) not in sys.path:
    sys.path.insert(0, str(CORE_DIR))

import gestor_bd
from importar_lecturas_xls import import_readings_xls


def make_readings_xls(path: Path, rows: list[tuple[int, str, float, float]]) -> Path:
    workbook = xlwt.Workbook()
    sheet = workbook.add_sheet("Lecturas")
    for column, value in enumerate(("Cod.", "Propiedad", "Nombre", "7/2024", "7/2025")):
        sheet.write(0, column, value)
    for row_index, (code, property_code, initial, final) in enumerate(rows, start=1):
        sheet.write(row_index, 0, code)
        sheet.write(row_index, 1, property_code)
        sheet.write(row_index, 2, f"VECINO {code}")
        sheet.write(row_index, 3, initial)
        sheet.write(row_index, 4, final)
    workbook.save(str(path))
    return path


class ReadingXlsImportTest(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory(ignore_cleanup_errors=True)
        self.database = Path(self.directory.name) / "gestion.db"
        self.xls_path = Path(self.directory.name) / "readings.xls"
        with redirect_stdout(StringIO()):
            gestor_bd.crear_bd(str(self.database))
        self.connection = gestor_bd.conectar(str(self.database))
        self.community_id = gestor_bd.obtener_o_crear_comunidad(
            self.connection, "TEST", "COMUNIDAD DE PRUEBA"
        )
        for code in ("PA2-1A", "PA2-1B"):
            self.connection.execute(
                """
                INSERT INTO propietarios
                    (id_comunidad, codigo_vivienda, nombre_propietario)
                VALUES (?, ?, ?)
                """,
                (self.community_id, code, f"VECINO {code}"),
            )
        self.period_id = self.connection.execute(
            """
            INSERT INTO periodos
                (id_comunidad, nombre, fecha_inicio, fecha_fin)
            VALUES (?, '2024-2025', '2024-08-01', '2025-07-31')
            """,
            (self.community_id,),
        ).lastrowid
        self.connection.commit()

    def tearDown(self):
        self.connection.close()
        self.directory.cleanup()

    def test_imports_cumulative_readings_and_is_idempotent(self):
        make_readings_xls(
            self.xls_path,
            [(101, "PA2-1A", 100, 112), (102, "PA2-1B", 200, 207)],
        )
        first = import_readings_xls(
            self.connection, self.community_id, self.period_id, self.xls_path
        )
        second = import_readings_xls(
            self.connection, self.community_id, self.period_id, self.xls_path
        )
        values = [
            row[0]
            for row in self.connection.execute(
                "SELECT valor_acumulado FROM lecturas_vecino ORDER BY id_lectura"
            )
        ]

        self.assertFalse(first.reset_detected)
        self.assertEqual(4, first.inserted)
        self.assertEqual(4, second.duplicates)
        self.assertEqual([100.0, 112.0, 200.0, 207.0], values)

    def test_detects_annual_reset_and_uses_final_value_as_consumption(self):
        make_readings_xls(
            self.xls_path,
            [(101, "PA2-1A", 100, 5), (102, "PA2-1B", 200, 7)],
        )
        summary = import_readings_xls(
            self.connection, self.community_id, self.period_id, self.xls_path
        )
        final_values = [
            row[0]
            for row in self.connection.execute(
                """
                SELECT valor_acumulado FROM lecturas_vecino
                WHERE fecha_lectura='2025-07-31' ORDER BY id_propietario
                """
            )
        ]

        self.assertTrue(summary.reset_detected)
        self.assertEqual([5.0, 7.0], final_values)

    def test_unknown_nonzero_property_rolls_back_all_readings(self):
        make_readings_xls(
            self.xls_path,
            [(101, "PA2-1A", 100, 112), (999, "DESCONOCIDA", 0, 4)],
        )
        with self.assertRaisesRegex(ValueError, "DESCONOCIDA"):
            import_readings_xls(
                self.connection, self.community_id, self.period_id, self.xls_path
            )
        count = self.connection.execute("SELECT COUNT(*) FROM lecturas_vecino").fetchone()[0]
        self.assertEqual(0, count)

    def test_isolated_lower_reading_requires_manual_review(self):
        make_readings_xls(
            self.xls_path,
            [(101, "PA2-1A", 100, 112), (102, "PA2-1B", 200, 0)],
        )
        with self.assertRaisesRegex(ValueError, "PA2-1B"):
            import_readings_xls(
                self.connection, self.community_id, self.period_id, self.xls_path
            )
        count = self.connection.execute("SELECT COUNT(*) FROM lecturas_vecino").fetchone()[0]
        self.assertEqual(0, count)


if __name__ == "__main__":
    unittest.main()
