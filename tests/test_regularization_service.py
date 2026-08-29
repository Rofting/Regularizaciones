import sys
import unittest
from contextlib import redirect_stdout
from decimal import Decimal
from io import StringIO
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
CORE_DIR = PROJECT_ROOT / "core"
if str(CORE_DIR) not in sys.path:
    sys.path.insert(0, str(CORE_DIR))

import gestor_bd
from regularization_service import allocate_cents, calculate_period
from tests.helpers import temporary_database


def prepare_period(connection):
    community_id = gestor_bd.obtener_o_crear_comunidad(
        connection, "TEST", "COMUNIDAD DE PRUEBA"
    )
    period_id = connection.execute(
        """INSERT INTO periodos
           (id_comunidad,nombre,fecha_inicio,fecha_fin)
           VALUES (?,'2024-2025','2024-08-01','2025-07-31')""",
        (community_id,),
    ).lastrowid
    owner_ids = []
    for index, consumption in enumerate((3.0, 1.0), start=1):
        owner_id = connection.execute(
            """INSERT INTO propietarios
               (id_comunidad,codigo_vivienda,nombre_propietario)
               VALUES (?,?,?)""",
            (community_id, f"V{index}", f"VECINO {index}"),
        ).lastrowid
        owner_ids.append(owner_id)
        connection.executemany(
            """INSERT INTO lecturas_vecino
               (id_propietario,id_periodo,tipo,fecha_lectura,valor_acumulado)
               VALUES (?,?, 'ACS', ?, ?)""",
            [
                (owner_id, period_id, "2024-08-01", 0),
                (owner_id, period_id, "2025-07-31", consumption),
            ],
        )
    batch_id = connection.execute(
        """INSERT INTO import_batches
           (id_comunidad,id_periodo,source_kind,source_path,source_sha256,status)
           VALUES (?,?,'excel_reference','fixture','hash','validated')""",
        (community_id, period_id),
    ).lastrowid
    for concept, billed, actual in (
        ("acs_fixed", "1.00", "1.02"),
        ("acs_variable", "2.00", "2.04"),
    ):
        connection.executemany(
            """INSERT INTO source_values
               (id_batch,entity_type,entity_key,field_name,normalized_value)
               VALUES (?,'concept',?,?,?)""",
            [
                (batch_id, concept, "billed_cents", billed),
                (batch_id, concept, "actual_cents", actual),
            ],
        )
    connection.commit()
    return period_id, batch_id, owner_ids


class RegularizationServiceTest(unittest.TestCase):
    def test_allocate_cents_preserves_positive_and_negative_totals(self):
        weights = {2: Decimal("1"), 1: Decimal("1"), 3: Decimal("1")}
        self.assertEqual({1: 34, 2: 33, 3: 33}, allocate_cents(100, weights))
        self.assertEqual({1: -34, 2: -33, 3: -33}, allocate_cents(-100, weights))

    def test_calculates_fixed_and_variable_results_from_reference_totals(self):
        with temporary_database() as (connection, path):
            with redirect_stdout(StringIO()):
                gestor_bd.crear_bd(str(path))
            period_id, batch_id, owner_ids = prepare_period(connection)
            count = calculate_period(connection, period_id, batch_id)
            rows = connection.execute(
                """SELECT id_propietario,concept_key,billed_cents,actual_cents,
                          difference_cents,consumption
                   FROM owner_concept_results ORDER BY concept_key,id_propietario"""
            ).fetchall()

        self.assertEqual(4, count)
        self.assertEqual(
            [
                (owner_ids[0], "acs_fixed", 50, 51, 1, None),
                (owner_ids[1], "acs_fixed", 50, 51, 1, None),
                (owner_ids[0], "acs_variable", 150, 153, 3, 3.0),
                (owner_ids[1], "acs_variable", 50, 51, 1, 1.0),
            ],
            [tuple(row) for row in rows],
        )


if __name__ == "__main__":
    unittest.main()
