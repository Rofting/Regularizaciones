import sqlite3
import sys
import unittest
from contextlib import redirect_stdout
from io import StringIO
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
CORE_DIR = PROJECT_ROOT / "core"
if str(CORE_DIR) not in sys.path:
    sys.path.insert(0, str(CORE_DIR))

import gestor_bd
from case_distribution import (
    DistributionBlockedError,
    allocate_concept_cents,
    calculate_case_distribution,
)
from excel_export_service import calculate_case_input_hash
from tests.helpers import temporary_database


class CaseDistributionTest(unittest.TestCase):
    def setUp(self):
        self.database = temporary_database()
        self.connection, path = self.database.__enter__()
        with redirect_stdout(StringIO()):
            gestor_bd.crear_bd(str(path))
        self.community_id = gestor_bd.obtener_o_crear_comunidad(
            self.connection, "658", "Comunidad sintética 658"
        )
        self.period_id = self.connection.execute(
            """INSERT INTO periodos
               (id_comunidad,nombre,fecha_inicio,fecha_fin,estado)
               VALUES (?,'2025-2026','2025-09-01','2026-08-31','abierto')""",
            (self.community_id,),
        ).lastrowid
        self.case_id = self.connection.execute(
            """INSERT INTO regularization_cases
               (id_comunidad,nombre,fecha_inicio,fecha_fin,estado,id_periodo)
               VALUES (?,'2025-2026','2025-09-01','2026-08-31',
                       'ready_for_calculation',?)""",
            (self.community_id, self.period_id),
        ).lastrowid
        self.profile_id = self.connection.execute(
            """INSERT INTO excel_template_profiles
               (id_comunidad,profile_key,profile_version,template_relative_path,
                template_sha256,profile_sha256,status)
               VALUES (?,'658_acs_v1','1',?,'template-hash','profile-hash','active')""",
            (self.community_id, "plantillas/comunidades/658/658_acs_v1.xlsx"),
        ).lastrowid
        self.connection.execute(
            """INSERT INTO excel_export_runs
               (id_case,id_periodo,id_template_profile,input_sha256,template_sha256,status)
               VALUES (?,?,?,?,?,'validated')""",
            (self.case_id, self.period_id, self.profile_id, "inputs", "template-hash"),
        )
        self.owner_one = self._add_owner("A", 1, 10, 30)
        self.owner_two = self._add_owner("B", 2, 20, 30)
        self.inactive_owner = self._add_owner("Z", 99, 1, 99, active=0)
        self._set_parameter("acs_fixed_actual", "120.01")
        self._set_parameter("acs_fixed_billed", "100.00")
        self._set_parameter("acs_variable_actual", "300.00")
        self._set_parameter("acs_variable_billed", "200.00")
        self._set_parameter("credit_actual", "-3.01")
        self.connection.commit()
        self._refresh_validated_export()

    def tearDown(self):
        self.database.__exit__(None, None, None)

    def _add_owner(self, code, coefficient, first, last, *, active=1):
        owner_id = self.connection.execute(
            """INSERT INTO propietarios
               (id_comunidad,codigo_vivienda,nombre_propietario,coeficiente,activo)
               VALUES (?,?,?,?,?)""",
            (self.community_id, code, f"Propietario {code}", coefficient, active),
        ).lastrowid
        self.connection.executemany(
            """INSERT INTO lecturas_vecino
               (id_propietario,id_periodo,tipo,fecha_lectura,valor_acumulado,estado)
               VALUES (?,?,'ACS',?,?,'real')""",
            [
                (owner_id, self.period_id, "2025-09-01", first),
                (owner_id, self.period_id, "2026-08-31", last),
            ],
        )
        return owner_id

    def _set_parameter(self, key, amount):
        self.connection.execute(
            """INSERT INTO period_parameters
               (id_comunidad,id_periodo,parameter_key,numeric_value,unit)
               VALUES (?,?,?,?, 'EUR')""",
            (self.community_id, self.period_id, key, amount),
        )

    def _refresh_validated_export(self):
        input_hash = calculate_case_input_hash(
            self.connection, id_case=self.case_id, project_root=PROJECT_ROOT,
        )
        self.connection.execute(
            """INSERT INTO excel_export_runs
               (id_case,id_periodo,id_template_profile,input_sha256,template_sha256,status)
               VALUES (?,?,?,?,?,'validated')""",
            (self.case_id, self.period_id, self.profile_id, input_hash, "template-hash"),
        )
        self.connection.commit()

    def _rows(self, concept):
        return self.connection.execute(
            """SELECT id_propietario,billed_cents,actual_cents,difference_cents,consumption
               FROM owner_concept_results WHERE id_periodo=? AND concept_key=?
               ORDER BY id_propietario""",
            (self.period_id, concept),
        ).fetchall()

    def test_allocates_fixed_variable_and_credit_exactly_to_the_cent(self):
        stages = []
        result = calculate_case_distribution(
            self.connection, id_case=self.case_id,
            progress=lambda stage, payload: stages.append(stage),
        )

        self.assertEqual(self.case_id, result.id_case)
        self.assertEqual(self.period_id, result.id_periodo)
        self.assertEqual(12001, result.concept_totals_cents["acs_fixed"])
        self.assertEqual(30000, result.concept_totals_cents["acs_variable"])
        self.assertEqual(-301, result.concept_totals_cents["credit"])
        self.assertEqual(
            [(self.owner_one, 5000, 6001, 1001, None),
             (self.owner_two, 5000, 6000, 1000, None)],
            [tuple(row) for row in self._rows("acs_fixed")],
        )
        self.assertEqual(
            [(self.owner_one, 13333, 20000, 6667, 20.0),
             (self.owner_two, 6667, 10000, 3333, 10.0)],
            [tuple(row) for row in self._rows("acs_variable")],
        )
        self.assertEqual(
            [(self.owner_one, 0, -100, -100, None),
             (self.owner_two, 0, -201, -201, None)],
            [tuple(row) for row in self._rows("credit")],
        )
        self.assertEqual(
            "cuadrado",
            self.connection.execute(
                """SELECT status FROM reconciliations
                   WHERE id_periodo=? AND concept_key='acs_variable'""",
                (self.period_id,),
            ).fetchone()[0],
        )
        self.assertEqual(6, result.owner_result_count)
        self.assertEqual(
            ["validate_export", "calculate_acs_fixed", "calculate_acs_variable",
             "calculate_credit", "reconcile", "complete"],
            stages,
        )

    def test_largest_remainder_breaks_ties_by_owner_identifier_for_negative_credits(self):
        self.assertEqual(
            {1: -34, 2: -33, 3: -33},
            allocate_concept_cents(-100, {3: 1, 2: 1, 1: 1}),
        )

    def test_requires_a_new_validated_export_before_rerun_after_parameter_change(self):
        calculate_case_distribution(self.connection, id_case=self.case_id)
        before = [tuple(row) for row in self._rows("acs_fixed")]
        self.connection.execute(
            "UPDATE period_parameters SET numeric_value=120.02 WHERE parameter_key='acs_fixed_actual'"
        )
        self.connection.commit()

        with self.assertRaisesRegex(DistributionBlockedError, "regenerar Excel"):
            calculate_case_distribution(self.connection, id_case=self.case_id)
        self.assertEqual(before, [tuple(row) for row in self._rows("acs_fixed")])

        self._refresh_validated_export()
        second = calculate_case_distribution(self.connection, id_case=self.case_id)

        self.assertEqual(6, second.owner_result_count)
        self.assertEqual(2, len(self._rows("acs_fixed")))
        self.assertFalse(any(row[0] == self.inactive_owner for row in self._rows("acs_fixed")))
        self.assertEqual(12002, sum(row[2] for row in self._rows("acs_fixed")))

    def test_blocks_reading_change_after_export_without_modifying_results(self):
        calculate_case_distribution(self.connection, id_case=self.case_id)
        before = [tuple(row) for row in self._rows("acs_variable")]
        self.connection.execute(
            """UPDATE lecturas_vecino SET valor_acumulado=31
               WHERE id_propietario=? AND fecha_lectura='2026-08-31'""",
            (self.owner_two,),
        )
        self.connection.commit()

        with self.assertRaisesRegex(DistributionBlockedError, "regenerar Excel"):
            calculate_case_distribution(self.connection, id_case=self.case_id)

        self.assertEqual(before, [tuple(row) for row in self._rows("acs_variable")])

    def test_blocks_owner_coefficient_or_activity_change_after_export(self):
        calculate_case_distribution(self.connection, id_case=self.case_id)
        before = [tuple(row) for row in self._rows("credit")]
        self.connection.execute(
            "UPDATE propietarios SET coeficiente=3,activo=0 WHERE id_propietario=?",
            (self.owner_two,),
        )
        self.connection.commit()

        with self.assertRaisesRegex(DistributionBlockedError, "regenerar Excel"):
            calculate_case_distribution(self.connection, id_case=self.case_id)

        self.assertEqual(before, [tuple(row) for row in self._rows("credit")])

    def test_absent_optional_concept_removes_stale_rows(self):
        self.connection.execute(
            "DELETE FROM period_parameters WHERE parameter_key='credit_actual'"
        )
        self.connection.execute(
            """INSERT INTO owner_concept_results
               (id_propietario,id_periodo,concept_key,billed_cents,actual_cents,difference_cents)
               VALUES (?,?,?,?,?,?)""",
            (self.owner_one, self.period_id, "credit", 0, -999, -999),
        )
        self.connection.commit()
        self._refresh_validated_export()

        result = calculate_case_distribution(self.connection, id_case=self.case_id)

        self.assertNotIn("credit", result.concept_totals_cents)
        self.assertEqual([], self._rows("credit"))

    def test_blocks_zero_weight_without_writing_partial_results(self):
        self.connection.execute(
            "UPDATE propietarios SET coeficiente=0 WHERE id_propietario=?", (self.owner_two,)
        )
        self.connection.commit()
        self._refresh_validated_export()

        with self.assertRaisesRegex(DistributionBlockedError, "peso"):
            calculate_case_distribution(self.connection, id_case=self.case_id)

        self.assertEqual(
            0,
            self.connection.execute("SELECT COUNT(*) FROM owner_concept_results").fetchone()[0],
        )

    def test_blocks_counter_reset_without_approved_estimation(self):
        self.connection.execute(
            """UPDATE lecturas_vecino SET valor_acumulado=5
               WHERE id_propietario=? AND fecha_lectura='2026-08-31'""",
            (self.owner_one,),
        )
        self.connection.commit()
        self._refresh_validated_export()

        with self.assertRaisesRegex(DistributionBlockedError, "contador"):
            calculate_case_distribution(self.connection, id_case=self.case_id)

        self.assertEqual(0, self.connection.execute(
            "SELECT COUNT(*) FROM owner_concept_results"
        ).fetchone()[0])

    def test_accepts_explicitly_approved_estimation(self):
        self.connection.execute(
            """UPDATE lecturas_vecino
               SET estado='estimado', approved_by='gestor', approved_at='2026-08-31T10:00:00'
               WHERE id_propietario=? AND fecha_lectura='2026-08-31'""",
            (self.owner_one,),
        )
        self.connection.commit()
        self._refresh_validated_export()

        result = calculate_case_distribution(self.connection, id_case=self.case_id)

        self.assertEqual(6, result.owner_result_count)

    def test_blocks_when_latest_export_is_not_validated(self):
        self.connection.execute(
            """INSERT INTO excel_export_runs
               (id_case,id_periodo,id_template_profile,input_sha256,template_sha256,status)
               VALUES (?,?,?,?,?,'failed')""",
            (self.case_id, self.period_id, self.profile_id, "new-inputs", "template-hash"),
        )
        self.connection.commit()

        with self.assertRaisesRegex(DistributionBlockedError, "Excel"):
            calculate_case_distribution(self.connection, id_case=self.case_id)


if __name__ == "__main__":
    unittest.main()
