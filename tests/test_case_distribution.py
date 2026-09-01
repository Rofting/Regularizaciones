import sqlite3
import hashlib
import shutil
import sys
import tempfile
import unittest
from contextlib import redirect_stdout
from io import StringIO
from pathlib import Path
from unittest.mock import patch


PROJECT_ROOT = Path(__file__).resolve().parents[1]
CORE_DIR = PROJECT_ROOT / "core"
if str(CORE_DIR) not in sys.path:
    sys.path.insert(0, str(CORE_DIR))

import gestor_bd
import document_review
from case_distribution import (
    DistributionBlockedError,
    allocate_concept_cents,
    calculate_case_distribution,
)
from excel_export_service import calculate_case_input_hash
from excel_profiles import ConceptRule, ExcelProfile
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
               VALUES (?,'658_acs_v1','1',?,'template-hash',?,'active')""",
            (
                self.community_id,
                "plantillas/comunidades/658/658_acs_v1.xlsx",
                hashlib.sha256(
                    (PROJECT_ROOT / "config" / "excel_profiles" / "658_acs_v1.json").read_bytes()
                ).hexdigest(),
            ),
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

    def test_heating_reading_change_invalidates_reparto_and_legacy_projection(self):
        profile = ExcelProfile(
            key="synthetic_heating_v1",
            version="1",
            community_code="658",
            template_relative_path="plantillas/synthetic.xlsx",
            active_modules=("CALEFACCION",),
            required_sheets=("LECTURAS CALEF KWH",),
            required_formula_cells=(),
            concepts=(
                ConceptRule(
                    key="heating_fixed", allocation_method="equal",
                    actual_source="period_parameters.heating_fixed_actual",
                    billed_source="period_parameters.heating_fixed_billed", required=True,
                ),
                ConceptRule(
                    key="heating_variable", allocation_method="consumption",
                    actual_source="period_parameters.heating_variable_actual",
                    billed_source="period_parameters.heating_variable_billed", required=True,
                ),
            ),
            source_sha256=self.connection.execute(
                "SELECT profile_sha256 FROM excel_template_profiles WHERE id_template_profile=?",
                (self.profile_id,),
            ).fetchone()[0],
        )
        self._set_parameter("heating_fixed_actual", "20.00")
        self._set_parameter("heating_fixed_billed", "10.00")
        self._set_parameter("heating_variable_actual", "60.00")
        self._set_parameter("heating_variable_billed", "30.00")
        for owner_id, first, last in (
            (self.owner_one, 100, 130),
            (self.owner_two, 200, 210),
        ):
            self.connection.executemany(
                """INSERT INTO lecturas_vecino
                   (id_propietario,id_periodo,tipo,fecha_lectura,valor_acumulado,estado)
                   VALUES (?,?,'CALEFACCION',?,?,'real')""",
                [
                    (owner_id, self.period_id, "2025-09-01", first),
                    (owner_id, self.period_id, "2026-08-31", last),
                ],
            )
        self.connection.commit()
        heating_hash = calculate_case_input_hash(
            self.connection, id_case=self.case_id, project_root=PROJECT_ROOT, profile=profile,
        )
        self.connection.execute(
            """INSERT INTO excel_export_runs
               (id_case,id_periodo,id_template_profile,input_sha256,template_sha256,status)
               VALUES (?,?,?,?,?,'validated')""",
            (self.case_id, self.period_id, self.profile_id, heating_hash, "template-hash"),
        )
        self.connection.commit()

        with patch("case_distribution.load_profile", return_value=profile):
            calculate_case_distribution(self.connection, id_case=self.case_id)
            before = [tuple(row) for row in self._rows("heating_variable")]
            before_legacy = [tuple(row) for row in self.connection.execute(
                """SELECT id_propietario,importe_cobrado,importe_real,diferencia
                   FROM repartos WHERE id_periodo=? AND tipo_suministro='CALEFACCION'
                   ORDER BY id_propietario""",
                (self.period_id,),
            )]
            self.connection.execute(
                """UPDATE lecturas_vecino SET valor_acumulado=212
                   WHERE id_propietario=? AND tipo='CALEFACCION'
                     AND fecha_lectura='2026-08-31'""",
                (self.owner_two,),
            )
            self.connection.commit()

            with self.assertRaisesRegex(DistributionBlockedError, "regenerar Excel"):
                calculate_case_distribution(self.connection, id_case=self.case_id)

        self.assertEqual(before, [tuple(row) for row in self._rows("heating_variable")])
        self.assertEqual(before_legacy, [tuple(row) for row in self.connection.execute(
            """SELECT id_propietario,importe_cobrado,importe_real,diferencia
               FROM repartos WHERE id_periodo=? AND tipo_suministro='CALEFACCION'
               ORDER BY id_propietario""",
            (self.period_id,),
        )])

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

    def test_rejects_validated_export_when_profile_bytes_changed_without_version_bump(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as directory:
            root = Path(directory)
            profile_path = root / "config" / "excel_profiles" / "658_acs_v1.json"
            profile_path.parent.mkdir(parents=True)
            shutil.copy2(
                PROJECT_ROOT / "config" / "excel_profiles" / "658_acs_v1.json",
                profile_path,
            )
            registered_hash = hashlib.sha256(profile_path.read_bytes()).hexdigest()
            self.connection.execute(
                "UPDATE excel_template_profiles SET profile_sha256=? WHERE id_template_profile=?",
                (registered_hash, self.profile_id),
            )
            input_hash = calculate_case_input_hash(
                self.connection, id_case=self.case_id, project_root=root
            )
            self.connection.execute(
                """INSERT INTO excel_export_runs
                   (id_case,id_periodo,id_template_profile,input_sha256,template_sha256,status)
                   VALUES (?,?,?,?,?,'validated')""",
                (self.case_id, self.period_id, self.profile_id, input_hash, "template-hash"),
            )
            profile_path.write_bytes(profile_path.read_bytes() + b"\n")
            self.connection.commit()

            with self.assertRaisesRegex(DistributionBlockedError, "huella"):
                calculate_case_distribution(
                    self.connection, id_case=self.case_id, project_root=root
                )

    def test_accepts_zero_consumption_as_a_zero_weight(self):
        self.connection.execute(
            """UPDATE lecturas_vecino SET valor_acumulado=10
               WHERE id_propietario=? AND fecha_lectura='2026-08-31'""",
            (self.owner_one,),
        )
        self.connection.commit()
        self._refresh_validated_export()

        result = calculate_case_distribution(self.connection, id_case=self.case_id)

        variable_rows = [tuple(row) for row in self._rows("acs_variable")]
        self.assertEqual(6, result.owner_result_count)
        self.assertEqual(0.0, variable_rows[0][4])
        self.assertEqual(0, variable_rows[0][2])
        self.assertEqual(30000, variable_rows[1][2])

    def test_fixed_only_profile_still_blocks_unapproved_final_counter_reading(self):
        fixed_only = ExcelProfile(
            key="synthetic_fixed_only_v1",
            version="1",
            community_code="658",
            template_relative_path="plantillas/synthetic-fixed.xlsx",
            active_modules=(),
            required_sheets=(),
            required_formula_cells=(),
            concepts=(
                ConceptRule(
                    key="acs_fixed", allocation_method="equal",
                    actual_source="period_parameters.acs_fixed_actual",
                    billed_source="period_parameters.acs_fixed_billed", required=True,
                ),
            ),
            source_sha256="profile-fixed",
        )
        profile_id = self.connection.execute(
            """INSERT INTO excel_template_profiles
               (id_comunidad,profile_key,profile_version,template_relative_path,
                template_sha256,profile_sha256,status)
               VALUES (?,?,'1',?,'template-fixed','profile-fixed','active')""",
            (self.community_id, fixed_only.key, fixed_only.template_relative_path),
        ).lastrowid
        self.connection.execute(
            """UPDATE lecturas_vecino SET estado='contador_averiado'
               WHERE id_propietario=? AND fecha_lectura='2026-08-31'""",
            (self.owner_one,),
        )
        fixed_hash = calculate_case_input_hash(
            self.connection, id_case=self.case_id, project_root=PROJECT_ROOT, profile=fixed_only,
        )
        self.connection.execute(
            """INSERT INTO excel_export_runs
               (id_case,id_periodo,id_template_profile,input_sha256,template_sha256,status)
               VALUES (?,?,?,?,?,'validated')""",
            (self.case_id, self.period_id, profile_id, fixed_hash, "template-fixed"),
        )
        self.connection.commit()

        with patch("case_distribution.load_profile", return_value=fixed_only):
            with self.assertRaisesRegex(DistributionBlockedError, "lectura final"):
                calculate_case_distribution(self.connection, id_case=self.case_id)

    def test_distribution_unblocks_after_every_counter_reset_is_approved(self):
        document_id = self.connection.execute(
            """INSERT INTO source_documents
               (id_case,original_name,archived_path,sha256,document_kind,status)
               VALUES (?,'lecturas.xlsx','archivo_lecturas.xlsx',
                       'lecturas-sinteticas','meter_readings','under_review')""",
            (self.case_id,),
        ).lastrowid
        self.connection.execute(
            """UPDATE lecturas_vecino SET valor_acumulado=?,estado='contador_averiado'
               WHERE id_propietario=? AND fecha_lectura='2026-08-31'""",
            (5, self.owner_one),
        )
        self.connection.execute(
            """UPDATE lecturas_vecino SET valor_acumulado=?,estado='contador_averiado'
               WHERE id_propietario=? AND fecha_lectura='2026-08-31'""",
            (6, self.owner_two),
        )
        self.connection.commit()
        first_issue = document_review.create_review_issue(
            self.connection, self.case_id, document_id, code="COUNTER_RESET",
            field_name="reading.A.ACS", message="Contador reiniciado", detected_value="10 -> 5",
        )
        second_issue = document_review.create_review_issue(
            self.connection, self.case_id, document_id, code="COUNTER_RESET",
            field_name="reading.B.ACS", message="Contador reiniciado", detected_value="20 -> 6",
        )
        self._refresh_validated_export()

        with self.assertRaisesRegex(DistributionBlockedError, "incidencias abiertas"):
            calculate_case_distribution(self.connection, id_case=self.case_id)

        document_review.approve_counter_reset_estimate(
            self.connection, first_issue.id_issue, consumption="12",
            reason="Sustitución confirmada", approved_by="gestora",
        )
        self._refresh_validated_export()
        with self.assertRaisesRegex(DistributionBlockedError, "incidencias abiertas"):
            calculate_case_distribution(self.connection, id_case=self.case_id)

        document_review.approve_counter_reset_estimate(
            self.connection, second_issue.id_issue, consumption="9",
            reason="Sustitución confirmada", approved_by="gestora",
        )
        self._refresh_validated_export()

        result = calculate_case_distribution(self.connection, id_case=self.case_id)

        self.assertEqual(6, result.owner_result_count)
        self.assertEqual(
            {"resolved"},
            {row[0] for row in self.connection.execute(
                "SELECT status FROM review_issues WHERE id_issue IN (?,?)",
                (first_issue.id_issue, second_issue.id_issue),
            )},
        )

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
