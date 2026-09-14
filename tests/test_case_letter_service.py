import shutil
import hashlib
import sqlite3
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest.mock import patch


PROJECT_ROOT = Path(__file__).resolve().parents[1]
CORE_DIR = PROJECT_ROOT / "core"
if str(CORE_DIR) not in sys.path:
    sys.path.insert(0, str(CORE_DIR))

import gestor_bd
from excel_export_service import calculate_case_input_hash


class CaseLetterServiceTest(unittest.TestCase):
    """Contrato de cartas auditables, independiente de datos privados."""

    def setUp(self):
        self.directory = tempfile.TemporaryDirectory(ignore_cleanup_errors=True)
        self.root = Path(self.directory.name) / "project"
        shutil.copytree(PROJECT_ROOT / "config", self.root / "config")
        (self.root / "plantillas").mkdir(parents=True)
        shutil.copy2(
            PROJECT_ROOT / "plantillas" / "Plantilla_Cartas.docx",
            self.root / "plantillas" / "Plantilla_Cartas.docx",
        )
        (self.root / "config" / "letter_identities.json").write_text(
            '{"communities":{"658":{"office_name":"Gestión Portable",'
            '"footer":"Atención de la comunidad",'
            '"signature":"Equipo gestor", "city":"Valencia"}}}',
            encoding="utf-8",
        )

        self.database_path = str(Path(self.directory.name) / "letters.db")
        gestor_bd.crear_bd(self.database_path)
        self.connection = gestor_bd.conectar(self.database_path)
        self.community_id = gestor_bd.obtener_o_crear_comunidad(
            self.connection, "658", "Comunidad de prueba"
        )
        self.period_id = self.connection.execute(
            """INSERT INTO periodos(id_comunidad,nombre,fecha_inicio,fecha_fin)
               VALUES (?,?,?,?)""",
            (self.community_id, "2025-2026", "2025-09-01", "2026-08-31"),
        ).lastrowid
        self.case_id = self.connection.execute(
            """INSERT INTO regularization_cases
               (id_comunidad,nombre,fecha_inicio,fecha_fin,estado,id_periodo)
               VALUES (?,?,?,?,?,?)""",
            (
                self.community_id, "Regularización 2025-2026", "2025-09-01",
                "2026-08-31", "reconciled", self.period_id,
            ),
        ).lastrowid
        self.owner_one = self._owner("A-1", "Ana Vecina", 1.0, 100, 120)
        self.owner_two = self._owner("B-2", "Bruno Vecino", 2.0, 200, 210)
        self._parameter("acs_fixed_actual", "120.00")
        self._parameter("acs_fixed_billed", "100.00")
        self._parameter("acs_variable_actual", "300.00")
        self._parameter("acs_variable_billed", "200.00")
        self._parameter("credit_actual", "-3.00")
        self._result_rows()
        self._reconcile_all()
        self._validated_export()
        self.connection.commit()

    def tearDown(self):
        self.connection.close()
        self.directory.cleanup()

    def _owner(self, code, name, coefficient, start, end):
        owner_id = self.connection.execute(
            """INSERT INTO propietarios
               (id_comunidad,codigo_vivienda,nombre_propietario,coeficiente,activo)
               VALUES (?,?,?,?,1)""",
            (self.community_id, code, name, coefficient),
        ).lastrowid
        self.connection.executemany(
            """INSERT INTO lecturas_vecino
               (id_propietario,id_periodo,tipo,fecha_lectura,valor_acumulado,estado)
               VALUES (?,?,'ACS',?,?,'real')""",
            [
                (owner_id, self.period_id, "2025-09-01", start),
                (owner_id, self.period_id, "2026-08-31", end),
            ],
        )
        return owner_id

    def _parameter(self, key, value):
        self.connection.execute(
            """INSERT INTO period_parameters
               (id_comunidad,id_periodo,parameter_key,numeric_value,unit)
               VALUES (?,?,?,?,?)""",
            (self.community_id, self.period_id, key, value, "EUR"),
        )

    def _result_rows(self):
        rows = []
        for owner_id, consumption, fixed_actual, variable_actual, credit_actual in (
            (self.owner_one, 20.0, 6000, 20000, -100),
            (self.owner_two, 10.0, 6000, 10000, -200),
        ):
            rows.extend([
                (owner_id, self.period_id, "acs_fixed", None, 5000, fixed_actual,
                 fixed_actual - 5000),
                (owner_id, self.period_id, "acs_variable", consumption,
                 10000 if owner_id == self.owner_one else 10000,
                 variable_actual, variable_actual - 10000),
                (owner_id, self.period_id, "credit", None, 0, credit_actual, credit_actual),
            ])
        self.connection.executemany(
            """INSERT INTO owner_concept_results
               (id_propietario,id_periodo,concept_key,consumption,billed_cents,
                actual_cents,difference_cents)
               VALUES (?,?,?,?,?,?,?)""",
            rows,
        )

    def _reconcile_all(self):
        for key, billed, actual in (
            ("acs_fixed", 10000, 12000),
            ("acs_variable", 20000, 30000),
            ("credit", 0, -300),
        ):
            self.connection.execute(
                """INSERT INTO reconciliations
                   (id_periodo,concept_key,reference_billed_cents,calculated_billed_cents,
                    reference_actual_cents,calculated_actual_cents,difference_cents,status)
                   VALUES (?,?,?,?,?,?,?,'cuadrado')""",
                (self.period_id, key, billed, billed, actual, actual, 0),
            )

    def _validated_export(self):
        profile_id = self.connection.execute(
            """INSERT INTO excel_template_profiles
               (id_comunidad,profile_key,profile_version,template_relative_path,
                template_sha256,profile_sha256,status)
               VALUES (?,?,?,?,?,?, 'active')""",
            (
                self.community_id, "658_acs_v1", "1",
                "plantillas/comunidades/658/658_acs_v1.xlsx",
                "excel-template", hashlib.sha256(
                    (self.root / "config" / "excel_profiles" / "658_acs_v1.json").read_bytes()
                ).hexdigest(),
            ),
        ).lastrowid
        input_hash = calculate_case_input_hash(
            self.connection, id_case=self.case_id, project_root=self.root
        )
        self.connection.execute(
            """INSERT INTO excel_export_runs
               (id_case,id_periodo,id_template_profile,input_sha256,template_sha256,status)
               VALUES (?,?,?,?,?,'validated')""",
            (self.case_id, self.period_id, profile_id, input_hash, "excel-template"),
        )

    def _run_rows(self, run_id):
        return self.connection.execute(
            """SELECT id_propietario,status,output_path,error_message
               FROM generated_letters WHERE id_letter_run=? ORDER BY id_propietario""",
            (run_id,),
        ).fetchall()

    def test_case_letters_use_active_concepts_and_record_each_owner(self):
        from case_letter_service import generate_case_letters

        result = generate_case_letters(
            self.database_path, id_case=self.case_id, project_root=self.root
        )

        self.assertEqual(2, result.generated_count)
        self.assertEqual((), result.failures)
        self.assertTrue(result.output_path.is_dir())
        rows = self._run_rows(result.id_letter_run)
        self.assertEqual({"generated"}, {row["status"] for row in rows})
        self.assertEqual(2, len(list(result.output_path.glob("*.docx"))))
        self.assertEqual(
            "completed",
            self.connection.execute(
                "SELECT status FROM letter_generation_runs WHERE id_letter_run=?",
                (result.id_letter_run,),
            ).fetchone()[0],
        )

    def test_case_letters_only_accept_selected_concepts_from_the_active_profile(self):
        from case_letter_service import LetterGenerationBlockedError, generate_case_letters

        result = generate_case_letters(
            self.database_path,
            id_case=self.case_id,
            project_root=self.root,
            selected_concepts=("acs_fixed", "credit"),
        )
        document = next(result.output_path.glob("*.docx"))
        with zipfile.ZipFile(document) as archive:
            text = "\n".join(
                archive.read(name).decode("utf-8", errors="ignore")
                for name in archive.namelist()
                if name.startswith("word/") and name.endswith(".xml")
            )
        self.assertIn("Cuota fija de ACS", text)
        self.assertIn("Abono", text)
        self.assertNotIn("Consumo de ACS", text)

        with self.assertRaisesRegex(LetterGenerationBlockedError, "no está activo"):
            generate_case_letters(
                self.database_path,
                id_case=self.case_id,
                project_root=self.root,
                selected_concepts=("heating_fixed",),
            )

    def test_letter_content_uses_only_active_non_heating_concepts_and_portable_identity(self):
        from case_letter_service import generate_case_letters

        result = generate_case_letters(
            self.database_path, id_case=self.case_id, project_root=self.root
        )
        document = next(result.output_path.glob("*.docx"))
        with zipfile.ZipFile(document) as archive:
            text = "\n".join(
                archive.read(name).decode("utf-8", errors="ignore")
                for name in archive.namelist() if name.startswith("word/") and name.endswith(".xml")
            )
        self.assertIn("Cuota fija de ACS", text)
        self.assertIn("Consumo de ACS", text)
        self.assertIn("Abono", text)
        self.assertNotIn("Calefacci", text)
        self.assertIn("Gesti", text)
        self.assertIn("Valencia", text)
        self.assertNotIn("Meditrade", text)
        self.assertNotIn("Zaragoza", text)

    def test_case_letters_refuse_unreconciled_results_without_writing_files(self):
        from case_letter_service import LetterGenerationBlockedError, generate_case_letters

        self.connection.execute(
            "UPDATE reconciliations SET status='descuadrado' WHERE concept_key='acs_variable'"
        )
        self.connection.commit()

        with self.assertRaisesRegex(LetterGenerationBlockedError, "concili"):
            generate_case_letters(
                self.database_path, id_case=self.case_id, project_root=self.root
            )
        self.assertFalse((self.root / "salidas").exists())
        self.assertEqual(0, self.connection.execute("SELECT COUNT(*) FROM letter_generation_runs").fetchone()[0])

    def test_one_owner_failure_is_audited_without_removing_other_letter(self):
        from case_letter_service import generate_case_letters

        def fail_only_for_bruno(datos, ruta_plantilla, ruta_salida):
            if datos["vecino"]["nombre"] == "Bruno Vecino":
                raise OSError("plantilla dañada para la prueba")
            Path(ruta_salida).write_bytes(b"DOCX-SINTETICO")
            return str(ruta_salida)

        with patch("case_letter_service.generar_carta", side_effect=fail_only_for_bruno):
            result = generate_case_letters(
                self.database_path, id_case=self.case_id, project_root=self.root
            )

        self.assertEqual(1, result.generated_count)
        self.assertEqual(1, len(result.failures))
        self.assertIn("Bruno Vecino", result.failures[0])
        rows = self._run_rows(result.id_letter_run)
        self.assertEqual(["generated", "failed"], [row["status"] for row in rows])
        self.assertTrue(Path(rows[0]["output_path"]).is_file())
        self.assertIsNone(rows[1]["output_path"])
        self.assertEqual(
            "incomplete",
            self.connection.execute(
                "SELECT status FROM letter_generation_runs WHERE id_letter_run=?",
                (result.id_letter_run,),
            ).fetchone()[0],
        )

    def test_failed_write_removes_partial_destination_and_temporary_file(self):
        from case_letter_service import generate_case_letters

        def write_partial_then_fail(datos, ruta_plantilla, ruta_salida):
            Path(ruta_salida).write_bytes(b"DOCUMENTO-PARCIAL")
            if datos["vecino"]["nombre"] == "Bruno Vecino":
                raise OSError("fallo tras escribir el documento")
            return str(ruta_salida)

        with patch("case_letter_service.generar_carta", side_effect=write_partial_then_fail):
            result = generate_case_letters(
                self.database_path, id_case=self.case_id, project_root=self.root
            )

        failed = next(row for row in self._run_rows(result.id_letter_run) if row["status"] == "failed")
        self.assertIsNone(failed["output_path"])
        self.assertFalse((result.output_path / "CARTA_B-2_BRUNO_VECINO.docx").exists())
        self.assertEqual([], list(result.output_path.glob("*.tmp")))

    def test_completed_identical_run_is_reused_without_writing_or_auditing_again(self):
        from case_letter_service import generate_case_letters

        first = generate_case_letters(
            self.database_path, id_case=self.case_id, project_root=self.root
        )
        before_runs = self.connection.execute("SELECT COUNT(*) FROM letter_generation_runs").fetchone()[0]
        before_letters = self.connection.execute("SELECT COUNT(*) FROM generated_letters").fetchone()[0]

        with patch("case_letter_service.generar_carta", side_effect=AssertionError("no debe reescribir")):
            second = generate_case_letters(
                self.database_path, id_case=self.case_id, project_root=self.root
            )

        self.assertEqual(first.id_letter_run, second.id_letter_run)
        self.assertEqual(first.output_path, second.output_path)
        self.assertEqual(first.generated_count, second.generated_count)
        self.assertEqual((), second.failures)
        self.assertEqual(before_runs, self.connection.execute("SELECT COUNT(*) FROM letter_generation_runs").fetchone()[0])
        self.assertEqual(before_letters, self.connection.execute("SELECT COUNT(*) FROM generated_letters").fetchone()[0])

    def test_different_concept_selection_starts_a_new_batch(self):
        from case_letter_service import generate_case_letters

        first = generate_case_letters(
            self.database_path, id_case=self.case_id, project_root=self.root,
            selected_concepts=("acs_fixed",),
        )
        second = generate_case_letters(
            self.database_path, id_case=self.case_id, project_root=self.root,
            selected_concepts=("acs_variable",),
        )

        self.assertNotEqual(first.id_letter_run, second.id_letter_run)

    def test_rejects_letters_when_validated_profile_bytes_changed_without_version_bump(self):
        from case_letter_service import LetterGenerationBlockedError, generate_case_letters

        profile_path = self.root / "config" / "excel_profiles" / "658_acs_v1.json"
        profile_path.write_bytes(profile_path.read_bytes() + b"\n")

        with self.assertRaisesRegex(LetterGenerationBlockedError, "huella"):
            generate_case_letters(
                self.database_path, id_case=self.case_id, project_root=self.root
            )

    def test_changed_historical_graph_input_starts_a_new_batch(self):
        from case_letter_service import generate_case_letters

        historical_period = self.connection.execute(
            """INSERT INTO periodos(id_comunidad,nombre,fecha_inicio,fecha_fin)
               VALUES (?, '2024-2025', '2024-09-01', '2025-08-31')""",
            (self.community_id,),
        ).lastrowid
        self.connection.executemany(
            """INSERT INTO lecturas_vecino
               (id_propietario,id_periodo,tipo,fecha_lectura,valor_acumulado,estado)
               VALUES (?,?,'ACS',?,?,'real')""",
            [
                (self.owner_one, historical_period, "2024-09-01", 50),
                (self.owner_one, historical_period, "2025-08-31", 60),
            ],
        )
        self.connection.commit()
        first = generate_case_letters(
            self.database_path, id_case=self.case_id, project_root=self.root
        )
        self.connection.execute(
            """UPDATE lecturas_vecino SET valor_acumulado=65
               WHERE id_propietario=? AND id_periodo=? AND fecha_lectura='2025-08-31'""",
            (self.owner_one, historical_period),
        )
        self.connection.commit()

        second = generate_case_letters(
            self.database_path, id_case=self.case_id, project_root=self.root
        )

        self.assertNotEqual(first.id_letter_run, second.id_letter_run)

    def test_generated_letters_receive_shared_boundary_graphs(self):
        import case_letter_service

        previous_period = self.connection.execute(
            """INSERT INTO periodos(id_comunidad,nombre,fecha_inicio,fecha_fin)
               VALUES (?, 'Anterior compartido', '2025-08-01', '2025-09-01')""",
            (self.community_id,),
        ).lastrowid
        initial_readings = self.connection.execute(
            "SELECT id_lectura FROM lecturas_vecino WHERE id_periodo=? AND fecha_lectura='2025-09-01'",
            (self.period_id,),
        ).fetchall()
        for reading in initial_readings:
            self.connection.execute("UPDATE lecturas_vecino SET id_periodo=? WHERE id_lectura=?",
                                    (previous_period, reading[0]))
            self.connection.execute("INSERT INTO reading_periods(id_lectura,id_periodo) VALUES (?,?)",
                                    (reading[0], self.period_id))
        self.connection.executemany(
            """INSERT INTO lecturas_vecino
               (id_propietario,id_periodo,tipo,fecha_lectura,valor_acumulado,estado)
               VALUES (?,?,'ACS','2025-08-01',?,'real')""",
            [(self.owner_one, previous_period, 80), (self.owner_two, previous_period, 190)],
        )
        previous_case = self.connection.execute(
            """INSERT INTO regularization_cases
               (id_comunidad,nombre,fecha_inicio,fecha_fin,estado,id_periodo)
               VALUES (?, 'Anterior compartido', '2025-08-01', '2025-09-01', 'reconciled', ?)""",
            (self.community_id, previous_period),
        ).lastrowid
        current_period = self.period_id
        self.period_id = previous_period
        try:
            self._result_rows()
            self._reconcile_all()
        finally:
            self.period_id = current_period
        self.connection.execute(
            """INSERT INTO period_parameters
               (id_comunidad,id_periodo,parameter_key,numeric_value,unit)
               SELECT id_comunidad,?,parameter_key,numeric_value,unit
               FROM period_parameters WHERE id_periodo=?""",
            (previous_period, self.period_id),
        )
        previous_hash = calculate_case_input_hash(
            self.connection, id_case=previous_case, project_root=self.root,
        )
        self.connection.execute(
            """INSERT INTO excel_export_runs
               (id_case,id_periodo,id_template_profile,input_sha256,template_sha256,status)
               SELECT ?,?,id_template_profile,?,template_sha256,'validated'
               FROM excel_export_runs WHERE id_case=?""",
            (previous_case, previous_period, previous_hash, self.case_id),
        )
        self.connection.commit()

        # The real writer still creates the documents; inspect the payload at
        # its production call boundary, without substituting graph calculation.
        for case_id, other_period_name in (
            (previous_case, "2025-2026"), (self.case_id, "Anterior compartido"),
        ):
            with self.subTest(case_id=case_id), patch(
                "case_letter_service.generar_carta", wraps=case_letter_service.generar_carta,
            ) as writer:
                result = case_letter_service.generate_case_letters(
                    self.database_path, id_case=case_id, project_root=self.root,
                )
                self.assertEqual(2, result.generated_count)
                self.assertEqual((), result.failures)
                graphs = {call.args[0]["vecino"]["vivienda"]: call.args[0]["consumo_grafica"]
                          for call in writer.call_args_list}
                self.assertEqual({"A-1", "B-2"}, set(graphs))
                for dwelling, consumption in (("A-1", 20.0), ("B-2", 10.0)):
                    self.assertEqual(consumption, graphs[dwelling]["owner_consumption"])
                    self.assertEqual([20.0, 10.0], graphs[dwelling]["neighbor_consumptions"])
                    self.assertEqual([(other_period_name, consumption)], graphs[dwelling]["history"])
        self.assertEqual(6, self.connection.execute("SELECT COUNT(*) FROM lecturas_vecino").fetchone()[0])

    def test_negative_historical_consumption_is_omitted_not_replaced_with_zero(self):
        from case_letter_service import _consumption_graphs

        historical_period = self.connection.execute(
            """INSERT INTO periodos(id_comunidad,nombre,fecha_inicio,fecha_fin)
               VALUES (?, 'reinicio', '2024-09-01', '2025-08-31')""",
            (self.community_id,),
        ).lastrowid
        self.connection.executemany(
            """INSERT INTO lecturas_vecino
               (id_propietario,id_periodo,tipo,fecha_lectura,valor_acumulado,estado)
               VALUES (?,?,'ACS',?,?,'real')""",
            [
                (self.owner_one, historical_period, "2024-09-01", 90),
                (self.owner_one, historical_period, "2025-08-31", 10),
            ],
        )
        self.connection.commit()
        case = self.connection.execute(
            """SELECT c.*,co.codigo,co.nombre AS community_name,per.nombre AS period_name
               FROM regularization_cases c JOIN comunidades co ON co.id_comunidad=c.id_comunidad
               JOIN periodos per ON per.id_periodo=c.id_periodo WHERE c.id_case=?""",
            (self.case_id,),
        ).fetchone()

        graphs = _consumption_graphs(
            self.connection, case, [{"id_propietario": self.owner_one}]
        )

        self.assertEqual([], graphs[self.owner_one]["history"])

    def test_missing_file_prevents_reuse_and_starts_new_batch(self):
        from case_letter_service import generate_case_letters

        first = generate_case_letters(
            self.database_path, id_case=self.case_id, project_root=self.root
        )
        next(first.output_path.glob("*.docx")).unlink()

        second = generate_case_letters(
            self.database_path, id_case=self.case_id, project_root=self.root
        )

        self.assertNotEqual(first.id_letter_run, second.id_letter_run)
        self.assertEqual(2, second.generated_count)

    def test_audit_failure_after_publish_compensates_final_document(self):
        from case_letter_service import generate_case_letters

        with patch(
            "case_letter_service._mark_generated",
            side_effect=sqlite3.OperationalError("fallo al confirmar auditoría"),
        ):
            result = generate_case_letters(
                self.database_path, id_case=self.case_id, project_root=self.root
            )

        self.assertEqual(0, result.generated_count)
        self.assertEqual(2, len(result.failures))
        self.assertEqual([], list(result.output_path.glob("*.docx")))
        self.assertEqual([], list(result.output_path.glob("*.tmp")))
        self.assertEqual({"failed"}, {row["status"] for row in self._run_rows(result.id_letter_run)})

    def test_new_run_uses_its_own_directory_without_reusing_old_owner_document(self):
        from carta_writer import generar_carta as real_generate
        from case_letter_service import generate_case_letters

        first = generate_case_letters(
            self.database_path, id_case=self.case_id, project_root=self.root
        )
        old_bruno = first.output_path / "CARTA_B-2_BRUNO_VECINO.docx"
        self.assertTrue(old_bruno.is_file())
        (self.root / "config" / "letter_identities.json").write_text(
            '{"communities":{"658":{"office_name":"Nueva gestión",'
            '"footer":"Atención de la comunidad",'
            '"signature":"Equipo gestor", "city":"Valencia"}}}',
            encoding="utf-8",
        )

        def fail_only_for_bruno(datos, ruta_plantilla, ruta_salida):
            if datos["vecino"]["nombre"] == "Bruno Vecino":
                raise OSError("fallo de la segunda ejecución")
            return real_generate(datos, ruta_plantilla, ruta_salida)

        with patch("case_letter_service.generar_carta", side_effect=fail_only_for_bruno):
            second = generate_case_letters(
                self.database_path, id_case=self.case_id, project_root=self.root
            )

        self.assertNotEqual(first.id_letter_run, second.id_letter_run)
        self.assertNotEqual(first.output_path, second.output_path)
        self.assertTrue(old_bruno.is_file())
        self.assertFalse((second.output_path / "CARTA_B-2_BRUNO_VECINO.docx").exists())
        self.assertEqual(1, second.generated_count)


if __name__ == "__main__":
    unittest.main()
