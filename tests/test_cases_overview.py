"""Mejora 15: panel de expedientes de todas las comunidades."""

import sqlite3
import sys
import tempfile
import unittest
from contextlib import redirect_stdout
from datetime import date
from io import StringIO
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "core"))

import cases_overview as overview
import expedient_service
import gestor_bd


class CasesOverviewTest(unittest.TestCase):
    def setUp(self):
        directory = tempfile.TemporaryDirectory(ignore_cleanup_errors=True)
        self.addCleanup(directory.cleanup)
        database = Path(directory.name) / "g.db"
        with redirect_stdout(StringIO()):
            gestor_bd.crear_bd(str(database))
        self.connection = sqlite3.connect(database)
        self.connection.row_factory = sqlite3.Row
        self.addCleanup(self.connection.close)
        self.cases = {}
        for code, name in (("658", "CP Las Flores"), ("701", "CP Mirador Ñandú"), ("702", "CP Inactiva")):
            community = gestor_bd.obtener_o_crear_comunidad(self.connection, code, name)
            for year in (2025, 2026):
                case = expedient_service.create_case(
                    self.connection, community, name=f"Año {year}",
                    start_date=date(year, 1, 1), end_date=date(year, 12, 31))
                self.cases[(code, year)] = case.id_case
        self.connection.execute("UPDATE comunidades SET activa=0 WHERE codigo='702'")
        self.connection.commit()

    def status(self, key, value):
        self.connection.execute("UPDATE regularization_cases SET estado=? WHERE id_case=?", (value, self.cases[key]))

    def issue(self, key, number):
        self.connection.execute("PRAGMA foreign_keys=OFF")
        document = self.connection.execute(
            "INSERT INTO source_documents(id_case,original_name,archived_path,sha256,document_kind,status) "
            "VALUES (?,?,?,?, 'invoice','registered')",
            (self.cases[key], f"f{number}.pdf", f"/x/f{number}.pdf", f"h{number}")).lastrowid
        self.connection.execute(
            "INSERT INTO review_issues(id_case,id_document,code,field_name,message,status) VALUES (?,?,?,?,?,'open')",
            (self.cases[key], document, "MISSING", "importe_total", "Falta"))

    def test_buckets_issues_and_last_output(self):
        self.issue(("658", 2026), 1)
        self.issue(("658", 2026), 2)
        self.status(("658", 2025), "deliveries_generated")
        self.status(("701", 2026), "ready_for_calculation")
        self.status(("701", 2025), "reconciled")
        self.connection.execute("PRAGMA foreign_keys=OFF")
        excel = self.connection.execute(
            "INSERT INTO excel_export_runs(id_case,id_periodo,id_template_profile,input_sha256,template_sha256,"
            "output_path,status,published_at) VALUES (?,1,1,'i','t','/s/658.xlsx','published','2026-02-01 10:00:00')",
            (self.cases[("658", 2025)],)).lastrowid
        self.connection.execute(
            "INSERT INTO letter_generation_runs(id_case,id_periodo,id_export_run,input_sha256,template_sha256,"
            "output_path,status,completed_at) VALUES (?,1,?,'i','t','/s/cartas','completed','2026-02-03 09:30:00')",
            (self.cases[("658", 2025)], excel))
        self.connection.commit()

        cases = overview.list_cases(self.connection)
        self.assertEqual(4, len(cases))  # la comunidad inactiva no aparece
        by_key = {(case.community_code, case.start_date.year): case for case in cases}
        self.assertEqual(overview.BLOCKED, by_key[("658", 2026)].bucket)
        self.assertEqual(2, by_key[("658", 2026)].open_issues)
        self.assertEqual(overview.LETTERS, by_key[("658", 2025)].bucket)
        self.assertEqual("Cartas · 03/02/2026 09:30", by_key[("658", 2025)].last_output_label)
        self.assertEqual(overview.READY, by_key[("701", 2026)].bucket)
        self.assertEqual(overview.CALCULATED, by_key[("701", 2025)].bucket)
        self.assertEqual("Sin salidas", by_key[("701", 2026)].last_output_label)
        self.assertEqual({"blocked": 1, "pending": 0, "ready": 1, "calculated": 1, "letters": 1, "closed": 0},
                         overview.bucket_counts(cases))
        # El más reciente primero dentro de cada comunidad.
        self.assertEqual([2026, 2025], [c.start_date.year for c in cases if c.community_code == "658"])

    def test_filters(self):
        self.issue(("701", 2025), 3)
        self.connection.commit()
        cases = overview.list_cases(self.connection)
        self.assertEqual({"701"}, {c.community_code for c in overview.filter_cases(cases, text="nandu")})
        self.assertEqual(2, len(overview.filter_cases(cases, text="año 2026")))
        self.assertEqual(2, len(overview.filter_cases(cases, latest_only=True)))
        blocked = overview.filter_cases(cases, bucket=overview.BLOCKED)
        self.assertEqual([("701", 2025)], [(c.community_code, c.start_date.year) for c in blocked])
        self.assertEqual((), overview.filter_cases(cases, bucket=overview.BLOCKED, latest_only=True))

    def test_resolved_issues_do_not_block_and_closed_wins(self):
        self.issue(("658", 2026), 4)
        self.connection.execute("UPDATE review_issues SET status='resolved'")
        self.status(("701", 2026), "closed")
        self.issue(("701", 2026), 5)
        self.connection.commit()
        by_key = {(c.community_code, c.start_date.year): c for c in overview.list_cases(self.connection)}
        self.assertEqual(overview.PENDING, by_key[("658", 2026)].bucket)
        self.assertEqual(overview.CLOSED, by_key[("701", 2026)].bucket)


class OpenCaseTest(unittest.TestCase):
    """«Abrir» activa la comunidad y el expediente elegidos, no el primero."""

    setUp = CasesOverviewTest.setUp

    def make_app(self):
        import types
        from unittest.mock import Mock
        import app as app_module

        class Var:
            def __init__(self):
                self.value = ""

            def set(self, value):
                self.value = value

            def get(self):
                return self.value

        fake = types.SimpleNamespace(
            _procesando=False, id_comunidad=None, id_expediente=None, id_periodo=None,
            ruta_bd_expedientes=self.connection.execute("PRAGMA database_list").fetchone()[2],
            comunidad_actual=Var(), expediente_seleccionado=Var(), expediente_actual=Var(),
            cb_comunidad=Mock(), cb_expediente=Mock(), log=Mock(), _refrescar_expediente=Mock(),
        )
        fake._ids_comunidad = {
            f"{row[1]} — {row[2]}": row[0]
            for row in self.connection.execute("SELECT id_comunidad,codigo,nombre FROM comunidades WHERE activa=1")
        }
        fake._limpiar_contexto_expediente = lambda: setattr(fake, "id_expediente", None)
        fake._cargar_comunidades = Mock()
        fake._refrescar_lista_expedientes = types.MethodType(
            app_module.AppGestionFincas._refrescar_lista_expedientes, fake)
        fake.abrir = types.MethodType(app_module.AppGestionFincas.abrir_expediente, fake)
        return fake

    def test_opens_the_chosen_older_case(self):
        fake = self.make_app()
        community = self.connection.execute("SELECT id_comunidad FROM comunidades WHERE codigo='701'").fetchone()[0]
        older = self.cases[("701", 2025)]
        self.assertTrue(fake.abrir(community, older))
        self.assertEqual((community, older), (fake.id_comunidad, fake.id_expediente))
        self.assertEqual("701 — CP Mirador Ñandú", fake.comunidad_actual.get())
        self.assertTrue(fake.expediente_seleccionado.get().startswith("Año 2025"))
        fake._refrescar_expediente.assert_called_once()

    def test_refuses_a_case_from_another_community(self):
        fake = self.make_app()
        community = self.connection.execute("SELECT id_comunidad FROM comunidades WHERE codigo='701'").fetchone()[0]
        self.assertFalse(fake.abrir(community, self.cases[("658", 2025)]))
        self.assertEqual(community, fake.id_comunidad)
        self.assertNotEqual(self.cases[("658", 2025)], fake.id_expediente)
        fake.log.assert_called()

    def test_inactive_community_is_not_opened(self):
        fake = self.make_app()
        community = self.connection.execute("SELECT id_comunidad FROM comunidades WHERE codigo='702'").fetchone()[0]
        self.assertFalse(fake.abrir(community, self.cases[("702", 2026)]))


if __name__ == "__main__":
    unittest.main()
