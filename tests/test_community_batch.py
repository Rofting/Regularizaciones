"""Mejora 13: alta masiva de comunidades desde subcarpetas."""

import sqlite3
import sys
import tempfile
import unittest
from contextlib import redirect_stdout
from datetime import date
from io import StringIO
from pathlib import Path

from openpyxl import Workbook

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "core"))

import community_batch
import gestor_bd


def _no_text(_path):
    return ""


class CommunityBatchTest(unittest.TestCase):
    def setUp(self):
        directory = tempfile.TemporaryDirectory(ignore_cleanup_errors=True)
        self.addCleanup(directory.cleanup)
        self.base = Path(directory.name).resolve()
        self.project = self.base / "proyecto"
        self.project.mkdir()
        self.batch = self.base / "lote"
        database = self.base / "gestion.db"
        with redirect_stdout(StringIO()):
            gestor_bd.crear_bd(str(database))
        self.connection = sqlite3.connect(database)
        self.connection.row_factory = sqlite3.Row
        self.addCleanup(self.connection.close)

    def community(self, folder, *, heating=False, owners=True):
        path = self.batch / folder
        path.mkdir(parents=True)
        if owners:
            (path / "propietarios.csv").write_text(
                "Fdenominacion;Nombre;Coeficiente\n1º A;Vecino uno;50\n1º B;Vecino dos;50\n",
                encoding="utf-8")
        workbook = Workbook()
        sheet = workbook.active
        sheet.append(("Lecturas de calefacción" if heating else "Lecturas de agua caliente sanitaria ACS",))
        sheet.append(("Vivienda", "Nº contador", "Lectura 01/01/2026", "Lectura 31/12/2026", "Consumo"))
        sheet.append(("1º A", "C-11", 100, 130, 30))
        sheet.append(("1º B", "C-12", 50, 60, 10))
        workbook.save(path / "lecturas.xlsx")
        workbook.close()
        return path

    def scan(self):
        return {entry.folder.name: entry for entry in community_batch.scan_batch(
            self.batch, connection=self.connection, project_root=self.project,
            work_directory=self.base / "importaciones", text_reader=_no_text)}

    def create(self, entries):
        return community_batch.create_batch(
            self.connection, entries.values(), project_root=self.project,
            archive_root=self.base / "expedientes", actor="Prueba")

    def test_folder_name_gives_code_and_name(self):
        self.assertEqual(("658", "CP LAS FLORES"), community_batch.code_and_name_from_folder("658 - CP Las Flores"))
        self.assertEqual(("658", None), community_batch.code_and_name_from_folder("658"))
        self.assertEqual((None, None), community_batch.code_and_name_from_folder("Varios"))

    def test_company_report_is_converted_and_the_summary_shows_a_sample(self):
        self.community("901 - CP Uno")
        entry = self.scan()["901 - CP Uno"]
        self.assertEqual(community_batch.READY, entry.status, entry.reasons)
        self.assertEqual((date(2026, 1, 1), date(2026, 12, 31)), (entry.start, entry.end))
        self.assertEqual("Año 2026", entry.period_name)
        self.assertIn("ACS: contador C-11, 31/12/2026, lectura 130", entry.summary)
        self.assertEqual("lecturas_lecturas.xlsx", entry.readings[0].name)

    def test_creates_each_community_with_owners_readings_and_case(self):
        self.community("901 - CP Uno")
        self.community("902 - CP Dos", heating=True)
        outcomes = self.create(self.scan())
        self.assertEqual(["created", "created"], [item.status for item in outcomes])
        self.assertEqual([(2, 4), (2, 4)], [(item.owner_count, item.reading_count) for item in outcomes])
        rows = self.connection.execute(
            """SELECT c.codigo, l.tipo, p.codigo_vivienda, l.fecha_lectura, l.valor_acumulado
               FROM lecturas_vecino l JOIN propietarios p USING(id_propietario)
               JOIN comunidades c ON c.id_comunidad=p.id_comunidad
               WHERE p.codigo_vivienda='1º A' ORDER BY 1, 4""").fetchall()
        self.assertEqual([("901", "ACS", "1º A", "2026-01-01", 100), ("901", "ACS", "1º A", "2026-12-31", 130),
                          ("902", "CALEFACCION", "1º A", "2026-01-01", 100),
                          ("902", "CALEFACCION", "1º A", "2026-12-31", 130)],
                         [tuple(row) for row in rows])
        self.assertEqual(2, self.connection.execute("SELECT count(*) FROM regularization_cases").fetchone()[0])

    def test_repeating_the_batch_does_not_duplicate(self):
        self.community("901 - CP Uno")
        self.create(self.scan())
        entries = self.scan()
        self.assertEqual(community_batch.EXISTING, entries["901 - CP Uno"].status)
        self.assertEqual((), self.create(entries))
        self.assertEqual(1, self.connection.execute("SELECT count(*) FROM comunidades").fetchone()[0])
        self.assertEqual(2, self.connection.execute("SELECT count(*) FROM propietarios").fetchone()[0])

    def test_incomplete_and_repeated_folders_wait_for_review(self):
        self.community("901 - CP Uno", owners=False)
        self.community("902 - CP Dos")
        self.community("902 - Duplicada")
        (self.batch / "Varios").mkdir()
        entries = self.scan()
        self.assertIn("listado de propietarios", " ".join(entries["901 - CP Uno"].reasons))
        for name in ("902 - CP Dos", "902 - Duplicada"):
            self.assertEqual(community_batch.REVIEW, entries[name].status)
            self.assertIn("varias carpetas", " ".join(entries[name].reasons))
        self.assertIn("no contiene documentos", " ".join(entries["Varios"].reasons))
        self.assertEqual((), self.create(entries))

    def test_a_failure_keeps_its_diagnosis_and_the_rest_continue(self):
        broken = self.community("901 - CP Uno")
        self.community("902 - CP Dos")
        entries = self.scan()
        # El listado cambia entre la revisión y la creación: su huella ya no coincide.
        (self.base / "importaciones").mkdir(exist_ok=True)
        owner_source = next(s.path for s in entries["901 - CP Uno"].draft.sources if s.kind == "owner_list")
        owner_source.write_text(owner_source.read_text(encoding="utf-8") + "2º A;Otro;0\n", encoding="utf-8")
        self.assertTrue(broken.is_dir())
        outcomes = {item.code: item for item in self.create(entries)}
        self.assertEqual("failed", outcomes["901"].status)
        self.assertTrue(outcomes["901"].message.startswith("No se pudo crear"))
        self.assertEqual("created", outcomes["902"].status)
        codes = [row[0] for row in self.connection.execute("SELECT codigo FROM comunidades")]
        self.assertEqual(["902"], codes)
        self.assertFalse((self.project / "plantillas" / "comunidades" / "901").exists())


if __name__ == "__main__":
    unittest.main()
