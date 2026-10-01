import csv
import sys
import tempfile
import unittest
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
CORE_DIR = PROJECT_ROOT / "core"
if str(CORE_DIR) not in sys.path:
    sys.path.insert(0, str(CORE_DIR))

from importar_propietarios_csv import _decimal
from owner_lists import is_canonical, normalise_owner_list, read_owner_rows


class OwnerListsTest(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.root = Path(self.directory.name)

    def tearDown(self):
        self.directory.cleanup()

    def test_excel_list_with_other_headers_is_converted_to_the_canonical_csv(self):
        from openpyxl import Workbook
        path = self.root / "Listado.xlsx"
        book = Workbook()
        sheet = book.active
        for row in (["Comunidad de propietarios Ejemplo"], [],
                    ["Piso", "Titular", "Participación (%)", "Correo electrónico"],
                    ["1º A", "Ana Pérez", 2.5, "ana@example.com"],
                    ["1º B", "Luis Gil", 3, None], ["TOTAL", None, 5.5, None]):
            sheet.append(row)
        book.save(path)
        canonical = normalise_owner_list(path, self.root / "staging")
        self.assertTrue(is_canonical(canonical))
        with canonical.open(encoding="utf-8") as handle:
            rows = list(csv.reader(handle, delimiter=";"))
        self.assertEqual(["Fdenominacion", "Nombre", "Coeficiente", "Email"], rows[0])
        self.assertEqual(["1º A", "Ana Pérez", "2,5", "ana@example.com"], rows[1])
        self.assertEqual(2.5, _decimal(rows[1][2]))  # el importador no lo convierte en 25

    def test_comma_csv_and_email_on_the_next_line(self):
        path = self.root / "propietarios.csv"
        path.write_text("Vivienda,Nombre,Coeficiente\n1A,Ana,2.50\n,,\n,ana@x.es,\n1B,Luis,3\n", encoding="cp1252")
        rows = read_owner_rows(path)
        self.assertEqual([("1A", "Ana", "2,50", "ana@x.es"), ("1B", "Luis", "3", "")],
                         [(r.vivienda, r.nombre, r.coeficiente, r.email) for r in rows])

    def test_canonical_files_are_used_as_they_are(self):
        path = self.root / "canon.csv"
        path.write_text("Fdenominacion;Nombre;Coeficiente;Email\n1A;Ana;1;\n", encoding="utf-8")
        self.assertEqual(path, normalise_owner_list(path, self.root / "staging"))

    def test_unrecognised_lists_explain_what_is_missing(self):
        path = self.root / "otra.csv"
        path.write_text("Fecha;Importe\n1;2\n", encoding="utf-8")
        with self.assertRaisesRegex(ValueError, "vivienda y nombre"):
            read_owner_rows(path)


if __name__ == "__main__":
    unittest.main()
