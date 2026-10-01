import sys
import tempfile
import unittest
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
CORE_DIR = PROJECT_ROOT / "core"
if str(CORE_DIR) not in sys.path:
    sys.path.insert(0, str(CORE_DIR))

from community_folder import community_name_from_text, propose_from_folder

INVOICE = """FACTURA Nº F-77 Fecha factura 05/02/2026
Cliente: COMUNIDAD DE PROPIETARIOS LAS ACACIAS 12   CIF H12345674
Base imponible 100,00 IVA 21,00 Total factura 121,00 €"""
QUOTE = "PRESUPUESTO Nº 4 para la comunidad. Total 900,00 €"
READING = """Informe de lecturas de agua caliente sanitaria
Periodo de lectura: del 01/01/2025 al 31/12/2025
Vivienda    Lectura anterior    Lectura actual
1A          10                  15
1B          20                  26
"""


class CommunityFolderTest(unittest.TestCase):
    def test_folder_is_classified_and_identity_is_deduced(self):
        from openpyxl import Workbook
        with tempfile.TemporaryDirectory() as directory:
            folder = Path(directory) / "731"
            folder.mkdir()
            book = Workbook()
            for row in (["Piso", "Titular", "Coeficiente"], ["1A", "Ana", 50], ["1B", "Luis", 50]):
                book.active.append(row)
            book.save(folder / "propietarios.xlsx")
            texts = {"factura_gas.pdf": INVOICE, "presupuesto.pdf": QUOTE, "lecturas.pdf": READING}
            for name in texts:
                (folder / name).write_bytes(b"%PDF sintetico")
            (folder / "~$temporal.xlsx").write_bytes(b"x")
            proposal = propose_from_folder(folder, text_reader=lambda path: texts[path.name])
        self.assertEqual("731", proposal.code)
        self.assertEqual("LAS ACACIAS 12", proposal.name)
        self.assertEqual("H12345674", proposal.cif)
        self.assertEqual("propietarios.xlsx", proposal.owners.name)
        self.assertEqual(["lecturas.pdf"], [path.name for path in proposal.readings])
        self.assertEqual(["factura_gas.pdf"], [path.name for path in proposal.invoices])
        self.assertEqual([("presupuesto.pdf", "presupuesto")],
                         [(path.name, reason) for path, reason in proposal.ignored])

    def test_community_name_patterns(self):
        for text, expected in (
            ("C.P. Residencial Sol 4 - CIF H12345674", "RESIDENCIAL SOL 4"),
            ("Cdad. Prop. Calle Mayor 10\nOtra línea", "CALLE MAYOR 10"),
            ("Comunidad de Propietarios: Edificio Luna  NIF H1", "EDIFICIO LUNA"),
        ):
            with self.subTest(text=text):
                self.assertEqual(expected, community_name_from_text(text))
        self.assertIsNone(community_name_from_text("Factura sin comunidad"))


if __name__ == "__main__":
    unittest.main()
