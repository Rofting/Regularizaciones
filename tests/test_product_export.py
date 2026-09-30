import importlib.util
import sqlite3
import sys
import tempfile
import unittest
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
CORE_DIR = PROJECT_ROOT / "core"
if str(CORE_DIR) not in sys.path:
    sys.path.insert(0, str(CORE_DIR))

spec = importlib.util.spec_from_file_location(
    "exportar_producto", PROJECT_ROOT / "scripts" / "exportar_producto.py"
)
exportar_producto = importlib.util.module_from_spec(spec)
sys.modules["exportar_producto"] = exportar_producto
spec.loader.exec_module(exportar_producto)


class ProductExportTest(unittest.TestCase):
    def test_export_contains_the_product_and_none_of_this_office(self):
        with tempfile.TemporaryDirectory() as directory:
            destination = Path(directory) / "producto"
            copied = set(exportar_producto.copy_product(PROJECT_ROOT, destination))
            self.assertIn("core/app.py", copied)
            self.assertIn("config/proveedores.json", copied)
            self.assertIn("plantillas/modelo/modelo_acs_v1.xlsx", copied)
            for private in ("config/proveedores_despacho.json", "config/letter_identities.json",
                            "core/private_658_validation.py"):
                self.assertNotIn(private, copied)
            self.assertFalse(any(path.startswith(("data/", "config/excel_profiles/", "tests/"))
                                 for path in copied))
            # Guardia permanente: el producto no debe llevar datos de nadie.
            self.assertEqual([], exportar_producto.scan_for_leaks(destination))

    def test_scan_detects_community_person_and_supply_identifiers(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            (root / "nota.txt").write_text(
                "Comunidad H-12.345.673? no; real H99258139, vecino 12345678Z, "
                "CUPS ES0031300718354001RR0F, ejemplo H12345674", encoding="utf-8")
            database = root / "g.db"
            connection = sqlite3.connect(database)
            connection.execute("CREATE TABLE comunidades (nombre TEXT, cif TEXT)")
            connection.execute("CREATE TABLE propietarios (nombre_propietario TEXT)")
            connection.execute("INSERT INTO comunidades VALUES ('CP LAS ACACIAS', NULL)")
            connection.execute("INSERT INTO propietarios VALUES ('ANA PEREZ GOMEZ')")
            connection.commit()
            connection.close()
            (root / "otra.json").write_text('{"x": "cp las acacias - Ana Perez Gomez"}', encoding="utf-8")
            kinds = {(leak.kind, leak.value) for leak in exportar_producto.scan_for_leaks(root, database)}
        self.assertIn(("CIF de comunidad", "H99258139"), kinds)
        self.assertIn(("NIF/NIE de persona", "12345678Z"), kinds)
        self.assertIn(("CUPS", "ES0031300718354001RR0F"), kinds)
        self.assertIn(("nombre de comunidad", "CP LAS ACACIAS"), kinds)
        self.assertIn(("propietario", "ANA PEREZ GOMEZ"), kinds)
        self.assertNotIn(("CIF de comunidad", "H12345674"), kinds)


if __name__ == "__main__":
    unittest.main()
