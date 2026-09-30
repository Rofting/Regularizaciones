import sys
import tempfile
import unittest
from contextlib import redirect_stdout
from io import StringIO
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
CORE_DIR = PROJECT_ROOT / "core"
if str(CORE_DIR) not in sys.path:
    sys.path.insert(0, str(CORE_DIR))

import gestor_bd
import office_settings
from letter_settings import load_community_letter_identity
from office_settings import OfficeSettings


class OfficeSettingsTest(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory(ignore_cleanup_errors=True)
        self.root = Path(self.directory.name)
        with redirect_stdout(StringIO()):
            gestor_bd.crear_bd(str(self.root / "data" / "g.db"))
        self.connection = gestor_bd.conectar(str(self.root / "data" / "g.db"))

    def tearDown(self):
        self.connection.close()
        self.directory.cleanup()

    def test_new_installation_starts_neutral_and_unconfigured(self):
        settings = office_settings.load_office_settings(self.connection)
        self.assertFalse(settings.configured)
        self.assertEqual("", settings.name)

    def test_saved_settings_round_trip_and_normalise_tax_id(self):
        office_settings.save_office_settings(self.connection, OfficeSettings(
            name=" Fincas Norte ", tax_id="b-50.000.009", city="Bilbao", fiscal_start_month=9,
        ))
        loaded = office_settings.load_office_settings(self.connection)
        self.assertEqual(("Fincas Norte", "B50000009", 9), (loaded.name, loaded.tax_id, loaded.fiscal_start_month))

    def test_invalid_settings_are_rejected(self):
        for settings in (OfficeSettings(name=""), OfficeSettings(name="X", tax_id="B50000000"),
                         OfficeSettings(name="X", fiscal_start_month=13)):
            with self.subTest(settings=settings), self.assertRaises(ValueError):
                office_settings.save_office_settings(self.connection, settings)

    def test_fiscal_month_is_inferred_from_existing_cases(self):
        community = gestor_bd.obtener_o_crear_comunidad(self.connection, "1", "C")
        for start in ("2024-09-01", "2025-09-01", "2026-01-01"):
            self.connection.execute(
                "INSERT INTO regularization_cases(id_comunidad,nombre,fecha_inicio,fecha_fin,estado) VALUES (?,?,?,?, 'draft')",
                (community, start, start, start),
            )
        self.assertEqual(9, office_settings.suggested_settings(self.connection).fiscal_start_month)

    def test_letters_use_office_identity_unless_a_community_overrides_it(self):
        office = OfficeSettings(name="Fincas Norte", city="Bilbao", address="Gran Vía 1",
                                phone="944000000")
        identity = load_community_letter_identity(self.root, "7", office=office)
        self.assertEqual(("Fincas Norte", "Bilbao", "Fincas Norte"),
                         (identity.office_name, identity.city, identity.signature))
        self.assertEqual("Gran Vía 1 · 944000000", identity.footer)
        (self.root / "config").mkdir()
        (self.root / "config" / "letter_identities.json").write_text(
            '{"communities":{"7":{"signature":"Presidencia"}}}', encoding="utf-8")
        overridden = load_community_letter_identity(self.root, "7", office=office)
        self.assertEqual(("Fincas Norte", "Presidencia"), (overridden.office_name, overridden.signature))

    def test_logo_is_copied_inside_the_installation(self):
        source = self.root / "mi_logo.PNG"
        source.write_bytes(b"\x89PNG fake")
        relative = office_settings.install_logo(self.root, source)
        self.assertEqual("data/despacho/logo.png", relative)
        resolved = office_settings.resolve_logo(self.root, OfficeSettings(name="X", logo_path=relative))
        self.assertTrue(resolved.is_file())
        with self.assertRaises(ValueError):
            office_settings.install_logo(self.root, self.root / "logo.gif")


if __name__ == "__main__":
    unittest.main()
