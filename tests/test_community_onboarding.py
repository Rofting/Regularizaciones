import sqlite3
import sys
import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook


PROJECT_ROOT = Path(__file__).resolve().parents[1]
CORE_DIR = PROJECT_ROOT / "core"
if str(CORE_DIR) not in sys.path:
    sys.path.insert(0, str(CORE_DIR))

import community_onboarding


def make_readings_workbook(path: Path, headers: tuple[str, ...]) -> Path:
    workbook = Workbook()
    sheet = workbook.active
    for column, header in enumerate(headers, start=1):
        sheet.cell(1, column, header)
    sheet.append(("A-01", "100", "120"))
    workbook.save(path)
    return path


def make_pdf(path: Path, text: str) -> Path:
    payload = (
        b"%PDF-1.4\n"
        b"1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n"
        b"2 0 obj\n<< /Type /Pages /Kids [3 0 R] /Count 1 >>\nendobj\n"
        b"3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792]"
        b" /Contents 4 0 R /Resources << /Font << /F1 5 0 R >> >> >>\nendobj\n"
    )
    stream = f"BT /F1 12 Tf 72 720 Td ({text}) Tj ET".encode("latin-1")
    payload += b"4 0 obj\n<< /Length " + str(len(stream)).encode() + b" >>\nstream\n"
    payload += stream + b"\nendstream\nendobj\n"
    payload += b"5 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>\nendobj\n"
    payload += b"xref\n0 6\n0000000000 65535 f \n"
    offsets = []
    offset = len(b"%PDF-1.4\n")
    for object_number in range(1, 6):
        marker = f"{object_number} 0 obj".encode()
        offsets.append(payload.index(marker))
    payload += b"".join(f"{value:010d} 00000 n \n".encode() for value in offsets)
    payload += b"trailer\n<< /Size 6 /Root 1 0 R >>\nstartxref\n"
    payload += str(payload.index(b"xref")).encode() + b"\n%%EOF\n"
    path.write_bytes(payload)
    return path


class CommunityOnboardingTest(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory(ignore_cleanup_errors=True)
        self.project_root = Path(self.directory.name) / "project"
        self.project_root.mkdir()
        self.database = self.project_root / "gestion.db"
        self.connection = sqlite3.connect(self.database)
        self.connection.execute("CREATE TABLE audit (id INTEGER PRIMARY KEY)")
        self.connection.commit()
        self.owners_csv = self.project_root / "owners.csv"
        self.owners_csv.write_text(
            "Codigo;Nombre;Fdenominacion\n1;VECINO SINTETICO;A-01\n",
            encoding="utf-8",
        )
        self.readings_xlsx = make_readings_workbook(
            self.project_root / "readings.xlsx", ("Vivienda", "Inicial", "Final")
        )
        self.readings_pdf = make_pdf(
            self.project_root / "readings.pdf", "LECTURAS ACS CONTADOR"
        )
        self.invoice_pdf = make_pdf(
            self.project_root / "invoice.pdf", "FACTURA SINTETICA"
        )
        self.ambiguous_draft = community_onboarding.OnboardingDraft(
            community_code="900",
            community_name="Comunidad prueba",
            sources=(),
            detected_modules=("ACS",),
            questions=(community_onboarding.OnboardingQuestion(
                key="reading_column",
                prompt="Selecciona la columna de lectura.",
                candidates=("Lectura inicial", "Lectura final"),
                required=True,
            ),),
        )
        self.module_draft = community_onboarding.OnboardingDraft(
            community_code="900",
            community_name="Comunidad prueba",
            sources=(),
            detected_modules=("ACS", "CALEFACCION"),
            questions=(),
        )

    def tearDown(self):
        self.connection.close()
        self.directory.cleanup()

    def _row_count(self):
        return self.connection.execute("SELECT COUNT(*) FROM audit").fetchone()[0]

    def test_analyse_sources_classifies_owner_list_excel_readings_and_pdf_invoice_without_writes(self):
        draft = community_onboarding.analyse_sources(
            community_code="900", community_name="Comunidad prueba",
            owner_list_path=self.owners_csv, reading_paths=(self.readings_xlsx,),
            invoice_paths=(self.invoice_pdf,), project_root=self.project_root,
        )
        self.assertEqual(("owner_list", "meter_reading_excel", "invoice_pdf"),
                         tuple(item.kind for item in draft.sources))
        self.assertEqual(0, self._row_count())
        self.assertFalse((self.project_root / "config/excel_profiles/900_v1.json").exists())

    def test_analyse_sources_accepts_pdf_readings_as_an_alternative(self):
        draft = community_onboarding.analyse_sources(
            community_code="900", community_name="Comunidad prueba",
            owner_list_path=self.owners_csv, reading_paths=(self.readings_pdf,),
            invoice_paths=(), project_root=self.project_root,
        )
        self.assertEqual("meter_reading_pdf", draft.sources[1].kind)

    def test_analyse_sources_records_sha256_and_requires_confirmation_for_ambiguous_reading_columns(self):
        ambiguous_readings = make_readings_workbook(
            self.project_root / "ambiguous.xlsx", ("Vivienda", "Lectura ACS", "Lectura Agua")
        )
        draft = community_onboarding.analyse_sources(
            community_code="900", community_name="Comunidad prueba",
            owner_list_path=self.owners_csv, reading_paths=(ambiguous_readings,),
            invoice_paths=(), project_root=self.project_root,
        )
        reading = draft.sources[1]
        self.assertEqual(64, len(reading.sha256))
        self.assertTrue(any(question.required for question in draft.questions))

    def test_analyse_sources_does_not_treat_identical_reading_columns_as_ambiguous(self):
        second_readings = make_readings_workbook(
            self.project_root / "same-readings.xlsx", ("Vivienda", "Lectura ACS")
        )
        first_readings = make_readings_workbook(
            self.project_root / "first-readings.xlsx", ("Vivienda", "Lectura ACS")
        )
        draft = community_onboarding.analyse_sources(
            community_code="900", community_name="Comunidad prueba",
            owner_list_path=self.owners_csv,
            reading_paths=(first_readings, second_readings),
            invoice_paths=(), project_root=self.project_root,
        )
        self.assertFalse(any(question.key == "reading_column" for question in draft.questions))

    def test_analyse_sources_prioritizes_acs_over_partial_water_match_but_keeps_real_service_ambiguity(self):
        acs_readings = make_readings_workbook(
            self.project_root / "acs.xlsx", ("Vivienda", "Lectura Agua Caliente ACS")
        )
        acs_draft = community_onboarding.analyse_sources(
            community_code="900", community_name="Comunidad prueba",
            owner_list_path=self.owners_csv, reading_paths=(acs_readings,),
            invoice_paths=(), project_root=self.project_root,
        )
        self.assertEqual(("ACS",), acs_draft.detected_modules)
        self.assertFalse(any(question.key == "service" for question in acs_draft.questions))

        ambiguous_readings = make_readings_workbook(
            self.project_root / "two-services.xlsx", ("Vivienda", "Lectura ACS", "Lectura Calefacción")
        )
        ambiguous_draft = community_onboarding.analyse_sources(
            community_code="900", community_name="Comunidad prueba",
            owner_list_path=self.owners_csv, reading_paths=(ambiguous_readings,),
            invoice_paths=(), project_root=self.project_root,
        )
        self.assertEqual(("ACS", "CALEFACCION"), ambiguous_draft.detected_modules)
        self.assertTrue(any(question.key == "service" and question.required
                            for question in ambiguous_draft.questions))

    def test_profile_payload_requires_an_answer_for_ambiguous_reading_column(self):
        with self.assertRaisesRegex(ValueError, "columna de lectura"):
            community_onboarding.build_profile_payload(self.ambiguous_draft, answers={})

    def test_profile_payload_contains_only_confirmed_modules(self):
        payload = community_onboarding.build_profile_payload(
            self.module_draft,
            answers={"reading_column": "Lectura", "module:ACS": True,
                     "module:CALEFACCION": False},
        )
        self.assertIn("ACS", payload["active_modules"])
        self.assertNotIn("CALEFACCION", payload["active_modules"])


if __name__ == "__main__":
    unittest.main()
