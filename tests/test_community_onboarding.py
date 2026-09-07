import json
import hashlib
import os
import sqlite3
import sys
import tempfile
import unittest
from contextlib import redirect_stdout
from datetime import date
from io import StringIO
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook, load_workbook


PROJECT_ROOT = Path(__file__).resolve().parents[1]
CORE_DIR = PROJECT_ROOT / "core"
if str(CORE_DIR) not in sys.path:
    sys.path.insert(0, str(CORE_DIR))

import community_onboarding
import case_workflow_actions
import document_review
import expedient_service
import excel_generator
import excel_profiles
import gestor_bd
from case_distribution import calculate_case_distribution
from excel_export_service import generate_official_excel
from office_recalculation import DeterministicRecalculator


def make_readings_workbook(
    path: Path,
    headers: tuple[str, ...],
    rows: tuple[tuple[str, ...], ...] = (("A-01", "100", "120"),),
) -> Path:
    workbook = Workbook()
    sheet = workbook.active
    for column, header in enumerate(headers, start=1):
        sheet.cell(1, column, header)
    for row in rows:
        sheet.append(row)
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
        with redirect_stdout(StringIO()):
            gestor_bd.crear_bd(str(self.database))
        self.connection = sqlite3.connect(self.database)
        self.connection.row_factory = sqlite3.Row
        self.connection.execute("PRAGMA foreign_keys = ON")
        self.connection.execute("CREATE TABLE audit (id INTEGER PRIMARY KEY)")
        self.connection.commit()
        self.archive_root = self.project_root / "expedientes"
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
        self.readings_with_fields_pdf = make_pdf(
            self.project_root / "readings-with-fields.pdf",
            (
                "CONTADOR: C-01; COLUMNA: Lectura ACS; FECHA: 2026-01-01; "
                "LECTURA: 100; SERVICIO: ACS"
            ),
        )
        self.readings_with_multiple_values_pdf = make_pdf(
            self.project_root / "readings-with-multiple-values.pdf",
            (
                "CONTADOR: C-01; COLUMNA: Lectura ACS; FECHA: 2026-01-01; "
                "LECTURA: 100; SERVICIO: ACS; LECTURA: 101"
            ),
        )
        self.acs_reading_pdf = make_pdf(
            self.project_root / "acs-reading.pdf",
            (
                "CONTADOR: C-ACS; COLUMNA: Lectura ACS; FECHA: 2026-01-01; "
                "LECTURA: 100; SERVICIO: ACS"
            ),
        )
        self.heating_reading_without_meter_or_date_pdf = make_pdf(
            self.project_root / "heating-reading-without-meter-or-date.pdf",
            "COLUMNA: Lectura Calefaccion; LECTURA: 200; SERVICIO: CALEFACCION",
        )
        self.invoice_pdf = make_pdf(
            self.project_root / "invoice.pdf", "FACTURA SINTETICA"
        )
        self.readings_without_labels = make_readings_workbook(
            self.project_root / "readings-without-labels.xlsx",
            ("Vivienda", "Consumo"),
        )
        self.heating_invoice = make_pdf(
            self.project_root / "heating-invoice.pdf",
            (
                "PROVEEDOR: ENERGIA NORTE; PERIODO: 2025-2026; "
                "IMPORTE: 123,45 EUR; CONCEPTO: CALEFACCION"
            ),
        )
        self.invoice_without_period = make_pdf(
            self.project_root / "invoice-without-period.pdf",
            "PROVEEDOR: ENERGIA NORTE; IMPORTE: 123,45 EUR; CONCEPTO: ACS",
        )
        self.readings_without_value = make_readings_workbook(
            self.project_root / "readings-without-value.xlsx",
            ("Vivienda", "Contador ACS", "Fecha", "Lectura ACS"),
            rows=(),
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
            sources=(community_onboarding.SourceCandidate(
                self.readings_xlsx, "meter_reading_excel",
                community_onboarding._sha256(self.readings_xlsx),
                headers=("Vivienda", "Lectura"),
            ),),
            detected_modules=("ACS", "CALEFACCION"),
            questions=(community_onboarding.OnboardingQuestion(
                key="service",
                prompt="Confirma el servicio de las lecturas.",
                candidates=("ACS", "CALEFACCION", "ACS+CALEFACCION", "NO_APLICA"),
                required=True,
            ),),
            reading_evidence=(community_onboarding.ReadingEvidence(
                source_sha256=community_onboarding._sha256(self.readings_xlsx),
                meters=("Contador confirmado",),
                columns=("Lectura",),
                dates=("2026-01-01",),
                readings=("100",),
                services=("ACS",),
            ),),
        )
        self.acs_heating_draft = community_onboarding.OnboardingDraft(
            community_code="900",
            community_name="Comunidad prueba",
            sources=(community_onboarding.SourceCandidate(
                self.readings_xlsx, "meter_reading_excel",
                community_onboarding._sha256(self.readings_xlsx),
                headers=("Vivienda", "Lectura ACS", "Lectura Calefacción"),
            ),),
            detected_modules=("ACS", "CALEFACCION"),
            questions=(
                community_onboarding.OnboardingQuestion(
                    key="service",
                    prompt="Confirma los servicios de las lecturas.",
                    candidates=("ACS", "CALEFACCION", "ACS+CALEFACCION", "NO_APLICA"),
                    required=True,
                ),
                community_onboarding.OnboardingQuestion(
                    key="reading_column:ACS",
                    prompt="Confirma la columna de lectura de ACS.",
                    candidates=("Lectura ACS", "Lectura Calefacción"),
                    required=True,
                ),
                community_onboarding.OnboardingQuestion(
                    key="reading_column:CALEFACCION",
                    prompt="Confirma la columna de lectura de calefacción.",
                    candidates=("Lectura ACS", "Lectura Calefacción"),
                    required=True,
                ),
            ),
        )
        self.valid_draft = community_onboarding.OnboardingDraft(
            community_code="900",
            community_name="Comunidad prueba",
            sources=(
                community_onboarding.SourceCandidate(
                    self.owners_csv, "owner_list",
                    community_onboarding._sha256(self.owners_csv),
                ),
                community_onboarding.SourceCandidate(
                    self.readings_xlsx, "meter_reading_excel",
                    community_onboarding._sha256(self.readings_xlsx),
                    headers=("Vivienda", "Lectura"),
                ),
                community_onboarding.SourceCandidate(
                    self.invoice_pdf, "invoice_pdf",
                    community_onboarding._sha256(self.invoice_pdf),
                ),
            ),
            detected_modules=("ACS",),
            questions=(community_onboarding.OnboardingQuestion(
                key="service",
                prompt="Confirma el servicio de las lecturas.",
                candidates=("ACS", "NO_APLICA"),
                required=True,
            ),),
            invoice_evidence=(community_onboarding.InvoiceEvidence(
                source_sha256=community_onboarding._sha256(self.invoice_pdf),
                providers=("Proveedor confirmado",),
                periods=("2026",),
                amounts=("100,00 EUR",),
                concepts=("ACS",),
            ),),
            reading_evidence=(community_onboarding.ReadingEvidence(
                source_sha256=community_onboarding._sha256(self.readings_xlsx),
                meters=("Contador confirmado",),
                columns=("Lectura",),
                dates=("2026-01-01",),
                readings=("100",),
                services=("ACS",),
            ),),
        )
        self.answers = {"service": "ACS"}

    def tearDown(self):
        self.connection.close()
        self.directory.cleanup()

    def _row_count(self):
        return self.connection.execute("SELECT COUNT(*) FROM audit").fetchone()[0]

    def _combined_answers(self) -> dict[str, str]:
        return {
            "service": "ACS+CALEFACCION",
            "reading_column:ACS": "Lectura ACS",
            "reading_meter:ACS": "Contador ACS",
            "reading_date:ACS": "2026-12-31",
            "reading_value:ACS": "130",
            "reading_column:CALEFACCION": "Lectura Calefacción",
            "reading_meter:CALEFACCION": "Contador Calefacción",
            "reading_date:CALEFACCION": "2026-12-31",
            "reading_value:CALEFACCION": "500",
        }

    def _combined_payload_with_layout(self, code: str = "900") -> dict[str, object]:
        draft = community_onboarding.OnboardingDraft(
            community_code=code,
            community_name="Comunidad combinada",
            sources=self.acs_heating_draft.sources,
            detected_modules=self.acs_heating_draft.detected_modules,
            questions=self.acs_heating_draft.questions,
        )
        configuration = community_onboarding.resolve_onboarding_configuration(
            draft, self._combined_answers()
        )
        payload = community_onboarding.build_profile_payload(draft, configuration)
        payload["workbook_layout"] = excel_generator.create_canonical_community_template(
            self.project_root / f"{code}-canonical.xlsx",
            community_code=code,
            community_name=draft.community_name,
            active_modules=configuration.active_modules,
        )
        return payload

    def _community_count(self):
        return self.connection.execute(
            "SELECT COUNT(*) FROM comunidades WHERE codigo='900'"
        ).fetchone()[0]

    def _registered_source_count(self):
        return self.connection.execute(
            "SELECT COUNT(*) FROM source_documents"
        ).fetchone()[0]

    def _confirmation_arguments(self):
        return {
            "draft": self.valid_draft,
            "answers": self.answers,
            "project_root": self.project_root,
            "archive_root": self.archive_root,
            "period_name": "2026",
            "start_date": date(2026, 1, 1),
            "end_date": date(2026, 12, 31),
            "actor": "Prueba",
        }

    def test_analyse_sources_classifies_owner_list_excel_readings_and_pdf_invoice_without_writes(self):
        draft = community_onboarding.analyse_sources(
            community_code="900", community_name="Comunidad prueba",
            owner_list_path=self.owners_csv, reading_paths=(self.readings_xlsx,),
            invoice_paths=(self.invoice_pdf,), project_root=self.project_root,
        )
        self.assertEqual(("owner_list", "meter_reading_excel", "invoice_pdf"),
                         tuple(item.kind for item in draft.sources))
        self.assertEqual(0, self._row_count())
        self.assertFalse((self.project_root / "config/excel_profiles/runtime/900_v1.json").exists())

    def test_analyse_sources_accepts_pdf_readings_as_an_alternative(self):
        draft = community_onboarding.analyse_sources(
            community_code="900", community_name="Comunidad prueba",
            owner_list_path=self.owners_csv, reading_paths=(self.readings_pdf,),
            invoice_paths=(), project_root=self.project_root,
        )
        self.assertEqual("meter_reading_pdf", draft.sources[1].kind)

    def test_pdf_reading_uses_its_own_fields_and_never_exposes_amounts(self):
        draft = community_onboarding.analyse_sources(
            community_code="900", community_name="Comunidad prueba",
            owner_list_path=self.owners_csv,
            reading_paths=(self.readings_with_fields_pdf,),
            invoice_paths=(), project_root=self.project_root,
        )

        evidence = draft.reading_evidence[0]
        self.assertEqual(("C-01",), evidence.meters)
        self.assertEqual(("Lectura ACS",), evidence.columns)
        self.assertEqual(("2026-01-01",), evidence.dates)
        self.assertEqual(("100",), evidence.readings)
        self.assertFalse(hasattr(evidence, "amounts"))
        self.assertFalse(any(
            question.key == "reading_column:ACS" for question in draft.questions
        ))
        configuration = community_onboarding.resolve_onboarding_configuration(
            draft, {"service": "ACS"}
        )
        self.assertEqual("Lectura ACS", configuration.reading_bindings[0].column)

    def test_multiple_reading_values_require_a_per_module_decision_and_persist_it(self):
        draft = community_onboarding.analyse_sources(
            community_code="900", community_name="Comunidad prueba",
            owner_list_path=self.owners_csv,
            reading_paths=(self.readings_with_multiple_values_pdf,),
            invoice_paths=(), project_root=self.project_root,
        )

        questions = {question.key: question for question in draft.questions}
        self.assertTrue(questions["reading_value:ACS"].required)
        with self.assertRaisesRegex(ValueError, "valor de lectura|reading_value:ACS"):
            community_onboarding.resolve_onboarding_configuration(
                draft,
                {
                    "service": "ACS",
                    "reading_column:ACS": "Lectura ACS",
                },
            )
        with self.assertRaises(ValueError):
            community_onboarding.resolve_onboarding_configuration(
                draft,
                {
                    "service": "ACS",
                    "reading_column:ACS": "Lectura ACS",
                    "reading_value:ACS": "999",
                },
            )

        configuration = community_onboarding.resolve_onboarding_configuration(
            draft,
            {
                "service": "ACS",
                "reading_column:ACS": "Lectura ACS",
                "reading_value:ACS": "100",
            },
        )

        self.assertEqual("100", configuration.reading_bindings[0].value)

    def test_missing_reading_value_blocks_publication_for_the_active_module(self):
        draft = community_onboarding.analyse_sources(
            community_code="900", community_name="Comunidad prueba",
            owner_list_path=self.owners_csv,
            reading_paths=(self.readings_without_value,),
            invoice_paths=(), project_root=self.project_root,
        )

        questions = {question.key: question for question in draft.questions}
        self.assertTrue(questions["reading_value:ACS"].required)
        with self.assertRaisesRegex(ValueError, "valor de lectura|reading_value:ACS"):
            community_onboarding.build_profile_payload(
                draft,
                answers={
                    "service": "ACS",
                    "reading_column:ACS": "Lectura ACS",
                    "reading_meter:ACS": "Contador ACS",
                    "reading_date:ACS": "Fecha",
                },
            )

    def test_meter_and_date_questions_are_scoped_to_each_module(self):
        draft = community_onboarding.analyse_sources(
            community_code="900", community_name="Comunidad prueba",
            owner_list_path=self.owners_csv,
            reading_paths=(
                self.acs_reading_pdf,
                self.heating_reading_without_meter_or_date_pdf,
            ),
            invoice_paths=(), project_root=self.project_root,
        )

        required_keys = {
            question.key for question in draft.questions if question.required
        }
        self.assertIn("reading_meter:CALEFACCION", required_keys)
        self.assertIn("reading_date:CALEFACCION", required_keys)
        with self.assertRaisesRegex(ValueError, "contador.*CALEFACCION"):
            community_onboarding.resolve_onboarding_configuration(
                draft,
                {
                    "service": "ACS+CALEFACCION",
                    "reading_column:ACS": "Lectura ACS",
                    "reading_column:CALEFACCION": "Lectura Calefaccion",
                },
            )

        configuration = community_onboarding.resolve_onboarding_configuration(
            draft,
            {
                "service": "ACS+CALEFACCION",
                "reading_column:ACS": "Lectura ACS",
                "reading_column:CALEFACCION": "Lectura Calefaccion",
                "reading_meter:CALEFACCION": "C-CALEFACCION",
                "reading_date:CALEFACCION": "2026-01-01",
            },
        )
        heating_binding = configuration.reading_bindings[1]
        self.assertEqual("C-CALEFACCION", heating_binding.meter)
        self.assertEqual("2026-01-01", heating_binding.date)

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
        self.assertTrue(any(question.key == "service" and question.required
                            for question in acs_draft.questions))

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

    def test_invoice_contributes_service_but_never_supplies_a_reading_column(self):
        readings = make_readings_workbook(
            self.project_root / "neutral-readings.xlsx", ("Vivienda", "Consumo")
        )
        invoice = make_pdf(
            self.project_root / "heating-invoice.pdf",
            "FACTURA CALEFACCION LECTURA CONTADOR",
        )

        draft = community_onboarding.analyse_sources(
            community_code="900", community_name="Comunidad prueba",
            owner_list_path=self.owners_csv, reading_paths=(readings,),
            invoice_paths=(invoice,), project_root=self.project_root,
        )

        self.assertEqual(("CALEFACCION",), draft.detected_modules)
        self.assertEqual("invoice_pdf", draft.sources[-1].kind)
        self.assertTrue(any(
            question.key == "reading_column:CALEFACCION" and question.required
            for question in draft.questions
        ))

    def test_invoice_detects_heating_but_is_not_a_meter_reading(self):
        draft = community_onboarding.analyse_sources(
            community_code="900", community_name="Comunidad prueba",
            owner_list_path=self.owners_csv,
            reading_paths=(self.readings_without_labels,),
            invoice_paths=(self.heating_invoice,), project_root=self.project_root,
        )

        self.assertIn("CALEFACCION", draft.detected_modules)
        self.assertFalse(any(
            source.kind == "invoice_pdf" and source.headers
            for source in draft.sources
        ))
        self.assertTrue(
            hasattr(draft, "invoice_evidence"),
            "El borrador debe conservar evidencia de factura separada",
        )
        self.assertEqual(("ENERGIA NORTE",), draft.invoice_evidence[0].providers)
        self.assertEqual(("2025-2026",), draft.invoice_evidence[0].periods)
        self.assertEqual(("123,45 EUR",), draft.invoice_evidence[0].amounts)
        self.assertEqual(("CALEFACCION",), draft.invoice_evidence[0].concepts)
        configuration = community_onboarding.resolve_onboarding_configuration(
            draft,
            {
                "service": "CALEFACCION",
                "reading_column:CALEFACCION": "Consumo",
                "reading_meter:CALEFACCION": "Contador confirmado",
                "reading_date:CALEFACCION": "Fecha confirmada",
                "reading_value:CALEFACCION": "100",
            },
        )
        self.assertEqual(
            ("ENERGIA NORTE", "2025-2026", "123,45 EUR", "CALEFACCION"),
            (
                configuration.invoice_decisions[0].provider,
                configuration.invoice_decisions[0].period,
                configuration.invoice_decisions[0].amount,
                configuration.invoice_decisions[0].concept,
            ),
        )

    def test_combined_modules_require_a_confirmed_reading_for_each_module(self):
        with self.assertRaisesRegex(ValueError, "columna de lectura"):
            community_onboarding.resolve_onboarding_configuration(
                self.acs_heating_draft,
                {
                    "service": "ACS+CALEFACCION",
                    "reading_column:ACS": "Lectura ACS",
                },
            )
        try:
            configuration = community_onboarding.resolve_onboarding_configuration(
                self.acs_heating_draft,
                {
                    "service": "ACS+CALEFACCION",
                    "reading_column:ACS": "Lectura ACS",
                    "reading_meter:ACS": "Contador ACS",
                    "reading_date:ACS": "2026-01-01",
                    "reading_value:ACS": "100",
                    "reading_column:CALEFACCION": "Lectura Calefacción",
                    "reading_meter:CALEFACCION": "Contador Calefacción",
                    "reading_date:CALEFACCION": "2026-01-01",
                    "reading_value:CALEFACCION": "200",
                },
            )
        except ValueError as error:
            self.fail(f"La configuración combinada debe resolverse por módulo: {error}")

        self.assertEqual(
            ("Lectura ACS", "Lectura Calefacción"),
            tuple(binding.column for binding in configuration.reading_bindings),
        )

    def test_no_aplica_is_rejected_for_active_and_traced_for_inactive_module(self):
        with self.assertRaisesRegex(ValueError, "NO_APLICA"):
            community_onboarding.resolve_onboarding_configuration(
                self.acs_heating_draft,
                {
                    "service": "ACS",
                    "reading_column:ACS": "NO_APLICA",
                },
            )

        configuration = community_onboarding.resolve_onboarding_configuration(
            self.acs_heating_draft,
            {
                "service": "ACS",
                "reading_column:ACS": "Lectura ACS",
                "reading_meter:ACS": "Contador ACS",
                "reading_date:ACS": "2026-01-01",
                "reading_value:ACS": "100",
                "reading_column:CALEFACCION": "NO_APLICA",
            },
        )
        self.assertTrue(
            hasattr(configuration, "not_applicable_modules"),
            "La configuración debe trazar módulos confirmados como NO_APLICA",
        )
        self.assertEqual(
            ("CALEFACCION",), configuration.not_applicable_modules
        )

    def test_missing_invoice_period_and_missing_reading_create_required_questions(self):
        draft = community_onboarding.analyse_sources(
            community_code="900", community_name="Comunidad prueba",
            owner_list_path=self.owners_csv,
            reading_paths=(self.readings_without_value,),
            invoice_paths=(self.invoice_without_period,), project_root=self.project_root,
        )

        required_keys = {
            question.key for question in draft.questions if question.required
        }
        self.assertTrue(
            {"invoice_period", "reading_column:ACS"}.issubset(required_keys)
        )
        self.assertTrue(all(
            question.required for question in draft.questions
            if question.key in {"invoice_period", "reading_column:ACS"}
        ))

    def test_absent_service_and_reading_column_require_explicit_decisions(self):
        readings = make_readings_workbook(
            self.project_root / "undetected-readings.xlsx", ("Vivienda", "Consumo")
        )

        draft = community_onboarding.analyse_sources(
            community_code="900", community_name="Comunidad prueba",
            owner_list_path=self.owners_csv, reading_paths=(readings,),
            invoice_paths=(), project_root=self.project_root,
        )

        questions = {question.key: question for question in draft.questions}
        self.assertTrue(questions["service"].required)
        self.assertTrue(questions["reading_column:ACS"].required)
        self.assertTrue(questions["reading_column:CALEFACCION"].required)
        with self.assertRaisesRegex(ValueError, "servicio"):
            community_onboarding.build_profile_payload(
                draft, answers={"reading_column:ACS": "Consumo"}
            )
        payload = community_onboarding.build_profile_payload(
            draft,
            answers={"service": "NO_APLICA"},
        )
        self.assertEqual([], payload["active_modules"])
        self.assertEqual(
            "NO_APLICA",
            payload["onboarding_configuration"]["service_decision"],
        )

    def test_service_and_column_answers_are_the_single_payload_configuration(self):
        readings = make_readings_workbook(
            self.project_root / "ambiguous-services.xlsx",
            ("Vivienda", "Lectura ACS", "Lectura Calefacción"),
        )
        draft = community_onboarding.analyse_sources(
            community_code="900", community_name="Comunidad prueba",
            owner_list_path=self.owners_csv, reading_paths=(readings,),
            invoice_paths=(), project_root=self.project_root,
        )

        acs = community_onboarding.build_profile_payload(
            draft,
            answers={
                "service": "ACS",
                "reading_column:ACS": "Lectura ACS",
                "reading_meter:ACS": "Contador confirmado",
                "reading_date:ACS": "Fecha confirmada",
                "reading_value:ACS": "100",
            },
        )
        heating = community_onboarding.build_profile_payload(
            draft,
            answers={
                "service": "CALEFACCION",
                "reading_column:CALEFACCION": "Lectura Calefacción",
                "reading_meter:CALEFACCION": "Contador confirmado",
                "reading_date:CALEFACCION": "Fecha confirmada",
                "reading_value:CALEFACCION": "120",
            },
        )

        self.assertEqual(["ACS"], acs["active_modules"])
        self.assertEqual(["acs_fixed", "acs_variable"], [
            concept["key"] for concept in acs["concepts"]
        ])
        self.assertEqual("Lectura ACS", acs["onboarding_configuration"]["reading_column"])
        self.assertEqual(["CALEFACCION"], heating["active_modules"])
        self.assertEqual(["heating_fixed", "heating_variable"], [
            concept["key"] for concept in heating["concepts"]
        ])
        self.assertEqual(
            "Lectura Calefacción",
            heating["onboarding_configuration"]["reading_bindings"][0]["column"],
        )
        self.assertEqual(
            "Contador confirmado",
            heating["onboarding_configuration"]["reading_bindings"][0]["meter"],
        )

    def test_a_detected_service_can_be_corrected_before_building_the_profile(self):
        readings = make_readings_workbook(
            self.project_root / "detected-acs.xlsx", ("Vivienda", "Lectura ACS")
        )
        draft = community_onboarding.analyse_sources(
            community_code="900", community_name="Comunidad prueba",
            owner_list_path=self.owners_csv, reading_paths=(readings,),
            invoice_paths=(), project_root=self.project_root,
        )

        payload = community_onboarding.build_profile_payload(
            draft,
            answers={
                "service": "CALEFACCION",
                "reading_column:CALEFACCION": "Lectura ACS",
                "reading_meter:CALEFACCION": "Contador confirmado",
                "reading_date:CALEFACCION": "2026-01-01",
                "reading_value:CALEFACCION": "100",
            },
        )

        self.assertEqual(["CALEFACCION"], payload["active_modules"])
        self.assertEqual(
            ["heating_fixed", "heating_variable"],
            [concept["key"] for concept in payload["concepts"]],
        )

    def test_profile_payload_requires_an_answer_for_ambiguous_reading_column(self):
        with self.assertRaisesRegex(ValueError, "columna de lectura"):
            community_onboarding.build_profile_payload(self.ambiguous_draft, answers={})

    def test_profile_payload_rejects_empty_false_or_unknown_required_answers(self):
        for answer in ("", False, "Otra columna"):
            with self.subTest(answer=answer):
                with self.assertRaisesRegex(ValueError, "columna de lectura"):
                    community_onboarding.build_profile_payload(
                        self.ambiguous_draft, answers={"reading_column": answer}
                    )

    def test_profile_payload_contains_only_confirmed_modules(self):
        payload = community_onboarding.build_profile_payload(
            self.module_draft,
            answers={"service": "ACS"},
        )
        self.assertIn("ACS", payload["active_modules"])
        self.assertNotIn("CALEFACCION", payload["active_modules"])

    def test_profile_validation_rejects_no_aplica_with_an_active_module(self):
        payload = community_onboarding.build_profile_payload(
            self.valid_draft, answers={"service": "ACS"}
        )
        payload["onboarding_configuration"]["service_decision"] = "NO_APLICA"

        with self.assertRaisesRegex(ValueError, "service_decision"):
            excel_profiles.validate_profile_payload(payload, self.project_root)

    def test_profile_validation_rejects_new_configuration_stripped_to_look_historical(self):
        payload = community_onboarding.build_profile_payload(
            self.valid_draft, answers={"service": "ACS"}
        )
        configuration = payload["onboarding_configuration"]
        configuration.pop("schema_version")
        configuration.pop("invoice_decisions")
        configuration.pop("not_applicable_modules")
        for binding in configuration["reading_bindings"]:
            binding.pop("meter")
            binding.pop("date")
            binding.pop("value")

        self.assertTrue(any(
            source["kind"] == "invoice_pdf"
            for source in configuration["sources"]
        ))
        with self.assertRaisesRegex(ValueError, "schema_version"):
            excel_profiles.validate_profile_payload(payload, self.project_root)

    def test_profile_validation_rejects_incomplete_new_invoice_configuration(self):
        payload = community_onboarding.build_profile_payload(
            self.valid_draft, answers={"service": "ACS"}
        )
        payload["onboarding_configuration"]["invoice_decisions"] = []

        with self.assertRaisesRegex(ValueError, "factura"):
            excel_profiles.validate_profile_payload(payload, self.project_root)

    def test_profile_validation_rejects_empty_new_reading_decisions(self):
        payload = community_onboarding.build_profile_payload(
            self.valid_draft, answers={"service": "ACS"}
        )
        payload["onboarding_configuration"]["reading_bindings"][0]["value"] = ""

        with self.assertRaisesRegex(ValueError, "Binding de lectura"):
            excel_profiles.validate_profile_payload(payload, self.project_root)

    def test_profile_validation_requires_nonempty_layout_for_onboarding(self):
        for layout in ("missing", "empty"):
            with self.subTest(layout=layout):
                payload = community_onboarding.build_profile_payload(
                    self.valid_draft, answers={"service": "ACS"}
                )
                if layout == "empty":
                    payload["workbook_layout"] = {}

                with self.assertRaisesRegex(
                    ValueError, "(?=.*onboarding)(?=.*workbook_layout)"
                ):
                    excel_profiles.validate_profile_payload(payload, self.project_root)

    def test_profile_validation_rejects_cross_module_canonical_concept_source(self):
        payload = self._combined_payload_with_layout()
        heating_variable = next(
            concept for concept in payload["concepts"]
            if concept["key"] == "heating_variable"
        )
        heating_variable["actual_source"] = "period_parameters.acs_variable_actual"

        with self.assertRaisesRegex(ValueError, "heating_variable"):
            excel_profiles.validate_profile_payload(payload, self.project_root)

    def test_profile_validation_rejects_shared_combined_parameter_binding(self):
        payload = self._combined_payload_with_layout()
        parameter_cells = payload["workbook_layout"]["parameter_cells"]
        parameter_cells["heating_variable_actual"] = parameter_cells[
            "acs_variable_actual"
        ]

        with self.assertRaisesRegex(ValueError, "heating_variable_actual"):
            excel_profiles.validate_profile_payload(payload, self.project_root)

    def test_profile_validation_rejects_final_onboarding_layout_without_bootstrap_marker(self):
        payload = self._combined_payload_with_layout()
        del payload["workbook_layout"]["bootstrap_template"]

        with self.assertRaisesRegex(ValueError, "bootstrap"):
            excel_profiles.validate_profile_payload(payload, self.project_root)

    def test_combined_onboarding_profile_has_canonical_concepts_and_bindings(self):
        payload = self._combined_payload_with_layout()
        profile = excel_profiles.validate_profile_payload(payload, self.project_root)

        self.assertEqual(
            {
                "acs_fixed": (
                    "equal", "period_parameters.acs_fixed_actual",
                    "period_parameters.acs_fixed_billed",
                ),
                "acs_variable": (
                    "consumption", "period_parameters.acs_variable_actual",
                    "period_parameters.acs_variable_billed",
                ),
                "heating_fixed": (
                    "equal", "period_parameters.heating_fixed_actual",
                    "period_parameters.heating_fixed_billed",
                ),
                "heating_variable": (
                    "consumption", "period_parameters.heating_variable_actual",
                    "period_parameters.heating_variable_billed",
                ),
            },
            {
                concept.key: (
                    concept.allocation_method,
                    concept.actual_source,
                    concept.billed_source,
                )
                for concept in profile.concepts
            },
        )
        self.assertEqual(
            {
                "acs_variable_actual": ("LECTURAS ACS M3", "H8"),
                "acs_fixed_actual": ("LECTURAS ACS M3", "I9"),
                "acs_variable_billed": ("LECTURAS ACS M3", "H4"),
                "acs_fixed_billed": ("LECTURAS ACS M3", "I4"),
                "heating_variable_actual": ("LECTURAS CALEF KWH", "H8"),
                "heating_fixed_actual": ("LECTURAS CALEF KWH", "I9"),
                "heating_variable_billed": ("LECTURAS CALEF KWH", "H4"),
                "heating_fixed_billed": ("LECTURAS CALEF KWH", "I4"),
            },
            {
                key: tuple(binding)
                for key, binding in profile.workbook_layout["parameter_cells"].items()
            },
        )
        self.assertEqual(
            {"ACS": "Lectura ACS", "CALEFACCION": "Lectura Calefacción"},
            {
                binding["module"]: binding["column"]
                for binding in profile.onboarding_configuration["reading_bindings"]
            },
        )
        marker = profile.workbook_layout["bootstrap_template"]
        template = self.project_root / "900-canonical.xlsx"
        self.assertEqual("fresh_onboarding", marker["state"])
        self.assertEqual(
            hashlib.sha256(template.read_bytes()).hexdigest(), marker["sha256"]
        )

    def test_profile_payload_accepts_safe_community_codes(self):
        draft = community_onboarding.OnboardingDraft(
            community_code="644", community_name="Comunidad prueba", sources=(),
            detected_modules=(), questions=(),
        )

        self.assertEqual("644_v1", community_onboarding.build_profile_payload(
            draft, answers={"service": "NO_APLICA"}
        )["key"])

    def test_profile_payload_rejects_unsafe_windows_community_codes(self):
        for community_code in (
            "900/archivo", "900\\archivo", "900:", "900*", "900?", "900<",
            "900>", "900|", '900"', "900.", "900 ", "CON", "PRN", "AUX",
            "NUL", "COM1", "COM9", "LPT1", "LPT9", "CON.txt", "COM1.csv",
            "LPT9.foo", "900\x00", "900\x1f",
        ):
            draft = community_onboarding.OnboardingDraft(
                community_code=community_code, community_name="Comunidad prueba",
                sources=(), detected_modules=(), questions=(),
            )
            with self.subTest(community_code=community_code):
                with self.assertRaisesRegex(ValueError, "código de comunidad"):
                    community_onboarding.build_profile_payload(
                        draft, answers={"service": "NO_APLICA"}
                    )

    def _exercise_meter_service_flow(self, service: str, code: str) -> tuple[str, ...]:
        reading_label = "Lectura ACS" if service == "ACS" else "Lectura Calefacción"
        readings = make_readings_workbook(
            self.project_root / f"{code}-readings.xlsx", ("Vivienda", reading_label)
        )
        draft = community_onboarding.analyse_sources(
            community_code=code, community_name=f"Comunidad {service}",
            owner_list_path=self.owners_csv, reading_paths=(readings,),
            invoice_paths=(), project_root=self.project_root,
        )
        result = community_onboarding.confirm_onboarding(
            self.connection,
            draft=draft,
            answers={
                "service": service,
                f"reading_meter:{service}": "Contador confirmado",
                f"reading_date:{service}": "Fecha confirmada",
            },
            project_root=self.project_root,
            archive_root=self.archive_root,
            period_name="2026",
            start_date=date(2026, 1, 1),
            end_date=date(2026, 12, 31),
            actor="Prueba",
        )
        owner = self.connection.execute(
            """INSERT INTO propietarios
               (id_comunidad,codigo_vivienda,nombre_propietario,coeficiente,activo)
               VALUES (?,?,?,?,1)""",
            (result.community_id, "A-01", "VECINO SINTETICO", 100),
        )
        self.connection.executemany(
            """INSERT INTO lecturas_vecino
               (id_propietario,id_periodo,tipo,fecha_lectura,valor_acumulado,estado)
               VALUES (?,?,?,?,?,'real')""",
            (
                (owner.lastrowid, result.period_id, service, "2026-01-01", 100),
                (owner.lastrowid, result.period_id, service, "2026-12-31", 125),
            ),
        )
        prefix = "acs" if service == "ACS" else "heating"
        persisted = json.loads(result.profile_path.read_text(encoding="utf-8"))
        self.assertEqual(
            {
                f"{prefix}_fixed_actual", f"{prefix}_fixed_billed",
                f"{prefix}_variable_actual", f"{prefix}_variable_billed",
            },
            set(persisted["workbook_layout"]["parameter_cells"]),
        )
        for key, value in (
            (f"{prefix}_fixed_actual", 20),
            (f"{prefix}_fixed_billed", 10),
            (f"{prefix}_variable_actual", 80),
            (f"{prefix}_variable_billed", 40),
        ):
            self.connection.execute(
                """INSERT INTO period_parameters
                   (id_comunidad,id_periodo,parameter_key,numeric_value,unit)
                   VALUES (?,?,?,?, 'EUR')""",
                (result.community_id, result.period_id, key, value),
            )
        self.connection.commit()
        document_review.validate_case_ready(self.connection, result.case_id)

        exported = generate_official_excel(
            self.connection,
            id_case=result.case_id,
            project_root=self.project_root,
            output_root=self.project_root / "salidas",
            recalculator=DeterministicRecalculator(),
        )
        self.assertTrue(exported.output_path.is_file())
        distributed = calculate_case_distribution(
            self.connection, id_case=result.case_id, project_root=self.project_root
        )
        return tuple(distributed.concept_totals_cents)

    def test_acs_onboarding_profile_reaches_canonical_export_and_distribution(self):
        self.assertEqual(
            ("acs_fixed", "acs_variable"),
            self._exercise_meter_service_flow("ACS", "901"),
        )

    def test_heating_onboarding_profile_reaches_canonical_export_and_distribution(self):
        self.assertEqual(
            ("heating_fixed", "heating_variable"),
            self._exercise_meter_service_flow("CALEFACCION", "902"),
        )

    def test_combined_onboarding_exports_and_distributes_distinct_module_readings(self):
        code = "903"
        draft = community_onboarding.OnboardingDraft(
            community_code=code,
            community_name="Comunidad combinada",
            sources=self.acs_heating_draft.sources,
            detected_modules=self.acs_heating_draft.detected_modules,
            questions=self.acs_heating_draft.questions,
        )
        result = community_onboarding.confirm_onboarding(
            self.connection,
            draft=draft,
            answers=self._combined_answers(),
            project_root=self.project_root,
            archive_root=self.archive_root,
            period_name="2026",
            start_date=date(2026, 1, 1),
            end_date=date(2026, 12, 31),
            actor="Prueba",
        )
        bootstrap, companions = case_workflow_actions.run_bootstrap_import(
            self.database,
            id_case=result.case_id,
            active_community_id=result.community_id,
            project_root=self.project_root,
            master_path=result.template_path,
            actor="Prueba",
        )
        self.assertIsNone(companions)
        self.assertEqual(0, bootstrap.open_issue_count)

        owner = self.connection.execute(
            """INSERT INTO propietarios
               (id_comunidad,codigo_vivienda,nombre_propietario,coeficiente,activo)
               VALUES (?,?,?,?,1)""",
            (result.community_id, "A-01", "VECINO SINTETICO", 100),
        )
        self.connection.executemany(
            """INSERT INTO lecturas_vecino
               (id_propietario,id_periodo,tipo,fecha_lectura,valor_acumulado,estado)
               VALUES (?,?,?,?,?,'real')""",
            (
                (owner.lastrowid, result.period_id, "ACS", "2026-01-01", 100),
                (owner.lastrowid, result.period_id, "ACS", "2026-12-31", 130),
                (owner.lastrowid, result.period_id, "CALEFACCION", "2026-01-01", 200),
                (owner.lastrowid, result.period_id, "CALEFACCION", "2026-12-31", 500),
            ),
        )
        for key, value in (
            ("acs_fixed_actual", 20), ("acs_fixed_billed", 10),
            ("acs_variable_actual", 80), ("acs_variable_billed", 40),
            ("heating_fixed_actual", 50), ("heating_fixed_billed", 25),
            ("heating_variable_actual", 120), ("heating_variable_billed", 60),
        ):
            self.connection.execute(
                """INSERT INTO period_parameters
                   (id_comunidad,id_periodo,parameter_key,numeric_value,unit)
                   VALUES (?,?,?,?, 'EUR')""",
                (result.community_id, result.period_id, key, value),
            )
        self.connection.commit()

        exported = generate_official_excel(
            self.connection,
            id_case=result.case_id,
            project_root=self.project_root,
            output_root=self.project_root / "salidas",
            recalculator=DeterministicRecalculator(),
        )
        distributed = calculate_case_distribution(
            self.connection, id_case=result.case_id, project_root=self.project_root
        )

        self.assertTrue(exported.output_path.is_file())
        self.assertEqual(
            {"acs_fixed", "acs_variable", "heating_fixed", "heating_variable"},
            set(distributed.concept_totals_cents),
        )
        self.assertEqual(
            {"acs_variable": 30, "heating_variable": 300},
            {
                row["concept_key"]: row["consumption"]
                for row in self.connection.execute(
                    """SELECT concept_key,consumption FROM owner_concept_results
                       WHERE id_periodo=? AND concept_key IN (?,?)""",
                    (result.period_id, "acs_variable", "heating_variable"),
                )
            },
        )

    def test_confirm_onboarding_writes_profile_template_archives_sources_and_creates_case(self):
        result = community_onboarding.confirm_onboarding(
            self.connection, **self._confirmation_arguments()
        )

        profile_path = self.project_root / "config/excel_profiles/runtime/900_v1.json"
        template_path = self.project_root / "plantillas/comunidades/900/900_v1.xlsx"
        self.assertIsNotNone(result.case_id)
        self.assertTrue(profile_path.is_file())
        self.assertTrue(template_path.is_file())
        self.assertEqual(3, self._registered_source_count())
        payload = json.loads(profile_path.read_text(encoding="utf-8"))
        workbook = load_workbook(template_path, read_only=True)
        try:
            self.assertEqual(
                ["DATOS", "LECTURAS ACS M3", "ANALISIS"], workbook.sheetnames
            )
        finally:
            workbook.close()
        self.assertEqual("LECTURAS ACS M3", payload["workbook_layout"]["tables"]["ACS"]["sheet"])
        self.assertEqual("900_v1", excel_profiles.load_profile(
            "900_v1", self.project_root
        ).key)
        self.assertEqual(
            "Lectura",
            excel_profiles.load_profile(
                "900_v1", self.project_root
            ).onboarding_configuration["reading_column"],
        )
        self.assertEqual("900_v1", case_workflow_actions.resolve_case_profile(
            self.connection,
            id_case=result.case_id,
            active_community_id=result.community_id,
            project_root=self.project_root,
        ).key)

    def test_onboarded_runtime_profile_installs_bootstrap_template_for_long_code(self):
        draft = community_onboarding.OnboardingDraft(
            community_code="1234",
            community_name="Comunidad prueba larga",
            sources=self.valid_draft.sources,
            detected_modules=self.valid_draft.detected_modules,
            questions=self.valid_draft.questions,
            invoice_evidence=self.valid_draft.invoice_evidence,
            reading_evidence=self.valid_draft.reading_evidence,
        )
        result = community_onboarding.confirm_onboarding(
            self.connection,
            **{**self._confirmation_arguments(), "draft": draft},
        )
        template = self.project_root / "plantillas/comunidades/1234/1234_v1.xlsx"

        bootstrap, companions = case_workflow_actions.run_bootstrap_import(
            self.database,
            id_case=result.case_id,
            active_community_id=result.community_id,
            project_root=self.project_root,
            master_path=template,
            actor="Prueba",
        )

        self.assertIsNone(companions)
        self.assertEqual(template, bootstrap.installed_template_path)
        self.assertTrue(template.is_file())

    def test_confirm_onboarding_rolls_back_json_database_template_and_archives_after_error(self):
        real_register = expedient_service.register_source_document
        calls = 0
        sentinel_path = None

        def fail_after_two_real_registrations(*args, **kwargs):
            nonlocal calls, sentinel_path
            calls += 1
            if calls == 1:
                case_id = args[1]
                sentinel_path = self.archive_root / str(case_id) / "fuentes" / "ajeno.txt"
                sentinel_path.parent.mkdir(parents=True, exist_ok=True)
                sentinel_path.write_bytes(b"no pertenece al alta")
            result = real_register(*args, **kwargs)
            if calls == 3:
                raise OSError("fallo sintético tardío")
            return result

        with patch(
            "community_onboarding.expedient_service.register_source_document",
            side_effect=fail_after_two_real_registrations,
        ):
            with self.assertRaisesRegex(OSError, "sintético tardío"):
                community_onboarding.confirm_onboarding(
                    self.connection, **self._confirmation_arguments()
                )

        self.assertFalse((self.project_root / "config/excel_profiles/runtime/900_v1.json").exists())
        self.assertFalse((self.project_root / "plantillas/comunidades/900/900_v1.xlsx").exists())
        self.assertFalse((self.project_root / "config/excel_profiles").exists())
        self.assertFalse((self.project_root / "plantillas/comunidades/900").exists())
        self.assertEqual(0, self._community_count())
        self.assertEqual(0, self._registered_source_count())
        self.assertIsNotNone(sentinel_path)
        self.assertEqual(b"no pertenece al alta", sentinel_path.read_bytes())
        self.assertEqual(
            [sentinel_path],
            [path for path in self.archive_root.rglob("*") if path.is_file()],
        )

    def test_confirm_onboarding_never_replaces_a_concurrent_template(self):
        template_path = self.project_root / "plantillas/comunidades/900/900_v1.xlsx"
        sentinel = b"plantilla concurrente"
        real_link = os.link

        def publish_after_concurrent_create(source, destination):
            destination = Path(destination)
            if destination == template_path:
                destination.write_bytes(sentinel)
            return real_link(source, destination)

        with patch("community_onboarding.os.link", side_effect=publish_after_concurrent_create):
            with self.assertRaises(FileExistsError):
                community_onboarding.confirm_onboarding(
                    self.connection, **self._confirmation_arguments()
                )

        self.assertEqual(sentinel, template_path.read_bytes())
        self.assertFalse((self.project_root / "config/excel_profiles/runtime/900_v1.json").exists())
        self.assertEqual(0, self._community_count())

    def test_confirm_onboarding_never_replaces_a_concurrent_profile(self):
        profile_path = self.project_root / "config/excel_profiles/runtime/900_v1.json"
        template_path = self.project_root / "plantillas/comunidades/900/900_v1.xlsx"
        sentinel = b'{"propietario":"ajeno"}'
        real_link = os.link
        calls = 0

        def publish_after_concurrent_create(source, destination):
            nonlocal calls
            calls += 1
            destination = Path(destination)
            if calls == 2:
                self.assertEqual(profile_path, destination)
                destination.write_bytes(sentinel)
            return real_link(source, destination)

        with patch("community_onboarding.os.link", side_effect=publish_after_concurrent_create):
            with self.assertRaises(FileExistsError):
                community_onboarding.confirm_onboarding(
                    self.connection, **self._confirmation_arguments()
                )

        self.assertEqual(sentinel, profile_path.read_bytes())
        self.assertFalse(template_path.exists())
        self.assertEqual(0, self._community_count())

    def test_confirm_onboarding_rejects_partial_period_dates_without_side_effects(self):
        arguments = self._confirmation_arguments()
        arguments["end_date"] = None

        with self.assertRaisesRegex(ValueError, "fechas.*completas"):
            community_onboarding.confirm_onboarding(self.connection, **arguments)

        self.assertEqual(0, self._community_count())
        self.assertFalse((self.project_root / "config/excel_profiles/runtime/900_v1.json").exists())

    def test_confirm_onboarding_cleans_the_actual_source_when_it_changes_during_registration(self):
        real_register = expedient_service.register_source_document
        changed_source = self.owners_csv

        def change_then_fail(*args, **kwargs):
            changed_source.write_bytes(b"fuente cambiada durante el registro")
            result = real_register(*args, **kwargs)
            raise OSError("fallo posterior al registro cambiado")

        with patch(
            "community_onboarding.expedient_service.register_source_document",
            side_effect=change_then_fail,
        ):
            with self.assertRaisesRegex(OSError, "registro cambiado"):
                community_onboarding.confirm_onboarding(
                    self.connection, **self._confirmation_arguments()
                )

        self.assertEqual(0, self._registered_source_count())
        self.assertEqual(
            [], [path for path in self.archive_root.rglob("*") if path.is_file()]
            if self.archive_root.exists() else [],
        )
        self.assertEqual(0, self._community_count())


if __name__ == "__main__":
    unittest.main()
