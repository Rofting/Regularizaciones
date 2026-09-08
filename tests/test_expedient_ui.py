import sys
import tempfile
import unittest
from dataclasses import replace
from pathlib import Path
from unittest.mock import Mock, patch


PROJECT_ROOT = Path(__file__).resolve().parents[1]
CORE_DIR = PROJECT_ROOT / "core"
if str(CORE_DIR) not in sys.path:
    sys.path.insert(0, str(CORE_DIR))

import expedient_ui
from community_onboarding import (
    InvoiceDecision,
    OnboardingConfiguration,
    OnboardingDraft,
    OnboardingQuestion,
    OnboardingResult,
    ReadingBinding,
    ReadingEvidence,
    SourceCandidate,
)


class FakeVariable:
    def __init__(self, master=None, value=""):
        self.value = value
        self.callbacks = []

    def get(self):
        return self.value

    def set(self, value):
        self.value = value
        for callback in self.callbacks:
            callback(None, None, None)

    def trace_add(self, mode, callback):
        self.callbacks.append(callback)


class FakeWidget:
    """Only replaces Tk drawing; application commands and routing remain real."""
    def __init__(self, master=None, **options):
        self.options = options
        self.children = []
        self.master = master
        if master is not None:
            master.children.append(self)

    def pack(self, **kwargs):
        pass

    grid = pack
    configure = lambda self, **kwargs: self.options.update(kwargs)
    focus_set = lambda self: None
    minsize = lambda self, *args: None
    resizable = minsize
    protocol = minsize
    grid_columnconfigure = lambda self, *args, **kwargs: None

    def winfo_children(self):
        return list(self.children)

    def destroy(self):
        if self.master is not None:
            self.master.children.remove(self)
        self.children.clear()

    def descendants(self):
        yield self
        for child in self.children:
            yield from child.descendants()


class OnboardingSummaryDataTest(unittest.TestCase):
    def setUp(self):
        self.configuration = OnboardingConfiguration(
            service_decision="ACS+CALEFACCION",
            active_modules=("ACS", "CALEFACCION"),
            reading_column="",
            reading_bindings=(
                ReadingBinding(
                    module="ACS",
                    column="Lectura ACS",
                    source_sha256s=("a" * 64,),
                    meter="Contador ACS",
                    date="2026-08-31",
                    value="125",
                ),
                ReadingBinding(
                    module="CALEFACCION",
                    column="Lectura Calefacción",
                    source_sha256s=("b" * 64,),
                    meter="Contador calefacción",
                    date="2026-08-31",
                    value="980",
                ),
            ),
            source_traces=(
                ("owners", "propietarios.csv", "c" * 64),
                ("meter_reading_excel", "lecturas.xlsx", "a" * 64),
                ("meter_reading_pdf", "calefaccion.pdf", "b" * 64),
                ("invoice_pdf", "factura.pdf", "d" * 64),
            ),
            invoice_decisions=(InvoiceDecision(
                source_sha256="d" * 64,
                provider="Proveedor Norte",
                period="01/08/2026 - 31/08/2026",
                amount="432.10",
                concept="CALEFACCION",
            ),),
        )

    def test_exposes_invoice_and_reading_decisions_per_module(self):
        summary = expedient_ui.onboarding_summary_data(self.configuration)

        self.assertEqual(
            {
                "reading_column": "Lectura ACS",
                "meter": "Contador ACS",
                "date": "2026-08-31",
                "value": "125",
                "source_sha256s": ("a" * 64,),
            },
            summary["modules"]["ACS"],
        )
        self.assertEqual(
            "Lectura Calefacción",
            summary["modules"]["CALEFACCION"]["reading_column"],
        )
        self.assertEqual("Proveedor Norte", summary["invoices"][0]["provider"])
        self.assertEqual("CALEFACCION", summary["invoices"][0]["concept"])
        self.assertTrue(summary["ready_to_publish"])

    def test_marks_required_confirmation_before_publication(self):
        incomplete_binding = ReadingBinding(
            module="CALEFACCION",
            column="Lectura Calefacción",
            source_sha256s=("b" * 64,),
            meter="Contador calefacción",
            date="2026-08-31",
            value="",
        )
        incomplete_invoice = replace(
            self.configuration.invoice_decisions[0], provider=""
        )

        incomplete_configurations = (
            replace(
                self.configuration,
                reading_bindings=(self.configuration.reading_bindings[0],),
            ),
            replace(
                self.configuration,
                reading_bindings=(
                    self.configuration.reading_bindings[0],
                    incomplete_binding,
                ),
            ),
            replace(
                self.configuration,
                invoice_decisions=(incomplete_invoice,),
            ),
            replace(self.configuration, invoice_decisions=()),
        )

        for configuration in incomplete_configurations:
            with self.subTest(configuration=configuration):
                summary = expedient_ui.onboarding_summary_data(configuration)
                self.assertFalse(summary["ready_to_publish"])


class SourceFolderAndIssueGuidanceTest(unittest.TestCase):
    def test_source_folder_ignores_unsupported_and_hidden_files(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            for name in ("factura.pdf", "lecturas.XLS", "vecinos.csv", "nota.txt", ".oculto.pdf", "~$lecturas.xlsx"):
                (root / name).write_text("x", encoding="utf-8")
            (root / "lecturas").mkdir()
            (root / "lecturas" / "cierre.xlsx").write_text("x", encoding="utf-8")
            (root / ".temporal").mkdir()
            (root / ".temporal" / "ignorar.pdf").write_text("x", encoding="utf-8")

            paths = expedient_ui.source_files_in_folder(root)

        self.assertEqual(
            ["factura.pdf", "lecturas.XLS", "lecturas/cierre.xlsx", "vecinos.csv"],
            [str(path.relative_to(root)).replace("\\", "/") for path in paths],
        )

    def test_issue_guidance_explains_invoice_start_date_for_a_human(self):
        guidance = expedient_ui.issue_guidance("fecha_inicio")

        self.assertEqual("Fecha de inicio del período facturado", guidance["label"])
        self.assertIn("Busca", guidance["what_to_find"])
        self.assertIn("dd/mm/aaaa", guidance["format"])
        self.assertIn("cálculo", guidance["why"])


class CommunityOnboardingDialogTest(unittest.TestCase):
    def setUp(self):
        self.dialog = FakeWidget()
        self.app = Mock(_procesando=False)
        self.app.ruta_bd_expedientes = "test.db"
        self.app.ruta_archivo_expedientes = "archive"
        self.app._ids_comunidad = {"TEST — Comunidad de prueba": 1}
        self.workers = []
        self.callbacks = []
        self.app._en_hilo.side_effect = self.workers.append
        self.app.after.side_effect = lambda delay, fn: self.callbacks.append(fn)
        for target, replacement in (
            ("_dialog", Mock(return_value=self.dialog)),
            ("tk.StringVar", FakeVariable), ("tk.BooleanVar", FakeVariable),
            ("UIM.fuente", Mock(return_value="font")),
            *((f"ctk.{kind}", FakeWidget) for kind in (
                "CTkFrame", "CTkScrollableFrame", "CTkLabel", "CTkEntry",
                "CTkButton", "CTkComboBox",
            )),
        ):
            p = patch(f"expedient_ui.{target}", replacement)
            p.start()
            self.addCleanup(p.stop)
        self.messages = self.patch("messagebox")
        self.files = self.patch("filedialog")
        self.checkboxes = self.patch("ctk.CTkCheckBox")
        self.analysis = self.patch("community_onboarding.analyse_sources")
        self.confirm = self.patch("community_onboarding.confirm_onboarding")
        self.patch("gestor_bd.conectar")
        self.analysis.return_value = OnboardingDraft(
            "TEST", "Comunidad de prueba", (
                SourceCandidate(Path("owners.csv"), "owners", "a" * 64),
                SourceCandidate(
                    Path("readings.xlsx"), "meter_reading_excel", "b" * 64,
                    headers=("Vivienda", "Lectura A", "Lectura B"),
                ),
            ), ("ACS",),
            (
                OnboardingQuestion(
                    "service", "Confirma el servicio",
                    ("ACS", "CALEFACCION", "ACS+CALEFACCION", "NO_APLICA"), True
                ),
                OnboardingQuestion(
                    "reading_column", "Elige columna", ("Lectura A", "Lectura B"), True
                ),
            ),
            reading_evidence=(ReadingEvidence(
                source_sha256="b" * 64,
                meters=("Contador A",),
                columns=("Lectura A", "Lectura B"),
                dates=("2026-01-01",),
                readings=("100",),
                services=("ACS",),
            ),),
        )
        expedient_ui.open_community_onboarding_dialog(self.app)
        entries = self.entries()
        self.code = entries[0].options["textvariable"]
        self.code.set("TEST")
        entries[1].options["textvariable"].set("Comunidad de prueba")
        self.period = [entry.options["textvariable"] for entry in entries[2:]]

    def patch(self, target):
        p = patch(f"expedient_ui.{target}")
        result = p.start()
        self.addCleanup(p.stop)
        return result

    def entries(self):
        return [w for w in self.dialog.descendants() if "placeholder_text" in w.options]

    def texts(self):
        return [w.options.get("text", "") for w in self.dialog.descendants()]

    def command(self, text):
        buttons = [w for w in self.dialog.descendants()
                   if w.options.get("text") == text and "command" in w.options]
        self.assertEqual(1, len(buttons), f"Button {text!r} missing: {self.texts()}")
        return buttons[0].options["command"]

    def click(self, text):
        self.command(text)()

    def select_sources(self, invoices=False):
        self.click("Continuar a fuentes")
        self.files.askopenfilename.return_value = "owners.csv"
        self.click("Elegir CSV")
        self.files.askopenfilenames.return_value = ("readings.pdf",)
        self.click("Elegir lecturas")
        if invoices:
            self.files.askopenfilenames.return_value = ("invoice.pdf",)
            self.click("Elegir facturas")

    def complete_worker(self):
        self.assertEqual(1, len(self.workers))
        self.workers.pop(0)()
        while self.callbacks:
            self.callbacks.pop(0)()

    def analyse(self):
        self.click("Analizar fuentes")
        self.complete_worker()

    def summary(self):
        self.click("Revisar confirmaciones")
        self.assertIn("Confirma lo detectado", self.texts())
        self.click("Preparar resumen")
        self.assertNotIn("Crear comunidad", self.texts())
        for combo in (w for w in self.dialog.descendants() if "values" in w.options):
            values = combo.options["values"]
            combo.options["variable"].set("ACS" if "ACS" in values else "Lectura A")
        self.click("Preparar resumen")
        self.assertIn("Crear comunidad", self.texts())

    def test_analysis_callback_reaches_confirmations_summary_and_publication(self):
        self.select_sources()
        self.analyse()
        self.summary()
        summary_text = " ".join(self.texts())
        self.assertIn("Lectura ACS", self.texts())
        self.assertIn("Contador: Contador A", summary_text)
        self.assertIn("Fecha: 2026-01-01", summary_text)
        self.assertIn("Valor: 100", summary_text)
        self.assertIn("Columna: Lectura A", summary_text)
        self.confirm.return_value = OnboardingResult(1, None, None, Path("profile"), Path("template"), ())
        self.click("Crear comunidad")
        self.complete_worker()
        configuration = self.confirm.call_args.kwargs["answers"]
        self.assertEqual(("ACS",), configuration.active_modules)
        self.assertEqual("Lectura A", configuration.reading_column)
        self.assertEqual(1, self.app._cargar_comunidades.call_count)

    def test_confirmations_use_the_service_decision_without_module_checkboxes(self):
        self.select_sources()
        self.analyse()
        self.click("Revisar confirmaciones")

        self.assertFalse(self.checkboxes.called)
        choices = [
            widget.options["values"]
            for widget in self.dialog.descendants()
            if "values" in widget.options
        ]
        self.assertIn(
            ["ACS", "CALEFACCION", "ACS+CALEFACCION", "NO_APLICA"], choices
        )

    def test_simple_service_reaches_summary_with_inactive_questions_unanswered(self):
        draft = self.analysis.return_value
        self.analysis.return_value = replace(
            draft,
            questions=(*draft.questions, OnboardingQuestion(
                "reading_meter:CALEFACCION", "Contador de calefacción", (), True,
            )),
        )
        self.select_sources()
        self.analyse()
        self.summary()
        self.assertIn("Crear comunidad", self.texts())

    def test_analysis_error_back_and_continue_returns_to_sources_then_retries(self):
        self.select_sources()
        self.analysis.side_effect = ValueError("bad source")
        self.analyse()
        self.click("Atrás")
        self.click("Continuar a fuentes")
        self.assertIn("Analizar fuentes", self.texts())
        self.analysis.side_effect = None
        self.analyse()
        self.assertIn("Revisar confirmaciones", self.texts())
        self.assertEqual(2, self.analysis.call_count)

    def test_reanalysis_blocks_old_publish_and_requires_fresh_answers(self):
        self.select_sources()
        self.analyse()
        self.summary()
        publish = self.command("Crear comunidad")
        self.click("Atrás")  # confirmations
        self.click("Atrás")  # detection
        self.click("Atrás")  # sources
        self.click("Analizar fuentes")
        publish()
        self.assertEqual(1, len(self.workers))  # only analysis was queued
        self.assertFalse(self.confirm.called)
        self.complete_worker()
        self.summary()  # its first attempt must still require an answer

    def test_detection_with_a_service_decision_still_visits_confirmations(self):
        draft = self.analysis.return_value
        readings = SourceCandidate(
            Path("readings.xlsx"), "meter_reading_excel", "b" * 64,
            headers=("Vivienda", "Lectura"),
        )
        self.analysis.return_value = OnboardingDraft(
            draft.community_code, draft.community_name, (draft.sources[0], readings), ("ACS",),
            (OnboardingQuestion(
                "service", "Confirma el servicio", ("ACS", "NO_APLICA"), True
            ),),
            reading_evidence=(ReadingEvidence(
                source_sha256="b" * 64,
                meters=("Contador A",),
                columns=("Lectura",),
                dates=("2026-01-01",),
                readings=("100",),
                services=("ACS",),
            ),),
        )
        self.select_sources()
        self.analyse()
        self.click("Revisar confirmaciones")
        self.assertIn("Confirma lo detectado", self.texts())
        combo = next(w for w in self.dialog.descendants() if "values" in w.options)
        combo.options["variable"].set("ACS")
        self.click("Preparar resumen")
        self.assertIn("Crear comunidad", self.texts())

    def test_reanalysis_hides_old_draft_navigation_and_failure_cannot_publish(self):
        self.select_sources()
        self.analyse()
        old_next = self.command("Revisar confirmaciones")
        self.click("Atrás")
        self.analysis.side_effect = ValueError("bad source")
        self.click("Analizar fuentes")
        self.assertNotIn("Revisar confirmaciones", self.texts())
        old_next()
        self.assertNotIn("Preparar resumen", self.texts())
        self.complete_worker()
        self.assertIn("Analizar fuentes", self.texts())
        self.assertNotIn("Crear comunidad", self.texts())

    def test_identity_changed_during_analysis_discards_stale_callback(self):
        self.select_sources()
        self.click("Analizar fuentes")
        self.code.set("CHANGED")
        self.complete_worker()
        self.assertNotIn("Revisar confirmaciones", self.texts())
        self.assertIn("Continuar a fuentes", self.texts())
        self.click("Continuar a fuentes")
        self.analyse()
        self.assertEqual("CHANGED", self.analysis.call_args.kwargs["community_code"])

    def test_period_changed_after_worker_before_callback_discards_result(self):
        self.select_sources()
        self.click("Analizar fuentes")
        self.workers.pop(0)()
        self.period[0].set("2026")
        self.callbacks.pop(0)()
        self.assertIn("Continuar a fuentes", self.texts())
        self.click("Continuar a fuentes")
        self.assertIn("Continuar a fuentes", self.texts())  # incomplete dates
        self.assertNotIn("Crear comunidad", self.texts())

    def test_source_selection_is_blocked_while_worker_runs(self):
        self.select_sources()
        select_readings = self.command("Elegir lecturas")
        self.click("Analizar fuentes")
        self.files.askopenfilenames.return_value = ("changed.pdf",)
        select_readings()
        self.complete_worker()
        self.click("Atrás")
        self.analyse()
        self.assertEqual((Path("readings.pdf"),), self.analysis.call_args.kwargs["reading_paths"])

    def test_removing_optional_invoices_invalidates_analysis(self):
        self.select_sources(invoices=True)
        self.analyse()
        old_next = self.command("Revisar confirmaciones")
        self.click("Atrás")
        self.click("Quitar facturas")
        old_next()
        self.assertIn("Analizar fuentes", self.texts())
        self.analyse()
        self.assertEqual((), self.analysis.call_args.kwargs["invoice_paths"])

    def test_summary_and_completion_report_actual_archival_with_and_without_period(self):
        for with_period in (False, True):
            with self.subTest(with_period=with_period):
                if with_period:
                    # Reopen a fresh dialog after the first publication closed it.
                    self.dialog.destroy()
                    expedient_ui.open_community_onboarding_dialog(self.app)
                    entries = self.entries()
                    for entry, value in zip(entries, ("TEST", "Comunidad de prueba", "2026", "01/01/2026", "31/12/2026")):
                        entry.options["textvariable"].set(value)
                self.select_sources()
                self.analyse()
                self.summary()
                summary_text = " ".join(self.texts())
                if not with_period:
                    self.assertIn("no se archivarán", summary_text)
                    self.assertIn("0 · sin expediente inicial", summary_text)
                else:
                    self.assertIn("Se archivarán copias", summary_text)
                    self.assertIn("2", self.texts())
                self.confirm.return_value = OnboardingResult(
                    1, 2 if with_period else None, 3 if with_period else None,
                    Path("profile"), Path("template"), (7, 8) if with_period else (),
                )
                self.click("Crear comunidad")
                self.complete_worker()
                message = self.messages.showinfo.call_args.args[1]
                self.assertIn("2 fuente(s)" if with_period else "No se han archivado fuentes", message)


class CommunityOnboardingRouteTest(unittest.TestCase):
    def test_selected_sources_do_not_allow_summary_without_completed_analysis(self):
        state = {
            "step": "summary", "identity_valid": True, "has_sources": True,
            "analysis_complete": False, "answers_complete": True,
        }
        self.assertEqual("sources", expedient_ui.onboarding_step_route(state))

    def test_onboarding_step_route_requires_identity_first(self):
        state = {"step": "summary", "identity_valid": False, "has_sources": True}

        self.assertEqual("identity", expedient_ui.onboarding_step_route(state))

    def test_onboarding_step_route_requires_sources_before_detection(self):
        state = {"step": "identity", "identity_valid": True, "has_sources": False}

        self.assertEqual("sources", expedient_ui.onboarding_step_route(state))

    def test_onboarding_step_route_requires_answers_before_summary(self):
        state = {
            "step": "detected",
            "analysis_complete": True,
            "identity_valid": True,
            "has_sources": True,
            "has_required_questions": True,
            "answers_complete": False,
        }

        self.assertEqual("confirmations", expedient_ui.onboarding_step_route(state))

    def test_onboarding_step_route_blocks_incomplete_summary(self):
        state = {
            "step": "summary",
            "analysis_complete": True,
            "identity_valid": True,
            "has_sources": True,
            "has_required_questions": True,
            "answers_complete": False,
        }

        self.assertEqual("confirmations", expedient_ui.onboarding_step_route(state))

    def test_onboarding_step_route_allows_complete_summary(self):
        state = {
            "step": "confirmations",
            "analysis_complete": True,
            "identity_valid": True,
            "has_sources": True,
            "has_required_questions": True,
            "answers_complete": True,
        }

        self.assertEqual("summary", expedient_ui.onboarding_step_route(state))


if __name__ == "__main__":
    unittest.main()
