import sys
import sqlite3
import tempfile
import unittest
from contextlib import closing
from datetime import date
from dataclasses import replace
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock, patch


PROJECT_ROOT = Path(__file__).resolve().parents[1]
CORE_DIR = PROJECT_ROOT / "core"
if str(CORE_DIR) not in sys.path:
    sys.path.insert(0, str(CORE_DIR))

import expedient_ui
from case_readiness import CaseReadinessReport, ReadinessBlocker
import ui_moderna
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
    bind = lambda self, *args, **kwargs: None
    insert = lambda self, *args, **kwargs: None

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


class IssuePageTest(unittest.TestCase):
    def test_issue_page_limits_first_view_to_ten_of_fifty_five(self):
        visible, page, pages = expedient_ui.issue_page(tuple(range(55)), page=1)

        self.assertEqual(10, len(visible))
        self.assertEqual((1, 6), (page, pages))

    def test_issue_page_normalizes_out_of_range_page(self):
        visible, page, pages = expedient_ui.issue_page(tuple(range(11)), page=9)

        self.assertEqual((2, 2), (page, pages))
        self.assertEqual(1, len(visible))


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
    def test_source_summary_groups_detected_documents_by_kind(self):
        self.assertTrue(hasattr(expedient_ui, "source_summary"))
        self.assertEqual(
            "1 factura detectada · 2 lecturas · 1 documento por revisar",
            expedient_ui.source_summary(("invoice", "reading", "reading", "unknown")),
        )

    def test_issue_context_describes_pdf_page_or_excel_cell(self):
        self.assertTrue(hasattr(expedient_ui, "issue_context_label"))
        self.assertIn("Página 2", expedient_ui.issue_context_label('{"page": 2, "excerpt": "TOTAL"}'))
        self.assertIn("TOTAL", expedient_ui.issue_context_label('{"page": 2, "excerpt": "TOTAL"}'))
        self.assertIn("Datos!B4", expedient_ui.issue_context_label('{"sheet": "Datos", "cell": "B4"}'))

    def test_issue_context_handles_old_text_and_missing_or_malformed_metadata(self):
        self.assertTrue(hasattr(expedient_ui, "issue_context_label"))
        for value in (None, "", "[]", "null", "{}"):
            self.assertIn("Abrir archivo", expedient_ui.issue_context_label(value))
        self.assertIn("Lectura final", expedient_ui.issue_context_label("Lectura final"))

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


class SourceActionsTest(unittest.TestCase):
    """Real ingestion/database, with only Tk drawing and PDF extraction replaced."""
    def setUp(self):
        import app
        self.app_module = app
        self.directory = tempfile.TemporaryDirectory()
        self.addCleanup(self.directory.cleanup)
        self.root = Path(self.directory.name)
        self.database_path = self.root / "gestion.db"
        # Explicitly close the fixture connection, including on Windows.
        with sqlite3.connect(self.database_path) as connection:
            for sql in expedient_ui.gestor_bd.TABLAS:
                connection.execute(sql)
            expedient_ui.gestor_bd.aplicar_migraciones(connection)
        connection.close()
        self.connection = expedient_ui.gestor_bd.conectar(str(self.database_path))
        self.addCleanup(self.connection.close)
        community = expedient_ui.gestor_bd.obtener_o_crear_comunidad(self.connection, "TEST", "Prueba")
        self.case = expedient_ui.expedient_service.create_case(
            self.connection, community, name="2026", start_date=date(2026, 1, 1), end_date=date(2026, 12, 31),
        )
        self.dialog = FakeWidget()
        self.app = Mock(_procesando=False)
        self.app.id_expediente = self.case.id_case
        self.app.id_comunidad = community
        self.app.id_periodo = 12
        self.app.ruta_bd_expedientes = self.database_path
        self.app.ruta_archivo_expedientes = self.root / "archive"
        self.workers, self.callbacks = [], []
        self.app._en_hilo.side_effect = self.workers.append
        self.app.after.side_effect = lambda delay, fn: self.callbacks.append(fn)
        for target, replacement in (
            ("_dialog", Mock(return_value=self.dialog)),
            ("tk.StringVar", FakeVariable), ("UIM.fuente", Mock(return_value="font")),
            *((f"ctk.{kind}", FakeWidget) for kind in (
                "CTkFrame", "CTkScrollableFrame", "CTkLabel", "CTkEntry", "CTkButton", "CTkComboBox", "CTkSegmentedButton",
            )),
        ):
            p = patch(f"expedient_ui.{target}", replacement)
            p.start()
            self.addCleanup(p.stop)
        p = patch("expedient_ui.messagebox")
        self.messages = p.start()
        self.addCleanup(p.stop)
        p = patch("app.messagebox", self.messages)
        p.start()
        self.addCleanup(p.stop)

    def click(self, text):
        buttons = [w for w in self.dialog.descendants() if w.options.get("text") == text and "command" in w.options]
        self.assertEqual(1, len(buttons), f"Missing button: {text}")
        buttons[0].options["command"]()

    def complete(self):
        while self.callbacks or self.workers:
            while self.callbacks:
                self.callbacks.pop(0)()
            if self.workers:
                self.workers.pop(0)()

    def ingest(self):
        from source_analysis import SourceAnalysis
        paths = [self.root / "invoice.pdf", self.root / "readings.csv", self.root / "unknown.csv"]
        for path, content in zip(paths, ("pdf fixture", "vivienda;tipo;fecha_ant;val_ant;fecha_act;val_act\nA;ACS;2026-01-01;100;2026-12-31;120", "foo;bar\nx;y")):
            path.write_text(content, encoding="utf-8")
        expedient_ui.open_add_sources_dialog(self.app, self.case.id_case)
        with patch("expedient_ui.filedialog.askopenfilenames", return_value=tuple(map(str, paths))), patch(
            "source_analysis.analyse_pdf", return_value=SourceAnalysis.invoice({"importe_total": "123"}),
        ):
            self.click("Añadir archivos")
            self.complete()

    def test_mixed_sources_are_classified_before_missing_fields_and_summarized(self):
        self.ingest()
        rows = self.connection.execute("SELECT document_kind FROM source_documents ORDER BY id_document").fetchall()
        self.assertEqual(["invoice", "reading", "unknown"], [row[0] for row in rows])
        issues = expedient_ui.document_review.list_open_issues(self.connection, self.case.id_case)
        self.assertEqual(["fecha_inicio", "fecha_fin", "document_kind"], [issue.field_name for issue in issues])
        self.assertIn("1 factura detectada · 1 lectura · 1 documento por revisar", self.messages.showinfo.call_args.args[1])

    def test_clean_folder_intake_advances_without_opening_confirmation_dialog(self):
        """Un lote limpio debe dejar visible el siguiente paso al terminar."""
        from source_analysis import SourceAnalysis

        source = self.root / "TEST_factura.pdf"
        source.write_text("pdf fixture", encoding="utf-8")
        analysis = SourceAnalysis.invoice(
            {
                "proveedor": "Naturgy Clientes S.A.U.",
                "tipo_suministro": "GAS",
                "fecha_inicio": "2026-01-01",
                "fecha_fin": "2026-01-31",
                "importe_total": "123.00",
            },
            provider_key="NATURGY_CLIENTES_GAS",
        )

        expedient_ui.open_add_sources_dialog(self.app, self.case.id_case)
        with patch(
            "expedient_ui.filedialog.askdirectory", return_value=str(self.root),
        ), patch("source_analysis.analyse_pdf", return_value=analysis):
            self.click("Añadir carpeta")
            self.complete()

        status = self.connection.execute(
            "SELECT estado FROM regularization_cases WHERE id_case=?",
            (self.case.id_case,),
        ).fetchone()[0]
        self.assertEqual("under_review", status)
        self.assertIn(
            "propietarios",
            self.messages.showinfo.call_args.args[1],
        )
        self.app._refrescar_lista_expedientes.assert_called_with(
            select_case_id=self.case.id_case
        )
        self.app._refrescar_expediente.assert_called()

    def test_classification_override_only_appears_for_unknown_and_creates_invoice_fields(self):
        self.ingest()
        issue = expedient_ui.document_review.list_open_issues(self.connection, self.case.id_case)[-1]
        expedient_ui.open_issue_dialog(self.app, issue)
        combos = [w for w in self.dialog.descendants() if "values" in w.options]
        self.assertEqual(1, len(combos))
        combos[0].options["variable"].set("Factura")
        with patch("expedient_ui._field") as field:
            self.dialog.destroy()
            field.return_value.get.return_value = "Comprobado en el original"
            expedient_ui.open_issue_dialog(self.app, issue)
            combo = next(w for w in self.dialog.descendants() if "values" in w.options)
            combo.options["variable"].set("Factura")
            self.click("Guardar clasificación")
        row = self.connection.execute("SELECT document_kind FROM source_documents WHERE id_document = ?", (issue.id_document,)).fetchone()
        self.assertEqual("invoice", row[0])
        pending = expedient_ui.document_review.list_open_issues(self.connection, self.case.id_case)
        self.assertEqual({"fecha_inicio", "fecha_fin", "importe_total"}, {i.field_name for i in pending if i.id_document == issue.id_document})

    def test_context_button_reads_saved_document_context_for_a_missing_field(self):
        self.ingest()
        issue = expedient_ui.document_review.list_open_issues(self.connection, self.case.id_case)[0]
        with self.connection:
            self.connection.execute("UPDATE extraction_candidates SET source_context = ? WHERE id_document = ?", ('{"page":2,"fragment":"TOTAL 123"}', issue.id_document))
        expedient_ui.open_issue_dialog(self.app, issue)
        self.click("Ver contexto")
        self.assertIn("Página 2", self.messages.showinfo.call_args.args[1])
        self.assertIn("TOTAL 123", self.messages.showinfo.call_args.args[1])
        self.assertTrue(any(w.options.get("text") == "Abrir archivo" for w in self.dialog.descendants()))

    def test_failed_manual_classification_keeps_the_issue_open_without_a_partial_correction(self):
        self.ingest()
        issue = expedient_ui.document_review.list_open_issues(self.connection, self.case.id_case)[-1]
        with patch("expedient_ui._field") as field, patch(
            "case_ingestion.reanalyze_case_documents", side_effect=OSError("Source unavailable"),
        ):
            field.return_value.get.return_value = "Comprobado en el original"
            expedient_ui.open_issue_dialog(self.app, issue)
            combo = next(w for w in self.dialog.descendants() if "values" in w.options)
            combo.options["variable"].set("Factura")
            try:
                self.click("Guardar clasificación")
            except OSError:
                pass  # Assert the durable state after a failed reanalysis.
        row = self.connection.execute("SELECT status FROM review_issues WHERE id_issue = ?", (issue.id_issue,)).fetchone()
        self.assertEqual("open", row[0])
        self.assertEqual(0, self.connection.execute("SELECT count(*) FROM manual_corrections WHERE id_issue = ?", (issue.id_issue,)).fetchone()[0])
        self.assertIn("Source unavailable", self.messages.showwarning.call_args.args[1])

    def test_recognized_issue_has_no_manual_classification_selector(self):
        self.ingest()
        issue = expedient_ui.document_review.list_open_issues(self.connection, self.case.id_case)[0]
        expedient_ui.open_issue_dialog(self.app, issue)
        self.assertFalse(any("values" in w.options for w in self.dialog.descendants()))

    def test_safe_reset_initializes_a_complete_empty_temporary_database_and_keeps_backup(self):
        self.assertTrue(hasattr(self.app_module.AppGestionFincas, "_accion_nueva_base_segura"))
        self.connection.close()  # Windows requires all live connections closed for replacement.
        self.messages.askyesno.return_value = True
        self.app_module.AppGestionFincas._accion_nueva_base_segura(self.app)
        self.complete()
        with closing(sqlite3.connect(self.database_path)) as connection:
            self.assertEqual(0, connection.execute("SELECT count(*) FROM comunidades").fetchone()[0])
            self.assertEqual(0, connection.execute("SELECT count(*) FROM regularization_cases").fetchone()[0])
        backups = list((self.root / "backups").glob("*.db"))
        self.assertEqual(1, len(backups))
        with closing(sqlite3.connect(backups[0])) as connection:
            self.assertEqual(1, connection.execute("SELECT count(*) FROM comunidades").fetchone()[0])
        self.assertIn(str(backups[0]), self.messages.showinfo.call_args.args[1])

    def test_reanalysis_action_updates_existing_case_and_reports_groups(self):
        self.ingest()
        from source_analysis import SourceAnalysis
        self.assertTrue(hasattr(self.app_module.AppGestionFincas, "_accion_reanalizar_fuentes"))
        with patch("source_analysis.analyse_pdf", return_value=SourceAnalysis.reading()):
            self.app_module.AppGestionFincas._accion_reanalizar_fuentes(self.app)
            self.complete()
        self.assertEqual(3, self.connection.execute("SELECT count(*) FROM source_documents").fetchone()[0])
        self.assertIn("2 lecturas", self.messages.showinfo.call_args.args[1])

    def test_resolve_incidents_action_opens_the_first_pending_issue(self):
        self.ingest()
        issue = expedient_ui.document_review.list_open_issues(self.connection, self.case.id_case)[0]
        review_ui = Mock()
        with patch.dict(self.app_module.MOD, {"expedient_ui": review_ui}):
            self.app_module.AppGestionFincas._accion_resolver_incidencias(self.app)

        review_ui.open_issue_dialog.assert_called_once_with(self.app, issue)

    def test_startup_creates_a_verified_database_backup(self):
        with patch.object(self.app_module, "RUTA_BD", self.database_path), \
             patch.object(self.app_module, "BASE_DIR", self.root), \
             patch.object(self.app_module, "RUTA_CARTAS", self.root / "cartas"):
            self.app_module.AppGestionFincas._verificar_estructura(self.app)
        copies = list((self.root / "backups").glob("*.db"))
        self.assertEqual(1, len(copies))
        with closing(sqlite3.connect(copies[0])) as con:
            self.assertEqual(1, con.execute("SELECT COUNT(*) FROM comunidades").fetchone()[0])

    def test_fixed_cost_blocker_routes_to_the_fixed_cost_form(self):
        from case_readiness import ReadinessBlocker, CaseReadinessReport
        state = expedient_ui.guided_workspace_state(
            has_case=True, document_count=1, open_issue_count=0,
            case_status="ready_for_calculation", readiness_report=CaseReadinessReport(
                1, (ReadinessBlocker("UNCONFIRMED_FIXED_COSTS", "excel", "Revisar B6", "review_fixed_costs"),),
            ),
        )
        self.assertEqual("revisar_gastos_fijos", state.next_action)
        self.app_module.AppGestionFincas._actualizar_workspace(self.app, state)
        self.assertEqual(self.app._accion_gastos_fijos, self.app._workspace_command)

    def test_coherence_blocker_routes_to_its_review_panel(self):
        from case_readiness import ReadinessBlocker, CaseReadinessReport
        state = expedient_ui.guided_workspace_state(
            has_case=True, document_count=1, open_issue_count=0, case_status='ready_for_calculation',
            readiness_report=CaseReadinessReport(1, (ReadinessBlocker('COHERENCE_OVERLAP', 'excel', 'Solape', 'review_coherence'),)),
        )
        self.assertEqual('revisar_coherencia', state.next_action)
        self.app_module.AppGestionFincas._actualizar_workspace(self.app, state)
        self.assertEqual(self.app._accion_coherencia, self.app._workspace_command)

    def test_coherence_panel_corrects_percentages_and_accepts_a_legitimate_overlap(self):
        from contextlib import ExitStack
        from excel_profiles import ExcelProfile, ConceptRule
        import case_coherence
        period = expedient_ui.expedient_service.link_case_to_period(self.connection, self.case.id_case)
        profile = ExcelProfile('ui', '1', 'TEST', 'modelo.xlsx', (), (), (), (
            ConceptRule('other', 'coefficient', 'period_parameters.other', None, True),))
        for code in ('A', 'B'):
            self.connection.execute('INSERT INTO propietarios(id_comunidad,codigo_vivienda,nombre_propietario,coeficiente) VALUES (?,?,?,49)', (self.app.id_comunidad, code, 'Vecino ' + code))
        for number in ('1', '2'):
            self.connection.execute('''INSERT INTO facturas(id_comunidad,id_periodo,tipo_suministro,
                num_factura,cups_o_referencia,fecha_inicio,fecha_fin,importe_total)
                VALUES (?,?,'GAS',?,'ES-PUNTO','2026-01-01','2026-02-01',100)''',
                (self.app.id_comunidad, period, number))
        self.connection.commit()
        case_coherence.save_settings(self.connection, self.case.id_case, dict(coefficient_mode='percent'))
        with ExitStack() as stack:
            stack.enter_context(patch('case_workflow_actions.resolve_case_profile', return_value=profile))
            stack.enter_context(patch.object(FakeWidget, 'set', lambda w, value: w.options.update(value=value), create=True))
            stack.enter_context(patch.object(FakeWidget, 'get', lambda w: w.options.get('value', ''), create=True))
            stack.enter_context(patch.object(FakeWidget, 'insert', lambda w, position, value: w.options.update(value=value)))
            expedient_ui.open_coherence_dialog(self.app)
            for widget in self.dialog.descendants():
                if widget.options.get('value') == '49.0':
                    widget.options['value'] = '50'
            self.click('Guardar configuración y coeficientes')
            for widget in self.dialog.descendants():
                if 'Motivo de la diferencia' in widget.options.get('placeholder_text', ''):
                    widget.options['value'] = 'Rectificación de la primera factura'
            self.click('Aceptar con motivo')
        report = case_coherence.evaluate(self.connection, self.case.id_case, profile)
        self.assertEqual((), report.pending)
        self.assertEqual(100, self.connection.execute('SELECT SUM(coeficiente) FROM propietarios').fetchone()[0])

    def test_fixed_cost_form_can_confirm_no_cost_for_this_period(self):
        period_id = expedient_ui.expedient_service.link_case_to_period(self.connection, self.case.id_case)
        expedient_ui.open_fixed_costs_dialog(self.app, period_id)
        with patch.object(FakeWidget, "get", return_value="", create=True):
            self.click("Guardar")
        values = expedient_ui.fixed_costs.load_fixed_costs(self.connection, self.app.id_comunidad, period_id)
        self.assertEqual(4, len(values))
        self.assertTrue(all(value == 0 for value in values.values()))

    def test_resolve_incidents_action_routes_counter_resets_to_the_grouped_view(self):
        self.ingest()
        document_id = self.connection.execute(
            "SELECT id_document FROM source_documents WHERE id_case=? ORDER BY id_document LIMIT 1",
            (self.case.id_case,),
        ).fetchone()[0]
        expedient_ui.document_review.create_review_issue(
            self.connection, self.case.id_case, document_id,
            code="COUNTER_RESET", field_name="reading.A.ACS",
            message="El contador disminuye y requiere una estimación aprobada",
        )
        review_ui = Mock()
        with patch.dict(self.app_module.MOD, {"expedient_ui": review_ui}):
            self.app_module.AppGestionFincas._accion_resolver_incidencias(self.app)

        review_ui.open_confirm_sources_dialog.assert_called_once_with(self.app, self.case.id_case)

    def test_resolve_incidents_action_routes_initial_zeroes_to_the_grouped_view(self):
        self.ingest()
        document_id = self.connection.execute(
            "SELECT id_document FROM source_documents WHERE id_case=? ORDER BY id_document LIMIT 1",
            (self.case.id_case,),
        ).fetchone()[0]
        expedient_ui.document_review.create_review_issue(
            self.connection, self.case.id_case, document_id,
            code="READING_ZERO_REVIEW", field_name="reading.A.ACS",
            message="La lectura cero no tiene una lectura anterior fiable para conservar",
        )
        review_ui = Mock()
        with patch.dict(self.app_module.MOD, {"expedient_ui": review_ui}):
            self.app_module.AppGestionFincas._accion_resolver_incidencias(self.app)

        review_ui.open_confirm_sources_dialog.assert_called_once_with(self.app, self.case.id_case)

    def test_background_workflow_error_is_shown_in_a_dialog(self):
        def fail():
            raise ValueError("Falta importar el modelo inicial")

        self.app_module.AppGestionFincas._ejecutar_hilo(self.app, fail)
        self.complete()

        self.messages.showerror.assert_called_once()
        title, detail = self.messages.showerror.call_args.args[:2]
        self.assertEqual("No se pudo completar", title)
        self.assertIn("Falta importar el modelo inicial", detail)
        self.assertIn("No se ha modificado", detail)

    def test_failed_excel_generation_refreshes_the_actionable_case_state(self):
        workflow = Mock()
        workflow.run_generate_excel.side_effect = ValueError("Falta la fecha de factura")

        with patch.dict(self.app_module.MOD, {"case_workflow_actions": workflow}):
            with self.assertRaisesRegex(ValueError, "fecha de factura"):
                self.app_module.AppGestionFincas._generar_excel_expediente_impl(self.app)

        self.app._refrescar_despues_de_accion.assert_called_once_with(self.case.id_case)

    def test_reset_cancellation_failure_and_success_preserve_or_clear_ui_at_the_right_time(self):
        from database_reset import DatabaseResetError, ResetResult
        self.assertTrue(hasattr(self.app_module.AppGestionFincas, "_accion_nueva_base_segura"))
        result = ResetResult(self.root / "backup.db", self.database_path)
        # Reset is deliberately substituted; no production database can be touched.
        with patch("database_reset.reset_database", side_effect=DatabaseResetError("backup failed")) as reset:
            self.messages.askyesno.return_value = False
            self.app_module.AppGestionFincas._accion_nueva_base_segura(self.app)
            self.complete()
            self.assertFalse(reset.called)
            self.messages.askyesno.return_value = True
            self.app_module.AppGestionFincas._accion_nueva_base_segura(self.app)
            self.complete()
            self.assertEqual(self.case.id_case, self.app.id_expediente)
            self.assertEqual(12, self.app.id_periodo)
            self.assertFalse(self.app._limpiar_contexto_expediente.called)
            reset.side_effect = None
            reset.return_value = result
            self.app_module.AppGestionFincas._accion_nueva_base_segura(self.app)
            self.assertEqual(self.case.id_case, self.app.id_expediente)
            self.complete()
            self.assertIsNone(self.app.id_comunidad)
            self.assertIsNone(self.app.id_periodo)
            self.assertTrue(self.app._limpiar_contexto_expediente.called)
            self.assertIn(str(result.backup_path), self.messages.showinfo.call_args.args[1])
            self.assertEqual(self.database_path, reset.call_args.args[0])


class GuidedWorkspaceStateTest(unittest.TestCase):
    def test_missing_owners_routes_to_importing_the_owner_list(self):
        report = CaseReadinessReport(7, (
            ReadinessBlocker(
                "MISSING_OWNERS", "excel",
                "Faltan propietarios activos para preparar el reparto.",
                "import_owners",
            ),
        ))

        state = expedient_ui.guided_workspace_state(
            has_case=True,
            document_count=5,
            open_issue_count=0,
            case_status="under_review",
            readiness_report=report,
        )

        self.assertEqual(("fuentes", "importar_propietarios"), (
            state.active_step, state.next_action,
        ))
        self.assertEqual("Importar propietarios", state.headline)
        self.assertEqual(report.blockers, state.blockers)

    def test_readiness_profile_blocker_offers_automatic_excel_preparation(self):
        report = CaseReadinessReport(7, (
            ReadinessBlocker(
                "MISSING_PROFILE", "excel",
                "Falta preparar el perfil Excel de la comunidad.",
                "prepare_excel",
            ),
        ))

        state = expedient_ui.guided_workspace_state(
            has_case=True,
            document_count=5,
            open_issue_count=0,
            case_status="under_review",
            readiness_report=report,
        )

        self.assertEqual(("reparto", "generar_excel"), (
            state.active_step, state.next_action,
        ))
        self.assertEqual("Preparar Excel oficial", state.headline)

    def test_blocker_free_calculated_case_routes_to_distribution(self):
        state = expedient_ui.guided_workspace_state(
            has_case=True,
            document_count=5,
            open_issue_count=0,
            case_status="calculated",
            readiness_report=CaseReadinessReport(7, ()),
        )

        self.assertEqual(("reparto", "calcular_reparto"), (
            state.active_step, state.next_action,
        ))

    def _issue(self, code, field_name, detected=None):
        return expedient_ui.ReviewIssue(
            id_issue=1, id_case=1, id_document=7, code=code, field_name=field_name,
            detected_value=detected, message="", archived_path=Path("factura.pdf"), status="open",
        )

    def test_issue_titles_are_readable_instead_of_technical_field_names(self):
        self.assertEqual("Proveedor sin identificar",
                         expedient_ui.issue_title(self._issue("PROVIDER_UNKNOWN", "document.provider")))
        self.assertEqual(
            "Falta: fecha de inicio del período facturado",
            expedient_ui.issue_title(self._issue("MISSING_REQUIRED_FIELD", "fecha_inicio")),
        )
        self.assertNotIn("_", expedient_ui.issue_title(self._issue("X", "campo_raro.algo")))

    def test_issue_suggestion_prefers_detected_value_then_best_evidence(self):
        connection = sqlite3.connect(":memory:")
        connection.row_factory = sqlite3.Row
        connection.execute(
            "CREATE TABLE source_field_evidence (id_document INTEGER, field_name TEXT, value TEXT, confidence TEXT)"
        )
        connection.executemany(
            "INSERT INTO source_field_evidence VALUES (7,'fecha_inicio',?,?)",
            [("2026-03-02", "low"), ("2026-03-01", "medium")],
        )
        self.assertEqual("01/03/2026", expedient_ui.issue_suggestion(
            connection, self._issue("MISSING_REQUIRED_FIELD", "fecha_inicio")))
        self.assertEqual("121,00", expedient_ui.issue_suggestion(
            connection, self._issue("X", "importe_total", detected="121,00")))
        self.assertIsNone(expedient_ui.issue_suggestion(
            connection, self._issue("X", "num_factura")))
        connection.close()

    def test_review_summary_treats_resets_from_one_source_as_one_action(self):
        issues = tuple(
            expedient_ui.ReviewIssue(
                id_issue=index,
                id_case=1,
                id_document=9,
                code="COUNTER_RESET",
                field_name=f"reading.{index}.ACS",
                detected_value=None,
                message="El contador disminuye",
                archived_path=Path("lecturas_anuales.xls"),
                status="open",
            )
            for index in range(1, 51)
        )

        summary = expedient_ui.review_summary(issues)

        self.assertEqual(50, summary.technical_count)
        self.assertEqual(1, summary.actionable_count)
        self.assertEqual(50, summary.groups[0]["count"])

    def test_review_summary_treats_initial_zeroes_from_one_source_as_one_action(self):
        issues = tuple(
            expedient_ui.ReviewIssue(
                id_issue=index,
                id_case=1,
                id_document=9,
                code="READING_ZERO_REVIEW",
                field_name=f"reading.{index}.ACS",
                detected_value="0",
                message="La lectura cero no tiene una lectura anterior fiable para conservar",
                archived_path=Path("lecturas_iniciales.xls"),
                status="open",
            )
            for index in range(1, 13)
        )

        summary = expedient_ui.review_summary(issues)

        self.assertEqual(12, summary.technical_count)
        self.assertEqual(1, summary.actionable_count)
        self.assertEqual(12, summary.groups[0]["count"])

    def test_under_review_without_issues_routes_to_source_confirmation(self):
        """Guards against a silent attempt to generate Excel before applying sources."""
        state = expedient_ui.guided_workspace_state(
            has_case=True,
            document_count=3,
            open_issue_count=0,
            case_status="under_review",
        )

        self.assertEqual("confirmar_fuentes", state.next_action)
        self.assertEqual("Confirmar fuentes", state.headline)

    def test_completed_stages_can_be_repeated_once_excel_exists(self):
        def repeats(status):
            return expedient_ui.guided_workspace_state(
                has_case=True, document_count=3, open_issue_count=0, case_status=status,
            ).repeat_actions

        self.assertEqual((), repeats("under_review"))
        self.assertIn("reevaluar_fuentes", repeats("calculated"))
        self.assertIn("generar_excel", repeats("calculated"))
        self.assertNotIn("generar_cartas", repeats("calculated"))
        self.assertIn("calcular_reparto", repeats("reconciled"))
        self.assertIn("generar_cartas", repeats("deliveries_generated"))

    def test_profile_mismatch_routes_to_an_actionable_revalidation(self):
        state = expedient_ui.guided_workspace_state(
            has_case=True,
            document_count=3,
            open_issue_count=0,
            case_status="reconciled",
            profile_issue="La huella del perfil Excel activo no coincide con su registro",
        )

        self.assertEqual(("validar", "revalidar_perfil"), (state.active_step, state.next_action))
        self.assertEqual("Revalidar configuración del Excel", state.headline)
        self.assertIn("huella", state.detail)

    def test_missing_profile_routes_to_automatic_excel_preparation(self):
        state = expedient_ui.guided_workspace_state(
            has_case=True,
            document_count=3,
            open_issue_count=0,
            case_status="reconciled",
            has_registered_template=False,
            profile_missing=True,
        )

        self.assertEqual(("reparto", "generar_excel"), (state.active_step, state.next_action))
        self.assertEqual("Prepara el Excel oficial", state.headline)
        self.assertIn("automáticamente", state.detail.lower())

    def test_no_case_prompts_case_creation(self):
        state = expedient_ui.guided_workspace_state(
            has_case=False, document_count=0, open_issue_count=0, case_status=""
        )

        self.assertEqual(
            ("fuentes", "crear_expediente"),
            (state.active_step, state.next_action),
        )
        self.assertEqual("Crear expediente", state.headline)

    def test_open_issue_routes_to_validation(self):
        state = expedient_ui.guided_workspace_state(
            has_case=True, document_count=4, open_issue_count=2,
            case_status="under_review",
        )

        self.assertEqual(
            ("validar", "resolver_incidencias"),
            (state.active_step, state.next_action),
        )
        self.assertEqual("Hay 2 decisión(es) pendiente(s) antes de continuar.", state.detail)

    def test_ready_case_routes_to_excel_generation(self):
        state = expedient_ui.guided_workspace_state(
            has_case=True, document_count=4, open_issue_count=0,
            case_status="ready_for_calculation",
        )

        self.assertEqual(
            ("reparto", "generar_excel"),
            (state.active_step, state.next_action),
        )

    def test_ready_case_without_registered_template_still_generates(self):
        """Sin plantilla propia se genera igual: la crea el modelo canónico.

        Exigir antes un Excel maestro elegido a mano dejaba parado un
        expediente que ya tenía todas sus fuentes validadas.
        """
        state = expedient_ui.guided_workspace_state(
            has_case=True,
            document_count=4,
            open_issue_count=0,
            case_status="ready_for_calculation",
            has_registered_template=False,
        )

        self.assertEqual(
            ("reparto", "generar_excel"),
            (state.active_step, state.next_action),
        )
        self.assertEqual("Genera el Excel oficial", state.headline)
        self.assertIn("modelo del despacho", state.detail)

    def test_reconciled_case_routes_to_letters(self):
        state = expedient_ui.guided_workspace_state(
            has_case=True, document_count=4, open_issue_count=0,
            case_status="reconciled",
        )

        self.assertEqual(
            ("cartas", "generar_cartas"),
            (state.active_step, state.next_action),
        )


class PackedModalLayoutTest(unittest.TestCase):
    def _app_and_widgets(self):
        grid_calls = []

        class TrackingWidget(FakeWidget):
            def grid(self, **kwargs):
                grid_calls.append(kwargs)

            def insert(self, *_args):
                pass

        dialog = TrackingWidget()
        app = type("DialogApp", (), {
            "_procesando": False,
            "_preparar_dialogo": lambda _self, *_args: dialog,
        })()
        return app, dialog, grid_calls, TrackingWidget

    def test_counter_reset_dialog_keeps_reason_and_action_in_the_packed_panel(self):
        app, dialog, grid_calls, TrackingWidget = self._app_and_widgets()
        with patch.object(expedient_ui.UIM, "fuente", return_value=None), patch.multiple(
            expedient_ui.ctk,
            CTkFrame=TrackingWidget,
            CTkLabel=TrackingWidget,
            CTkEntry=TrackingWidget,
            CTkButton=TrackingWidget,
        ):
            expedient_ui.open_counter_reset_carry_forward_dialog(
                app, case_id=1, document_id=2, count=50, source_name="lecturas.xls",
            )

        self.assertFalse(grid_calls)
        self.assertIn(
            "Aplicar a los 50 contadores",
            [widget.options.get("text") for widget in dialog.descendants()],
        )

    def test_initial_zero_dialog_keeps_reason_and_action_in_the_packed_panel(self):
        app, dialog, grid_calls, TrackingWidget = self._app_and_widgets()
        with patch.object(expedient_ui.UIM, "fuente", return_value=None), patch.multiple(
            expedient_ui.ctk,
            CTkFrame=TrackingWidget,
            CTkLabel=TrackingWidget,
            CTkEntry=TrackingWidget,
            CTkButton=TrackingWidget,
        ):
            expedient_ui.open_initial_zero_confirmation_dialog(
                app, case_id=1, document_id=2, count=12, source_name="lecturas.xls",
            )

        self.assertFalse(grid_calls)
        self.assertIn(
            "Aplicar a los 12 ceros",
            [widget.options.get("text") for widget in dialog.descendants()],
        )


class WorkflowVisualStyleTest(unittest.TestCase):
    def test_secondary_button_style_is_neutral_and_bordered(self):
        style = ui_moderna.secondary_button_kwargs()

        self.assertEqual("transparent", style["fg_color"])
        self.assertEqual(1, style["border_width"])
        self.assertEqual(ui_moderna.C["borde"], style["border_color"])

    def test_every_workflow_state_has_a_style(self):
        for status in ("pending", "active", "blocked", "ready", "done"):
            with self.subTest(status=status):
                self.assertIn(status, ui_moderna.WORKFLOW_STEP_STYLES)

    def test_completed_workflow_step_can_be_opened_again(self):
        invoked = []
        row = SimpleNamespace(_command=lambda: invoked.append("opened"), _status="done")

        ui_moderna.WorkflowStepRow._invoke(row)

        self.assertEqual(["opened"], invoked)


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
        self.click("Elegir archivo")
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
