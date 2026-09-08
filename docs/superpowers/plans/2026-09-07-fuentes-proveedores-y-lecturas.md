# Fuentes de proveedores y lecturas Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Hacer que el alta guiada combine facturas de proveedores y lecturas de contador con confirmaciones manuales completas, perfil/Excel canónico y reparto válido.

**Architecture:** El análisis conserva evidencia de factura y de lectura por separado. Una única normalización de respuestas humanas produce la configuración por módulo que consume el resumen, el perfil, la plantilla y el motor. La excepción de bootstrap se limita a una plantilla inicial reconocible, para que los maestros posteriores incompletos sigan bloqueándose.

**Tech Stack:** Python 3, openpyxl, pdfplumber, CustomTkinter y unittest.

**Spec:** `docs/superpowers/specs/2026-09-07-fuentes-proveedores-y-lecturas-design.md`

## Global Constraints

- Facturas de proveedores aportan proveedor, período, importe y concepto; nunca lecturas de contador.
- Lecturas Excel/PDF aportan contador, columna, fecha y lectura; nunca importes facturados.
- Cada ausencia, ambigüedad o contradicción crea una confirmación manual obligatoria.
- ACS y calefacción pueden tener bindings de lectura diferentes.
- El perfil canónico generado usa `acs_fixed`, `acs_variable`, `heating_fixed` y `heating_variable`.
- Un maestro posterior vacío debe fallar validación; no se versionan documentos privados, base de datos ni perfiles runtime.
- Mantener alta rápida e interfaz v3; no crear Tk en pruebas.

---

## File structure

| Archivo | Responsabilidad |
| --- | --- |
| `core/community_onboarding.py` | Evidencia separada, preguntas y configuración confirmada por módulo. |
| `core/excel_profiles.py` | Validación de configuración y bindings por módulo. |
| `core/excel_generator.py` | Plantilla/layout canónico con parámetros y bindings. |
| `core/excel_bootstrap_importer.py` | Excepción inicial limitada y validación posterior estricta. |
| `core/excel_export_service.py` | Consume los bindings de perfil sin asumir una columna global. |
| `core/expedient_ui.py` | Confirmaciones y resumen verificable por módulo. |
| `tests/test_community_onboarding.py` | Fuentes, preguntas, payload, reparto y exportación. |
| `tests/test_expedient_ui.py` | Estado/resumen puro sin Tk. |
| `GUIA_PROCESAR_TODO.txt` | Manual operativo de activación y uso. |

### Task 1: Evidencia separada y decisiones manuales por módulo

**Files:**
- Modify: `core/community_onboarding.py`
- Modify: `core/excel_profiles.py`
- Modify: `tests/test_community_onboarding.py`

**Interfaces:**
- `analyse_sources(...)` conserva evidencia de proveedor y lectura por separado.
- `resolve_onboarding_configuration(draft, answers)` devuelve un `ReadingBinding` por módulo activo y trazas de factura/lectura.

- [ ] **Step 1: Write failing tests**

```python
def test_invoice_detects_heating_but_is_not_a_meter_reading(self):
    draft = analyse_sources(..., reading_paths=(self.readings_without_labels,),
                            invoice_paths=(self.heating_invoice,))
    self.assertIn("CALEFACCION", draft.detected_modules)
    self.assertEqual((), tuple(source for source in draft.sources
                               if source.kind == "invoice_pdf" and source.headers))

def test_combined_modules_require_a_confirmed_reading_for_each_module(self):
    configuration = resolve_onboarding_configuration(
        self.acs_heating_draft,
        {"service": "ACS+CALEFACCION", "reading_column:ACS": "Lectura ACS",
         "reading_column:CALEFACCION": "Lectura Calefacción"},
    )
    self.assertEqual(("Lectura ACS", "Lectura Calefacción"),
                     tuple(binding.column for binding in configuration.reading_bindings))

def test_missing_invoice_period_and_missing_reading_create_required_questions(self):
    draft = analyse_sources(..., invoice_paths=(self.invoice_without_period,),
                            reading_paths=(self.readings_without_value,))
    self.assertTrue(all(question.required for question in draft.questions
                        if question.key in {"invoice_period", "reading_column:ACS"}))
```

- [ ] **Step 2: Run tests to verify RED**

Run: `$env:PYTHONUTF8='1'; & $env:REGULARIZACION_PYTHON -m unittest tests.test_community_onboarding -q`

Expected: FAIL because the current configuration has one global `reading_column` and invoice evidence is not modelled.

- [ ] **Step 3: Implement the smallest domain model**

```python
def _invoice_evidence(invoices: tuple[SourceCandidate, ...]) -> tuple[InvoiceEvidence, ...]:
    """Extract only proposal-level provider, period, amount and service clues."""

def resolve_onboarding_configuration(draft, answers):
    active_modules = _confirmed_modules(draft, answers)
    return OnboardingConfiguration(
        service_decision=..., active_modules=active_modules,
        reading_bindings=tuple(_binding_for(module, draft, answers)
                               for module in active_modules),
        source_traces=..., invoice_decisions=...,)
```

Feed invoice text into service/concept detection only. Add required questions for missing/ambiguous provider, concept, amount or period when an invoice is selected, and for every missing/ambiguous reading decision needed by an active module. Permit `NO_APLICA` only for an inactive module with a recorded answer. Validate that each active module has exactly one non-empty binding and source fingerprint.

- [ ] **Step 4: Run tests to verify GREEN**

Run: `$env:PYTHONUTF8='1'; & $env:REGULARIZACION_PYTHON -m unittest tests.test_community_onboarding tests.test_excel_profiles -q`

Expected: PASS.

- [ ] **Step 5: Commit**

Commit message: `feat: confirma facturas y lecturas por modulo`

### Task 2: Perfil, plantilla y bootstrap canónicos

**Files:**
- Modify: `core/community_onboarding.py`
- Modify: `core/excel_generator.py`
- Modify: `core/excel_profiles.py`
- Modify: `core/excel_export_service.py`
- Modify: `core/excel_bootstrap_importer.py`
- Modify: `tests/test_community_onboarding.py`
- Modify: `tests/test_excel_bootstrap_importer.py`

**Interfaces:**
- Consume `OnboardingConfiguration` de Task 1.
- Produce perfiles con conceptos y `workbook_layout.parameter_cells` canónicos.

- [ ] **Step 1: Write failing tests**

```python
def test_combined_onboarding_profile_has_canonical_concepts_and_parameter_bindings(self):
    payload = build_profile_payload(self.acs_heating_draft, self.confirmed_configuration)
    self.assertEqual({"acs_fixed", "acs_variable", "heating_fixed", "heating_variable"},
                     {concept["key"] for concept in payload["concepts"]})
    layout = build_canonical_layout(payload)
    self.assertEqual({"acs_fixed", "acs_variable", "heating_fixed", "heating_variable"},
                     set(layout["parameter_cells"]))

def test_onboarding_with_distinct_readings_exports_and_calculates_distribution(self):
    result = self._confirm_import_export_and_distribute(
        {"reading_column:ACS": "Lectura ACS",
         "reading_column:CALEFACCION": "Lectura Calefacción"})
    self.assertTrue(result.workbook_path.is_file())
    self.assertTrue(result.distribution_rows)

def test_later_empty_master_reports_missing_required_field(self):
    outcome = run_bootstrap_import(self.empty_later_master, self.onboarding_case)
    self.assertTrue(any(issue.code == "MISSING_REQUIRED_FIELD" for issue in outcome.issues))
```

- [ ] **Step 2: Run tests to verify RED**

Run: `$env:PYTHONUTF8='1'; & $env:REGULARIZACION_PYTHON -m unittest tests.test_community_onboarding tests.test_excel_bootstrap_importer -q`

Expected: FAIL because the current profile shares one selection and bootstrap exempts a later empty workbook too broadly.

- [ ] **Step 3: Implement canonical publication**

```python
MODULE_CONCEPTS = {
    "ACS": ("acs_fixed", "acs_variable"),
    "CALEFACCION": ("heating_fixed", "heating_variable"),
}

def _is_fresh_onboarding_template(path: Path, profile: ExcelProfile) -> bool:
    """Recognize only the freshly installed canonical template state."""
```

Build fixed/equal and variable/consumption concepts with existing names,
source contracts and parameter cells. Store per-module bindings and invoice
decisions in `onboarding_configuration`; validate them in `excel_profiles`.
Use the initial-template recognizer instead of a profile-wide empty-cell
exception. Ensure later masters emit `MISSING_REQUIRED_FIELD` and keep public
profiles unchanged.

- [ ] **Step 4: Run tests to verify GREEN**

Run: `$env:PYTHONUTF8='1'; & $env:REGULARIZACION_PYTHON -m unittest tests.test_community_onboarding tests.test_excel_bootstrap_importer tests.test_excel_export_service -q`

Expected: PASS.

- [ ] **Step 5: Commit**

Commit message: `feat: genera excel canonico desde facturas y lecturas`

### Task 3: Resumen verificable y manual operativo

**Files:**
- Modify: `core/expedient_ui.py`
- Modify: `GUIA_PROCESAR_TODO.txt`
- Modify: `docs/validacion-expedientes.md`
- Modify: `tests/test_expedient_ui.py`

**Interfaces:**
- Consume la configuración normalizada de Task 1.
- Produce una vista de resumen por módulo sin crear controles Tk en las pruebas.

- [ ] **Step 1: Write failing tests**

```python
def test_onboarding_summary_exposes_invoice_and_reading_decisions_per_module(self):
    summary = expedient_ui.onboarding_summary_data(self.combined_configuration)
    self.assertEqual("Lectura ACS", summary["modules"]["ACS"]["reading_column"])
    self.assertEqual("Lectura Calefacción",
                     summary["modules"]["CALEFACCION"]["reading_column"])
    self.assertIn("provider", summary["invoices"][0])

def test_onboarding_summary_marks_required_confirmation_before_publication(self):
    summary = expedient_ui.onboarding_summary_data(self.incomplete_configuration)
    self.assertFalse(summary["ready_to_publish"])
```

- [ ] **Step 2: Run tests to verify RED**

Run: `$env:PYTHONUTF8='1'; & $env:REGULARIZACION_PYTHON -m unittest tests.test_expedient_ui -q`

Expected: FAIL because the current summary omits confirmed reading and invoice decisions.

- [ ] **Step 3: Implement presentation and documentation**

```python
def onboarding_summary_data(configuration: OnboardingConfiguration) -> dict[str, object]:
    return {
        "modules": {binding.module: {"reading_column": binding.column}
                    for binding in configuration.reading_bindings},
        "invoices": list(configuration.invoice_decisions),
        "ready_to_publish": _configuration_is_complete(configuration),
    }
```

Render that data in the existing v3 summary, without recomputing decisions in
the dialog. Document how to launch the app, select owner list/readings/invoices,
answer confirmation screens, import the master, generate the official Excel,
calculate distribution and generate letters. Explain typical manual
confirmation cases and that the source files remain local.

- [ ] **Step 4: Run tests to verify GREEN**

Run: `$env:PYTHONUTF8='1'; & $env:REGULARIZACION_PYTHON -m unittest tests.test_expedient_ui tests.test_community_onboarding -q`

Expected: PASS.

- [ ] **Step 5: Run full verification and commit**

Run: `$env:PYTHONUTF8='1'; & $env:REGULARIZACION_PYTHON -m unittest discover -s tests -t . -q`

Expected: PASS.

Run: `& $env:REGULARIZACION_PYTHON -m compileall -q core tests`

Expected: exit 0.

Run: `git diff --check`

Expected: no output.

Commit message: `docs: explica uso de facturas y lecturas`

## Self-review

- Task 1 implements source separation, all uncertainty questions and independent module readings.
- Task 2 consumes the confirmed configuration through profile, template, export, reparto and bootstrap validation.
- Task 3 exposes exactly that configuration before publication and documents activation/use.
- No task relies on a manual Excel, fixed community, fixed provider or office identity.
