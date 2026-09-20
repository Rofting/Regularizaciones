# Cálculo definitivo, coeficientes y cartas Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Convertir lecturas, facturas y participaciones en un reparto reproducible, un Excel oficial y cartas de una página, conservando observaciones e historiales.

**Architecture:** La lectura recibida permanece inmutable en `reading_observations` y un proyector cronológico publica la lectura efectiva en `lecturas_vecino`. El reparto mantiene la vista canónica actual y crea una instantánea versionada con coeficiente registral, base elegible y participación aplicada. Las cartas consumen esa instantánea, agrupan conceptos técnicos y generan cinco bandas configurables.

**Tech Stack:** Python 3, SQLite, `unittest`, CustomTkinter, openpyxl, python-docx y matplotlib.

**Spec:** `docs/superpowers/specs/2026-09-20-calculo-definitivo-coeficientes-y-cartas-design.md`

## Global Constraints

- Conservar todos los cambios locales existentes y no restaurar bases de datos, perfiles ni plantillas.
- Guardar cada valor original antes de proyectar una lectura efectiva.
- No generar consumo negativo ni inventar una lectura inicial.
- Reconciliar los importes al céntimo.
- Mantener una sola página por carta y la identidad visual existente.
- La comunidad define sus conceptos activos; las filas vacías no se muestran.

---

### Task 1: Proyección cronológica de lecturas

**Files:**
- Modify: `core/importar_lecturas_metrigest.py`
- Modify: `core/document_review.py`
- Test: `tests/test_unified_ingestion_regressions.py`

**Interfaces:**
- Consumes: observaciones `{owner_id, service, reading_date, observed_value}` y lecturas canónicas previas.
- Produces: `project_effective_reading(...) -> EffectiveReadingProjection`, lecturas `real` o `estimado/carry_forward_zero`, y una sola incidencia inicial sin antecedente.

- [ ] **Step 1: Write failing tests for `144, 0, 0, 160`, a positive decrease and an initial zero.**

```python
def test_sequence_144_zero_zero_160_recovers_at_160():
    self.assertEqual([144, 144, 144, 160], effective_values)
    self.assertEqual(16, effective_values[-1] - effective_values[0])

def test_positive_decrease_carries_last_reliable_without_issue():
    self.assertEqual((579, "estimado", "carry_forward_decrease"), effective)
```

- [ ] **Step 2: Run `.\.venv-fase1\Scripts\python.exe -m unittest tests.test_unified_ingestion_regressions -v` and verify the new tests fail for missing behavior.**
- [ ] **Step 3: Implement `EffectiveReadingProjection` and one chronological projection path. Zero and positive decrease carry the last reliable reading; a later value at or above it resumes as real.**
- [ ] **Step 4: Run ingestion, flow and review suites and verify they pass.**
- [ ] **Step 5: Commit with `git commit -m "feat: proyectar lecturas efectivas cronológicas"`.**

### Task 2: Participaciones registrales y aplicadas con historial

**Files:**
- Modify: `core/db_migrations.py`
- Modify: `core/source_analysis.py`
- Modify: `core/case_ingestion.py`
- Modify: `core/case_distribution.py`
- Test: `tests/test_db_migrations.py`
- Test: `tests/test_source_analysis.py`
- Test: `tests/test_case_distribution.py`

**Interfaces:**
- Consumes: columnas `Coeficiente`, `Enteros Participación` o `Participación` y propietarios elegibles.
- Produces: migración 13, `distribution_runs`, `owner_distribution_snapshots` y participación aplicada histórica.

- [ ] **Step 1: Write failing tests for the 644 header alias, coefficient conflicts and `2.026 / 80.0002`.**
- [ ] **Step 2: Run the migration/source/distribution tests and verify those tests fail.**
- [ ] **Step 3: Add migration 13. A distribution run stores input hash/status; each snapshot stores raw coefficient, eligible total, applied ratio, consumption and cents.**
- [ ] **Step 4: Normalize accents and spaces in owner headers and accept `enteros participacion`. Keep an existing coefficient when a source differs by more than `0.0001` and create one grouped review.**
- [ ] **Step 5: Persist `raw`, `eligible_total` and `raw / eligible_total` for coefficient concepts, then run focused tests.**
- [ ] **Step 6: Commit with `git commit -m "feat: versionar participaciones aplicadas"`.**

### Task 3: Coste real derivado y conceptos activos

**Files:**
- Create: `core/concept_pricing.py`
- Modify: `core/case_distribution.py`
- Modify: `core/excel_profiles.py`
- Test: `tests/test_concept_pricing.py`
- Test: `tests/test_case_distribution.py`

**Interfaces:**
- Consumes: `invoice_components`, consumo efectivo, servicio y fechas del expediente.
- Produces: `derive_unit_price(connection, case, concept) -> DerivedConceptPrice`.

- [ ] **Step 1: Write failing tests for net unit price, zero consumption and missing invoice component.**
- [ ] **Step 2: Run `tests.test_concept_pricing` and verify the API is absent.**
- [ ] **Step 3: Add `derived_invoices.<service>.<fixed|variable>` as an explicit profile source while keeping `period_parameters.*` compatible. Prorate overlapping invoices by day.**
- [ ] **Step 4: Allocate calculated cents with the existing largest-remainder function and reconcile to the selected invoices.**
- [ ] **Step 5: Run pricing, distribution and reconciliation tests.**
- [ ] **Step 6: Commit with `git commit -m "feat: derivar costes variables desde facturas"`.**

### Task 4: Datos de carta agrupados y coeficientes visibles

**Files:**
- Modify: `core/case_letter_service.py`
- Modify: `core/carta_writer.py`
- Test: `tests/test_case_letter_service.py`
- Test: `tests/test_carta_writer_new_results.py`

**Interfaces:**
- Consumes: instantáneas y conceptos atómicos activos.
- Produces: filas `ACS`, `Calefacción`, `Cuota fija`, `participacion_registral` y `coeficiente_aplicado`.

- [ ] **Step 1: Write failing tests that four technical concepts become exactly three ordinary rows and both percentages reach the writer.**
- [ ] **Step 2: Run letter tests and verify they fail on atomic rows/missing applied ratio.**
- [ ] **Step 3: Add deterministic grouping. Fixed ACS and heating sum into `Cuota fija`; absent services produce no blank row; extras remain separate.**
- [ ] **Step 4: Render both participation labels without changing the corporate layout and run letter tests.**
- [ ] **Step 5: Commit with `git commit -m "feat: agrupar conceptos y mostrar participación aplicada"`.**

### Task 5: Cinco bandas y gráficas con datos efectivos

**Files:**
- Modify: `core/consumption_charts.py`
- Modify: `core/case_letter_service.py`
- Test: `tests/test_consumption_charts.py`
- Test: `tests/test_case_letter_service.py`

**Interfaces:**
- Consumes: consumo efectivo del resultado, histórico y límites del perfil.
- Produces: cinco bandas `Muy bajo`, `Bajo`, `Medio`, `Alto`, `Muy alto`.

- [ ] **Step 1: Write failing tests for exactly five ACS bands `[10,20,30,40]` and unavailable comparison without valid neighbors.**
- [ ] **Step 2: Run chart tests and verify dynamic-band behavior fails the new contract.**
- [ ] **Step 3: Render configured five-band charts and build history from `owner_concept_results.consumption`, never from negative raw differences.**
- [ ] **Step 4: Run chart and letter suites.**
- [ ] **Step 5: Commit with `git commit -m "feat: generar gráficas de consumo en cinco bandas"`.**

### Task 6: Repetición de pasos, invalidación y verificación integral

**Files:**
- Modify: `core/case_workflow_actions.py`
- Modify: `core/expedient_ui.py`
- Modify: `core/excel_export_service.py`
- Test: `tests/test_case_workflow_actions.py`
- Test: `tests/test_expedient_ui.py`
- Test: `tests/test_onboarding_final_flow.py`

**Interfaces:**
- Consumes: nuevas ejecuciones de Excel, reparto y cartas.
- Produces: acciones repetibles que invalidan sólo salidas posteriores y conservan ejecuciones previas.

- [ ] **Step 1: Write failing workflow tests for regenerating Excel after delivery and recalculating after letters.**
- [ ] **Step 2: Run workflow/UI tests and verify backward execution is blocked.**
- [ ] **Step 3: Permit explicit reruns without deleting run rows or files; update the case status to the latest valid stage.**
- [ ] **Step 4: Run all affected suites.**
- [ ] **Step 5: Run `.\.venv-fase1\Scripts\python.exe -m unittest discover -s tests -v`; record the pre-existing OOXML validator failure separately.**
- [ ] **Step 6: Commit with `git commit -m "feat: completar flujo reproducible de regularización"`.**
