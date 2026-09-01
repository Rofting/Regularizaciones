# Excel maestro, reparto y cartas Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (- [ ]) syntax for tracking.

**Goal:** Convert the validated sources of a regularization case into an audited official Excel workbook, cent-exact owner distribution, and individual Word letters for community 658.

**Architecture:** A versioned profile maps normalized database data to an immutable community workbook template. The initial 658 workbook, owner list, and reading files bootstrap the same normalized records that later PDF ingestion will write. A gated orchestration service validates the case, exports and recalculates the workbook, allocates every active concept in cents, reconciles results, then generates letters only when every prior stage is valid.

**Tech Stack:** Python 3, SQLite, existing CustomTkinter UI, existing openpyxl/OOXML workbook handling, LibreOffice headless recalculation, python-docx, existing consumption-chart helpers, unittest.

**Spec:** docs/superpowers/specs/2026-08-30-generacion-excel-maestro-design.md

## Global Constraints

- The first real profile is 658_acs_v1: GAS, ELECTRICIDAD, AGUA, OTROS GASTOS, ACS; no heating module.
- The real 658 files are private inputs and must never be committed to Git or copied into test fixtures.
- Bootstrap sources are the real 658 study workbook plus its owner list and reading files; every source is archived and traced to its case.
- A case must have exactly one linked logical period before export, distribution, or letters.
- A source with open review issues blocks official Excel, distribution, and letters.
- Generated output must be temporary-first, recalculated, structurally validated, reconciled to cents, then atomically published with a backup.
- The exact workbook template is never modified; only a temporary copy is written.
- Repartition uses configured active concepts, deterministic largest-remainder cent allocation, and approved reset-counter estimations only.
- Letters are generated locally but never emailed by this delivery.
- Preserve the current v3 palette, typography, shadows, progress feedback, and CustomTkinter compatibility. Do not pass justify to CTkButton.
- Do not launch Tk headlessly. GUI acceptance is a manual Windows sequence.
- Keep legacy Excel, reparto, and letter entry points compatible while the new case-based services take over.
- Before creating or editing a workbook in tests or validation, inspect its structure and preserve its existing formatting conventions.

---

## File Structure

| File | Responsibility |
| --- | --- |
| core/db_migrations.py | Schema version 3: case-to-period link, template profile registry, invoice components, period parameters, export and letter audit runs. |
| core/excel_profiles.py | Immutable profile dataclasses, JSON loading, template hashing, and required-module validation. |
| config/excel_profiles/658_acs_v1.json | Declarative 658 sheet anchors, table columns, active concepts, validations, and output rules. |
| core/excel_bootstrap_importer.py | Imports the real master Excel plus private companion sources into normal database tables with source-cell audit data and review issues. |
| core/excel_export_service.py | Builds the official temporary workbook from one ready case, writes profile rows, validates, recalculates, backs up, and publishes. |
| core/office_recalculation.py | Locates and invokes LibreOffice headlessly with a bounded timeout and understandable failure messages. |
| core/excel_validation.py | Structural fingerprint, formula-error scan, input totals, and publication preconditions. |
| core/case_distribution.py | Calculates per-owner concept results from a validated exported case and reconciles every concept. |
| core/case_letter_service.py | Generates and audits all letters for a reconciled case using the existing Word and chart code. |
| core/app.py | Connects profile/bootstrap/export/distribution/letter actions to the guided UI with incremental progress. |
| core/expedient_service.py | Creates or associates the logical period when a case is created and exposes the linked case state. |
| core/carta_writer.py | Receives case-based results while preserving its existing public batch API. |
| tests/test_excel_profiles.py | Profile schema, required sheets, hash, and active-module tests. |
| tests/test_excel_bootstrap_importer.py | Idempotent sanitized workbook import, source traceability, and issue creation tests. |
| tests/test_excel_export_service.py | Atomic export, backup, structural validation, blocking behavior, and cents reconciliation tests. |
| tests/test_case_distribution.py | Active concepts, fixed/variable allocation, credit, counter reset, and rounding tests. |
| tests/test_case_letter_service.py | Case-gated batch letter generation, optional concepts, and per-owner error isolation tests. |
| tests/test_db_migrations.py | Version 3 tables, columns, indexes, and idempotence checks. |
| docs/validacion-excel-reparto-cartas.md | Private-data-safe manual validation sequence for the real 658 sources. |

## Shared Interfaces

Task 1 creates the public contracts used by later tasks:

```python
@dataclass(frozen=True)
class ExcelProfile:
    key: str
    version: str
    community_code: str
    template_relative_path: str
    active_modules: tuple[str, ...]
    required_sheets: tuple[str, ...]
    required_formula_cells: tuple[tuple[str, str], ...]
    concepts: tuple["ConceptRule", ...]

@dataclass(frozen=True)
class ConceptRule:
    key: str
    allocation_method: str
    actual_source: str
    billed_source: str | None
    required: bool

@dataclass(frozen=True)
class ExportResult:
    id_export_run: int
    output_path: Path
    backup_path: Path | None
    stage: str
    input_hash: str

class ExportBlockedError(ValueError):
    pass

def load_profile(profile_key: str, project_root: Path) -> ExcelProfile: ...
def link_case_to_period(connection: sqlite3.Connection, id_case: int) -> int: ...
```

Task 4 creates the export entry point:

```python
def generate_official_excel(
    connection: sqlite3.Connection,
    *,
    id_case: int,
    project_root: Path,
    output_root: Path,
    progress: Callable[[str, dict], None] | None = None,
) -> ExportResult: ...
```

Task 5 creates the distribution entry point:

```python
@dataclass(frozen=True)
class DistributionResult:
    id_case: int
    id_periodo: int
    concept_totals_cents: dict[str, int]
    owner_result_count: int

def calculate_case_distribution(
    connection: sqlite3.Connection,
    *,
    id_case: int,
    progress: Callable[[str, dict], None] | None = None,
) -> DistributionResult: ...
```

Task 6 creates the letter entry point:

```python
@dataclass(frozen=True)
class LetterBatchResult:
    id_letter_run: int
    output_directory: Path
    generated_count: int
    failures: tuple[str, ...]

def generate_case_letters(
    database_path: str | Path,
    *,
    id_case: int,
    project_root: Path,
    progress: Callable[[str, dict], None] | None = None,
) -> LetterBatchResult: ...
```

### Task 1: Schema v3, periods, and declarative Excel profiles

**Files:**
- Create: core/excel_profiles.py
- Create: config/excel_profiles/658_acs_v1.json
- Modify: core/db_migrations.py
- Modify: core/expedient_models.py
- Modify: core/expedient_service.py
- Modify: tests/test_db_migrations.py
- Create: tests/test_excel_profiles.py

**Consumes:** schema version 2, regularization_cases, periodos, and the 658 workbook layout named in the spec.

**Produces:** schema version 3, ExcelProfile, ConceptRule, a case-period link, and profile configuration consumed by Tasks 2–7.

- [ ] **Step 1: Write failing migration and profile tests**

```python
def test_migration_v3_links_case_to_period_and_creates_export_audit_tables(self):
    migrate(connection)
    tables = {row[0] for row in connection.execute(
        "SELECT name FROM sqlite_master WHERE type='table'"
    )}
    self.assertTrue({
        "invoice_components", "period_parameters",
        "excel_template_profiles", "excel_export_runs",
        "letter_generation_runs", "generated_letters",
    }.issubset(tables))
    columns = {row[1] for row in connection.execute(
        "PRAGMA table_info(regularization_cases)"
    )}
    self.assertIn("id_periodo", columns)

def test_658_profile_declares_only_its_active_modules(tmp_path):
    profile = load_profile("658_acs_v1", project_root)
    self.assertEqual(("GAS", "ELECTRICIDAD", "AGUA", "OTROS_GASTOS", "ACS"),
                     profile.active_modules)
    self.assertNotIn("CALEFACCION", profile.active_modules)
    self.assertIn(("ANALISIS", "H23"), profile.required_formula_cells)
```

- [ ] **Step 2: Run the tests and verify RED**

Run:
```powershell
$env:PYTHONUTF8='1'; & '<project-root>\.venv-fase1\Scripts\python.exe' -m unittest tests.test_db_migrations tests.test_excel_profiles -v
```

Expected: profile import fails and the migration assertions show that version 3 tables or the id_periodo column do not exist.

- [ ] **Step 3: Implement the version 3 migration and profile loader**

Add migration 3 with these exact persistent contracts:

```sql
ALTER TABLE regularization_cases
    ADD COLUMN id_periodo INTEGER REFERENCES periodos(id_periodo);

CREATE UNIQUE INDEX idx_case_period
    ON regularization_cases(id_comunidad, id_periodo)
    WHERE id_periodo IS NOT NULL;

CREATE TABLE invoice_components (
    id_component INTEGER PRIMARY KEY AUTOINCREMENT,
    id_factura INTEGER NOT NULL REFERENCES facturas(id_factura) ON DELETE CASCADE,
    component_key TEXT NOT NULL,
    amount REAL NOT NULL,
    unit TEXT,
    source_sheet TEXT,
    source_cell TEXT,
    UNIQUE(id_factura, component_key)
);

CREATE TABLE period_parameters (
    id_parameter INTEGER PRIMARY KEY AUTOINCREMENT,
    id_comunidad INTEGER NOT NULL REFERENCES comunidades(id_comunidad),
    id_periodo INTEGER NOT NULL REFERENCES periodos(id_periodo),
    parameter_key TEXT NOT NULL,
    numeric_value REAL,
    text_value TEXT,
    unit TEXT,
    source_sheet TEXT,
    source_cell TEXT,
    UNIQUE(id_comunidad, id_periodo, parameter_key)
);
```

Also create excel_template_profiles, excel_export_runs, letter_generation_runs, and generated_letters with foreign keys, SHA-256 input/template fields, path, timestamps, and CHECK-constrained lifecycle states. The profile loader must reject missing JSON fields, non-relative template paths, duplicate concept keys, unsupported allocation methods, and absent required sheets.

Add link_case_to_period so it reads the case name/dates, finds an existing period with the same community and dates, otherwise inserts it, then stores the resulting id_periodo on the case. A same-name period with incompatible dates raises ValueError instead of silently reusing it.

- [ ] **Step 4: Run tests and verify GREEN**

Run the Task 1 tests again and then:

```powershell
$env:PYTHONUTF8='1'; & '<project-root>\.venv-fase1\Scripts\python.exe' -m unittest discover -s tests -t . -q
```

Expected: all tests pass; migration is idempotent; profile has exactly the five 658 modules; case-to-period linking preserves dates.

- [ ] **Step 5: Commit**

```powershell
git add core/db_migrations.py core/excel_profiles.py core/expedient_models.py core/expedient_service.py config/excel_profiles/658_acs_v1.json tests/test_db_migrations.py tests/test_excel_profiles.py
git commit -m "Añade perfiles Excel y auditoría de exportación"
```

### Task 2: Bootstrap auditable del modelo 658 y fuentes de reparto

**Files:**
- Create: core/excel_bootstrap_importer.py
- Modify: core/expedient_service.py
- Modify: core/document_review.py
- Modify: tests/helpers.py
- Create: tests/test_excel_bootstrap_importer.py

**Consumes:** Task 1 profile and migration contracts; source_documents; facturas; invoice_components; period_parameters; propietarios; lecturas_vecino.

**Produces:** import_master_excel and BootstrapImportResult. Task 4 reads only this normalized data; Tasks 5–6 read its period, owners, and readings.

- [ ] **Step 1: Write failing behavior tests with a sanitized workbook**

Create a test workbook with the exact sheet names and header labels of profile 658, but invented community, supplier, owner, invoice, and reading values. Include a GAS row, ELECTRICIDAD row, AGUA components, one other expense, two ACS readings, and one intentionally empty required water component.

```python
def test_bootstrap_import_is_idempotent_and_traces_each_value(self):
    first = import_master_excel(connection, id_case=case_id,
                                workbook_path=self.workbook, profile=profile,
                                actor="Prueba")
    second = import_master_excel(connection, id_case=case_id,
                                 workbook_path=self.workbook, profile=profile,
                                 actor="Prueba")
    self.assertEqual(first.id_batch, second.id_batch)
    self.assertEqual(3, connection.execute(
        "SELECT COUNT(*) FROM facturas"
    ).fetchone()[0])
    self.assertGreater(connection.execute(
        "SELECT COUNT(*) FROM source_values WHERE id_batch=?", (first.id_batch,)
    ).fetchone()[0], 8)

def test_missing_required_component_creates_issue_and_blocks_case(self):
    result = import_master_excel(connection, id_case=case_id,
                                 workbook_path=self.workbook, profile=profile,
                                 actor="Prueba")
    self.assertGreater(result.open_issue_count, 0)
    self.assertEqual("under_review", get_case(connection, case_id).state.value)
```

- [ ] **Step 2: Run the focused test module and verify RED**

Run:
```powershell
$env:PYTHONUTF8='1'; & '<project-root>\.venv-fase1\Scripts\python.exe' -m unittest tests.test_excel_bootstrap_importer -v
```

Expected: ModuleNotFoundError for excel_bootstrap_importer or missing import_master_excel.

- [ ] **Step 3: Implement bootstrap parsing and review integration**

Implement:

```python
@dataclass(frozen=True)
class BootstrapImportResult:
    id_batch: int
    id_periodo: int
    imported_invoice_count: int
    imported_owner_count: int
    imported_reading_count: int
    open_issue_count: int

def import_master_excel(
    connection: sqlite3.Connection,
    *,
    id_case: int,
    workbook_path: Path,
    profile: ExcelProfile,
    actor: str,
) -> BootstrapImportResult:
    ...
```

The importer must archive and hash the workbook as a source document before persistence; call link_case_to_period; locate every sheet by profile name and headers by label; write source_values with sheet/cell coordinates; upsert invoice rows and invoice components; upsert period parameters; and create review issues for blank required fields, unmatched headers, incompatible dates, or totals that do not match component sums.

Add import_companion_sources for the owner list and reading file paths. It must use the existing parsers when their formats match, then validate every active owner has period-bound start/end readings. A restart/reset must be represented as contador_averiado plus a review issue until an estimation is approved; never turn a negative difference into a real consumption.

The hash duplicate path must return the original import batch and must not insert duplicate invoices, readings, source values, or issues. Preserve resolved corrections and all existing owner records.

- [ ] **Step 4: Run tests and verify GREEN**

Run the bootstrap module, then the full test suite. Add one test where the same spreadsheet bytes are imported twice and one where a changed workbook hash adds only its new values.

Expected: idempotent import passes, open issue blocks readiness, and no personal data appears in repository fixtures.

- [ ] **Step 5: Commit**

```powershell
git add core/excel_bootstrap_importer.py core/expedient_service.py core/document_review.py tests/helpers.py tests/test_excel_bootstrap_importer.py
git commit -m "Importa fuentes maestras de Excel de forma auditable"
```

### Task 3: Validación de libro, recalculado y exportación oficial atómica

**Files:**
- Create: core/office_recalculation.py
- Create: core/excel_validation.py
- Create: core/excel_export_service.py
- Modify: core/excel_generator.py
- Create: tests/test_excel_export_service.py

**Consumes:** Task 1 profiles and export-run tables; Task 2 normalized inputs and closed review issues.

**Produces:** generate_official_excel and ExportResult. Task 5 requires a validated export run for its case.

- [ ] **Step 1: Write failing export tests**

Use the sanitized 658-shape workbook from Task 2 as a template. Populate a ready case with matched invoices, components, parameters, owners, and readings.

```python
def test_export_writes_profile_data_preserves_template_and_creates_backup(self):
    result = generate_official_excel(connection, id_case=self.ready_case,
                                     project_root=self.project_root,
                                     output_root=self.output_root)
    self.assertTrue(result.output_path.exists())
    self.assertTrue(result.backup_path is None)
    self.assertEqual("validated", export_run_status(connection, result.id_export_run))
    self.assert_required_sheets_formulas_and_print_areas(
        result.output_path, self.profile
    )

def test_export_refuses_open_issue_without_replacing_last_valid_book(self):
    old_bytes = self.current_output.read_bytes()
    with self.assertRaisesRegex(ExportBlockedError, "incidencia"):
        generate_official_excel(connection, id_case=self.case_with_issue,
                                project_root=self.project_root,
                                output_root=self.output_root)
    self.assertEqual(old_bytes, self.current_output.read_bytes())
```

Add a test that makes the recalculator report a formula error and assert no official output replacement. Add one that writes a second valid output and confirms a timestamped backup contains the preceding bytes.

- [ ] **Step 2: Run export tests and verify RED**

Run:
```powershell
$env:PYTHONUTF8='1'; & '<project-root>\.venv-fase1\Scripts\python.exe' -m unittest tests.test_excel_export_service -v
```

Expected: ModuleNotFoundError or missing generation function.

- [ ] **Step 3: Implement workbook validation, recalculation, and exporter**

Implement a Recalculator interface with a production LibreOffice implementation and a deterministic test double:

```python
class WorkbookRecalculator(Protocol):
    def recalculate(self, workbook_path: Path, work_directory: Path) -> None: ...

class LibreOfficeRecalculator:
    def recalculate(self, workbook_path: Path, work_directory: Path) -> None:
        # run soffice --headless with a bounded timeout; raise RecalculationError on failure
        ...

def validate_workbook(path: Path, profile: ExcelProfile,
                      expected_totals_cents: Mapping[str, int]) -> None:
    # check sheets, print areas, declared formula cells, formula error cells,
    # profile input totals, and output totals
    ...
```

The exporter must load one immutable registered template into a temporary sibling path, clear only profile-declared mutable row ranges, write typed dates/numbers and formula families, preserve all unrelated OOXML formatting, invoke the recalculator, validate, create an export-run record, then atomically replace Excels_Maestros/Comunidad_<codigo>.xlsx. If a predecessor exists, copy it to backups before replacement. On every exception, mark the run failed and leave the predecessor untouched.

Use the 658 profile to write these semantic inputs: DATOS metadata; GAS invoice entries and cost terms; ELECTRICIDAD entries; AGUA provider variable/fixed components; OTROS GASTOS; and aggregate ACS readings/parameters. Derive formulas from profile formula patterns rather than copying values cached in the real workbook.

Adapt the legacy regenerar_excel_comunidad entry point to call the new service when an eligible case is supplied; preserve its old community-only behavior for legacy callers.

- [ ] **Step 4: Run tests and verify GREEN**

Run export tests, then:

```powershell
$env:PYTHONUTF8='1'; & '<project-root>\.venv-fase1\Scripts\python.exe' -m unittest discover -s tests -t . -q
$env:PYTHONUTF8='1'; & '<project-root>\.venv-fase1\Scripts\python.exe' -m py_compile core\office_recalculation.py core\excel_validation.py core\excel_export_service.py core\excel_generator.py
```

Expected: all workbook data are numeric/date typed, a failing workbook cannot replace official output, and a valid replacement produces a backup.

- [ ] **Step 5: Commit**

```powershell
git add core/office_recalculation.py core/excel_validation.py core/excel_export_service.py core/excel_generator.py tests/test_excel_export_service.py
git commit -m "Genera Excel oficial desde expedientes validados"
```

### Task 4: Reparto final configurable y conciliado por propietario

**Files:**
- Create: core/case_distribution.py
- Modify: core/regularization_service.py
- Modify: core/reconciliation.py
- Modify: core/motor_reparto.py
- Create: tests/test_case_distribution.py

**Consumes:** Task 1 ConceptRule and case-period link; Task 2 owner/readings data; Task 3 validated export run.

**Produces:** calculate_case_distribution, DistributionResult, owner_concept_results for every active concept, and exact reconciliation records consumed by Task 6.

- [ ] **Step 1: Write failing distribution tests**

```python
def test_distribution_allocates_fixed_variable_and_credit_to_the_cent(self):
    result = calculate_case_distribution(connection, id_case=self.ready_case)
    rows = results_by_concept(connection, self.period_id)
    self.assertEqual(12_001, sum(row.actual_cents for row in rows["acs_fixed"]))
    self.assertEqual(30_000, sum(row.actual_cents for row in rows["acs_variable"]))
    self.assertEqual(-301, sum(row.actual_cents for row in rows["credit"]))
    self.assertEqual("cuadrado", reconciliation_status(connection, self.period_id, "acs_variable"))

def test_distribution_blocks_unresolved_counter_reset(self):
    mark_reading_as_reset_without_approved_estimation(connection, self.owner_id)
    with self.assertRaisesRegex(DistributionBlockedError, "contador"):
        calculate_case_distribution(connection, id_case=self.ready_case)
```

Add tests for zero total weights, one inactive owner, owner ordering that exercises largest-remainder tie-breaking, a concept disabled by the profile, and a calculated distribution rerun that updates rather than duplicates rows.

- [ ] **Step 2: Run distribution tests and verify RED**

Run:
```powershell
$env:PYTHONUTF8='1'; & '<project-root>\.venv-fase1\Scripts\python.exe' -m unittest tests.test_case_distribution -v
```

Expected: ModuleNotFoundError for case_distribution or missing calculate_case_distribution.

- [ ] **Step 3: Implement the generic concept distribution service**

Implement:

```python
class DistributionBlockedError(ValueError):
    pass

def allocate_concept_cents(
    total_cents: int,
    owner_weights: Mapping[int, Decimal],
) -> dict[int, int]:
    # Use deterministic largest remainder; sum(result.values()) == total_cents.
    ...

def calculate_case_distribution(
    connection: sqlite3.Connection,
    *,
    id_case: int,
    progress: Callable[[str, dict], None] | None = None,
) -> DistributionResult:
    ...
```

Require a validated export run and no open case issues. For every active profile concept, obtain total actual/billed cents from normalized inputs, select only active owners, build weights from the declared method, allocate both amounts independently, and upsert owner_concept_results. Reconcile every concept and the total general in the same transaction. Store a projection in legacy repartos only for ACS and CALEFACCION fields that existing letters still consume.

Counter rules: real monotonic readings use their difference; a reset/repair must use an approved manual estimation; missing or unapproved readings block the entire distribution. Never infer a negative consumption as zero.

- [ ] **Step 4: Run tests and verify GREEN**

Run the focused distribution tests and the existing tests for regularization_service, reconciliation, carta_writer result lookup, and motor_reparto.

Expected: every concept total matches exactly in cents, disabled concepts generate no result rows, and blocked input creates no partial distribution.

- [ ] **Step 5: Commit**

```powershell
git add core/case_distribution.py core/regularization_service.py core/reconciliation.py core/motor_reparto.py tests/test_case_distribution.py
git commit -m "Calcula reparto conciliado por concepto y propietario"
```

### Task 5: Cartas por expediente reconciliado

**Files:**
- Create: core/case_letter_service.py
- Modify: core/carta_writer.py
- Modify: core/letter_settings.py
- Modify: tests/test_carta_writer_new_results.py
- Create: tests/test_case_letter_service.py

**Consumes:** Task 4 canonical owner_concept_results and reconciliations; existing Word template and consumption_charts helpers.

**Produces:** generate_case_letters, LetterBatchResult, letter_generation_runs, generated_letters, and portable community identity fields.

- [ ] **Step 1: Write failing letter-service tests**

```python
def test_case_letters_use_active_concepts_and_record_each_owner(self):
    result = generate_case_letters(self.database_path, id_case=self.reconciled_case,
                                   project_root=self.project_root)
    self.assertEqual(2, result.generated_count)
    self.assertEqual(0, len(result.failures))
    rows = generated_letter_rows(self.connection, result.id_letter_run)
    self.assertEqual({"generated"}, {row["status"] for row in rows})

def test_case_letters_refuse_unreconciled_results(self):
    with self.assertRaisesRegex(LetterGenerationBlockedError, "concili"):
        generate_case_letters(self.database_path, id_case=self.unreconciled_case,
                              project_root=self.project_root)
```

Add a test with ACS fixed/variable and credit but no heating, asserting the output data excludes heating. Add a test that makes one owner template write fail and verifies the other letter is retained while the run is incomplete.

- [ ] **Step 2: Run letter tests and verify RED**

Run:
```powershell
$env:PYTHONUTF8='1'; & '<project-root>\.venv-fase1\Scripts\python.exe' -m unittest tests.test_case_letter_service -v
```

Expected: ModuleNotFoundError for case_letter_service or missing generate_case_letters.

- [ ] **Step 3: Implement case letter orchestration and adaptable data mapping**

Implement:

```python
class LetterGenerationBlockedError(ValueError):
    pass

def generate_case_letters(
    database_path: str | Path,
    *,
    id_case: int,
    project_root: Path,
    progress: Callable[[str, dict], None] | None = None,
) -> LetterBatchResult:
    ...
```

Verify validated export and all reconciliation rows are cuadrado. Load active concept labels and owner results from the case period, then adapt them to the existing carta_writer contract. The adapter must omit inactive services, pass ACS historical consumption and neighbor distribution for the existing charts, show fixed and variable values separately, and use community-configured office name/logo/signature instead of any hard-coded office identity.

Create output at salidas/cartas/<codigo>/<periodo>/, create one generated_letters audit row per owner before writing, mark success or failure per row, and mark the parent letter run completed only when all rows succeed. Preserve the existing one-page document design, color/shadow system, and graph placement; do not send email.

- [ ] **Step 4: Run tests and verify GREEN**

Run the focused letter tests, then render one sanitized generated document using the repository document test tooling or existing renderer. Verify it remains one page and charts/tables do not overflow.

Expected: full batch success is audited, a single failure is isolated and reported, and unreconciled cases generate no files.

- [ ] **Step 5: Commit**

```powershell
git add core/case_letter_service.py core/carta_writer.py core/letter_settings.py tests/test_carta_writer_new_results.py tests/test_case_letter_service.py
git commit -m "Genera cartas desde repartos conciliados"
```

### Task 6: Acciones guiadas, progreso y mensajes de bloqueo

**Files:**
- Modify: core/app.py
- Modify: core/expedient_ui.py
- Modify: tests/test_expedient_flow.py
- Create: tests/test_case_workflow_actions.py

**Consumes:** Tasks 1–5 public service interfaces.

**Produces:** UI state and controllers that make the complete flow visible and recoverable without headless Tk tests.

- [ ] **Step 1: Write failing controller tests without constructing Tk**

Extract controller functions from AppGestionFincas so they accept services and a progress callback.

```python
def test_generate_excel_action_reports_stages_and_refreshes_case_state(self):
    events = []
    result = generate_excel_action(
        services, id_case=case_id, progress=lambda stage, data: events.append(stage)
    )
    self.assertEqual(["validate_case", "prepare_template", "write_GAS",
                      "write_ELECTRICIDAD", "write_AGUA", "recalculate",
                      "reconcile", "publish"], events)
    self.assertEqual("validated", result.stage)

def test_distribution_action_does_not_run_when_export_is_not_validated(self):
    with self.assertRaisesRegex(DistributionBlockedError, "Excel"):
        calculate_distribution_action(services, id_case=case_id)
```

- [ ] **Step 2: Run controller tests and verify RED**

Run:
```powershell
$env:PYTHONUTF8='1'; & '<project-root>\.venv-fase1\Scripts\python.exe' -m unittest tests.test_case_workflow_actions -v
```

Expected: action controller functions do not yet exist.

- [ ] **Step 3: Implement guided UI integration**

Add, within the current expediente section:

- profile display and source readiness summary;
- action to import an Excel master source only for an initial bootstrap;
- Generate Excel oficial, Calcular reparto final, and Generar cartas actions;
- per-stage progress and a readable log message for each service event;
- open-result controls only after a successful published run;
- concise blocking copy that names the count and type of data still pending.

Each long action runs using the existing background-thread and finally-based restoration pattern. The active case/community ownership check must run before every action. Refresh the persistent case selector and issue tray after bootstrap, export, distribution, and letter completion. Do not reorganize the surrounding v3 design or add a second application page.

- [ ] **Step 4: Run tests and manual GUI smoke sequence**

Run full unittest and compile app.py. Then on Windows manually:

1. select community 658 and its case;
2. import private bootstrap sources into a test database;
3. confirm any missing data appears in the issue tray;
4. resolve test issues and watch Generate Excel oficial show stage progress;
5. confirm subsequent buttons unlock only after their prerequisite;
6. verify the UI retains the selected case after changing away and returning to 658.

No Tk headless launch is permitted.

- [ ] **Step 5: Commit**

```powershell
git add core/app.py core/expedient_ui.py tests/test_expedient_flow.py tests/test_case_workflow_actions.py
git commit -m "Integra Excel reparto y cartas en el flujo guiado"
```

### Task 7: Validación privada de la 658, documentación y release check

**Files:**
- Create: docs/validacion-excel-reparto-cartas.md
- Modify: GUIA_PROCESAR_TODO.txt
- Modify: tests/test_dependencies.py
- Create: tests/test_private_658_validation.py

**Consumes:** all Tasks 1–6 and the external private 658 sources.

**Produces:** a repeatable, privacy-safe release procedure and evidence that the full real community workflow can be checked before deployment.

- [ ] **Step 1: Write failing private-validation guard tests**

```python
def test_private_validation_requires_explicit_source_environment(self):
    with patch.dict(os.environ, {}, clear=True):
        with self.assertRaisesRegex(RuntimeError, "REGULARIZACION_658"):
            resolve_private_658_sources()

def test_real_source_paths_are_not_tracked_by_git(self):
    tracked = subprocess.check_output(["git", "ls-files"], text=True)
    self.assertNotIn("<private-master-name>", tracked)
    self.assertNotIn("Listado propietarios", tracked)
```

- [ ] **Step 2: Run private guard tests and verify RED**

Run:
```powershell
$env:PYTHONUTF8='1'; & '<project-root>\.venv-fase1\Scripts\python.exe' -m unittest tests.test_private_658_validation -v
```

Expected: private validation helper or documented environment contract does not exist.

- [ ] **Step 3: Implement release validation and documentation**

Document exact environment variables for the three private bootstrap sources and a separate temporary validation database/output root. The private validation runner must:

1. create a fresh database;
2. register template/profile without copying private files into Git;
3. import workbook, owner list, and readings;
4. require manual issue resolution where source values are ambiguous;
5. generate an Excel under a validation-only output root;
6. compare required sheets, formulas, print areas, and reconciled totals against the source workbook;
7. calculate distribution and compare concept totals to owner totals;
8. generate all letters;
9. render one non-identifying selected letter and verify one page;
10. print a compact success/failure report without names, emails, addresses, or account data.

The guide must distinguish this initial bootstrap from later PDF-driven runs. It must state that the user sends email manually.

- [ ] **Step 4: Run full release verification**

Run:

```powershell
$env:PYTHONUTF8='1'; & '<project-root>\.venv-fase1\Scripts\python.exe' -m unittest discover -s tests -t . -q
$env:PYTHONUTF8='1'; & '<project-root>\.venv-fase1\Scripts\python.exe' -m py_compile core\app.py core\excel_profiles.py core\excel_bootstrap_importer.py core\office_recalculation.py core\excel_validation.py core\excel_export_service.py core\case_distribution.py core\case_letter_service.py
git diff --check
```

Then run the documented private validation only in the user-approved local source folder and inspect the generated Excel plus one letter visually in Windows.

- [ ] **Step 5: Commit**

```powershell
git add docs/validacion-excel-reparto-cartas.md GUIA_PROCESAR_TODO.txt tests/test_dependencies.py tests/test_private_658_validation.py
git commit -m "Documenta validación integral de la 658"
```

## Plan Self-Review

| Spec requirement | Implementing task |
| --- | --- |
| Immutable, versioned templates and portable profiles | Task 1 |
| 658 workbook bootstrap, companion readings, audit, and review issues | Task 2 |
| Typed Excel generation, recalculation, validation, backup, and atomic publication | Task 3 |
| Configurable cent-exact owner allocation, reset counter gating, and reconciliation | Task 4 |
| One-page portable letters from reconciled owner results | Task 5 |
| Guided actions and visible staged progress | Task 6 |
| Real 658 private validation and user-facing guidance | Task 7 |

Self-review results: every specification section maps to a task; all public names used by downstream tasks are declared in Shared Interfaces; private sources remain outside version control; no GUI test requires headless Tk.
