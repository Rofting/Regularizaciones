# Ingesta unificada de fuentes y reinicio seguro — Plan de implementación

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Cargar una carpeta mixta en un expediente, clasificar y extraer cada fuente antes de crear incidencias, y permitir reiniciar la base con copia verificable.

**Architecture:** Un nuevo módulo `source_analysis` transforma cada archivo en una clasificación, candidatos normalizados y referencias de contexto. `case_ingestion` persiste ese resultado, crea solo incidencias específicas y puede reanalizar archivos ya archivados sin sobrescribir correcciones manuales. La UI solo orquesta acciones y muestra el resumen; la persistencia de facturas y lecturas se centraliza en un adaptador de aplicación.

**Tech Stack:** Python 3, SQLite, CustomTkinter, `lector_pdf`, `xlrd`, `openpyxl`, `unittest`.

**Spec:** `docs/superpowers/specs/2026-09-13-unified-source-ingestion-design.md`

## Global Constraints

- La base existente se copia y verifica antes de sustituir `data/gestion.db`.
- No se usan valores cero para suplir datos extraídos con duda.
- Las correcciones manuales confirmadas sobreviven a la reanálisis.
- PDF, XLS, XLSX y CSV son fuentes de entrada compatibles.
- El visor de Windows puede no admitir ir a página; la referencia se muestra siempre que exista.
- No se codifican reglas exclusivas de la comunidad 658.

---

## Estructura de archivos

- Crear `core/source_analysis.py`: contratos `SourceAnalysis`/`SourceLocator`, clasificación y extracción por formato.
- Crear `core/database_reset.py`: copia SQLite verificada e inicialización atómica de una base limpia.
- Modificar `core/db_migrations.py`: migración 6 para contexto de candidato y tipo/confianza de documento.
- Modificar `core/document_review.py`: incidencias de clasificación, limpieza segura de incidencias automáticas y conservación de correcciones manuales.
- Modificar `core/case_ingestion.py`: persistencia de análisis y reanálisis idempotente de archivos archivados.
- Modificar `core/expedient_ui.py`: carga automática, resumen, botón Reanalizar fuentes y contexto de incidencia.
- Modificar `core/app.py`: acción de nueva base segura, refresco de comunidad/expediente y puente a datos canónicos.
- Crear `tests/test_source_analysis.py`, `tests/test_database_reset.py`.
- Modificar `tests/test_expedient_flow.py`, `tests/test_document_review.py`, `tests/test_expedient_ui.py`.

### Task 1: Modelo de análisis de fuentes

**Files:**
- Create: `core/source_analysis.py`
- Test: `tests/test_source_analysis.py`

**Interfaces:**
- Produces: `SourceAnalysis(kind, confidence, candidates, required_fields, locator, review_message)`.
- Consumes: `Path`, opcional `pdf_processor(path, community_code)` y lectores tabulares.

- [ ] **Step 1: Write the failing tests**

```python
def test_invoice_result_requires_only_missing_invoice_fields():
    result = source_analysis.analyse_pdf(
        Path("invoice.pdf"),
        pdf_processor=lambda *_: {"ok": True, "tipo": "FACTURA", "datos": {
            "fecha_inicio": "2026-01-01", "fecha_fin": "2026-01-31", "importe_total": 42.5,
        }},
    )
    assert result.kind == "invoice"
    assert result.required_fields == ("fecha_inicio", "fecha_fin", "importe_total")
    assert result.candidates["importe_total"] == "42.5"

def test_reading_result_never_requires_invoice_fields():
    result = source_analysis.analyse_pdf(
        Path("meters.pdf"),
        pdf_processor=lambda *_: {"ok": True, "tipo": "LECTURA_METRIGEST", "datos": {"tipo": "ACS"}},
    )
    assert result.kind == "reading"
    assert result.required_fields == ()

def test_unknown_tabular_file_creates_one_classification_review():
    result = source_analysis.classify_headers(("Referencia", "Observacion"), suffix=".csv")
    assert result.kind == "unknown"
    assert result.review_message.startswith("No se ha podido identificar")
```

- [ ] **Step 2: Run the focused test to verify it fails**

Run: `..\\..\\.venv-fase1\\Scripts\\python.exe -m unittest tests.test_source_analysis -v`

Expected: FAIL because `source_analysis` and `SourceAnalysis` do not exist.

- [ ] **Step 3: Implement the minimal analyser**

```python
@dataclass(frozen=True)
class SourceAnalysis:
    kind: str
    confidence: str
    candidates: Mapping[str, str | None]
    required_fields: tuple[str, ...]
    locator: SourceLocator | None = None
    review_message: str | None = None

def analyse_source(path: Path, *, community_code: str, pdf_processor=None) -> SourceAnalysis:
    if path.suffix.lower() == ".pdf":
        return analyse_pdf(path, pdf_processor=pdf_processor, community_code=community_code)
    return analyse_tabular(path)
```

Map `FACTURA` to `invoice` and only its three canonical required fields;
map `LECTURA_METRIGEST` to `reading` with no invoice requirements. Inspect
XLS/XLSX/CSV headers for `lectura`, `contador`, `propiedad`, `vivienda` and
`propietario`; return `unknown` with one classification message if no strong
signal exists. Preserve extractor page/fragment or sheet/cell in `SourceLocator`.

- [ ] **Step 4: Run the focused tests to verify they pass**

Run: `..\\..\\.venv-fase1\\Scripts\\python.exe -m unittest tests.test_source_analysis -v`

Expected: PASS.

- [ ] **Step 5: Commit**

```powershell
git add core/source_analysis.py tests/test_source_analysis.py
git commit -m "feat: classify sources before review"
```

### Task 2: Persist classification, context and specific review issues

**Files:**
- Modify: `core/db_migrations.py`
- Modify: `core/document_review.py`
- Modify: `core/case_ingestion.py`
- Test: `tests/test_expedient_flow.py`
- Test: `tests/test_document_review.py`

**Interfaces:**
- Consumes: `SourceAnalysis` from Task 1.
- Produces: `add_analysed_document_to_case(...)` and `reanalyze_case_documents(...)`.

- [ ] **Step 1: Write the failing tests**

```python
def test_reading_document_creates_no_invoice_missing_field_issues(self):
    result = case_ingestion.add_analysed_document_to_case(
        self.connection, self.case_id, source_path=self.reading_file,
        archive_root=self.archive_root, analysis=SourceAnalysis.reading(),
    )
    self.assertEqual(0, result.open_issue_count)

def test_unknown_document_creates_one_classification_issue(self):
    result = case_ingestion.add_analysed_document_to_case(
        self.connection, self.case_id, source_path=self.unknown_file,
        archive_root=self.archive_root, analysis=SourceAnalysis.unknown(),
    )
    self.assertEqual(1, result.open_issue_count)
    self.assertEqual("DOCUMENT_CLASSIFICATION_REQUIRED", self.open_issue_code())

def test_reanalysis_keeps_manually_confirmed_candidate(self):
    document = self.add_invoice_with_manual_total("10.00")
    case_ingestion.reanalyze_case_documents(self.connection, self.case_id, analyser=self.changed_analyser)
    self.assertEqual("10.00", self.candidate_value(document.id_document, "importe_total"))
```

- [ ] **Step 2: Run focused tests to verify they fail**

Run: `..\\..\\.venv-fase1\\Scripts\\python.exe -m unittest tests.test_expedient_flow tests.test_document_review -v`

Expected: FAIL because analysed ingestion and classification issue support do not exist.

- [ ] **Step 3: Implement minimal schema and persistence**

Add migration 6:

```sql
ALTER TABLE source_documents ADD COLUMN classification_confidence TEXT;
ALTER TABLE extraction_candidates ADD COLUMN source_context TEXT;
```

Persist `kind`, confidence, candidates and compact JSON context. Add
`DOCUMENT_CLASSIFICATION_REQUIRED` as an idempotent document-level issue.
During reanalysis, only delete open automatically generated missing-field or
classification issues; retain `manual_corrections`, validated candidates and
all user-originated review outcomes.

- [ ] **Step 4: Run focused tests to verify they pass**

Run: `..\\..\\.venv-fase1\\Scripts\\python.exe -m unittest tests.test_expedient_flow tests.test_document_review -v`

Expected: PASS.

- [ ] **Step 5: Commit**

```powershell
git add core/db_migrations.py core/document_review.py core/case_ingestion.py tests/test_expedient_flow.py tests/test_document_review.py
git commit -m "feat: persist analysed sources and safe reanalysis"
```

### Task 3: Incorporar los datos confirmados a las tablas canónicas

**Files:**
- Modify: `core/case_ingestion.py`
- Modify: `core/gestor_bd.py`
- Modify: `core/importar_lecturas_metrigest.py`
- Test: `tests/test_expedient_flow.py`

**Interfaces:**
- Consumes: documento analizado y valores confirmados de Task 2.
- Produces: `apply_confirmed_source(connection, case_id, document_id)`.

- [ ] **Step 1: Write failing tests**

```python
def test_confirmed_invoice_is_available_to_case_excel_export(self):
    document = self.add_confirmed_invoice(total="128.10")
    case_ingestion.apply_confirmed_source(self.connection, self.case_id, document.id_document)
    self.assertEqual(1, self.connection.execute("SELECT COUNT(*) FROM facturas").fetchone()[0])

def test_confirmed_reading_is_available_as_reading_not_invoice(self):
    document = self.add_confirmed_reading()
    case_ingestion.apply_confirmed_source(self.connection, self.case_id, document.id_document)
    self.assertEqual(0, self.connection.execute("SELECT COUNT(*) FROM facturas").fetchone()[0])
```

- [ ] **Step 2: Run focused tests to verify they fail**

Run: `..\\..\\.venv-fase1\\Scripts\\python.exe -m unittest tests.test_expedient_flow -v`

Expected: FAIL because the application adapter does not exist.

- [ ] **Step 3: Implement the canonical-data adapter**

Use the existing invoice insertion and reading import rules rather than a
parallel schema. Write only validated data. Reuse counter-reset behavior from
the reading importer; it must retain the prior reading or create a dedicated
review issue instead of calculating a negative consumption. Mark the source
`validated` only after canonical insertion succeeds idempotently.

- [ ] **Step 4: Run focused tests to verify they pass**

Run: `..\\..\\.venv-fase1\\Scripts\\python.exe -m unittest tests.test_expedient_flow -v`

Expected: PASS.

- [ ] **Step 5: Commit**

```powershell
git add core/case_ingestion.py core/gestor_bd.py core/importar_lecturas_metrigest.py tests/test_expedient_flow.py
git commit -m "feat: apply confirmed sources to canonical data"
```

### Task 4: Reinicio seguro de la base

**Files:**
- Create: `core/database_reset.py`
- Test: `tests/test_database_reset.py`

**Interfaces:**
- Produces: `reset_database(database_path: Path, *, backup_root: Path, initialise: Callable[[Path], None]) -> ResetResult`.

- [ ] **Step 1: Write failing tests**

```python
def test_reset_copies_and_verifies_old_database_before_replacing_it(self):
    result = database_reset.reset_database(self.database, backup_root=self.backups, initialise=self.initialise)
    self.assertTrue(result.backup_path.exists())
    self.assertEqual("new", self.read_marker(self.database))

def test_reset_keeps_existing_database_when_backup_verification_fails(self):
    with self.assertRaises(database_reset.DatabaseResetError):
        database_reset.reset_database(self.database, backup_root=self.backups, initialise=self.broken_initialise)
    self.assertEqual("old", self.read_marker(self.database))
```

- [ ] **Step 2: Run the focused test to verify it fails**

Run: `..\\..\\.venv-fase1\\Scripts\\python.exe -m unittest tests.test_database_reset -v`

Expected: FAIL because `database_reset` does not exist.

- [ ] **Step 3: Implement reset with rollback protection**

Copy using SQLite backup APIs to a unique timestamped filename, open the copy
and run `PRAGMA integrity_check`, initialize a temporary replacement and run
its migrations, then replace only after both steps succeed. Return the backup
and new paths. Raise an actionable error before any replacement on failure.

- [ ] **Step 4: Run the focused tests to verify they pass**

Run: `..\\..\\.venv-fase1\\Scripts\\python.exe -m unittest tests.test_database_reset -v`

Expected: PASS.

- [ ] **Step 5: Commit**

```powershell
git add core/database_reset.py tests/test_database_reset.py
git commit -m "feat: add verified database reset"
```

### Task 5: Integrar acciones y contexto en la interfaz guiada

**Files:**
- Modify: `core/expedient_ui.py`
- Modify: `core/app.py`
- Test: `tests/test_expedient_ui.py`

**Interfaces:**
- Consumes: `analyse_source`, `reanalyze_case_documents`, `reset_database`.
- Produces: carga automática, resumen de tipos, reanálisis, contexto de incidencia y acción de nueva base.

- [ ] **Step 1: Write failing UI-behavior tests**

```python
def test_source_summary_groups_detected_documents_by_kind(self):
    summary = expedient_ui.source_summary(("invoice", "reading", "reading", "unknown"))
    self.assertEqual("1 factura detectada · 2 lecturas · 1 documento por revisar", summary)

def test_issue_context_describes_pdf_page_or_excel_cell(self):
    self.assertIn("Página 2", expedient_ui.issue_context_label('{"page": 2, "excerpt": "TOTAL"}'))
    self.assertIn("Datos!B4", expedient_ui.issue_context_label('{"sheet": "Datos", "cell": "B4"}'))
```

- [ ] **Step 2: Run focused tests to verify they fail**

Run: `..\\..\\.venv-fase1\\Scripts\\python.exe -m unittest tests.test_expedient_ui -v`

Expected: FAIL because the UI helpers do not exist.

- [ ] **Step 3: Implement the UI actions**

Replace global source type selection with automatic analysis and an optional
manual override only for `unknown`. Show grouped summary after ingestion.
Add **Reanalizar fuentes** on the expediente and a guarded **Nueva base segura**
action in Ajustes that displays the backup location on success. Add **Ver
contexto** to review dialogs; display page/fragment or sheet/cell and retain
the universal **Abrir archivo** button.

- [ ] **Step 4: Run focused tests to verify they pass**

Run: `..\\..\\.venv-fase1\\Scripts\\python.exe -m unittest tests.test_expedient_ui -v`

Expected: PASS.

- [ ] **Step 5: Commit**

```powershell
git add core/expedient_ui.py core/app.py tests/test_expedient_ui.py
git commit -m "feat: guide source review and safe reset"
```

### Task 6: End-to-end verification and documentation

**Files:**
- Modify: `README.md` or `docs/MANUAL_USO.md`
- Test: relevant existing test suites

- [ ] **Step 1: Write an end-to-end regression test for a mixed batch**

```python
def test_mixed_batch_creates_specific_reviews_not_three_per_source(self):
    result = self.ingest_fixture_folder("mixed_sources")
    self.assertEqual({"invoice": 1, "reading": 2, "unknown": 1}, result.kind_counts)
    self.assertEqual(1, result.open_issue_count)
```

- [ ] **Step 2: Run it to verify it fails before the final integration**

Run: `..\\..\\.venv-fase1\\Scripts\\python.exe -m unittest tests.test_expedient_flow -v`

Expected: FAIL until Tasks 1–5 are all wired.

- [ ] **Step 3: Complete integration and update the manual**

Document: create a community, create an expediente, add a folder, understand
the source summary, resolve a context-bearing incident, reanalyze sources,
generate Excel/reparto/cartas, and create a secure new database. Include the
fact that the old database is backed up rather than deleted.

- [ ] **Step 4: Run complete verification**

Run: `..\\..\\.venv-fase1\\Scripts\\python.exe -m unittest discover -s tests -q`

Expected: all tests pass with zero failures.

- [ ] **Step 5: Manual acceptance on a copy of the 658 sources**

1. Create a clean database with backup.
2. Create the 658 community and a dated expediente.
3. Load the mixed source folder.
4. Verify that the summary has categories and no 204 repeated invoice issues.
5. Resolve only genuine uncertainties; use context where provided.
6. Generate Excel, calculate distribution and generate letters.

- [ ] **Step 6: Commit**

```powershell
git add docs tests core
git commit -m "docs: describe unified source workflow"
```
