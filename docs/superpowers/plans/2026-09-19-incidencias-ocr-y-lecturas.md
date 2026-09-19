# Incidencias, OCR y lecturas Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Resolver incidencias repetidas, hacer accionables los grupos de contadores, conservar observaciones de lecturas y mejorar OCR y proveedores.

**Architecture:** `lecturas_vecino` sigue siendo la proyección canónica. `reading_observations` guarda de manera inmutable todo valor de origen y enlaza la decisión efectiva. Una migración permite historial cerrado de incidencias y mantiene exclusividad solo para incidencias abiertas. RapidOCR se invoca de forma opcional antes del respaldo actual de Tesseract.

**Tech Stack:** Python 3, SQLite, CustomTkinter, unittest, pdfplumber, RapidOCR, pytesseract opcional.

**Spec:** `docs/superpowers/specs/2026-09-19-incidencias-ocr-y-lecturas-design.md`

## Global Constraints

- No se borra una lectura observada, incluso si su valor es `0`.
- El arrastre solo usa una lectura anterior fiable del mismo propietario y suministro.
- Solo las incidencias abiertas son únicas; el historial resuelto se conserva.
- La interfaz nunca deja `Abrir archivo` como única acción resolutiva.
- Cada cambio empieza por una prueba que falla.

---

### Task 1: Migración de incidencias y observaciones

**Files:**

- Modify: `core/db_migrations.py`
- Modify: `core/gestor_bd.py`
- Test: `tests/test_db_migrations.py`

**Interfaces:**

- Produces `CURRENT_SCHEMA_VERSION = 10`.
- Produces `reading_observations(id_observation, id_propietario, tipo, fecha_lectura, observed_value, source_path, id_document, status, effective_reading_id, created_at)`.
- Produces `idx_review_issues_open_unique`, a partial unique index over open issue identity.

- [ ] **Step 1: Write the failing migration tests.**

```python
def test_migration_ten_allows_closed_history_but_rejects_duplicate_open_issue(self):
    db_migrations.MIGRATIONS[10](self.connection)
    issue = self.connection.execute("SELECT id_case,id_document,code,field_name FROM review_issues LIMIT 1").fetchone()
    self.connection.execute("UPDATE review_issues SET status='resolved' WHERE id_document=?", (issue["id_document"],))
    values = (issue["id_case"], issue["id_document"], issue["code"], issue["field_name"], "Nueva comprobación")
    self.connection.execute("INSERT INTO review_issues (id_case,id_document,code,field_name,message,status) VALUES (?,?,?,?,?,'open')", values)
    with self.assertRaises(sqlite3.IntegrityError):
        self.connection.execute("INSERT INTO review_issues (id_case,id_document,code,field_name,message,status) VALUES (?,?,?,?,?,'open')", values)

def test_migration_ten_creates_reading_observations(self):
    db_migrations.MIGRATIONS[10](self.connection)
    columns = {row[1] for row in self.connection.execute("PRAGMA table_info(reading_observations)")}
    self.assertTrue({"id_propietario", "tipo", "fecha_lectura", "observed_value", "status"}.issubset(columns))
```

- [ ] **Step 2: Run the migration tests.**

Run: `python -m unittest tests.test_db_migrations`

Expected: FAIL because migration 10 and its table/index do not exist.

- [ ] **Step 3: Implement migration 10.**

Drop the legacy unique index created from `UNIQUE(id_document,code,field_name,status)`, create `idx_review_issues_open_unique` with `WHERE status='open'`, and create the observation table plus a `(id_propietario,tipo,fecha_lectura)` index. Register migration 10 and update version assertions.

- [ ] **Step 4: Run the migration tests again.**

Run: `python -m unittest tests.test_db_migrations`

Expected: PASS.

- [ ] **Step 5: Commit.**

Run: `git add core/db_migrations.py core/gestor_bd.py tests/test_db_migrations.py; git commit -m "feat: conservar observaciones e historial de incidencias"`

### Task 2: Proyectar con seguridad las lecturas cero

**Files:**

- Modify: `core/importar_lecturas_metrigest.py`
- Modify: `core/case_ingestion.py`
- Test: `tests/test_unified_ingestion_regressions.py`

**Interfaces:**

- Produces `record_reading_observation(connection: sqlite3.Connection, *, owner_id: int, service: str, reading_date: str, observed_value: float, source_path: str, document_id: int) -> int`.
- Produces `apply_zero_carry_forward(connection: sqlite3.Connection, *, owner_id: int, service: str, reading_date: str, period_id: int, observation_id: int, source_path: str) -> bool`.

- [ ] **Step 1: Write failing ingestion tests.**

```python
def test_zero_preserves_observation_and_carries_previous_value(self):
    self.add_confirmed_reading(initial=100, final=0)
    observed = self.connection.execute("SELECT observed_value,status FROM reading_observations").fetchone()
    effective = self.connection.execute("SELECT valor_acumulado,estado,metodo_estimacion FROM lecturas_vecino WHERE fecha_lectura='2026-01-31'").fetchone()
    self.assertEqual((0, "carried_forward"), tuple(observed))
    self.assertEqual((100, "estimado", "carry_forward_zero"), tuple(effective))

def test_zero_without_prior_reading_creates_review(self):
    self.add_reading_with_only_zero_final()
    self.assertTrue(document_review.list_open_issues(self.connection, self.case.id_case))
```

- [ ] **Step 2: Run the ingestion tests.**

Run: `python -m unittest tests.test_unified_ingestion_regressions`

Expected: FAIL because observations and zero carry-forward do not exist.

- [ ] **Step 3: Implement the raw-to-canonical projection.**

Record each incoming row before upserting the canonical reading. For a final `0`, choose only the latest earlier reading of the same owner/service whose state is real or an approved estimate. Persist the current-date effective row as `estimado/carry_forward_zero`, link it to its observation and prior row, and retain the raw zero. If no predecessor, an incompatible same-date real row, or a meter-replacement marker exists, create `READING_ZERO_REVIEW` without changing the canonical row.

- [ ] **Step 4: Run focused reading tests.**

Run: `python -m unittest tests.test_unified_ingestion_regressions tests.test_document_review`

Expected: PASS, including existing counter-reset tests.

- [ ] **Step 5: Commit.**

Run: `git add core/importar_lecturas_metrigest.py core/case_ingestion.py tests/test_unified_ingestion_regressions.py; git commit -m "feat: conservar y arrastrar lecturas cero seguras"`

### Task 3: Mostrar y cerrar grupos de incidencias

**Files:**

- Modify: `core/document_review.py`
- Modify: `core/app.py`
- Modify: `core/expedient_ui.py`
- Test: `tests/test_document_review.py`
- Test: `tests/test_expedient_ui.py`

**Interfaces:**

- `resolve_issue` closes a new open row even if an equivalent historic row is resolved.
- `_accion_resolver_incidencias` routes reset groups to a dialog containing their primary resolution action.

- [ ] **Step 1: Write failing regressions.**

```python
def test_reopened_issue_closes_after_an_equal_resolved_history_row(self):
    first = self._create_missing_date_issue()
    document_review.resolve_issue(self.connection, first.id_issue, value="2026-01-01", reason="Original")
    second = document_review.create_missing_field_issues(
        self.connection, first.id_case, first.id_document, (first.field_name,)
    )[0]
    resolved = document_review.resolve_issue(self.connection, second.id_issue, value="2026-01-02", reason="Nueva fuente")
    self.assertEqual("resolved", resolved.status)

def test_grouped_reset_dialog_has_primary_resolution_action(self):
    expedient_ui.open_confirm_sources_dialog(self.app, self.case.id_case)
    self.assertTrue(any(w.options.get("text", "").startswith("Mantener lecturas previas") for w in self.dialog.descendants()))
```

- [ ] **Step 2: Run the review/UI tests.**

Run: `python -m unittest tests.test_document_review tests.test_expedient_ui`

Expected: duplicate resolution fails before Task 1; the group action is explicitly exercised.

- [ ] **Step 3: Implement atomic routing.**

Keep correction audit, candidate update, issue status update and reapplication inside the existing transaction. Make every grouped `COUNTER_RESET` display `Mantener lecturas previas (n)` and preserve `Abrir archivo` as a secondary control. Ensure the main resolver opens this route whenever reset groups exist.

- [ ] **Step 4: Run the review/UI tests again.**

Run: `python -m unittest tests.test_document_review tests.test_expedient_ui`

Expected: PASS.

- [ ] **Step 5: Commit.**

Run: `git add core/document_review.py core/app.py core/expedient_ui.py tests/test_document_review.py tests/test_expedient_ui.py; git commit -m "fix: resolver grupos e incidencias reabiertas"`

### Task 4: OCR autocontenido y proveedor con evidencia

**Files:**

- Modify: `requirements.txt`
- Modify: `core/lector_pdf.py`
- Modify: `config/proveedores.json`
- Test: `tests/test_source_analysis.py`
- Test: `tests/test_dependencies.py`

**Interfaces:**

- Produces `_rapidocr_text(path) -> str` and structured `OCRResult`.
- `identificar_proveedor` returns a profile only when its evidence score is unique and sufficient.

- [ ] **Step 1: Write failing mocked OCR/provider tests.**

```python
def test_rapidocr_is_used_before_tesseract(self):
    with patch("lector_pdf._rapidocr_text", return_value="NATURGY CLIENTES GAS"):
        result = lector_pdf.extraer_texto_ocr_con_diagnostico("scan.pdf")
    self.assertEqual("rapidocr", result.status)

def test_provider_alias_normalization_accepts_unique_match_and_rejects_tie(self):
    self.assertEqual("TOTALENERGIES_GAS", lector_pdf.identificar_proveedor("TOTAL ENERGIES IBERIA", "factura.pdf", self.providers)[0])
    self.assertEqual((None, None), lector_pdf.identificar_proveedor("ENERGIA", "factura.pdf", self.ambiguous_providers))
```

- [ ] **Step 2: Run OCR/source tests.**

Run: `python -m unittest tests.test_source_analysis tests.test_dependencies`

Expected: FAIL because RapidOCR and ambiguity scoring do not exist.

- [ ] **Step 3: Implement bounded OCR and scoring.**

Add RapidOCR to requirements, import it lazily, render only the first page, and produce short diagnostic states. Preserve the existing Tesseract path only as fallback. Normalize accents, whitespace and punctuation; score configurable filename aliases and document signatures; reject weak or tied candidates. Add configuration aliases only when existing fixtures support them.

- [ ] **Step 4: Run OCR/source tests again.**

Run: `python -m unittest tests.test_source_analysis tests.test_dependencies`

Expected: PASS with OCR mocked when its optional runtime is unavailable.

- [ ] **Step 5: Commit.**

Run: `git add requirements.txt core/lector_pdf.py config/proveedores.json tests/test_source_analysis.py tests/test_dependencies.py; git commit -m "feat: reforzar OCR y deteccion de proveedores"`

### Task 5: Verificación e integración

**Files:**

- Verify: all files changed above

- [ ] **Step 1: Install declared dependencies in the isolated environment.**

Run: `python -m pip install -r requirements.txt`

Expected: project dependencies and RapidOCR are available.

- [ ] **Step 2: Run complete verification.**

Run: `python -m unittest discover -s tests`

Run: `python -m compileall -q core`

Run: `git diff --check master..HEAD`

Expected: zero failures, compilation success and no whitespace errors.

- [ ] **Step 3: Merge only the verified branch.**

Run: `git status --short; git log --oneline master..HEAD`

Expected: only feature commits. Merge with `git merge --no-ff feature/incidencias-ocr-lecturas` from the main worktree after Step 2 passes.
