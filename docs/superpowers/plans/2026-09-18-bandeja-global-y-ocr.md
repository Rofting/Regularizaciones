# Bandeja global, OCR y continuidad de reparto Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Procesar una carpeta mixta de facturas y lecturas sin contexto previo y conducir cada expediente de forma visible hasta reparto y cartas.

**Architecture:** La detección global se concentra en un servicio puro que devuelve propuestas, sin mover archivos ni escribir en SQLite. La interfaz muestra las propuestas y solo publica comunidades/expedientes después de confirmación. El OCR devuelve diagnóstico estructurado; la UI traduce ese diagnóstico sin mensajes técnicos. El estado del expediente es la única fuente de verdad para decidir el siguiente botón.

**Tech Stack:** Python 3, SQLite, CustomTkinter, pdfplumber, pdf2image, pytesseract, openpyxl, unittest.

**Spec:** `docs/superpowers/specs/2026-09-18-bandeja-global-y-ocr-design.md`

## Global Constraints

- Windows local; los originales no se mueven ni se eliminan durante la detección.
- Tesseract y Poppler se detectan en tiempo de ejecución; `eng` es alternativa válida cuando `spa` no exista.
- Una comunidad nueva exige una señal inequívoca; los conflictos de CIF se bloquean.
- La interfaz actual y la base SQLite existente se conservan; no se fusionan versiones antiguas completas.
- Cada cambio de comportamiento se implementa con una prueba que primero falle.

---

### Task 1: Enrutado visible del paso Reparto

**Files:**
- Modify: `core/app.py:229-257`, `core/app.py:1038-1060`, `core/app.py:1294-1312`
- Modify: `core/expedient_ui.py:141-202`
- Test: `tests/test_expedient_ui.py`

**Interfaces:**
- Consumes: `guided_workspace_state(has_case, document_count, open_issue_count, case_status)`.
- Produces: `AppGestionFincas._accion_siguiente_expediente()` and visible action text.

- [ ] **Step 1: Write the failing test**

```python
def test_workspace_keeps_source_confirmation_as_next_action_when_no_issues_exist():
    state = guided_workspace_state(
        has_case=True, document_count=3, open_issue_count=0,
        case_status="under_review",
    )
    assert state.next_action == "confirmar_fuentes"
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m unittest tests.test_expedient_ui`

Expected: failure because `under_review` currently routes to `resolver_incidencias`.

- [ ] **Step 3: Implement the smallest state-to-action map**

```python
"under_review": ("Confirmar fuentes", self._accion_confirmar_fuentes)
```

Make the lateral Reparto action execute `self._accion_siguiente_expediente`, which consults current workspace state rather than calling Excel generation unconditionally. On workflow exceptions, show a user-visible dialog containing the action required next.

- [ ] **Step 4: Run focused tests**

Run: `python -m unittest tests.test_expedient_ui tests.test_case_workflow_actions`

Expected: pass.

- [ ] **Step 5: Commit**

```powershell
git add core/app.py core/expedient_ui.py tests/test_expedient_ui.py
git commit -m "fix: guiar reparto al siguiente paso pendiente"
```

### Task 2: OCR autocontenido con diagnóstico útil

**Files:**
- Modify: `core/lector_pdf.py:531-585`, `core/lector_pdf.py:1038-1053`
- Modify: `core/source_analysis.py:140-180`
- Modify: `core/expedient_ui.py:247-320`
- Test: `tests/test_source_analysis.py`

**Interfaces:**
- Consumes: `extraer_texto_ocr(ruta_archivo)`.
- Produces: `OCRExtraction(text: str, status: str, detail: str)` and `SourceAnalysis.review_message` suitable for users.

- [ ] **Step 1: Write failing OCR diagnostic tests**

```python
def test_ocr_uses_english_when_spanish_language_is_not_installed():
    result = lector_pdf.ocr_text_from_pdf(Path("scan.pdf"), engine=engine_without_spa)
    self.assertEqual("eng", result.language)

def test_empty_ocr_returns_actionable_status_not_install_instructions():
    result = lector_pdf.ocr_text_from_pdf(Path("scan.pdf"), engine=empty_engine)
    self.assertEqual("no_readable_text", result.status)
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m unittest tests.test_source_analysis`

Expected: failure because structured OCR result does not exist.

- [ ] **Step 3: Implement structured OCR**

```python
@dataclass(frozen=True)
class OCRExtraction:
    text: str
    status: str
    detail: str
    language: str | None
```

Detect engine, Poppler and available languages; attempt the first page with `spa` when installed otherwise `eng`. Preserve existing string wrapper for compatible extractors. Replace `SIN_TEXTO` download copy with the structured failure reason and render it through `issue_guidance`.

- [ ] **Step 4: Run focused tests and one real scanned fixture**

Run: `python -m unittest tests.test_source_analysis`

Run the extractor against `Flujo_Regularizacion/entrada/Pruebas_Facturas/644 agua.pdf` and confirm an OCR result is produced.

- [ ] **Step 5: Commit**

```powershell
git add core/lector_pdf.py core/source_analysis.py core/expedient_ui.py tests/test_source_analysis.py
git commit -m "feat: informar OCR integrado en fuentes escaneadas"
```

### Task 3: Clasificación de nombres recibidos por correo

**Files:**
- Modify: `core/community_discovery.py`
- Modify: `core/source_analysis.py`
- Modify: `config/proveedores.json`
- Test: `tests/test_source_analysis.py`

**Interfaces:**
- Consumes: a `Path` such as `644 - Limpieza mar 2026.pdf`.
- Produces: `FilenameEvidence(community_code, category, supply_hint, month, year, duplicate_index)`.

- [ ] **Step 1: Write failing filename-evidence tests**

```python
def test_mail_filename_extracts_community_and_service_hint():
    evidence = filename_evidence(Path("644 - Limpieza mar 2026.pdf"))
    self.assertEqual("644", evidence.community_code)
    self.assertEqual("LIMPIEZA", evidence.supply_hint)
    self.assertEqual((3, 2026), (evidence.month, evidence.year))
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m unittest tests.test_source_analysis`

Expected: failure because filename evidence does not exist.

- [ ] **Step 3: Implement conservative name evidence**

Recognise code prefixes and category folders (`Habituales`, `Extraordinarias`, `Sin_Comunidad`). Map name keywords to hints without inventing invoice totals or dates. Treat suffixes ` (2)` as duplicate hints only; SHA-256 remains duplicate authority. Add only provider/config entries evidenced by the old project or real mail set.

- [ ] **Step 4: Run focused tests**

Run: `python -m unittest tests.test_source_analysis tests.test_unified_ingestion_regressions`

Expected: pass.

- [ ] **Step 5: Commit**

```powershell
git add core/community_discovery.py core/source_analysis.py config/proveedores.json tests/test_source_analysis.py
git commit -m "feat: clasificar nombres habituales de facturas por correo"
```

### Task 4: Bandeja global y publicación segura

**Files:**
- Modify: `core/community_discovery.py`
- Modify: `core/expedient_ui.py:229-380`, `core/app.py:2938-3010`
- Modify: `core/expedient_service.py`
- Test: `tests/test_community_onboarding.py`, `tests/test_expedient_ui.py`

**Interfaces:**
- Consumes: `source_files_in_folder(folder)` and `discover_communities(paths)`.
- Produces: `GlobalIntakeProposal` grouped by community, period evidence and route status.

- [ ] **Step 1: Write failing global-intake test**

```python
def test_global_intake_groups_two_communities_without_selected_context(tmp_path):
    proposal = build_global_intake((tmp_path / "644 - Factura mar 2026.pdf", tmp_path / "658 - Factura abr 2026.pdf"))
    self.assertEqual(("644", "658"), tuple(item.community_code for item in proposal.groups))
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m unittest tests.test_community_onboarding tests.test_expedient_ui`

Expected: failure because global proposals do not exist.

- [ ] **Step 3: Implement the review-first global intake**

Build groups before persistence. Show source count, community evidence, period evidence, hints and blocking reason. Accepting a group creates/reuses its community and creates/reuses an expedition only when both dates are verified. Route unbound documents to a visible review row; never use a selected community as fallback.

- [ ] **Step 4: Run focused workflow tests**

Run: `python -m unittest tests.test_community_onboarding tests.test_expedient_ui tests.test_case_workflow_actions`

Expected: pass.

- [ ] **Step 5: Commit**

```powershell
git add core/community_discovery.py core/expedient_ui.py core/app.py core/expedient_service.py tests/test_community_onboarding.py tests/test_expedient_ui.py
git commit -m "feat: incorporar fuentes mixtas sin contexto previo"
```

### Task 5: Auditoría de reglas anteriores y verificación final

**Files:**
- Modify only files justified by comparison with `C:/Users/Jose/Proyectos/Flujo_Regularizacion`
- Modify: `README.md` or create `docs/manual-operacion.md`
- Test: affected existing test files

**Interfaces:**
- Consumes: provider profiles and real filenames from the historical folders.
- Produces: documented accepted/rejected legacy rules and user operating instructions.

- [ ] **Step 1: Compare provider and extraction rules**

List files that differ under `core/` and `config/`; identify only rules absent from current code and supported by a real sample. Do not copy UI files or database schema from the old project.

- [ ] **Step 2: Add a test for each accepted rule before production code**

Use the fixture or a minimal string representative of the evidence. Every accepted provider pattern must name the supplier or filename format it supports.

- [ ] **Step 3: Implement accepted rules and write operating manual**

Document: global folder intake, community proposal, blocked rows, OCR status, source confirmation, Excel, distribution and letters.

- [ ] **Step 4: Run full relevant suite and compile**

Run: `python -m unittest discover -s tests`

Run: `python -m compileall -q core`

Run: `git diff --check`

Expected: no failures and no whitespace errors.

- [ ] **Step 5: Commit**

```powershell
git add core config docs tests README.md
git commit -m "feat: consolidar entrada global de regularizaciones"
```
