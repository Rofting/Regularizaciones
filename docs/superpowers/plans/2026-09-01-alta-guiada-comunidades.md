# Alta guiada de comunidades Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Dar de alta comunidades desde propietarios, lecturas y facturas sin requerir un Excel perfecto previo.

**Architecture:** Un servicio puro analiza las fuentes y entrega un borrador con propuestas y preguntas. La confirmación publica un perfil validado y crea/archiva todo de forma transaccional. La interfaz v3 muestra cinco pasos y no contiene lógica de lectura o persistencia.

**Tech Stack:** Python 3, sqlite3, openpyxl, pdfplumber, CustomTkinter, unittest.

**Spec:** `docs/superpowers/specs/2026-09-01-alta-guiada-comunidades-design.md`

## Global Constraints

- El Excel manual es opcional; propietarios y lecturas son las fuentes mínimas esperadas.
- Las lecturas Excel y PDF son alternativas del mismo dato.
- Todo dato incierto crea una pregunta obligatoria; no se inventan conceptos, consumos, importes, períodos o columnas.
- Los documentos privados, base de datos y plantillas instaladas no se versionan.
- El perfil tiene versión inicial `1`, ruta relativa segura y pasa la validación de `excel_profiles`.
- Una excepción de confirmación no deja filas, perfil, plantilla ni archivos archivados parciales.
- Mantener la interfaz v3, sin marca del despacho y sin crear Tk en pruebas.

---

## File structure

| Archivo | Responsabilidad |
| --- | --- |
| `core/community_onboarding.py` | Borrador, análisis, preguntas, perfil y confirmación atómica. |
| `core/excel_profiles.py` | Validación pública de un payload antes de publicar JSON. |
| `core/excel_generator.py` | Generación del libro base canónico para perfiles creados por el asistente. |
| `core/expedient_ui.py` | Ruta pura y diálogo visual de cinco pasos. |
| `core/app.py` | Carga del servicio, apertura y refresco de selectores. |
| `tests/test_community_onboarding.py` | Análisis, perfil, confirmación, fuentes alternativas y rollback. |
| `tests/test_expedient_ui.py` | Enrutado visual sin controles Tk. |
| `GUIA_PROCESAR_TODO.txt` | Alta operativa sin Excel maestro manual. |
| `.gitignore` | Excluir perfiles generados localmente sin ocultar el perfil público 658. |

### Task 1: Borrador y análisis no destructivo

**Files:**
- Create: `core/community_onboarding.py`
- Create: `tests/test_community_onboarding.py`

**Interfaces:**
- Produce `SourceCandidate`, `OnboardingQuestion`, `OnboardingDraft` y `analyse_sources`.
- Recibe rutas explícitas, no una carpeta abierta; reutiliza `lector_pdf`, openpyxl y los lectores XLS existentes.

- [ ] **Step 1: Write failing tests**

```python
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
```

- [ ] **Step 2: Run tests to verify RED**

Run: `$env:PYTHONUTF8='1'; & 'C:\Users\Jose\Proyectos\Soporte\Soporte\Flujo calculo regularizaciones\.venv-fase1\Scripts\python.exe' -m unittest tests.test_community_onboarding -q`

Expected: FAIL because `community_onboarding` does not exist.

- [ ] **Step 3: Implement minimal analysis**

```python
@dataclass(frozen=True)
class OnboardingDraft:
    community_code: str
    community_name: str
    sources: tuple[SourceCandidate, ...]
    detected_modules: tuple[str, ...]
    questions: tuple[OnboardingQuestion, ...]

def analyse_sources(*, community_code: str, community_name: str,
                    owner_list_path: Path, reading_paths: tuple[Path, ...],
                    invoice_paths: tuple[Path, ...], project_root: Path) -> OnboardingDraft:
    sources = _classify_source_paths(owner_list_path, reading_paths, invoice_paths)
    return _create_draft(community_code, community_name, sources, project_root)
```

Validate file existence/extensions and calculate SHA-256. Inspect Excel headers and PDF text. If multiple reading columns or service candidates exist, emit `OnboardingQuestion(required=True)` rather than choose one. Do not write files or database rows.

- [ ] **Step 4: Run tests to verify GREEN**

Run: `$env:PYTHONUTF8='1'; & 'C:\Users\Jose\Proyectos\Soporte\Soporte\Flujo calculo regularizaciones\.venv-fase1\Scripts\python.exe' -m unittest tests.test_community_onboarding -q`

Expected: PASS.

- [ ] **Step 5: Commit**

Commit message: `feat: analiza fuentes para alta guiada`

### Task 2: Perfil confirmado y validación reutilizable

**Files:**
- Modify: `core/community_onboarding.py`
- Modify: `core/excel_profiles.py`
- Modify: `tests/test_community_onboarding.py`
- Modify: `tests/test_excel_profiles.py`

**Interfaces:**
- Consume `OnboardingDraft` y `answers: Mapping[str, str | bool]`.
- Produce `build_profile_payload(draft, answers)` y `validate_profile_payload(payload, project_root)`.

- [ ] **Step 1: Write failing tests**

```python
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
```

- [ ] **Step 2: Run tests to verify RED**

Run: `$env:PYTHONUTF8='1'; & 'C:\Users\Jose\Proyectos\Soporte\Soporte\Flujo calculo regularizaciones\.venv-fase1\Scripts\python.exe' -m unittest tests.test_community_onboarding tests.test_excel_profiles -q`

Expected: FAIL because the builder and validator do not exist.

- [ ] **Step 3: Implement minimal profile APIs**

```python
def validate_profile_payload(payload: Mapping[str, Any], project_root: Path) -> ExcelProfile:
    """Validate generated JSON without publishing it under config/."""

def build_profile_payload(draft: OnboardingDraft,
                          answers: Mapping[str, str | bool]) -> dict[str, Any]:
    """Build version 1 from explicit human answers only."""
```

Create `<community_code>_v1`, a safe relative template path and only confirmed modules/concepts. Reuse existing checks for safe paths, allocation methods, required sheets and A1 cells.

- [ ] **Step 4: Run tests to verify GREEN**

Run: `$env:PYTHONUTF8='1'; & 'C:\Users\Jose\Proyectos\Soporte\Soporte\Flujo calculo regularizaciones\.venv-fase1\Scripts\python.exe' -m unittest tests.test_community_onboarding tests.test_excel_profiles -q`

Expected: PASS.

- [ ] **Step 5: Commit**

Commit message: `feat: construye perfiles confirmados de comunidades`

### Task 3: Confirmación atómica y expediente inicial

**Files:**
- Modify: `core/community_onboarding.py`
- Modify: `core/excel_generator.py`
- Modify: `.gitignore`
- Modify: `tests/test_community_onboarding.py`

**Interfaces:**
- Consume Tasks 1 and 2.
- Produce `OnboardingResult` and `confirm_onboarding(...)`.

- [ ] **Step 1: Write failing tests**

```python
def test_confirm_onboarding_writes_profile_archives_sources_and_creates_case(self):
    result = community_onboarding.confirm_onboarding(
        self.connection, draft=self.valid_draft, answers=self.answers,
        project_root=self.project_root, archive_root=self.archive_root,
        period_name="2026", start_date=date(2026, 1, 1), end_date=date(2026, 12, 31),
        actor="Prueba",
    )
    self.assertIsNotNone(result.case_id)
    self.assertTrue((self.project_root / "config/excel_profiles/900_v1.json").is_file())
    self.assertTrue((self.project_root / "plantillas/comunidades/900/900_v1.xlsx").is_file())
    self.assertEqual(3, self._registered_source_count())

def test_confirm_onboarding_rolls_back_json_database_and_archives_after_error(self):
    with self.assertRaises(OSError):
        community_onboarding.confirm_onboarding(self.connection, **self.failing_arguments())
    self.assertFalse((self.project_root / "config/excel_profiles/900_v1.json").exists())
    self.assertEqual(0, self._community_count())
    self.assertEqual(0, self._registered_source_count())
```

- [ ] **Step 2: Run tests to verify RED**

Run: `$env:PYTHONUTF8='1'; & 'C:\Users\Jose\Proyectos\Soporte\Soporte\Flujo calculo regularizaciones\.venv-fase1\Scripts\python.exe' -m unittest tests.test_community_onboarding -q`

Expected: FAIL because `confirm_onboarding` does not exist.

- [ ] **Step 3: Implement atomically**

```python
def confirm_onboarding(connection: sqlite3.Connection, *, draft: OnboardingDraft,
                       answers: Mapping[str, str | bool], project_root: Path,
                       archive_root: Path, period_name: str | None,
                       start_date: date | None, end_date: date | None,
                       actor: str) -> OnboardingResult:
    profile = build_profile_payload(draft, answers)
    return _publish_onboarding(connection, draft, profile, archive_root, actor)
```

Require both dates or neither. Publish the profile via a same-directory temporary file after validation. Use `excel_generator` to create a canonical workbook with only the confirmed module sheets and the matching declarative layout; install it under the private template path before committing success. Create community/period/case with existing services and archive every source with SHA. `register_source_document` owns its SQLite transaction, so compensate explicitly on exception: delete only the source rows/archive paths, JSON and template created by this call. Add `.gitignore` rules for newly generated local profile JSON files while retaining the committed `658_acs_v1.json`.

- [ ] **Step 4: Run tests to verify GREEN**

Run: `$env:PYTHONUTF8='1'; & 'C:\Users\Jose\Proyectos\Soporte\Soporte\Flujo calculo regularizaciones\.venv-fase1\Scripts\python.exe' -m unittest tests.test_community_onboarding tests.test_expedient_service -q`

Expected: PASS.

- [ ] **Step 5: Commit**

Commit message: `feat: confirma altas de comunidad de forma atómica`

### Task 4: Asistente visual v3 y guía

**Files:**
- Modify: `core/expedient_ui.py`
- Modify: `core/app.py`
- Create: `tests/test_expedient_ui.py`
- Modify: `GUIA_PROCESAR_TODO.txt`
- Modify: `docs/validacion-expedientes.md`

**Interfaces:**
- Consume `analyse_sources` and `confirm_onboarding`.
- Produce `onboarding_step_route(state)` and `open_community_onboarding_dialog(app)`.

- [ ] **Step 1: Write failing route tests**

```python
def test_onboarding_step_route_requires_sources_before_detection(self):
    state = {"step": "identity", "identity_valid": True, "has_sources": False}
    self.assertEqual("sources", expedient_ui.onboarding_step_route(state))

def test_onboarding_step_route_requires_answers_before_summary(self):
    state = {"step": "detected", "has_required_questions": True, "answers_complete": False}
    self.assertEqual("confirmations", expedient_ui.onboarding_step_route(state))
```

- [ ] **Step 2: Run tests to verify RED**

Run: `$env:PYTHONUTF8='1'; & 'C:\Users\Jose\Proyectos\Soporte\Soporte\Flujo calculo regularizaciones\.venv-fase1\Scripts\python.exe' -m unittest tests.test_expedient_ui -q`

Expected: FAIL because the route function does not exist.

- [ ] **Step 3: Implement visual flow and wiring**

```python
def onboarding_step_route(state: Mapping[str, Any]) -> str:
    return _next_onboarding_step(state)

def open_community_onboarding_dialog(app: "AppGestionFincas") -> None:
    _build_onboarding_dialog(app)
```

Make **+ Comunidad** offer **Registro rápido** and **Alta guiada desde fuentes**. Use five labelled stages, explicit file selectors, confirmation fields, the app's existing background progress mechanism and selector refresh on success. Keep quick registration unchanged. Unit tests call only the route function, never Tk.

- [ ] **Step 4: Document and verify GREEN**

Document the five stages, minimum sources, optional Excel and privacy rule. Run:

Run: `$env:PYTHONUTF8='1'; & 'C:\Users\Jose\Proyectos\Soporte\Soporte\Flujo calculo regularizaciones\.venv-fase1\Scripts\python.exe' -m unittest tests.test_expedient_ui tests.test_community_onboarding -q`

Expected: PASS.

- [ ] **Step 5: Run complete verification and commit**

Run: `$env:PYTHONUTF8='1'; & 'C:\Users\Jose\Proyectos\Soporte\Soporte\Flujo calculo regularizaciones\.venv-fase1\Scripts\python.exe' -m unittest discover -s tests -t . -q`

Expected: PASS.

Run: `& 'C:\Users\Jose\Proyectos\Soporte\Soporte\Flujo calculo regularizaciones\.venv-fase1\Scripts\python.exe' -m compileall -q core tests`

Expected: exit 0.

Run: `git diff --check`

Expected: no output.

Commit message: `feat: añade asistente visual de alta de comunidades`

## Review checklist

- Tasks 1–3 cover sources, questions, profile, transaction and archives.
- Task 4 covers the v3 user path, documentation and regression suite.
- `OnboardingDraft` moves from Task 1 through Task 3; Task 4 consumes only public services.
- No task assumes an Excel perfecto, a fixed community, a fixed provider or a specific office.
