# Guided Home Redesign Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Replace the crowded main window with a portable, four-step guided workspace without changing any regularization data or business-service contract.

**Architecture:** A pure projection converts the selected case into one active workflow step and one allowed next action. `AppGestionFincas` renders that projection through reusable CustomTkinter primitives while existing ingestion, review, Excel, distribution and letter services remain the only source of truth.

**Tech Stack:** Python 3, CustomTkinter, tkinter, SQLite, `unittest`.

**Spec:** `docs/superpowers/specs/2026-09-09-guided-home-redesign-design.md`

## Global Constraints

- Do not alter database schemas or service contracts.
- Keep branding neutral to any despacho and preserve light/dark themes.
- Turquoise is primary/advance, green is save/confirm, amber is review, red is destructive or blocking, and secondary controls keep a neutral visible border.
- Preserve the approved folder ingestion, human incident guidance, meter-reset approval and all service-level gates.
- Do not stage or edit `config/ui_prefs.json`, the loose text document, or the private Word template.

---

### Task 1: Derive a pure four-step workspace state

**Files:**
- Modify: `core/expedient_ui.py:22-105`
- Modify: `tests/test_expedient_ui.py:173-205`

**Interfaces:**
- Produces: immutable `GuidedStep(key: str, label: str, status: str)`.
- Produces: immutable `GuidedWorkspaceState(active_step: str, next_action: str, headline: str, detail: str, steps: tuple[GuidedStep, ...])`.
- Produces: `guided_workspace_state(*, has_case: bool, document_count: int, open_issue_count: int, case_status: str) -> GuidedWorkspaceState`.
- Consumes: `draft`, `gathering_sources`, `under_review`, `ready_for_calculation`, `calculated`, `reconciled`, `deliveries_generated`, and `closed` case statuses.

- [ ] **Step 1: Write failing unit tests**

```python
class GuidedWorkspaceStateTest(unittest.TestCase):
    def test_no_case_prompts_case_creation(self):
        state = expedient_ui.guided_workspace_state(
            has_case=False, document_count=0, open_issue_count=0, case_status=""
        )
        self.assertEqual(("fuentes", "crear_expediente"), (state.active_step, state.next_action))

    def test_open_issue_routes_to_validation(self):
        state = expedient_ui.guided_workspace_state(
            has_case=True, document_count=4, open_issue_count=2, case_status="under_review"
        )
        self.assertEqual(("validar", "resolver_incidencias"), (state.active_step, state.next_action))

    def test_ready_case_routes_to_excel_generation(self):
        state = expedient_ui.guided_workspace_state(
            has_case=True, document_count=4, open_issue_count=0,
            case_status="ready_for_calculation"
        )
        self.assertEqual(("reparto", "generar_excel"), (state.active_step, state.next_action))
```

- [ ] **Step 2: Run them red**

Run: `& .\.venv-fase1\Scripts\python.exe -m unittest tests.test_expedient_ui.GuidedWorkspaceStateTest -v`

Expected: FAIL because `guided_workspace_state` is absent.

- [ ] **Step 3: Implement the state projection**

```python
@dataclass(frozen=True)
class GuidedStep:
    key: str
    label: str
    status: str

@dataclass(frozen=True)
class GuidedWorkspaceState:
    active_step: str
    next_action: str
    headline: str
    detail: str
    steps: tuple[GuidedStep, ...]
```

Implement `guided_workspace_state` with exactly `fuentes`, `validar`, `reparto`, `cartas`. Map no case to `crear_expediente`; no documents/draft to `anadir_fuentes`; open issues to `resolver_incidencias`; `ready_for_calculation` and `calculated` to `generar_excel`; `reconciled` to `generar_cartas`; deliveries/closed to `abrir_salidas`. Unknown statuses return Fuentes with explanatory detail, never raise.

- [ ] **Step 4: Add terminal-state tests and run green**

```python
def test_reconciled_case_routes_to_letters(self):
    state = expedient_ui.guided_workspace_state(
        has_case=True, document_count=4, open_issue_count=0, case_status="reconciled"
    )
    self.assertEqual(("cartas", "generar_cartas"), (state.active_step, state.next_action))
```

Run: `& .\.venv-fase1\Scripts\python.exe -m unittest tests.test_expedient_ui -v`

Expected: PASS.

- [ ] **Step 5: Commit**

```powershell
git add core/expedient_ui.py tests/test_expedient_ui.py
git commit -m "feat: derive guided workspace state"
```

### Task 2: Formalize semantic visual primitives

**Files:**
- Modify: `core/ui_moderna.py:24-58`
- Modify: `core/ui_moderna.py:after PuntoEstado`
- Modify: `tests/test_expedient_ui.py`

**Interfaces:**
- Produces: `C["secundario_hover"]`, `C["revision"]`, `C["revision_hover"]`.
- Produces: `WORKFLOW_STEP_STYLES` keyed by `pending`, `active`, `blocked`, `ready`, `done`.
- Produces: `secondary_button_kwargs() -> dict[str, object]`.
- Produces: `WorkflowStepRow(master, *, number: str, label: str, status: str, command)`, with `set_status(status: str) -> None`.

- [ ] **Step 1: Write failing style tests**

```python
def test_secondary_button_style_is_neutral_and_bordered(self):
    style = ui_moderna.secondary_button_kwargs()
    self.assertEqual("transparent", style["fg_color"])
    self.assertEqual(1, style["border_width"])
    self.assertEqual(ui_moderna.C["borde"], style["border_color"])

def test_every_workflow_state_has_a_style(self):
    for status in ("pending", "active", "blocked", "ready", "done"):
        self.assertIn(status, ui_moderna.WORKFLOW_STEP_STYLES)
```

- [ ] **Step 2: Run them red**

Run: `& .\.venv-fase1\Scripts\python.exe -m unittest tests.test_expedient_ui.WorkflowVisualStyleTest -v`

Expected: FAIL because the helpers are absent.

- [ ] **Step 3: Implement tokens, shared secondary style and step row**

```python
def secondary_button_kwargs():
    return {
        "fg_color": "transparent", "hover_color": C["secundario_hover"],
        "border_width": 1, "border_color": C["borde"],
        "text_color": C["texto_sec"],
    }
```

The row displays number, label and state marker. Only active/ready rows receive a command. It must not make a business decision; `app.py` supplies its state.

- [ ] **Step 4: Run focused verification and commit**

Run:

```powershell
& .\.venv-fase1\Scripts\python.exe -m unittest tests.test_expedient_ui -v
& .\.venv-fase1\Scripts\python.exe -m py_compile core\ui_moderna.py core\expedient_ui.py
```

Expected: PASS and exit 0.

```powershell
git add core/ui_moderna.py core/expedient_ui.py tests/test_expedient_ui.py
git commit -m "feat: add guided workflow visual primitives"
```

### Task 3: Rebuild the main shell as the approved guided layout

**Files:**
- Modify: `core/app.py:176-339`
- Modify: `tests/test_expedient_ui.py`

**Interfaces:**
- Consumes: `guided_workspace_state` and `UIM.WorkflowStepRow`.
- Produces: `self.workflow_step_rows`, `self.workspace_headline`, `self.workspace_detail`, `self.workspace_primary`, `self.workspace_summary`, and `self.workspace_log_toggle`.
- Preserves: `cb_comunidad`, `cb_periodo`, `cb_expediente`, `botones`, `log_area`, status bar and every existing `_accion_*` method.

- [ ] **Step 1: Write failing composition tests with existing fake widgets**

```python
def test_guided_home_has_exactly_four_step_rows(self):
    self.build_v3_home()
    self.assertEqual({"fuentes", "validar", "reparto", "cartas"}, set(self.app.workflow_step_rows))

def test_guided_home_keeps_context_and_settings_actions(self):
    self.build_v3_home()
    self.assertIn("Nueva comunidad", self.button_texts())
    self.assertIn("Ajustes", self.button_texts())
    self.assertIn("Crear expediente · principal", self.app.botones)
```

- [ ] **Step 2: Run them red**

Run: `& .\.venv-fase1\Scripts\python.exe -m unittest tests.test_expedient_ui.GuidedHomeCompositionTest -v`

Expected: FAIL because the new widget references are absent.

- [ ] **Step 3: Replace `_crear_ui_v3` with three regions**

```text
Header: Regularizaciones | Comunidad | Período | Expediente | Nueva comunidad | Ajustes | Tema
Body:   four WorkflowStepRow items | one workspace headline/detail/action | compact case summary
Footer: persistent status/progress; activity log only opens through a secondary toggle
```

Reuse the current combo callbacks, `_nueva_comunidad`, `_nuevo_periodo`, `_configurar_rutas`, `_abrir_salidas`, `self.log_area`, and current action methods. Remove the duplicate large period banner, old five-dot tracker, seven colored navigation cards and permanently dominant activity card.

- [ ] **Step 4: Keep workflow actions in their appropriate steps**

Fuentes exposes `_accion_anadir_fuentes` and existing "Añadir archivos"/"Añadir carpeta" dialog. Validar displays `_refrescar_bandeja_incidencias` in the central area and `_accion_resolver_incidencias`. Reparto exposes Excel then distribution only from the computed state. Cartas exposes concept selection/generation only when reconciled.

- [ ] **Step 5: Verify and commit**

Run:

```powershell
& .\.venv-fase1\Scripts\python.exe -m unittest tests.test_expedient_ui -v
& .\.venv-fase1\Scripts\python.exe .\core\app.py
```

Expected: tests PASS; app opens without CustomTkinter unsupported-argument errors. Manually resize and switch theme to verify no selector is clipped.

```powershell
git add core/app.py tests/test_expedient_ui.py
git commit -m "feat: redesign guided regularization home"
```

### Task 4: Bind the live case refresh and gates to the workspace

**Files:**
- Modify: `core/app.py:1033-1245`
- Modify: `core/app.py:2210-2240`
- Modify: `tests/test_expedient_ui.py`
- Test: `tests/test_case_workflow_actions.py`

**Interfaces:**
- Produces: `_actualizar_workspace(state, *, document_count: int, open_issue_count: int, date_range: str, profile_label: str) -> None`.
- Consumes: the case, document count, issues and profile already loaded by `_refrescar_expediente`.
- Preserves: `case_workflow_actions` as the final gate; a disabled UI never bypasses service validation.

- [ ] **Step 1: Write failing live-refresh tests**

```python
def test_refresh_renders_incident_resolution_as_the_primary_action(self):
    state = expedient_ui.guided_workspace_state(
        has_case=True, document_count=3, open_issue_count=1, case_status="under_review"
    )
    self.app._actualizar_workspace(
        state, document_count=3, open_issue_count=1,
        date_range="01/09/2025 – 31/08/2026", profile_label="658_acs_v1",
    )
    self.assertEqual("Resolver incidencias", self.app.workspace_primary.cget("text"))
```

- [ ] **Step 2: Run it red**

Run: `& .\.venv-fase1\Scripts\python.exe -m unittest tests.test_expedient_ui.GuidedHomeRefreshTest -v`

Expected: FAIL because `_actualizar_workspace` is absent.

- [ ] **Step 3: Implement refresh binding**

At the end of `_refrescar_expediente`, compute `guided_workspace_state` and call `_actualizar_workspace` with the current counts, date range and profile. Use only this action map:

```python
{
    "crear_expediente": self._accion_crear_expediente,
    "anadir_fuentes": self._accion_anadir_fuentes,
    "resolver_incidencias": self._accion_resolver_incidencias,
    "generar_excel": self._accion_generar_excel_expediente,
    "generar_cartas": self._accion_generar_cartas_expediente,
    "abrir_salidas": self._abrir_salidas,
}
```

Update `_limpiar_contexto_expediente`, `_en_hilo`, `_deshabilitar_botones` and `_habilitar_botones` so the central action and step rows refresh or disable together. Summary text must say exactly what blocks a later step: missing case, missing sources, open incidents, profile pending or calculation pending.

- [ ] **Step 4: Run full verification and manual acceptance**

Run:

```powershell
& .\.venv-fase1\Scripts\python.exe -m unittest tests.test_expedient_ui tests.test_case_workflow_actions -v
& .\.venv-fase1\Scripts\python.exe -m unittest discover -s tests -q
& .\.venv-fase1\Scripts\python.exe -m py_compile core\app.py core\expedient_ui.py core\ui_moderna.py
git diff --check
```

Expected: all tests PASS, compilation exits 0, and `git diff --check` has no errors.

Manually use a separate test database to create `PRUEBA-001`, make a short case, load a nested source folder, resolve a missing date, verify the Excel/reparto gate and then the letters gate. Repeat in both themes and at reduced window width.

- [ ] **Step 5: Commit**

```powershell
git add core/app.py core/expedient_ui.py core/ui_moderna.py tests/test_expedient_ui.py tests/test_case_workflow_actions.py
git commit -m "feat: connect guided workspace to case progress"
```
