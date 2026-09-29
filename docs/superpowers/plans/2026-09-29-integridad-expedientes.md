# Integridad de expedientes Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Impedir estados falsamente listos y dirigir cada expediente a una acción concreta según los datos que realmente necesita la siguiente etapa.

**Architecture:** Un servicio puro `case_readiness` calculará un informe único desde SQLite. Las acciones de Excel, reparto y cartas lo usarán como puerta de entrada; la interfaz sólo representará ese informe y nunca volverá a inferir preparación a partir de “cero incidencias”.

**Tech Stack:** Python 3.12, SQLite, unittest, CustomTkinter.

**Spec:** `docs/superpowers/specs/2026-09-29-cierre-produccion-portabilidad-y-correo-design.md`

## Global Constraints

- No se inventan datos económicos, lecturas ni coeficientes.
- La base de datos es la fuente de verdad y las decisiones se agrupan por causa.
- La base local, las preferencias y las fuentes privadas quedan fuera de cada commit.
- Cada cambio de comportamiento sigue RED→GREEN y termina con la suite completa.

## Review Focus

- Un expediente sin incidencias pero sin propietarios no puede generar Excel ni figurar listo.
- Una comunidad con conceptos sólo fijos no debe bloquearse por lecturas que no utiliza.
- Una exportación validada para una huella antigua no autoriza el reparto actual.
- Un reparto conciliado antiguo no autoriza cartas para entradas modificadas.
- Repetir un paso no elimina ejecuciones históricas ni deja un estado adelantado tras fallar.

---

### Task 1: Modelo único de preparación

**Files:**
- Create: `core/case_readiness.py`
- Create: `tests/test_case_readiness.py`

**Interfaces:**
- Consumes: conexión `sqlite3.Connection`, `id_case: int`, `project_root: Path`.
- Produces: `ReadinessBlocker(code, stage, message, action, count)`, `CaseReadinessReport`, `evaluate_case_readiness(...)` y `require_stage(...)`.

- [ ] **Step 1: Write the failing tests**

Crear casos SQLite temporales que demuestren: expediente sin propietarios bloqueado con `MISSING_OWNERS`; concepto por coeficiente sin coeficientes bloqueado con `MISSING_COEFFICIENTS`; concepto fijo sin lecturas no bloqueado por lecturas; y caso completo con `excel_ready=True`.

```python
report = evaluate_case_readiness(connection, case_id, PROJECT_ROOT)
self.assertEqual(("MISSING_OWNERS",), tuple(item.code for item in report.for_stage("excel")))
self.assertFalse(report.excel_ready)
```

- [ ] **Step 2: Verify RED**

Run: `..\..\.venv-fase1\Scripts\python.exe -m unittest tests.test_case_readiness -v`

Expected: `ModuleNotFoundError: No module named 'case_readiness'`.

- [ ] **Step 3: Implement the report**

Añadir dataclasses inmutables y consultas explícitas para documentos, incidencias, propietarios, perfil activo, reglas de concepto, coeficientes, lecturas efectivas, exportación validada, reparto completado y lote de cartas. `for_stage(stage)` devolverá sólo bloqueos de esa etapa y sus precondiciones.

```python
@dataclass(frozen=True)
class ReadinessBlocker:
    code: str
    stage: str
    message: str
    action: str
    count: int = 1

@dataclass(frozen=True)
class CaseReadinessReport:
    case_id: int
    blockers: tuple[ReadinessBlocker, ...]

    def for_stage(self, stage: str) -> tuple[ReadinessBlocker, ...]: ...
```

- [ ] **Step 4: Verify GREEN**

Run: `..\..\.venv-fase1\Scripts\python.exe -m unittest tests.test_case_readiness -v`

Expected: all tests in the module pass.

- [ ] **Step 5: Commit**

Run: `git add core/case_readiness.py tests/test_case_readiness.py && git commit -m "feat: centralizar preparacion de expedientes"`

### Task 2: Gates de Excel, reparto y cartas

**Files:**
- Modify: `core/case_workflow_actions.py`
- Modify: `core/document_review.py`
- Modify: `tests/test_case_workflow_actions.py`
- Modify: `tests/test_expedient_flow.py`

**Interfaces:**
- Consumes: `require_stage(connection, case_id, stage, project_root)` de Task 1.
- Produces: las acciones públicas existentes con un único criterio de bloqueo y mensajes accionables.

- [ ] **Step 1: Write failing workflow tests**

Añadir pruebas que llamen a `run_generate_excel` sin propietarios, `run_calculate_distribution` sin exportación vigente y `run_generate_letters` sin reparto vigente; exigir `WorkflowBlockedError` con la acción concreta y que el estado original permanezca intacto.

- [ ] **Step 2: Verify RED**

Run: `..\..\.venv-fase1\Scripts\python.exe -m unittest tests.test_case_workflow_actions -v`

Expected: at least the new missing-owner test fails because the action currently advances from state alone.

- [ ] **Step 3: Replace state-only gates**

Invocar `require_stage` antes de toda escritura. Mantener `_prepare_explicit_rerun` únicamente después de superar el informe. Hacer que `document_review.validate_case_ready` derive el estado máximo permitido por el informe y retroceda a `under_review` cuando falten datos estructurales.

- [ ] **Step 4: Verify GREEN and regressions**

Run: `..\..\.venv-fase1\Scripts\python.exe -m unittest tests.test_case_workflow_actions tests.test_expedient_flow tests.test_unified_ingestion_regressions -v`

Expected: all selected tests pass.

- [ ] **Step 5: Commit**

Run: `git add core/case_workflow_actions.py core/document_review.py tests/test_case_workflow_actions.py tests/test_expedient_flow.py && git commit -m "fix: bloquear etapas con requisitos reales"`

### Task 3: Navegación accionable en la interfaz

**Files:**
- Modify: `core/expedient_ui.py`
- Modify: `core/app.py`
- Modify: `tests/test_expedient_ui.py`

**Interfaces:**
- Consumes: `CaseReadinessReport` de Task 1.
- Produces: `guided_workspace_state(..., readiness_report=...)` y tarjetas agrupadas que ejecutan `action`.

- [ ] **Step 1: Write failing UI-state tests**

Probar que `MISSING_OWNERS` muestra “Importar propietarios”, que un perfil ausente ofrece “Preparar Excel”, que cero bloqueos y estado calculado ofrece reparto, y que al terminar un análisis de carpeta se refrescan comunidad, período, expediente y acción principal.

- [ ] **Step 2: Verify RED**

Run: `..\..\.venv-fase1\Scripts\python.exe -m unittest tests.test_expedient_ui -v`

Expected: new readiness-report arguments/actions are unsupported.

- [ ] **Step 3: Render the report**

Extender `GuidedWorkspaceState` con bloqueos resumidos y mapear `action` a los diálogos existentes. Tras el callback de incorporación de carpeta, recargar selectores, seleccionar el expediente procesado y llamar a `_refrescar_workspace()` en el hilo Tk.

- [ ] **Step 4: Verify GREEN**

Run: `..\..\.venv-fase1\Scripts\python.exe -m unittest tests.test_expedient_ui tests.test_onboarding_final_flow -v`

Expected: all selected tests pass.

- [ ] **Step 5: Commit**

Run: `git add core/expedient_ui.py core/app.py tests/test_expedient_ui.py && git commit -m "feat: guiar requisitos pendientes en la interfaz"`

### Task 4: Versión de esquema y verificación integral

**Files:**
- Modify: `core/db_migrations.py`
- Modify: `tests/test_db_migrations.py`
- Modify: `docs/manual-operacion.md`

**Interfaces:**
- Consumes: `MIGRATIONS` existente.
- Produces: `CURRENT_SCHEMA_VERSION == max(MIGRATIONS)` y documentación del nuevo gate.

- [ ] **Step 1: Write the failing invariant test**

```python
self.assertEqual(max(db_migrations.MIGRATIONS), db_migrations.CURRENT_SCHEMA_VERSION)
```

- [ ] **Step 2: Verify RED**

Run: `..\..\.venv-fase1\Scripts\python.exe -m unittest tests.test_db_migrations -v`

Expected: `13 != 14`.

- [ ] **Step 3: Fix the public schema version and document recovery**

Fijar la constante en 14 y explicar que “sin incidencias” no sustituye propietarios, lecturas, coeficientes ni un perfil activo; detallar la acción que ofrece cada bloqueo.

- [ ] **Step 4: Run full verification**

Run: `..\..\.venv-fase1\Scripts\python.exe -m unittest discover -s tests -q`

Expected: all tests pass, including the initial 498 plus the new readiness tests.

- [ ] **Step 5: Commit**

Run: `git add core/db_migrations.py tests/test_db_migrations.py docs/manual-operacion.md && git commit -m "fix: alinear esquema y documentar preparacion"`
