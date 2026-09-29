# Aceptación real y proveedores Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Ejecutar 658 y 644 en bases aisladas, producir controles saneados y medir el reconocimiento documental sin filtrar datos privados.

**Architecture:** El runner específico 658 se convertirá en un motor dirigido por un manifiesto local. La auditoría de proveedores reutilizará el mismo análisis y devolverá métricas agregadas con umbrales de publicación.

**Tech Stack:** Python, SQLite, JSON, unittest, LibreOffice headless, RapidOCR.

**Spec:** `docs/superpowers/specs/2026-09-29-cierre-produccion-portabilidad-y-correo-design.md`

## Global Constraints

- Los manifiestos, rutas privadas, bases y resultados identificables permanecen fuera de Git.
- Los originales se copian a una ejecución nueva y nunca se modifican.
- El informe público contiene sólo códigos, conteos, totales de control y huellas.
- Ninguna diferencia se fuerza para imitar el Excel manual.

## Review Focus

- Un manifiesto que apunta dentro del proyecto se rechaza antes de escribir.
- Dos comunidades en un lote nunca comparten facturas, lecturas ni propietarios.
- Reanudar una ejecución no duplica Excel, reparto ni cartas.
- Un proveedor desconocido no se aprende automáticamente con una firma débil.
- Un PDF defectuoso cuenta como error aislado y no aborta el corpus.

---

### Task 1: Manifiesto genérico de aceptación

**Files:**
- Create: `core/private_validation.py`
- Modify: `core/private_658_validation.py`
- Create: `tests/test_private_validation.py`
- Modify: `tests/test_private_658_validation.py`

**Interfaces:**
- Produces: `ValidationManifest`, `ValidationCase`, `load_manifest(path)`, `prepare_validation_run(...)`; el módulo 658 queda como adaptador compatible.

- [ ] Escribir pruebas RED para manifiesto con 658 y 644, rutas inseguras, código duplicado y campos desconocidos; ejecutar `..\..\.venv-fase1\Scripts\python.exe -m unittest tests.test_private_validation -v` y comprobar `ModuleNotFoundError`.
- [ ] Implementar dataclasses, validación estricta y copiado exclusivo de recursos públicos; repetir la prueba hasta GREEN.
- [ ] Adaptar `private_658_validation.py` sin romper su CLI; ejecutar ambos módulos de prueba.
- [ ] Commit: `git add core/private_validation.py core/private_658_validation.py tests/test_private_validation.py tests/test_private_658_validation.py && git commit -m "feat: generalizar validacion privada por comunidad"`.

### Task 2: Pipeline completo por caso

**Files:**
- Modify: `core/private_validation.py`
- Modify: `tests/test_private_validation.py`

**Interfaces:**
- Consumes: acciones públicas `run_generate_excel`, `run_calculate_distribution`, `run_generate_letters`.
- Produces: `ValidationCaseResult(code, status, issue_counts, invoice_count, reading_count, owner_count, distribution_total_cents, generated_letters)`.

- [ ] Escribir una aceptación sintética RED que cree SQLite vacía, incorpore dos comunidades y exija resultados separados, reanudación idempotente e informe sin nombres/rutas.
- [ ] Ejecutar la prueba y confirmar que falta `run_validation_manifest`.
- [ ] Implementar fases registradas `prepared`, `ingested`, `excel_validated`, `distributed`, `letters_rendered`; detenerse con estado explícito cuando el informe de preparación bloquee.
- [ ] Ejecutar `tests.test_private_validation` y `tests.test_onboarding_final_flow`; commit `feat: validar flujo completo en entorno aislado`.

### Task 3: Controles 658 y 644

**Files:**
- Modify: `core/private_validation.py`
- Modify: `tests/test_private_validation.py`
- Create: `docs/validacion-privada-comunidades.md`

**Interfaces:**
- Produce comparaciones tolerancia-cero en céntimos y conteos esperados declarados en manifiesto.

- [ ] Añadir pruebas RED para 62 unidades 658, suma registral/eligible separada y 644 con ACS/calefacción; incluir lectura cero que reutiliza la última lectura fiable y ausencia inicial que bloquea.
- [ ] Implementar comparadores independientes y render de muestra sin incluir datos en el informe.
- [ ] Ejecutar `tests.test_private_validation tests.test_case_distribution tests.test_case_letter_service`.
- [ ] Documentar las variables/rutas locales y commit `feat: añadir controles reales de comunidades`.

### Task 4: Umbrales del catálogo de proveedores

**Files:**
- Modify: `scripts/audit_provider_catalog.py`
- Modify: `core/provider_registry.py`
- Modify: `tests/test_provider_catalog_audit.py`
- Modify: `tests/test_provider_registry.py`

**Interfaces:**
- Produces: `AuditThresholds`, `AuditReport.passes(thresholds)` y salida JSON saneada.

- [ ] Escribir pruebas RED para reconocimiento mínimo 90 %, menos de 60 decisiones, cero falsos positivos y cero cruces de comunidad.
- [ ] Implementar cálculo y código de salida no cero al incumplir; mantener CIF/firma fuerte como único alta automática.
- [ ] Ejecutar las pruebas de proveedor, análisis y texto documental; commit `feat: convertir auditoria de proveedores en gate`.

### Task 5: Ejecución privada y suite

**Files:**
- Modify: `docs/validacion-privada-comunidades.md`

**Interfaces:**
- Consumes: manifiesto privado fuera de Git.
- Produces: informes locales verificables para 658 y 644.

- [ ] Ejecutar la suite completa y registrar el conteo en el ledger.
- [ ] Ejecutar el manifiesto real sólo con variables locales ya configuradas; si falta una fuente, conservar el estado `blocked` y el requisito exacto sin inventar nada.
- [ ] Inspeccionar visualmente Excel y una carta por comunidad con LibreOffice; documentar el procedimiento sin copiar resultados privados.
- [ ] Commit sólo de documentación pública: `git add docs/validacion-privada-comunidades.md && git commit -m "docs: documentar aceptacion privada multicomunidad"`.
