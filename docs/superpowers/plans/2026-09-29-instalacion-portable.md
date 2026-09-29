# Instalación portable Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Separar código y datos, permitir un primer arranque neutro y construir una distribución Windows `onedir` reproducible.

**Architecture:** `app_paths` resolverá recursos inmutables y un hogar de datos escribible. El arranque usará esas rutas por inyección; un asistente persistirá configuración no secreta y el empaquetado incluirá sólo recursos públicos.

**Tech Stack:** Python, pathlib, CustomTkinter, PyInstaller, unittest, Windows Credential Manager para secretos futuros.

**Spec:** `docs/superpowers/specs/2026-09-29-cierre-produccion-portabilidad-y-correo-design.md`

## Global Constraints

- `%LOCALAPPDATA%\Regularizaciones` es el hogar predeterminado.
- `REGULARIZACIONES_HOME` lo reemplaza de forma explícita.
- El paquete nunca contiene bases, fuentes, salidas, preferencias privadas ni credenciales.
- No se instala ninguna dependencia en tiempo de ejecución.

## Review Focus

- Rutas con espacios y caracteres no ASCII funcionan.
- Un hogar no escribible falla antes de abrir la ventana con mensaje accionable.
- Migrar una base existente crea copia verificada antes de modificarla.
- La configuración exportada excluye base, fuentes, correo y credenciales.
- Ejecutar desde un directorio distinto no cambia la resolución de recursos.

---

### Task 1: Resolución central de rutas

**Files:**
- Create: `core/app_paths.py`
- Create: `tests/test_app_paths.py`
- Modify: `core/app.py`
- Modify: `core/ui_moderna.py`

**Interfaces:**
- Produces: `ApplicationPaths.resolve(environ, executable, local_app_data)` y propiedades `database`, `sources`, `outputs`, `backups`, `logs`, `resources`.

- [ ] Escribir pruebas RED para ruta predeterminada, override portable, recursos empaquetados y directorio no escribible.
- [ ] Ejecutar `tests.test_app_paths` y confirmar que el módulo falta.
- [ ] Implementar resolución/creación; reemplazar constantes escribibles de `app.py`; mover `ui_prefs.json` al hogar.
- [ ] Ejecutar `tests.test_app_paths tests.test_expedient_ui`; commit `feat: separar rutas de aplicacion y datos`.

### Task 2: Primer arranque y configuración exportable

**Files:**
- Create: `core/installation_settings.py`
- Create: `tests/test_installation_settings.py`
- Modify: `core/app.py`
- Modify: `core/letter_settings.py`

**Interfaces:**
- Produces: `OfficeSettings`, `load_settings`, `save_settings_atomic`, `export_portable_settings`.

- [ ] Escribir pruebas RED para configuración neutra, escritura atómica, logo opcional, ruta LibreOffice y exportación sin secretos/rutas privadas.
- [ ] Implementar JSON versionado y asistente mostrado sólo cuando falta configuración válida.
- [ ] Ejecutar pruebas de ajustes, cartas y UI; commit `feat: añadir primer arranque portable`.

### Task 3: Copias y migración segura

**Files:**
- Create: `core/backup_service.py`
- Create: `tests/test_backup_service.py`
- Modify: `core/gestor_bd.py`

**Interfaces:**
- Produces: `create_verified_backup(database, backup_dir)` y `open_and_migrate_database(paths)`.

- [ ] Escribir pruebas RED para copia SQLite válida, copia corrupta rechazada y migración que conserva filas.
- [ ] Implementar backup mediante API SQLite y `PRAGMA integrity_check`; migrar sólo tras verificar la copia.
- [ ] Ejecutar `tests.test_backup_service tests.test_db_migrations`; commit `feat: proteger migraciones con copias verificadas`.

### Task 4: Build `onedir`

**Files:**
- Create: `Regularizaciones.spec`
- Create: `scripts/build_windows.ps1`
- Create: `tests/test_packaging_manifest.py`
- Modify: `requirements.txt`

**Interfaces:**
- Produces: `dist/Regularizaciones/Regularizaciones.exe` con recursos públicos y sin datos privados.

- [ ] Escribir una prueba RED que inspeccione el manifiesto de datos permitido y rechace `gestion.db`, `ui_prefs.json`, `entrada`, `salidas` y credenciales.
- [ ] Añadir PyInstaller como dependencia de desarrollo documentada, eliminar la instalación automática de CustomTkinter y definir recursos explícitos.
- [ ] Ejecutar prueba de manifiesto y construir con `powershell -ExecutionPolicy Bypass -File scripts/build_windows.ps1`.
- [ ] Lanzar el ejecutable contra un `REGULARIZACIONES_HOME` temporal, verificar creación inicial y cerrar sin modificar el repositorio.
- [ ] Commit `build: añadir distribucion Windows portable`.

### Task 5: Manual portable

**Files:**
- Create: `docs/instalacion-portable.md`
- Modify: `README.md`

**Interfaces:**
- Produce instrucciones de instalar, configurar, actualizar, copiar y recuperar.

- [ ] Documentar comandos exactos, estructura de datos, LibreOffice, backup y traslado a otro despacho.
- [ ] Ejecutar suite completa y `git diff --check`.
- [ ] Commit `docs: explicar instalacion y recuperacion portable`.
