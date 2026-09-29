# Integración y publicación Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Convertir los cuatro bloques funcionales en una versión reproducible, verificable y recuperable sin perder datos locales.

**Architecture:** Un comando de verificación agregará suite, auditorías y build; la publicación se hará sólo después de una revisión fresca del diff completo y una prueba limpia. La integración conserva historia y nunca fuerza `master`.

**Tech Stack:** Git, PowerShell, unittest, PyInstaller, LibreOffice.

**Spec:** `docs/superpowers/specs/2026-09-29-cierre-produccion-portabilidad-y-correo-design.md`

## Global Constraints

- No se hace merge, push, etiqueta ni publicación sin autorización explícita para ese efecto externo.
- Nunca se añaden `gestion.db`, fuentes, salidas, preferencias, informes privados o credenciales.
- Todos los gates deben ejecutarse de nuevo desde el commit candidato.
- Los archivos locales del usuario se preservan durante la integración.

## Review Focus

- El verificador falla si omite una suite o auditoría requerida.
- El paquete limpio no depende del checkout ni del entorno virtual del desarrollador.
- Una actualización conserva una base antigua tras migrarla.
- La documentación coincide con botones y rutas reales.
- El diff de publicación no contiene archivos privados ignorados por error.

---

### Task 1: Comando de verificación reproducible

**Files:**
- Create: `scripts/verify_release.ps1`
- Create: `tests/test_release_manifest.py`

**Interfaces:**
- Produce un código de salida único que agrega suite, compileall, auditoría sintética, build y escaneo de artefactos.

- [ ] Escribir prueba RED del manifiesto permitido y exclusiones privadas.
- [ ] Implementar script con comandos explícitos y parada al primer fallo.
- [ ] Ejecutar la prueba y el script sin aceptación privada; commit `build: añadir gate reproducible de publicacion`.

### Task 2: Prueba de instalación limpia y migrada

**Files:**
- Create: `tests/test_release_smoke.py`
- Modify: `scripts/verify_release.ps1`

**Interfaces:**
- Consume el paquete `onedir`; produce dos hogares temporales verificados.

- [ ] Escribir prueba RED para arranque con base vacía y copia v13/v14.
- [ ] Añadir modo de smoke no interactivo que inicializa rutas, abre/migra DB y valida recursos sin mostrar Tk.
- [ ] Ejecutar prueba y build; commit `test: verificar instalacion limpia y migrada`.

### Task 3: Manual final

**Files:**
- Modify: `README.md`
- Modify: `docs/manual-operacion.md`
- Modify: `docs/instalacion-portable.md`

**Interfaces:**
- Produce un índice único de instalación, alta, expediente, incidencias, Excel, reparto, cartas, correo, backup y recuperación.

- [ ] Recorrer la aplicación y corregir nombres de botones/rutas en documentos.
- [ ] Ejecutar `git diff --check` y el escaneo de privacidad.
- [ ] Commit `docs: completar manual operativo de regularizaciones`.

### Task 4: Revisión final y correcciones

**Files:**
- Modify: sólo archivos requeridos por hallazgos Critical/Important.

**Interfaces:**
- Produce un paquete de revisión desde el merge-base y un ledger de decisiones.

- [ ] Crear el review package y solicitar una revisión fresca del diff completo contra los cinco planes y la especificación.
- [ ] Regraduar hallazgos por impacto; corregir Critical/Important con prueba RED→GREEN y suite completa; registrar Minor sin alterar código.
- [ ] Ejecutar `scripts/verify_release.ps1`, aceptación privada disponible y `git diff --check`.
- [ ] Commit de correcciones sólo si existen, con mensaje `fix: resolver hallazgos de publicacion`.

### Task 5: Integración autorizada

**Files:**
- No source changes.

**Interfaces:**
- Produce rama integrada, etiqueta y PR/remote actualizados sólo después de autorización.

- [ ] Mostrar commits, verificaciones, rulings, minors y artefactos al usuario.
- [ ] Tras autorización explícita, actualizar remoto, integrar sin force, ejecutar de nuevo el gate desde `master` y crear etiqueta versionada.
- [ ] Adjuntar o actualizar el Pull Request y entregar comandos de prueba local.
