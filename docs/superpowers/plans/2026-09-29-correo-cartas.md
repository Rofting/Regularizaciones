# Correo de cartas Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Crear borradores EML auditables y habilitar envío SMTP opcional sin duplicados ni credenciales persistidas en archivos de la aplicación.

**Architecture:** Un modelo de lote/entrega persistirá identidad y estado; transportes inyectables generarán EML o enviarán SMTP. La interfaz prepara, revisa y confirma lotes, y las pruebas usan siempre un transporte falso.

**Tech Stack:** Python email/smtplib, SQLite, unittest, Windows Credential Manager mediante keyring.

**Spec:** `docs/superpowers/specs/2026-09-29-cierre-produccion-portabilidad-y-correo-design.md`

## Global Constraints

- EML es el transporte predeterminado y no requiere credenciales.
- SMTP no se conecta sin configuración y confirmación explícita.
- Contraseñas nunca se guardan en Git, JSON, SQLite ni logs.
- La huella de destinatario+carta+asunto+cuerpo impide duplicados.

## Review Focus

- Correos vacíos o inválidos se omiten sin bloquear las cartas.
- Dos propietarios con el mismo correo aparecen como duplicados revisables.
- Reintentar un lote no reenvía entregas `sent`.
- Un adjunto modificado produce una huella y un borrador nuevos.
- Un error SMTP no deja el lote en `sending` para siempre.

---

### Task 1: Migración y modelo de correo

**Files:**
- Modify: `core/db_migrations.py`
- Create: `core/mail_models.py`
- Create: `tests/test_mail_migrations.py`

**Interfaces:**
- Produces tablas `mail_runs` y `mail_deliveries`, estados tipados y versión de esquema 15.

- [ ] Escribir pruebas RED que migren base vacía y base v14, verifiquen restricciones, claves y que `CURRENT_SCHEMA_VERSION == 15`.
- [ ] Implementar migración 15 y dataclasses; ejecutar la prueba hasta GREEN.
- [ ] Ejecutar `tests.test_db_migrations tests.test_mail_migrations`; commit `feat: persistir lotes de correo`.

### Task 2: Generador EML idempotente

**Files:**
- Create: `core/mail_service.py`
- Create: `tests/test_mail_service.py`

**Interfaces:**
- Produces: `prepare_mail_run(connection, id_case, output_root, subject_template, body_template)` y `generate_eml_drafts(...)`.

- [ ] Escribir pruebas RED con dos cartas, un correo válido y otro ausente; comprobar MIME, adjunto DOCX, estado `draft`, omisión auditable y reutilización por huella.
- [ ] Implementar normalización, plantillas limitadas, SHA-256 y publicación atómica de `.eml`.
- [ ] Ejecutar `tests.test_mail_service`; commit `feat: generar borradores eml auditables`.

### Task 3: SMTP seguro e inyectable

**Files:**
- Create: `core/mail_transport.py`
- Create: `tests/test_mail_transport.py`
- Modify: `requirements.txt`

**Interfaces:**
- Produces: `MailTransport` protocol, `SmtpTransport`, `FakeMailTransport`, `send_mail_run(..., confirmed_by)`.

- [ ] Escribir pruebas RED para ausencia de confirmación, recuperación de contraseña por identificador, fallo de un destinatario, reintento y no reenvío.
- [ ] Implementar transporte con `keyring`, timeouts, TLS explícito y errores saneados; no incluir servidor real en pruebas.
- [ ] Ejecutar pruebas de correo; commit `feat: enviar correo opcional con confirmacion`.

### Task 4: Interfaz de preparación y confirmación

**Files:**
- Modify: `core/expedient_ui.py`
- Modify: `core/app.py`
- Modify: `tests/test_expedient_ui.py`

**Interfaces:**
- Consumes servicios de Tasks 2-3.
- Produce resumen de válidos/ausentes/duplicados, muestra, exclusión y confirmación final con conteo.

- [ ] Escribir pruebas RED de estado y comandos sin abrir Tk real.
- [ ] Añadir paso de correo posterior a cartas; mantener “Generar borradores” disponible aunque SMTP no esté configurado.
- [ ] Ejecutar `tests.test_expedient_ui tests.test_mail_service tests.test_mail_transport`.
- [ ] Commit `feat: añadir correo al flujo guiado`.

### Task 5: Manual y suite

**Files:**
- Modify: `docs/manual-operacion.md`

**Interfaces:**
- Produce procedimiento de borrador, envío, reintento y credenciales.

- [ ] Documentar que pruebas y aceptación nunca envían correo real.
- [ ] Ejecutar suite completa y revisar que logs/resultados no contienen contraseñas ni cuerpos.
- [ ] Commit `docs: documentar entrega segura de cartas`.
