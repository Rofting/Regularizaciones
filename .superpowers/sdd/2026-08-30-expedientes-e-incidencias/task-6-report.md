# Informe de la Tarea 6: robustez del expediente y navegación de incidencias

## Ciclos RED/GREEN

### Estado al añadir una fuente nueva

- RED: `test_new_incomplete_source_returns_ready_case_to_review` creó un expediente listo, añadió una segunda factura sin `importe_total` y falló porque el estado real seguía siendo `ready_for_calculation` en vez de `under_review`.
- GREEN: el registro de una fuente realmente nueva conserva `draft → gathering_sources` y devuelve a `under_review` los estados avanzados que admiten esa transición. La regresión y la prueba de prohibición de saltar desde `draft` a `ready_for_calculation` pasan.

### Recuperación idempotente de ingestión parcial

- RED: `test_retry_after_archiving_failure_completes_review_without_duplicates` inyectó un fallo en `record_candidates` después de registrar y archivar la fuente. En el reintento, la consulta devolvió cero candidatos en lugar de los dos esperados porque el camino de SHA-256 duplicado retornaba antes de completar la revisión.
- GREEN: el reintento inserta sólo candidatos ausentes y vuelve a construir idempotentemente las incidencias. Se verificó una única fuente, archivo intacto, dos candidatos sin duplicar, una única incidencia abierta, estado `under_review` y `created=False` en el reintento. También pasa `test_readding_resolved_source_preserves_manual_value_and_correction`: no se sobrescriben valores manuales ni se duplican correcciones o incidencias resueltas.

### Responsable de auditoría

- RED 1: `test_resolving_issue_normalizes_responsible_person_before_auditing` observó `"  gestora principal  "` en auditoría en vez de `"gestora principal"`.
- GREEN 1: `resolved_by` se recorta antes de persistir.
- RED 2: `test_resolving_issue_rejects_blank_responsible_before_auditing` no obtuvo el `ValueError` esperado y la incidencia podía resolverse con un responsable formado sólo por espacios.
- GREEN 2: el responsable vacío se rechaza antes de abrir la transacción; no se crea corrección y la incidencia permanece abierta.

### Índices de migración

- `test_migration_two_creates_case_and_review_tables` consulta `sqlite_master` sobre una base migrada y comprueba `idx_cases_community`, `idx_documents_case` e `idx_issues_case_open`. No inspecciona texto de código fuente.

## UI persistente y progreso

- El selector `Elegir expediente` se alimenta con `list_cases(id_comunidad)`, usa el `id_case` real, valida la pertenencia y selecciona el expediente persistido más reciente al cargar una comunidad.
- Cambiar de comunidad borra antes el contexto anterior y vuelve a cargar únicamente sus expedientes.
- La bandeja scrollable muestra campo, archivo y explicación, con acciones `Abrir archivo` y `Resolver` ligadas a cada incidencia concreta. La acción general ya no abre siempre la primera incidencia.
- Los estados vacíos guían a elegir expediente o indican `Sin incidencias pendientes`. El resumen incorpora un carril lateral con los colores ya existentes para borrador, revisión y listo.
- La incorporación múltiple activa el progreso antes de arrancar el hilo, muestra y registra `n de total`, continúa tras errores individuales, conserva cada error en el log y resume nuevas, duplicadas y fallidas. La infraestructura existente restaura los botones y oculta el progreso en `finally`.

## Verificación

- Suite completa: `46` pruebas, `OK`.
- Compilación de los seis módulos solicitados: salida vacía y código `0`.
- `git diff --check`: sin errores; Git sólo informa de la conversión futura LF/CRLF y del fichero global de exclusiones sin permiso de lectura.
- Inspección AST de llamadas `CTkButton`: ninguna contiene `justify=`.
- Búsqueda de `issues[0]`/`incidencias[0]` en la UI: sin coincidencias.

## Limitaciones de comprobación manual

No se lanzó Tk en modo headless, según el brief. Queda pendiente una comprobación manual en Windows con pantalla: dimensiones reales de la tarjeta al redimensionar, scroll con muchas incidencias, apertura del archivo archivado y secuencia de foco/diálogos. La lógica de dominio, persistencia y compilación sí está cubierta por las verificaciones automatizadas.

## Commit

- Mensaje único: `Refuerza expedientes y bandeja de incidencias`.
- El SHA final se comunicará en la respuesta de cierre, porque este informe forma parte del propio commit.
