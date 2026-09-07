# Informe de Tarea 2 — perfil, plantilla y bootstrap canónicos

## Resultado

La publicación de onboarding genera y valida los conceptos canónicos
`acs_fixed`, `acs_variable`, `heating_fixed` y `heating_variable`. Para una
comunidad combinada, cada módulo conserva sus propios bindings de lectura y
sus cuatro parámetros económicos en hojas independientes. El exportador toma
los tipos de contador de la configuración normalizada y el reparto usa las
lecturas normalizadas de ACS y calefacción sin cruzarlas.

La excepción de parámetros vacíos se limita ahora a los bytes exactos de la
plantilla canónica recién creada. El layout guarda un estado
`fresh_onboarding` y la huella SHA-256 del libro. Un maestro posterior que ya
no coincide con esa huella recorre la validación ordinaria y produce
`MISSING_REQUIRED_FIELD` por cada parámetro obligatorio ausente.

## Evidencia TDD

### Conceptos canónicos

Se escribió primero
`test_profile_validation_rejects_cross_module_canonical_concept_source`. La
prueba cambia la fuente real de `heating_variable` por
`period_parameters.acs_variable_actual`.

RED observado:

```text
Ran 1 test
FAILED (failures=1)
AssertionError: ValueError not raised
```

El GREEN mínimo centralizó `MODULE_CONCEPTS` y las reglas de método/fuentes en
`excel_profiles`, y restringió esa validación a perfiles que contienen
`onboarding_configuration`. El focal del test y la suite de perfiles públicos
terminó con `Ran 10 tests ... OK`.

### Bindings de parámetros por módulo

Se añadió
`test_profile_validation_rejects_shared_combined_parameter_binding`, que hace
apuntar `heating_variable_actual` a la celda de ACS.

RED observado:

```text
Ran 1 test
FAILED (failures=1)
AssertionError: ValueError not raised
```

El GREEN añadió la tabla literal de celdas canónicas y exige que el conjunto y
cada coordenada coincidan para un perfil final de onboarding. El focal junto a
las pruebas de perfiles terminó con `Ran 11 tests ... OK`.

### Bootstrap vacío limitado a la plantilla inicial

Se escribió
`test_later_empty_onboarding_master_reports_missing_required_field`. La prueba
importa primero la plantilla canónica vacía, crea después otro maestro con
metadatos de período pero sin importes y vuelve a importarlo.

RED observado:

```text
Ran 1 test
FAILED (failures=1)
AssertionError: 0 not greater than 0
```

El GREEN hizo que el generador publique en el layout la huella y el estado de
bootstrap, y que el importador sólo aplique la excepción cuando la huella del
archivo archivado coincide. El mismo test terminó con `Ran 1 test ... OK`.

### Marcador obligatorio del perfil final

Se añadió
`test_profile_validation_rejects_final_onboarding_layout_without_bootstrap_marker`.

RED observado:

```text
Ran 1 test
FAILED (failures=1)
AssertionError: ValueError not raised
```

Tras exigir estado `fresh_onboarding` y un SHA-256 hexadecimal de 64 caracteres,
el focal de marcador, perfil combinado y maestro posterior terminó con
`Ran 3 tests ... OK`.

## Pruebas de aceptación

- `test_combined_onboarding_profile_has_canonical_concepts_and_bindings`
  verifica literalmente los cuatro conceptos, métodos, fuentes, ocho celdas,
  columnas de lectura independientes y huella de la plantilla.
- `test_combined_onboarding_exports_and_distributes_distinct_module_readings`
  recorre confirmar, importar la plantilla inicial, exportar y repartir. La
  salida oficial existe y el reparto conserva consumos distintos: 30 para
  `acs_variable` y 300 para `heating_variable`.
- `test_later_empty_onboarding_master_reports_missing_required_field` confirma
  que el libro posterior queda bajo revisión por campos obligatorios ausentes.

Los dos primeros tests de aceptación ejecutados juntos terminaron con
`Ran 2 tests ... OK`. La suite focal pedida por el plan terminó con:

```text
Ran 82 tests in 16.739s
OK
```

## Compatibilidad y decisiones

- `canonical_concepts_for_module` es la única fábrica usada por la publicación;
  el mismo módulo valida después esas reglas.
- El layout de onboarding exige exactamente cuatro parameter cells por módulo.
  ACS usa `LECTURAS ACS M3` y calefacción usa `LECTURAS CALEF KWH`.
- `excel_export_service` deriva los tipos de lectura de `reading_bindings` para
  perfiles de onboarding. Los perfiles históricos conservan la inferencia
  anterior por módulos y nombres de concepto.
- Las nuevas restricciones se activan sólo cuando existe
  `onboarding_configuration`; no se modificó ningún JSON de perfil público.
- No se tocó la extracción de evidencia ni la UI de confirmación de Task 1.

## Auto-revisión

- Una mutación de fuente de concepto, método, obligatoriedad, conjunto de
  conceptos, conjunto de parámetros o coordenada canónica queda cubierta por
  las regresiones de validación y la aceptación literal.
- Una mutación de cualquiera de los dos tipos de lectura altera el E2E y la
  huella de entrada que usa exportación/reparto.
- Quitar o malformar el marcador bloquea la carga del perfil final; cambiar
  cualquier byte del libro impide usar la excepción de plantilla fresca.
- El importador conserva el comportamiento idempotente de los mismos bytes de
  la plantilla inicial, pero no trata un libro modificado como plantilla vacía.
- Los perfiles públicos sin configuración de onboarding no pasan por las
  reglas nuevas.

## Incidencia de suite heredada

La primera ejecución global terminó con `Ran 214 tests` y un único fallo en el
guard de privacidad. La causa fue que `task-1-report.md`, incorporado por el
commit anterior, contenía rutas absolutas locales. El test ya existente fue el
RED; se sustituyeron sólo esas rutas por `<project-root>`. El focal del guard
terminó después con `Ran 1 test ... OK`.

## Verificación previa al cierre

```text
unittest discover -s tests -t . -q: Ran 214 tests in 30.879s, OK
compileall -q core tests: exit 0
git diff --check: exit 0
```

Los avisos de conversión LF/CRLF son propios de la configuración Git de
Windows y no representan errores de whitespace.

## Preocupaciones

- La huella es deliberadamente exacta: incluso un cambio legítimo de metadatos
  convierte el archivo en un maestro posterior y obliga a completar sus
  parámetros. Esto evita excepciones silenciosas.
- Una copia byte a byte conserva el mismo estado inicial y se considera un
  reintento idempotente de la plantilla; cualquier edición invalida la huella.
- El flujo sigue dependiendo de que los importes se normalicen en
  `period_parameters` antes de la exportación oficial; el bootstrap inicial no
  inventa importes a partir de celdas vacías.
