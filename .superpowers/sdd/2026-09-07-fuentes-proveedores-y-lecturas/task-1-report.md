# Informe de Tarea 1 — evidencia y decisiones por módulo

## Trabajo retomado

La tarea se retomó con cambios sin commit en `core/community_onboarding.py`,
`core/excel_profiles.py` y `tests/test_community_onboarding.py`. Se conservaron y
revisaron esos cambios; no se hizo reset ni se descartó trabajo heredado.

La implementación separa la evidencia de factura y de lectura, formula las
confirmaciones obligatorias, construye un binding por módulo activo y publica
una única configuración resuelta. Durante la revisión se añadió una regresión
de compatibilidad para configuraciones de onboarding ya persistidas.

## Evidencia TDD

### RED heredado

El traspaso del turno anterior confirma que se observaron en RED, antes de la
implementación de producción, las tres regresiones vinculantes:

- `test_invoice_detects_heating_but_is_not_a_meter_reading`
- `test_combined_modules_require_a_confirmed_reading_for_each_module`
- `test_missing_invoice_period_and_missing_reading_create_required_questions`

La salida literal de aquella ejecución no quedó en el worktree, por lo que no
se reconstruye ni se inventa aquí. El fallo correspondía al contrato anterior:
no había evidencia de factura modelada y existía una sola columna global de
lectura.

### RED adicional de compatibilidad

Comando ejecutado:

```powershell
& '<project-root>\.venv-fase1\Scripts\python.exe' -m unittest tests.test_community_onboarding.CommunityOnboardingTest.test_profile_validation_accepts_existing_onboarding_configuration -v
```

Resultado: `FAILED (errors=1)`. La carga de una configuración persistida con el
esquema anterior falló con `ValueError: Faltan campos en
onboarding_configuration: invoice_decisions, not_applicable_modules`.

Tras hacer opcionales sólo en lectura los campos incorporados por esta tarea,
el mismo comando terminó con `Ran 1 test ... OK`.

## GREEN y verificación

Comandos ejecutados y resultados:

```powershell
& '<project-root>\.venv-fase1\Scripts\python.exe' -m unittest tests.test_community_onboarding tests.test_excel_profiles -v
```

Resultado: `Ran 39 tests ... OK`.

```powershell
& '<project-root>\.venv-fase1\Scripts\python.exe' -m unittest tests.test_expedient_ui -v
```

Resultado: `Ran 17 tests ... OK`.

```powershell
& '<project-root>\.venv-fase1\Scripts\python.exe' -m unittest discover -s tests -t . -q
```

Resultado: `Ran 203 tests in 28.321s` y `OK`.

```powershell
& '<project-root>\.venv-fase1\Scripts\python.exe' -m compileall -q core/community_onboarding.py core/excel_profiles.py tests/test_community_onboarding.py
git diff --check
```

Resultado: ambos comandos terminaron con código 0. `git diff --check` no informó
errores de whitespace; Git mostró únicamente los avisos de conversión LF/CRLF
del entorno Windows.

El alias `python` de WindowsApps no era ejecutable y el Python integrado de
Codex no incluía `xlrd`; por ello todas las verificaciones válidas se hicieron
con el intérprete `.venv-fase1` del proyecto.

## Interfaces añadidas o modificadas

- `OnboardingDraft` conserva `invoice_evidence` y `reading_evidence` por
  separado, ambos con valores por defecto para borradores existentes.
- `InvoiceEvidence` contiene sólo huella de fuente, proveedor, período,
  importe y concepto; no tiene lecturas ni cabeceras de contador.
- `ReadingEvidence` contiene sólo huella, contador, columnas, fechas,
  lecturas y candidatos de suministro; no tiene importes de factura.
- `ReadingBinding` mantiene `module`, `column` y `source_sha256s`, y añade las
  decisiones confirmadas `meter` y `date` con valores por defecto compatibles.
- `InvoiceDecision` normaliza por factura `source_sha256`, `provider`,
  `period`, `amount` y `concept`.
- `OnboardingConfiguration` añade `invoice_decisions` y
  `not_applicable_modules`, ambos opcionales para construcciones anteriores.
- `resolve_onboarding_configuration(draft, answers)` es el único punto que
  convierte respuestas en módulos activos, bindings, decisiones de factura y
  trazas de `NO_APLICA`.
- Las lecturas combinadas usan `reading_column:ACS` y
  `reading_column:CALEFACCION`; una columna compartida debe responderse en las
  dos claves.
- El validador de perfiles comprueba la coherencia entre `service_decision` y
  módulos, un binding por módulo, decisiones de factura y nombres de fuente
  sin rutas. Al leer perfiles anteriores admite la ausencia de los nuevos
  campos, sin relajar la validación de perfiles recién generados.

## Auto-revisión

- La factura sólo contribuye evidencia de proveedor/período/importe/concepto y
  candidatos de servicio; nunca se incorpora a `reading_sources` ni aporta una
  columna de lectura.
- La lectura sólo contribuye contador/columna/fecha/valor/suministro; el modelo
  no expone un campo de importe.
- Ausencias y ambigüedades generan preguntas `required`; las preguntas de
  módulos inactivos sólo admiten `NO_APLICA`, que queda trazado.
- Cada módulo activo exige una fuente de lecturas, una columna no vacía y una
  huella de fuente. En la decisión combinada no existe fallback a la columna
  global.
- Se conservaron los constructores anteriores mediante campos finales con
  valores por defecto y se añadió una prueba que carga el esquema de perfil
  previo.
- Las trazas persistidas usan nombre base y SHA-256, no rutas absolutas ni el
  texto completo extraído de los documentos.

## Preocupaciones

- La extracción es deliberadamente conservadora y basada en etiquetas y
  cabeceras reconocibles. Diseños de proveedor nuevos o PDFs poco estructurados
  producirán preguntas manuales obligatorias en vez de valores inferidos.
- En archivos `.xls` se detectan cabeceras, pero los valores de lectura no se
  enumeran en esta tarea; la ausencia se mantiene como confirmación obligatoria
  y no se convierte en importe ni en lectura procedente de factura.
- Los avisos LF/CRLF proceden de la configuración Git de Windows y no son
  errores de diff.

## Ronda 1 de corrección

### Incidencias verificadas

La revisión de `cf36f32` confirmó tres causas:

- `ReadingEvidence.readings` se usaba sólo para decidir si preguntar por la
  columna y no se convertía en una decisión persistida del binding.
- Las preguntas de contador y fecha se calculaban con candidatos agregados de
  todos los módulos. Por ello, la evidencia única de ACS ocultaba la ausencia
  de contador y fecha de CALEFACCION, y el resolver podía publicar cadenas
  vacías.
- El validador leía `invoice_decisions` con un valor por defecto vacío sin
  distinguir un perfil histórico de un payload generado con el contrato
  actual.

### RED de la corrección

Se añadieron primero cinco regresiones focales y se ejecutaron con:

```powershell
& '<project-root>\.venv-fase1\Scripts\python.exe' -m unittest tests.test_community_onboarding.CommunityOnboardingTest.test_multiple_reading_values_require_a_per_module_decision_and_persist_it tests.test_community_onboarding.CommunityOnboardingTest.test_missing_reading_value_blocks_publication_for_the_active_module tests.test_community_onboarding.CommunityOnboardingTest.test_meter_and_date_questions_are_scoped_to_each_module tests.test_community_onboarding.CommunityOnboardingTest.test_profile_validation_rejects_incomplete_new_invoice_configuration tests.test_community_onboarding.CommunityOnboardingTest.test_profile_validation_rejects_empty_new_reading_decisions -v
```

Resultado observado antes de modificar producción: `Ran 5 tests` y
`FAILED (failures=3, errors=2)`. En concreto, no existían las preguntas
`reading_value:ACS`, no aparecían `reading_meter:CALEFACCION` ni
`reading_date:CALEFACCION`, y el validador aceptaba tanto una lista de
decisiones de factura vacía como un valor de lectura vacío en un perfil nuevo.

### GREEN y regresión

El mismo comando focal terminó después con `Ran 5 tests ... OK`.

La suite de onboarding y perfiles terminó con:

```text
Ran 44 tests in 4.633s
OK
```

La primera ejecución de la suite completa mostró cinco fallos en los tests de
UI: sus borradores sintéticos no incluían evidencia de contador, fecha ni
valor y, correctamente, el resolver ya no permitía avanzar. Se completaron
esos fixtures sin cambiar el código de UI. El focal de UI terminó con
`Ran 17 tests ... OK` y la suite completa posterior con:

```text
Ran 208 tests in 29.760s
OK
```

### Interfaces y validación resultantes

- `ReadingBinding` añade `value`; el payload persiste `column`, `meter`,
  `date` y `value` por módulo.
- La respuesta manual nueva es `reading_value:<MODULO>`. Ausencia o múltiples
  candidatos generan una pregunta obligatoria; una respuesta contradictoria
  con candidatos detectados se rechaza.
- Contador, fecha, valor y columnas admisibles se obtienen mediante candidatos
  filtrados por módulo. Un binding activo nunca puede resolverse con contador,
  fecha o valor vacíos.
- Los payloads nuevos incluyen `onboarding_configuration.schema_version = 2`.
  En ese esquema son obligatorios `invoice_decisions`,
  `not_applicable_modules` y los tres detalles `meter`, `date`, `value` de cada
  binding.
- Para el esquema 2, las huellas de `invoice_decisions` deben coincidir
  exactamente con las fuentes `invoice_pdf`, de modo que cada factura tenga
  una decisión completa y ninguna decisión quede sin fuente.
- Un perfil histórico sigue siendo válido si carece por completo del marcador
  y de todos los campos incorporados por este contrato. La presencia parcial
  de campos actuales no activa el camino histórico ni relaja la validación.

### Auto-revisión de la corrección

- `InvoiceEvidence` e `InvoiceDecision` continúan sin campos de lectura;
  `ReadingEvidence` y `ReadingBinding` continúan sin importes de factura.
- La columna combinada sigue requiriendo respuestas independientes para ACS y
  CALEFACCION, incluso si ambas respuestas eligen la misma columna.
- `NO_APLICA` conserva su semántica: se rechaza para módulos activos y sólo se
  traza cuando se responde para un módulo inactivo.
- El diálogo ya recorría dinámicamente todas las preguntas; no fue necesario
  añadir checkboxes ni un segundo canal de decisiones.

### Preocupación residual

Un documento anterior al esquema 2 no contiene una marca temporal fiable. La
compatibilidad se reconoce estructuralmente sólo cuando faltan a la vez el
marcador y todos los campos actuales. Todos los payloads creados por esta
versión llevan el marcador 2, por lo que no pueden omitir parcialmente
decisiones nuevas sin que el validador los rechace.

### Verificación final de la ronda

La ejecución fresca final de `unittest discover` terminó con
`Ran 208 tests in 28.884s` y `OK`. A continuación, `compileall -q` sobre
`core/community_onboarding.py`, `core/excel_profiles.py`,
`tests/test_community_onboarding.py` y `tests/test_expedient_ui.py` terminó con
código 0. `git diff --check` también terminó con código 0; sólo mostró los
avisos LF/CRLF propios del entorno Windows.

## Ronda 2 de corrección — sin bypass de esquema

### Causa y RED

El reconocimiento anterior de compatibilidad trataba una configuración como
histórica cuando se habían eliminado conjuntamente `schema_version`,
`invoice_decisions`, `not_applicable_modules` y los detalles de cada binding.
Eso permitía degradar un payload nuevo, aun conservando en `sources` una
fuente `invoice_pdf`, y eludía la comprobación de que cada factura tuviera su
decisión completa.

Se cambió primero la regresión por
`test_profile_validation_rejects_new_configuration_stripped_to_look_historical`.
El test elimina exactamente esos campos de un perfil nuevo, comprueba que la
traza `invoice_pdf` sigue presente y exige un error sobre `schema_version`.

Comando RED ejecutado:

```powershell
& '<project-root>\.venv-fase1\Scripts\python.exe' -m unittest tests.test_community_onboarding.CommunityOnboardingTest.test_profile_validation_rejects_new_configuration_stripped_to_look_historical -v
```

Resultado antes del cambio: `Ran 1 test ... FAILED (failures=1)` con
`AssertionError: ValueError not raised`.

### Cambio mínimo y GREEN

`excel_profiles._onboarding_configuration` ahora diferencia sólo dos casos:

- Sin `onboarding_configuration` (`None`): perfil público histórico válido,
  que mantiene la carga de perfiles existentes como `658_acs_v1`.
- Con `onboarding_configuration`: contrato actual obligatorio. Exige
  `schema_version == 2`, `invoice_decisions`, `not_applicable_modules` y
  `meter`, `date` y `value` en cada binding; además, siempre compara las
  huellas de las decisiones con todas las fuentes `invoice_pdf`.

El mismo RED y la carga explícita de un perfil público terminaron con
`Ran 2 tests ... OK`. La suite focal de onboarding y perfiles terminó con
`Ran 44 tests in 4.556s` y `OK`.

### Revisión de alcance

No se modificó la generación de payloads: ya publica el esquema 2 completo.
El cambio se limita al borde de validación, donde existía el bypass. No se
aceptan configuraciones de onboarding pre-esquema; la compatibilidad histórica
queda reservada de forma vinculante para perfiles sin configuración de
onboarding.

### Verificación final de la ronda

La ejecución fresca de `unittest discover -s tests -t . -q` finalizó sin
errores. `compileall -q core\\excel_profiles.py
tests\\test_community_onboarding.py` y `git diff --check` también terminaron
sin errores; los únicos mensajes de Git fueron avisos locales de conversión
LF/CRLF.
