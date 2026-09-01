# Task 3 report — Confirmación atómica y expediente inicial

## Estado

Implementados `OnboardingResult` y `confirm_onboarding(...)`. La confirmación
exige fechas completas o ninguna, verifica que las fuentes no hayan cambiado
desde el análisis, genera y valida perfil y plantilla antes de persistir la
comunidad, y crea período/expediente mediante los servicios existentes.

El JSON se escribe primero en un temporal hermano, se vuelve a leer y se
valida antes de publicarlo. La plantilla canónica se genera desde
`excel_generator`, contiene `DATOS` y únicamente las hojas de los módulos
confirmados, y publica un `workbook_layout` con las mismas hojas y rangos.

Las fuentes se registran después de validar ambos artefactos. No se envuelve
`register_source_document` en una transacción externa. Ante error tardío se
compensan por identificador las filas de fuente, expediente, período y
comunidad creadas por la llamada, además de sus archivos archivados, perfil,
plantilla, temporales y directorios locales vacíos creados por el alta.

## TDD

- RED inicial: 3 errores esperados por ausencia de `confirm_onboarding`.
- GREEN inicial: 13 pruebas de onboarding correctas.
- RED de compensación reforzada: quedó `config/excel_profiles/` vacío tras un
  fallo tardío.
- RED de fallo posregistro: una fuente ya confirmada por el servicio pero no
  devuelta al llamante dejaba su archivo; la compensación final elimina ahora
  el directorio exclusivo del expediente creado por el alta.
- GREEN final focal: 21 pruebas de onboarding y expediente correctas.

## Privacidad y alcance

Las pruebas crean únicamente comunidad 900 y fuentes sintéticas temporales.
No se han leído fuentes reales ni modificado el perfil o flujo de 658. El
`.gitignore` ignora perfiles JSON generados localmente y conserva una excepción
explícita para `658_acs_v1.json`.

## Verificación

- Focales: 21 pruebas correctas.
- Suite completa: 169 pruebas ejecutadas; 1 fallo preexistente del guard de
  privacidad porque el plan versionado de esta rama ya contiene rutas locales.
  Tanto el guard como esas rutas están presentes en `HEAD`, fuera de los
  archivos autorizados de Task 3.
- Compilación: `python -m compileall -q core tests` correcta.
- `git diff --check`: correcto.

## Preocupaciones

La suite global no puede quedar verde sin corregir el documento de plan ya
versionado que contiene rutas locales. No se modifica aquí porque Task 3 limita
expresamente el alcance a los cuatro archivos funcionales y este informe.
