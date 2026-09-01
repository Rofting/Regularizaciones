# Task 3 report — Confirmación atómica y expediente inicial

## Estado

Implementados `OnboardingResult` y `confirm_onboarding(...)`. La confirmación
exige fechas completas o ninguna, verifica que las fuentes no hayan cambiado
desde el análisis, genera y valida perfil y plantilla antes de persistir la
comunidad, y crea período/expediente mediante los servicios existentes.

El JSON se escribe primero en un temporal hermano, se vuelve a leer y se
valida antes de publicarlo. Perfil y plantilla se publican mediante creación
atómica exclusiva: un destino aparecido concurrentemente provoca conflicto y
nunca se sustituye. La plantilla canónica se genera desde
`excel_generator`, contiene `DATOS` y únicamente las hojas de los módulos
confirmados, y publica un `workbook_layout` con las mismas hojas y rangos.

Las fuentes se registran después de validar ambos artefactos. No se envuelve
`register_source_document` en una transacción externa. Ante error tardío se
compensan por identificador las filas de fuente, expediente, período y
comunidad creadas por la llamada, además de sus archivos archivados, perfil,
plantilla, temporales y directorios locales vacíos creados por el alta. No se
borra recursivamente el expediente: una ruta se elimina sólo si procede de una
fila registrada por la fuente en curso.

## TDD

- RED inicial: 3 errores esperados por ausencia de `confirm_onboarding`.
- GREEN inicial: 13 pruebas de onboarding correctas.
- RED de compensación reforzada: quedó `config/excel_profiles/` vacío tras un
  fallo tardío.
- RED de fallo posregistro: una fuente ya confirmada por el servicio pero no
  devuelta al llamante dejaba su archivo; la compensación recupera su fila por
  SHA y elimina sólo esa ruta.
- RED de concurrencia: plantilla y JSON concurrentes eran sustituidos por
  `os.replace`; ambos sentinels sobreviven ahora al conflicto exclusivo.
- RED de rollback selectivo: un sentinel ajeno dentro del archivo del
  expediente era eliminado por `rmtree`; ahora sobrevive intacto.
- GREEN final focal: 23 pruebas de onboarding y expediente correctas.

## Privacidad y alcance

Las pruebas crean únicamente comunidad 900 y fuentes sintéticas temporales.
No se han leído fuentes reales ni modificado el perfil o flujo de 658. El
`.gitignore` ignora sólo la forma runtime de tres dígitos `<NNN>_v1.json`;
perfiles públicos descriptivos como `658_acs_v1.json` y nuevas variantes no
coinciden con el patrón.

## Verificación

- Focales: 23 pruebas correctas.
- Suite completa: 171 pruebas correctas.
- Compilación: `python -m compileall -q core tests` correcta.
- `git diff --check`: correcto.

## Preocupaciones

La creación exclusiva usa enlaces duros entre temporal y destino en el mismo
directorio. En un sistema de archivos que no admita enlaces duros, el alta
abortará sin sustituir ni publicar parcialmente ningún destino.
