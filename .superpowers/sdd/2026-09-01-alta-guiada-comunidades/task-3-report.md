# Task 3 report — Confirmación atómica y expediente inicial

## Estado

Implementados `OnboardingResult` y `confirm_onboarding(...)`. La confirmación
exige fechas completas o ninguna, verifica que las fuentes no hayan cambiado
desde el análisis, genera y valida perfil y plantilla antes de persistir la
comunidad, y crea período/expediente mediante los servicios existentes.

El JSON runtime se instala en `config/excel_profiles/runtime/`, se escribe
primero en un temporal hermano, se vuelve a leer y se
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

Si la fuente cambia después de la verificación inicial, cada registro toma una
instantánea de IDs del expediente. Un error posterior recupera las filas nuevas
por `id_case` y diferencia de IDs, por lo que limpia la ruta y el SHA realmente
registrados sin depender del SHA anterior del borrador.

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
- RED de fuente intercalada: cambiar los bytes entre la verificación y
  `register_source_document` dejaba el archivo del SHA nuevo; ahora no quedan
  filas ni archivos propios.
- RED de perfiles runtime generales: el patrón Git sólo cubría tres dígitos;
  el subdirectorio runtime admite cualquier código validado y mantiene la
  carga, resolución del expediente y selección del exportador.
- GREEN final focal ampliado: 60 pruebas de los consumidores modificados.

## Privacidad y alcance

Las pruebas crean únicamente comunidad 900 y fuentes sintéticas temporales.
No se han leído fuentes reales ni modificado el perfil o flujo de 658. El
`.gitignore` ignora únicamente `config/excel_profiles/runtime/`, de modo que
perfiles runtime de códigos como `1234` o `ABC` quedan locales y cualquier
perfil público colocado en `config/excel_profiles/` sigue visible para Git.
`load_profile` prioriza el perfil público y usa runtime como fallback, sin
cambiar la clave `<codigo>_v1`.

## Verificación

- Focales ampliadas: 60 pruebas correctas.
- Suite completa: 172 pruebas correctas.
- Compilación: `python -m compileall -q core tests` correcta.
- `git diff --check`: correcto.

## Preocupaciones

La creación exclusiva usa enlaces duros entre temporal y destino en el mismo
directorio. En un sistema de archivos que no admita enlaces duros, el alta
abortará sin sustituir ni publicar parcialmente ningún destino.
