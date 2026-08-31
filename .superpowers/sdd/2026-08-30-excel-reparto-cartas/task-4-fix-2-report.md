# Informe — corrección 2 de la Tarea 4

## Cambio

La huella canónica de entradas se forma ahora a partir de los módulos y conceptos activos del perfil. Incluye lecturas ACS cuando el perfil usa ACS y lecturas CALEFACCION cuando usa calefacción, junto con sus aprobaciones de estimación.

## Prueba

Se añadió un perfil sintético de calefacción y se valida que modificar una lectura tras el Excel validado bloquea el reparto sin alterar `owner_concept_results` ni la proyección heredada `repartos`.

## Verificación

- Focalizado de reparto: 11 pruebas correctas.
- Suite completa: 98 pruebas correctas.
- Compilación: `excel_export_service.py`, `case_distribution.py` y la prueba correctos.
- `git diff --check`: correcto (solo aviso LF/CRLF no bloqueante).

## Continuidad

El agente que añadió los cambios agotó su cuota antes de informe/commit. El coordinador inspeccionó el diff y ejecutó la verificación completa antes de este cierre.
