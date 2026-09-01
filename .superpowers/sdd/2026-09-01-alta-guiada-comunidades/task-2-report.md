# Task 2 report — Perfil confirmado y validación reutilizable

## Estado

Implementados `build_profile_payload(draft, answers)` y
`validate_profile_payload(payload, project_root)`. Ambos operan estrictamente
en memoria: no escriben JSON ni plantillas ni publican perfiles.

La construcción exige responder una columna de lectura obligatoria y sólo
incluye en `active_modules` y `concepts` los módulos con respuesta explícita
`True`; `False` y las respuestas ausentes no aplican. Genera una clave
`<community_code>_v1`, versión `"1"` y una ruta de plantilla relativa.

La validación reutilizable es ahora el núcleo empleado también por
`load_profile`, por lo que conserva las validaciones de método de reparto,
hojas requeridas por módulo, referencias A1 y ruta segura.

## TDD

- RED: las focales fallaron primero con `AttributeError` porque no existían
  `build_profile_payload` ni `validate_profile_payload`.
- GREEN: las focales pasan después de implementar las APIs y sus pruebas.

## Verificación

- Focales: `python -m unittest tests.test_community_onboarding tests.test_excel_profiles -q` — 16 pruebas correctas.
- Suite completa: `python -m unittest discover -s tests -q` — correcta.
- Compilación: `python -m py_compile core/community_onboarding.py core/excel_profiles.py` — correcta.
- `git diff --check` — correcto.

## Privacidad y alcance

Las pruebas usan únicamente borradores y perfiles sintéticos temporales. No se
han leído fuentes reales ni creado archivos de perfil o plantilla.

## Corrección de revisión 1

- Toda pregunta obligatoria exige una respuesta de texto no vacío; si declara
  candidatos, la respuesta debe ser uno de ellos. Las respuestas vacías,
  booleanas y ajenas se rechazan, incluyendo la columna de lectura.
- El código de comunidad se valida como componente seguro de Windows antes de
  derivar clave y ruta: rechaza separadores, caracteres reservados, finales con
  punto o espacio y dispositivos reservados. Los códigos normales siguen
  generando claves como `644_v1` y `900_v1`.
