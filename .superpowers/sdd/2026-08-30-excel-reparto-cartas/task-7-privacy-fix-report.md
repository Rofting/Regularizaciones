# Corrección de privacidad de Tarea 7

## Cambios

- La ejecución aislada ya no copia `config/` completo: sólo incorpora el
  perfil público `config/excel_profiles/658_acs_v1.json` y la plantilla Word
  pública. Identidades de despacho, proveedores y archivos locales no se
  copian implícitamente.
- La raíz de validación se rechaza antes de escribir si es una unidad, la
  carpeta de usuario o su padre inmediato, además de las guardas anteriores
  para proyecto y fuentes.
- La guía indica que se debe crear una subcarpeta acotada bajo Temp o una zona
  privada, nunca una raíz amplia.

## TDD y verificación

- RED: las nuevas pruebas demostraron que se copiaban ficheros extra y que la
  carpeta de usuario podía recibir una ejecución.
- GREEN focal: 6 pruebas de `test_private_658_validation` correctas.
- Suite completa: 129 pruebas correctas.
- `py_compile core/private_658_validation.py` correcto.
- `git diff --check` sin errores.

No se ejecutaron fuentes reales ni se lanzó la interfaz.
