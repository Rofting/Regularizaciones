# Tarea 7 — validación privada y guía

## Entregado

- Runner explícito `core/private_658_validation.py`, fuera de la interfaz y de
  la suite ordinaria.
- Guardas para las cuatro variables obligatorias, fuentes que deben ser
  archivos y raíz de validación aislada, nueva/vacía y fuera de las fuentes y
  del proyecto.
- Copia de ejecución limitada a la configuración y plantilla Word públicas;
  base, archivo de fuentes, plantilla privada, salidas y render quedan dentro
  de una subcarpeta única de validación.
- Informe saneado: estado, conteos por tipo y ruta de salida, sin rutas de
  fuentes, nombres de propietarios ni errores individualizados.
- Flujo preparado para importar, bloquear ante incidencias abiertas, generar
  Excel/reparto/cartas cuando esté limpio y renderizar una sola página con
  LibreOffice. No se ejecutó ninguna fuente real ni se lanzó Tk.
- Guía v3 y documento específico de ejecución y revisión manual.

## TDD y verificación

- RED confirmado: las guardas fallaban porque el runner no existía.
- Verde: `tests.test_private_658_validation` (5 pruebas).
- Suite completa: 128 pruebas correctas.
- `py_compile core/private_658_validation.py` correcto.
- Dependencia de render comprobada: `pypdfium2` disponible.
- `git diff --check` sin errores.

## Próximo paso controlado

El operador define las cuatro variables en una consola privada y ejecuta el
runner. Si devuelve incidencias abiertas, se corrigen en la app antes de volver
a validar; si termina correctamente, se revisan visualmente el Excel y la
carta renderizada en Windows. No se envía correo automáticamente.
