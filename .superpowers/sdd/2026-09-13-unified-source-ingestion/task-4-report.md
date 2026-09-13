# Task 4 — Reinicio seguro de base de datos

## Cambios

- Se añadió `core/database_reset.py` con `reset_database`, `ResetResult` y
  `DatabaseResetError`.
- El respaldo se genera con `sqlite3.Connection.backup`, recibe un nombre con
  marca UTC y UUID, y se valida con `PRAGMA integrity_check` antes de crear el
  reemplazo.
- La nueva base se inicializa en un fichero temporal contiguo a la original,
  se valida, y sólo entonces se publica con `os.replace`.
- Los errores de copia, validación o inicialización se traducen a
  `DatabaseResetError`; el fichero original no se reemplaza en esas rutas.
- Se añadieron pruebas aisladas con `TemporaryDirectory`; no se usa ninguna
  base de datos real del proyecto.

## TDD y verificación

La primera ejecución de `..\\..\\.venv-fase1\\Scripts\\python.exe -m unittest
tests.test_database_reset -v` falló durante la carga con
`ModuleNotFoundError: No module named 'database_reset'`, confirmando el estado
rojo antes de implementar el módulo.

Tras implementar el módulo, una ejecución reveló bloqueos de fichero de
Windows: el gestor de contexto de `sqlite3.Connection` confirma o revierte la
transacción, pero no cierra la conexión. Se cambió el código de producción y
las utilidades de prueba para cerrar explícitamente las conexiones con
`contextlib.closing`.

La ejecución final del mismo comando completó cuatro pruebas correctamente:

- copia verificada y reemplazo correcto;
- preservación de la original si falla la verificación del respaldo;
- preservación de la original si falla la inicialización nueva;
- rutas de respaldo únicas.

También se ejecutó `git diff --check` sin errores de espacios.
