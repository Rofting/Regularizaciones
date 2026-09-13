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

## Addendum de revisión

La revisión detectó dos rutas de seguridad no cubiertas inicialmente. Se
añadieron primero sus regresiones y la ejecución dirigida falló como se
esperaba: una colisión forzada reutilizaba y sobrescribía el respaldo ya
existente, y un fallo de `mkstemp` escapaba como `PermissionError`.

La reserva de respaldos ahora crea el nombre candidato en modo exclusivo
(`"xb"`) y vuelve a generar un candidato al encontrar uno existente, antes de
que SQLite abra el fichero de destino. La asignación del fichero de reemplazo
se incluye ahora en el bloque que convierte errores a `DatabaseResetError`, y
la limpieza sólo se intenta cuando se llegó a asignar una ruta temporal.

Las dos regresiones comprueban respectivamente que un respaldo irremplazable
conserva su contenido tras una colisión forzada, y que un fallo de asignación
temporal devuelve `DatabaseResetError`, conserva un respaldo válido y no cambia
la base original. La verificación dirigida final ejecutó seis pruebas con éxito.
