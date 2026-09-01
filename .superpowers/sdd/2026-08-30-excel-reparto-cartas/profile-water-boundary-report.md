# Informe — límite de la tabla AGUA del perfil 658

## Alcance

Se ajustó exclusivamente el perfil declarativo `658_acs_v1`: el límite final
de la tabla `AGUA` pasa de la fila 21 a la 20. No se usaron datos privados ni
se modificó lógica de negocio o perfiles de otras comunidades.

## Evidencia RED

Se añadió primero la prueba focal:

```text
tests/test_excel_profiles.py::ExcelProfileTest.test_658_agua_invoice_table_ends_at_row_20
```

Comando ejecutado antes de cambiar el perfil:

```text
<python> -m unittest tests.test_excel_profiles.ExcelProfileTest.test_658_agua_invoice_table_ends_at_row_20
```

Resultado observado:

```text
FAIL
AssertionError: 20 != 21
Ran 1 test in 0.001s
FAILED (failures=1)
```

El fallo confirma que la prueba detectaba el valor incorrecto existente.

## Cambio mínimo

En `config/excel_profiles/658_acs_v1.json`, dentro de `workbook_layout.tables.AGUA`,
se cambió únicamente `end_row` de `21` a `20`.

## Evidencia GREEN

Comando focal ejecutado después del cambio:

```text
<python> -m unittest tests.test_excel_profiles.ExcelProfileTest.test_658_agua_invoice_table_ends_at_row_20
```

Resultado:

```text
.
Ran 1 test in 0.001s
OK
```

Compilación:

```text
<python> -m compileall -q core tests
```

Resultado: código de salida 0.

Comprobación de espacios:

```text
git diff --check
```

Resultado: código de salida 0 (solo avisos de conversión de finales de línea
LF/CRLF de Git).

## Suite completa

### Verificación final con el entorno del proyecto

Comando exacto ejecutado desde el worktree, con UTF-8:

```text
PYTHONUTF8=1 <project-root>\\.venv-fase1\\Scripts\\python.exe -m unittest discover -s tests -t . -q
```

Resultado:

```text
Ran 143 tests in 20.032s
OK
```

Código de salida: 0.

### Ejecución anterior con runtime incompleto

Comando:

```text
<python> -m unittest discover -s tests
```

Resultado: `Ran 103 tests`; 24 errores y 1 fallo por dependencias no
disponibles en el intérprete (`matplotlib`, `xlrd`, `xlwt`, `customtkinter`) y
el fallo de generación de gráficos asociado. La prueba focal del perfil sí
queda verde; estos bloqueos son ajenos al cambio.
