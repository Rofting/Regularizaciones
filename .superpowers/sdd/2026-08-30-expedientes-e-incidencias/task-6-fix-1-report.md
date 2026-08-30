# Informe — Tarea 6, corrección visual 1

## Cambios

- Añadida `height=1` al carril `carril_estado_expediente`.
- Añadida `height=1` al marcador `marker` de cada incidencia.
- Se conserva `fill="y"`, sin cambios de paleta, tipografía, textos ni comportamiento.

## Verificaciones

- `py_compile core\\app.py`: OK.
- `git diff --check`: OK.
- Comprobación AST: OK; `carril_estado_expediente` y `marker` tienen `height=1` explícito.

## Limitación

No se lanzó Tk headless. La comprobación visual final requiere revisión manual en Windows.

SHA del commit: se comunica al finalizar el commit.
