# Tarea 6 — acciones guiadas y estado del expediente

## Entregado

- Se añadió `core/case_workflow_actions.py`, un borde sin Tk que verifica en
  cada operación que el expediente pertenece a la comunidad seleccionada y
  resuelve el perfil Excel activo por código comunitario.
- La pantalla v3 mantiene su estética y su carril de cinco etapas. El carril
  `EL FLUJO` ahora lleva, en orden, desde crear expediente e importar el
  modelo inicial hasta Excel oficial, reparto final y cartas.
- El bootstrap pide el libro maestro y, de forma opcional, el par de fuentes
  complementarias. Un bootstrap limpio valida el expediente para habilitar la
  exportación; si faltan datos, conserva las incidencias para su revisión.
- Las acciones de Excel, reparto y cartas usan exclusivamente los servicios
  por expediente. El reparto cambia el estado a conciliado y un lote completo
  de cartas a entregas generadas.
- Las cartas preguntan los conceptos antes de empezar, pero sólo muestra los
  conceptos calculados y activos del perfil. `case_letter_service` rechaza una
  selección fuera de ese conjunto, por lo que no puede forzarse calefacción
  inactiva ni otra partida ajena al expediente.
- La actividad traduce hitos de cálculo, escritura, validación y publicación
  a frases operativas; el indicador de progreso existente se conserva y los
  selectores quedan bloqueados durante una operación.

## TDD y verificación

- RED observado: primero faltaba `case_workflow_actions`; después la API de
  cartas no aceptaba `selected_concepts`; y se observó el estado sin cambiar
  tras un reparto correcto.
- Pruebas nuevas sin Tk: `tests/test_case_workflow_actions.py`; se amplió
  `tests/test_case_letter_service.py` para la selección limitada.
- Focales: 28 pruebas de controladores, cartas y reparto, correctas.
- Suite completa: `python -m unittest discover -s tests -t . -q` → **116 OK**.
- Compilación: `py_compile core/app.py core/case_workflow_actions.py
  core/case_letter_service.py` → correcta.
- `git diff --check` → correcto.

## Humo Windows para la entrega final

1. Seleccionar una comunidad, un período y su expediente.
2. Ejecutar **Importar modelo inicial** con una fuente de prueba y comprobar
   que una incidencia permite abrir el archivo archivado y resolver el campo.
3. Con todo validado, comprobar el progreso de **Generar Excel oficial** y
   abrir sólo el archivo publicado al finalizar.
4. Ejecutar **Calcular reparto final** y comprobar que el estado pasa a
   conciliado.
5. Ejecutar **Generar cartas**, seleccionar partidas activas, y comprobar que
   la carpeta ofrecida corresponde al lote terminado.
