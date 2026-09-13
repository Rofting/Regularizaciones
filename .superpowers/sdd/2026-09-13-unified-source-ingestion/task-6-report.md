# Task 6 — Verificación integral y manual

## Resultado

Se añadió una regresión end-to-end de un lote mixto en
`tests/test_expedient_flow.py`. Registra una factura completa, dos lecturas y
una fuente desconocida en la misma base SQLite temporal. Comprueba el resumen
persistido `{"invoice": 1, "reading": 2, "unknown": 1}`, una sola incidencia
abierta y que esa incidencia pertenece únicamente a la fuente desconocida con
el código `DOCUMENT_CLASSIFICATION_REQUIRED`. Así detecta la regresión en la
que las lecturas heredasen los tres campos de revisión de una factura.

`docs/validacion-expedientes.md` incluye el recorrido para crear o seleccionar
una comunidad, abrir el expediente, añadir una carpeta mixta, interpretar el
resumen clasificado, consultar contexto y resolver incidencias, reanalizar las
copias archivadas y producir Excel, reparto y cartas. También documenta la
aceptación con una base limpia: cerrar cualquier sesión de la app antes de
reiniciar, usar exclusivamente **Nueva base segura** al volver a abrirla y no
manipular `data/gestion.db` desde el explorador. El proceso verifica una copia
en `backups` antes de reemplazar la base; nunca elimina la base anterior de
forma silenciosa.

## TDD y regresión

La prueba focal se añadió antes de modificar código de producción. Su primera
ejecución detectó que el contrato de la incidencia de clasificación usa el
campo `document_kind`; se corrigió la expectativa de la prueba y pasó contra la
implementación ya disponible de las tareas anteriores.

Para comprobar que la prueba protege el comportamiento real, se aplicó una
mutación temporal y reversible de `_required_fields_for_kind` que hacía que
una lectura recibiera requisitos de factura. La prueba falló con `1 != 7`
incidencias abiertas. Tras restaurar el código original, la misma prueba pasó.
No queda ningún cambio en `core/case_ingestion.py`.

```text
..\\..\\.venv-fase1\\Scripts\\python.exe -m unittest \
  tests.test_expedient_flow.ExpedientFlowTest.test_mixed_classified_sources_create_only_the_unknown_review -v

Ran 1 test — OK
```

## Verificación final

```text
..\\..\\.venv-fase1\\Scripts\\python.exe -m unittest discover -s tests -q
Exit code 0 (sin salida con -q)

Suite descubierta: 277 pruebas

git diff --check
Exit code 0
```

La suite usa SQLite temporal para este flujo. No se abrió, reinició ni modificó
la base real `data/gestion.db`.

## Observaciones

- No se realizó una aceptación visual de la interfaz Tk ni un reinicio contra
  datos reales: ambas acciones quedan fuera de la regresión y evitan poner en
  riesgo datos operativos.
- El botón **Nueva base segura** requiere que la app esté en ejecución para
  mostrar su confirmación; la guía especifica cerrar cualquier sesión previa y
  no manipular la base directamente antes de volver a abrir la app y usar el
  botón.
