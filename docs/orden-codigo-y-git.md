# Inventario para la mejora 20

La interfaz y el flujo guiado usan `case_workflow_actions` y
`excel_export_service` para el Excel oficial. El flujo por línea de comandos
`pipeline.py` sigue usando la regeneración por comunidad de
`excel_generator.py`, que llama a `excel_writer.py` cuando no recibe
`id_case`. Por eso el escritor histórico se carga sólo en esa ruta; retirarlo
exige migrar antes `pipeline.py` y su salida para todos los períodos.

| Módulo heredado | Consumidor vigente | Decisión |
| --- | --- | --- |
| `importar_comunidad.py` | Ingesta de `pipeline.py` | Conservar hasta migrar el pipeline. |
| `importar_excel_referencia.py` | `excel_bootstrap_importer.py` y `regularization_flow.py` | Conservar. |
| `importar_lecturas_metrigest.py` | `case_ingestion.py` y `pipeline.py` | Conservar. |
| `importar_lecturas_xls.py` | `source_analysis.py` y `regularization_flow.py` | Conservar. |
| `importar_propietarios_csv.py` | `source_analysis.py` y `regularization_flow.py` | Conservar. |
| `importar_excel_maestro.py` | Entrada directa por consola | Conservar como herramienta independiente; la interfaz no necesita cargarla al arrancar. |
| `importar_ledger_calderas.py` | Entrada directa por consola | Conservar hasta decidir si se sustituye su función. |
| `importar_listado_comunidades.py` | Entrada directa por consola | Conservar hasta decidir si se sustituye su función. |

## Limpieza de Git

Se comprobó cada rama remota y se eliminaron `feature/global-provider-detection`
y otras ocho ramas obsoletas. Se conservó la única rama alternativa con cambios
propios.

Se guardaron un bundle verificable de todos los refs y una copia íntegra de
`data/gestion.db` fuera del repositorio. Se reescribieron `main` y la rama
alternativa para retirar esa ruta del historial publicado de ambas. La CI
impide incorporar nuevos archivos SQLite al árbol actual. Los refs internos
`refs/pull/1/head` a `refs/pull/5/head` conservan commits antiguos: GitHub
Support debe purgarlos, junto con las vistas cacheadas y objetos afectados.
Los clones locales con ramas propias requieren una renovación cuidadosa tras
guardar su trabajo.
