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

## Limpieza de Git pendiente

La rama remota `feature/global-provider-detection` apunta a un commit que ya
figura en la historia de `main`. Se puede eliminar tras sincronizar el remoto;
las demás ramas requieren la misma comprobación individual.

La base `data/gestion.db` está ausente del árbol actual, pero aparece en
commits antiguos. Reescribir el historial publicado cambiaría los SHA de
`main` y obligaría a renovar los clones. Antes de esa operación hay que
obtener una copia completa de los refs remotos, guardar un bundle y una copia
verificada de la base fuera del repositorio, revisar todas las referencias a
la ruta y acordar el momento de la migración. Este entorno no puede hacer
`git fetch` por el proxy, así que aún no puede crear ese bundle completo.
La CI impide incorporar nuevos archivos SQLite al árbol actual.
