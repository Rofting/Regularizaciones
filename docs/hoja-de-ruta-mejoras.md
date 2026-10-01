# Hoja de ruta de las 20 mejoras

Fecha: 1 de octubre de 2026. Revisión de partida: `a60926d`.

Este repositorio es la instalación personal del usuario. Los CUPS y la
configuración de sus comunidades son reales y se conservan. La aceptación
automática usa documentos y bases temporales sintéticos; la aceptación con
documentos reales se ejecuta sobre una copia verificada de la base.

El alcance de esta entrega es corregir el clon limpio, añadir CI y preparar
esta hoja de ruta. Las demás filas describen trabajo pendiente; no representan
funciones ya entregadas.

Actualización de la siguiente entrega: se implementan las mejoras 19 y 8,
documentadas en [Copias automáticas y gastos heredados](copias-y-gastos-heredados.md).
La mejora 10 también está implementada, documentada en
[Coherencia antes del Excel](coherencia-antes-del-excel.md).
La mejora 9 está implementada en
[Comparación con el ejercicio anterior](comparacion-con-ejercicio-anterior.md).
La mejora 12 tiene su prueba de aceptación en `tests/test_libreoffice_journey.py`
(ver la fila 12); queda pendiente comprobar su resultado en Windows en Actions
y fijar la versión de LibreOffice de referencia.

### Revisión de 9 y 10 (entrega B)

- Al aplicar una factura, `cups`, `consumo_kwh`/`consumo_m3` y el proveedor
  detectado no llegaban a `facturas` (`cups_o_referencia`, `consumo_total`,
  `unidad_consumo`, `proveedor`). Sin esos datos la comparación histórica no
  comparaba ningún importe y la coherencia trataba como posible solape dos
  suministros distintos del mismo tipo. Corregido en `case_ingestion`; un valor
  confirmado con el nombre canónico prevalece. Las facturas ya aplicadas antes
  de la corrección conservan sus columnas vacías hasta que se vuelvan a aplicar.
- Una factura de un solo día (servicio puntual) era un bloqueo sin aceptación
  posible; ahora es un aviso aceptable. Inicio posterior al fin sigue bloqueando.
- En comunidades de más de 10 viviendas las notas sin novedad de la comparación
  histórica se resumen en una línea para no tapar los avisos.

## Primera entrega: pruebas que se pueden repetir

- **11, corregido:** `tests/test_case_workflow_actions.py` construye su proyecto
  y su plantilla sintética en el directorio temporal de cada prueba. Conserva
  las comprobaciones reales de preparación, transición de estados,
  revalidación, repetición de etapas e invalidación de salidas posteriores.
  No lee ni escribe el maestro privado de la 658.
- **20, CI añadido:** `.github/workflows/tests.yml` instala todas las
  dependencias de `requirements.txt` y ejecuta la suite en Ubuntu 24.04 y
  Windows Server 2022 con Python 3.12 en cada push y pull request. Instala
  LibreOffice, comprueba que esté disponible y registra su versión. Por tanto,
  la prueba de exportación real no queda omitida por faltar el programa.
  Los registros se conservan 14 días como artefactos del workflow.
  En Windows se instala LibreOffice antes de preparar Python: el workflow
  comprueba que las pruebas usan Python 3.12 del runner y no el Python interno
  que LibreOffice trae en su carpeta de programa.
- `tests/test_dependencies.py` comprueba también `rapidfuzz`, ya declarado en
  los requisitos. Evita confundir una instalación incompleta con un fallo del
  reconocimiento aproximado de proveedores.
- **Corrección de Windows detectada por CI:** la restauración del diseño del
  Excel cierra los archivos ZIP antes de reemplazar el libro recalculado.
  Evita el `PermissionError` que Windows producía al mantener abierto el
  destino. Una prueba conserva la fórmula y su resultado y comprueba el cierre
  de los archivos. Las pruebas afectadas también normalizan su directorio
  temporal para comparar correctamente rutas cortas y largas de Windows.

La versión de LibreOffice instalada por los gestores de paquetes de CI puede
cambiar: Ubuntu instala la rama de su distribución (24.2.7.2) y en Windows se
usaba la más reciente (26.2.6.3). El recorrido pasó con ambas, así que se fija
una versión mínima (24.2) en lugar de una exacta, que dependería de que el
instalador antiguo siguiera publicado. Windows instala la rama estable
(`libreoffice-still`) y el paso «Comprobar LibreOffice real» ejecuta
`scripts/comprobar_libreoffice.py`, que falla con una versión anterior.

## Orden de ejecución

| Entrega | Mejoras | Motivo y resultado |
| --- | --- | --- |
| A. Base reproducible | 11 y CI de 20 | Detectar regresiones en un clon limpio. Es el alcance de esta entrega. |
| B. Protección del cálculo | 19, 8, 10, 9, 12 | Crear copias verificadas, detectar importes heredados y comprobar el cálculo antes de ampliar automatismos. |
| C. Entrada de documentos | 6, 5, 2, 3, 4, 1, 7 | Aprender identidades confirmadas, repartir por comunidad y mejorar extracción conservando trazabilidad. |
| D. Trabajo diario | 14, 15, 13 | Revisar documentos dentro de la aplicación, ver todos los expedientes y completar altas en lote. |
| E. Salidas y distribución | 17, 16, 18 | Obtener PDF, preparar/enviar cartas con control por destinatario y construir el instalador Windows. |
| F. Mantenimiento | Resto de 20 | Retirar código sin consumidores y limpiar Git con una copia recuperable. |

La entrega B precede a la automatización de altas y envíos. El PDF de cartas
precede al envío por correo. El instalador requiere resolver primero las rutas
de recursos, datos, copias y herramientas externas.

## Extracción automática de datos

| Nº | Estado comprobado y trabajo pendiente | Criterio de aceptación |
| --- | --- | --- |
| 1. Aprender formatos de lecturas | `reading_tables.py`, `keywords.py` y `dwelling_matching.py` ya reconocen columnas y viviendas por su contenido. Falta guardar un mapeo confirmado por empresa y firma de cabeceras, con su versión de extractor. Integrarlo en análisis y confirmación de fuentes. | Confirmar un informe enseña el formato; otro de la misma empresa y otra comunidad reutiliza el mapeo. Un cambio de columnas o un dato incoherente abre revisión y conserva el motivo. |
| 2. Preprocesar escaneos | RapidOCR tiene respaldo de Tesseract. La portada general se renderiza a 200 dpi y hay relecturas específicas; falta una preparación general de imagen. Añadir enderezado, reducción de ruido y contraste, manteniendo el original y evitando empeorar imágenes ya legibles. | Un conjunto de escaneos torcidos, tenues y con ruido mejora la extracción de fechas/importes; los PDF digitales siguen por su camino habitual. Registrar el tratamiento y comparar resultados en documentos reales. |
| 3. Otra estrategia de tablas PDF | `reading_tables.py` ya usa `pdfplumber.extract_tables()` y lectura de texto. Añadir estrategias explícitas de líneas y texto cuando la primera tabla no produce filas válidas, sin duplicar resultados. | Tablas con líneas y sin líneas producen viviendas, fechas y consumos correctos. Se detectan filas de totales y se comprueba el cuadre sin contarlas como otra vivienda. |
| 4. OCR de páginas posteriores | La extracción general intenta OCR de portada; algunos proveedores pueden solicitar páginas concretas. Falta continuar por los campos obligatorios ausentes. Añadir lectura por página con un límite de tiempo/páginas y caché de resultados. | Facturas con consumo o desglose en páginas 3 y 4 quedan completas. Se detiene cuando hay evidencia suficiente o se alcanza el límite; una portada completa no obliga a procesar todos los anexos. |
| 5. Bandeja global por contenido | La bandeja actual agrupa principalmente por código en nombre/ruta. Ampliar `community_discovery.py` para resolver por CIF de comunidad, CUPS aprendido y código confirmado, usando el texto cacheado. | Una carpeta mixta se reparte entre comunidades aunque los archivos no tengan código. Evidencias contradictorias o insuficientes quedan sin asignar y se pueden resolver; cada asignación conserva su evidencia. |
| 6. Aprender CUPS de comunidades | Existe mapeo manual en `config/proveedores_despacho.json` y aprendizaje de CIF de proveedores en SQLite. Falta persistir en la base el vínculo CUPS/comunidad/suministro al confirmar facturas. | Una factura confirmada enseña el CUPS; la siguiente se identifica sin depender del nombre del archivo. Un CUPS asociado a otra comunidad bloquea su reasignación automática y muestra ambas evidencias. |
| 7. Adjuntos `.eml` y `.msg` | Las fuentes admitidas actualmente son PDF, Excel y CSV. Añadir extracción de adjuntos con `email` para EML y un lector de MSG, y entregar los adjuntos al mismo proceso de archivado/análisis. | Se archivan el correo y sus adjuntos con trazabilidad. Adjuntos repetidos no duplican facturas y un correo sin adjuntos compatibles se informa sin interrumpir el resto del lote. |

## Fiabilidad del cálculo

| Nº | Estado comprobado y trabajo pendiente | Criterio de aceptación |
| --- | --- | --- |
| 8. Plantillas antiguas con gastos heredados | Implementado: se revisan B6:B9 y C6:C9 de las filas conocidas antes del Excel. La interfaz ofrece «Revisar gastos fijos». Guardar vacío confirma cero y reconstruye el anual mensual × 12 en el libro generado. | Las pruebas comprueban el bloqueo, la confirmación y la sustitución de un total anual heredado, también con LibreOffice real. La plantilla original no se altera. `993,44 €` es un indicio que revisar, nunca una razón para borrar un gasto legítimo automáticamente. |
| 9. Comparación con el ejercicio previo | Implementado en «Coherencia»: último período anterior de la misma comunidad, importes por suministro/punto/unidad e intervalo y consumos por vivienda, con factor ×3 configurable y tolerancia de duración. Avisos motivados y huellas que incluyen el histórico. Los futuros repartos conservan nombre del titular y código de vivienda. | Pruebas de subidas/bajadas, ceros, referencias ausentes o ambiguas, rectificaciones, cambios de titular e invalidación al editar el histórico. Sin referencia comparable se informa y no se prorratea. Los repartos antiguos sin identidad histórica se señalan como tales; no se inventa el titular anterior. |
| 10. Coherencia antes del Excel | Implementado: panel «Coherencia», control de porcentajes/pesos relativos, fechas y solapes por suministro/punto, y comparaciones de consumos declaradas por el usuario. Bloqueo en preparación y exportación directa, con aceptación motivada de avisos y revisión invalidada cuando cambian los datos. | Pruebas sintéticas de bloqueos, corrección desde la interfaz, abonos, unidades/intervalos incompatibles, lecturas estimadas y aceptación caducada. Los porcentajes deben sumar 100 % con tolerancia 0,01; los pesos se normalizan. Consumo sólo con unidades iguales, un punto identificado e intervalo exacto y continuo; no se prorratea. |
| 11. Pruebas independientes de la 658 | Corregido en esta entrega: proyecto y workbook temporales sintéticos en las pruebas del flujo guiado. | La suite pasa desde un checkout que nunca ha tenido `plantillas/comunidades/658`. Las pruebas no dejan archivos privados ni salidas en el árbol de trabajo. |
| 12. Recorrido real con LibreOffice | `tests/test_libreoffice_journey.py`: comunidad nueva desde el modelo común, dos propietarios, facturas y lecturas confirmadas, cuotas, Excel recalculado por LibreOffice, reparto y cartas. El ensayo destapó dos fallos ya corregidos: la validación posterior a LibreOffice exigía área de impresión en todas las hojas (el modelo sólo la tiene en ANALISIS) y la preparación pedía los importes de ACS como parámetros aunque hubiera cuotas, cuando el exportador los calcula. Verificado en CI (run 36865907823, commit 4b4a589): ubuntu-24.04 con LibreOffice 24.2.7.2 y windows-2022 con LibreOffice 26.2.6.3; en ambos el recorrido se ejecutó (no se omitió) y pasaron las 691 pruebas. Versión mínima fijada en 24.2 (`MINIMUM_LIBREOFFICE_VERSION` en `core/office_recalculation.py`); `scripts/comprobar_libreoffice.py` la exige en CI y sirve para revisar una instalación nueva. Windows instala ahora la rama estable (`libreoffice-still`). | Con dos propietarios y documentos sintéticos, el mismo recorrido conserva impresión/diseño, cuadra en céntimos y genera las cartas esperadas. Versiones registradas y mínima exigida en CI. Una prueba con un motor simulado no sustituye a esta aceptación. |

## Uso diario e interfaz

| Nº | Estado comprobado y trabajo pendiente | Criterio de aceptación |
| --- | --- | --- |
| 13. Alta masiva | Implementado: `core/community_batch.py` analiza cada subcarpeta con la detección desde carpeta (código y nombre también de «658 - CP Las Flores»), normaliza propietarios y lecturas de cualquier empresa al formato del alta guiada, responde sólo lo inequívoco (servicio y una lectura de ejemplo que se muestra para aprobar) y aparta las facturas dudosas a la bandeja. «Nueva comunidad» → «Alta masiva por subcarpetas» muestra el resumen aprobable y crea comunidad, perfil, expediente, propietarios y lecturas. | Probado en `tests/test_community_batch.py`: dos comunidades (ACS y calefacción) quedan con propietarios, lecturas y expediente; repetir el lote no duplica; incompletas y códigos repetidos esperan revisión con su motivo; un fallo se compensa sin dejar restos y las demás continúan. La prueba destapó que dos subcarpetas con archivos del mismo nombre se pisaban al convertirse: cada comunidad convierte ahora en su carpeta. |
| 14. Vista previa de incidencias | Implementado: `core/source_preview.py` reúne valor, fragmento, página y celda guardados y busca el texto en la capa de texto del PDF (también las formas impresas del valor: `31/01/2026`, `1.234,56`). «Ver en la fuente», en el diálogo de cada incidencia, abre la página con el dato resaltado, navegación entre páginas e «Ir al resaltado»; en Excel/CSV muestra el entorno de la celda marcada. | Probado en `tests/test_source_preview.py`. Se resalta sólo lo localizado: el valor dentro de su fragmento, el fragmento si el valor no aparece en él, o un valor distintivo que aparece una única vez. Un valor repetido («130») sin fragmento, un PDF escaneado o un texto que no está en el archivo no se resaltan: se muestra el fragmento tal cual. |
| 15. Panel de todas las comunidades | La pantalla principal trabaja con la comunidad activa. Añadir una consulta global de expedientes con estado, incidencias y última salida, filtros y apertura del caso elegido. | El panel muestra pendientes, bloqueados, calculados y cartas listas, se actualiza tras una acción y abre el expediente correcto sin cambiar de período por error. |
| 16. Correo de cartas | Existe un plan detallado en `docs/superpowers/plans/2026-09-29-correo-cartas.md`; no hay servicio de envío implementado. Empezar por borradores EML auditables, después transporte SMTP con credenciales del sistema y confirmación visible del lote. | Falta de email, adjunto cambiado, error y envío exitoso quedan registrados por propietario. Reintentar no reenvía los éxitos. Las pruebas usan un transporte simulado y no envían correos reales. |
| 17. Cartas en PDF | El servicio actual publica Word y registra su identidad. Añadir conversión con LibreOffice Writer y vincular DOCX/PDF al mismo propietario y lote. | Dos propietarios producen dos DOCX y dos PDF con sus importes correctos. Se conserva el diseño, se comprueba el número de páginas y un fallo de conversión impide presentar ese PDF como listo. |

## Mantenimiento e instalación

| Nº | Estado comprobado y trabajo pendiente | Criterio de aceptación |
| --- | --- | --- |
| 18. Instalador Windows | Hay plan portable y `scripts/exportar_producto.py`, pero no un ejecutable construido. Resolver rutas de datos escribibles, empaquetar con PyInstaller, modelos OCR y herramientas documentales, y generar el instalador en Windows. | Instalar y arrancar en Windows sin Python ni descargas en tiempo de ejecución. Probar rutas con espacios/acentos, primera base vacía, importación y cartas. Actualizar conserva los datos y permite recuperar una copia. |
| 19. Copias automáticas | Implementado en `database_backup.py`: arranque, migraciones con datos y generación de Excel/reparto/cartas del flujo guiado, además de reevaluación de fuentes desde la interfaz. Retención configurable, 20 por defecto, manifiesto y registro. | Las pruebas recuperan datos confirmados en WAL, comprueban integridad y restauración, y bloquean migraciones/regeneraciones si falla la copia. La limpieza conserva como mínimo 3 copias verificadas y excluye copias manuales, alteradas y de otras bases. |
| 20. Orden del código y Git | CI se añade en esta entrega. `excel_generator.py` aún importa `excel_writer.py` y existen consumidores de importadores antiguos: inventariar llamadas antes de retirarlos. La limpieza de ramas e historial queda para una entrega específica. | Retirar una pieza sólo tras migrar sus consumidores y pasar la suite. Para quitar la base histórica: guardar un bundle y una copia verificada fuera del repositorio, reescribir la ruta de la BD en todas las referencias que se vayan a publicar, comprobar su ausencia y publicar con control del estado remoto. La rama `feature/global-provider-detection` se elimina tras comprobar que sus cambios están integrados. Los clones de Codex/Claude se renuevan después de reescribir el historial. |

## Cómo repetir la primera entrega

Desde un clon nuevo, en la raíz del repositorio:

```bash
python -m venv .venv
# Linux/macOS:
.venv/bin/python -m pip install -r requirements.txt
.venv/bin/python -m unittest discover -s tests -v
# Windows, equivalente:
# .venv\Scripts\python.exe -m pip install -r requirements.txt
# .venv\Scripts\python.exe -m unittest discover -s tests -v
```

Para ejecutar también la exportación real hay que instalar LibreOffice y
poner `soffice` en el PATH. Actions hace esa instalación y comprueba el
ejecutable antes de ejecutar la suite. Sin LibreOffice, la comprobación
específica de ese motor se omite en una ejecución local.

Cada entrega futura debe incluir pruebas del comportamiento nuevo, una
aceptación con documentos reales sobre una copia de la base y un resultado
publicado verificable. Ninguna entrega se considera completa por el número
de archivos añadidos o por disponer sólo de un plan escrito.
