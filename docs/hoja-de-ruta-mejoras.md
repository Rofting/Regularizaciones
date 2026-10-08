# Hoja de ruta de las 20 mejoras

Fecha: 1 de octubre de 2026. Revisión de partida: `a60926d`.

Este repositorio es la instalación personal del usuario. Los CUPS y la
configuración de sus comunidades son reales y se conservan. La aceptación
automática usa documentos y bases temporales sintéticos; la aceptación con
documentos reales se ejecuta sobre una copia verificada de la base.

Esta hoja conserva el plan inicial y registra las entregas posteriores en cada
fila. «Implementado» indica que existe código y aceptación sintética; la
aceptación con documentos reales se registra por separado.

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
| 1. Aprender formatos de lecturas | Implementado: `reading_formats.py` guarda el mapeo de columnas por empresa, firma de cabeceras y versión del extractor, con la persona y fuente que lo confirmó. El análisis de PDF, Excel y CSV consulta los formatos aprendidos. El primer formato necesita confirmación humana; la aplicación automática no enseña formatos. Se releen fechas y viviendas en cada informe. | `tests/test_unified_ingestion_regressions.py` confirma un CSV, reutiliza el formato en un PDF y aplica un Excel de otra comunidad y otro período. Un cambio de columnas o un solo consumo incoherente abre revisión con el motivo; la confirmación bloqueada no aprende. Versiones antiguas y mapeos dañados no se reutilizan. Pendiente: aceptación con informes reales sobre una copia de la base. Véase [Aprendizaje de lecturas](aprendizaje-formatos-lecturas.md). |
| 2. Preprocesar escaneos | Implementado: `core/scan_preprocessing.py` mide giro, ruido y contraste tinta/papel y sólo corrige lo que falla (enderezado ±5°, mediana, autocontraste) sobre una copia. El OCR de portada (RapidOCR y Tesseract) y el de páginas posteriores leen el original y, si la página tiene problemas, la versión preparada, y se quedan con la que da más fechas, importes y CIF. El tratamiento queda en `OCRResult.preprocessing` y en el diagnóstico de la caché de texto (versión `document-text-v4`, que relee los escaneos ya cacheados). `scripts/comparar_ocr.py` compara con y sin preparación en una carpeta de documentos reales. | Probado en `tests/test_scan_preprocessing.py`: una página limpia no se trata y se lee una vez; un escaneo torcido, tenue y con ruido no daba datos y tras la preparación se leen fechas e importes con RapidOCR y Poppler reales; la preparada sólo se usa si lee más; el original no se modifica. Pendiente: pasar el comparador por escaneos reales del despacho. |
| 3. Otra estrategia de tablas PDF | Implementado: `reading_tables.py` compara la extracción predeterminada de `pdfplumber` con estrategias explícitas de líneas y texto, y elige la que devuelve más filas de lectura válidas sin mezclar duplicados. Las filas de totales no se cuentan como viviendas; si incluyen consumo, se comprueba el cuadre y se reduce la confianza cuando falla. | `tests/test_reading_tables.py` genera PDF sintéticos con y sin cuadrícula y comprueba viviendas, lecturas y consumo. También verifica un total correcto y otro incoherente. Falta aceptación con informes reales de distintos proveedores. |
| 4. OCR de páginas posteriores | Implementado: después de la portada y del detalle configurado para el proveedor, el análisis de facturas busca campos ausentes por página hasta la 4. Cada página tiene caché por huella y número; el OCR corre en un proceso con límite de 25 s por página y un presupuesto total de 50 s. Se detiene cuando encuentra fechas, importe y, en facturas escaneadas, consumo y desglose esperados. | `tests/test_source_analysis.py` comprueba consumo en la página 3, fechas e importe en la 4, caché al repetir y ausencia de OCR adicional cuando la portada está completa. Falta aceptación con facturas escaneadas reales de varios proveedores. |
| 5. Bandeja global por contenido | Implementado: `community_discovery.identify_source` reúne evidencias (código en nombre o carpeta, CIF de comunidad registrada en el texto cacheado, CUPS confirmado y mismo archivo ya archivado por SHA-256) y sólo asigna si coinciden. La bandeja muestra la evidencia por grupo y el motivo de cada archivo sin asignar. Una contradicción se puede resolver eligiendo comunidad y escribiendo una justificación. `core/intake_routing.py` comprueba de nuevo las evidencias, conserva la decisión en una incidencia del documento y bloquea la aplicación automática hasta revisar la comunidad. «Repartir en expedientes» incorpora cada grupo al expediente abierto de su comunidad con el mismo análisis que «Añadir fuentes». Añadir fuentes también aparta lo que por contenido es de otra comunidad. | Probado en `tests/test_intake_by_content.py` con PDF sintéticos: carpeta mixta repartida por CIF, nombre y archivo renombrado; contradicción nombre/CIF y ausencia de evidencia sin asignar con motivo; decisión manual auditada y pendiente de revisión, justificación obligatoria, evidencia cambiada rechazada y reparto repetido sin duplicar; comunidad sin alta o sin expediente abierto bloqueada. |
| 6. Aprender CUPS de comunidades | Implementado: una factura confirmada registra en SQLite el CUPS normalizado, comunidad, suministro y documento de origen. La detección de carpetas usa ese vínculo para agrupar un PDF sin código y el alta en un expediente rechaza los CUPS de otra comunidad. Un CUPS contradictorio impide confirmar la segunda factura e identifica las dos fuentes. Las referencias genéricas no se aprenden. | `tests/test_community_cups.py` cubre el aprendizaje, la siguiente factura sin código, el rechazo de otra comunidad, la contradicción con el nombre y las fuentes sin confirmación. Falta aceptación con facturas reales sobre una copia de la base. |
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
| 15. Panel de todas las comunidades | Implementado: `core/cases_overview.py` consulta todos los expedientes de comunidades activas con su grupo (bloqueado si tiene incidencias abiertas; pendiente, listo, calculado, cartas listas, cerrado según su estado), incidencias y última salida publicada (Excel o cartas). El botón «Todas las comunidades» abre un panel no modal con contadores que filtran, búsqueda y «sólo el más reciente por comunidad». «Abrir» usa `AppGestionFincas.abrir_expediente`, que activa la comunidad y el expediente elegidos (y su período) en lugar del primero del desplegable. | Probado en `tests/test_cases_overview.py`: grupos, incidencias resueltas que no bloquean, última salida, filtros con acentos, comunidad inactiva excluida, apertura de un expediente antiguo sin saltar al más reciente y rechazo de un expediente de otra comunidad. El panel se refresca tras cada acción del expediente. |
| 16. Correo de cartas | Implementado: borradores EML con PDF por propietario, registro de ausentes y duplicados, huellas del contenido y adjunto, transporte SMTP opcional con STARTTLS y credencial del sistema, y confirmación visible antes de enviar. | `tests/test_mail_service.py` comprueba MIME, adjunto, huellas, falta de correo, duplicados, error aislado y reintento sin reenviar éxitos. Las pruebas usan transporte simulado. |
| 17. Cartas en PDF | Implementado: cada carta Word se convierte con LibreOffice Writer, se valida el PDF y se registran las dos rutas y las páginas en el mismo propietario y lote. Un fallo deja ese propietario como fallido y retira ambos archivos parciales. Los lotes Word anteriores se regeneran para añadir PDF. | `tests/test_case_letter_service.py` comprueba con Writer real dos DOCX y dos PDF de una página con importes individuales correctos, además del fallo de conversión. |

## Mantenimiento e instalación

| Nº | Estado comprobado y trabajo pendiente | Criterio de aceptación |
| --- | --- | --- |
| 18. Instalador Windows | Preparado el build `onedir` con PyInstaller, modelos OCR y Poppler, hogar de datos escribible, migración de base antigua y prueba de arranque/actualización en Windows. LibreOffice 24.2+ se instala aparte para Calc y Writer. La distribución se publica como artefacto del workflow **Distribución Windows portable**. | `tests/test_app_paths.py`, `tests/test_packaging_manifest.py` y la prueba de CI comprueban rutas con espacios y acentos, base vacía, conservación de datos y ausencia de archivos privados. Véase [Instalación portable](instalacion-portable.md). |
| 19. Copias automáticas | Implementado en `database_backup.py`: arranque, migraciones con datos y generación de Excel/reparto/cartas del flujo guiado, además de reevaluación de fuentes desde la interfaz. El flujo desatendido `pipeline.py` crea una copia antes de la ingesta y la regeneración. Retención configurable, 20 por defecto, manifiesto y registro. | Las pruebas recuperan datos confirmados en WAL, comprueban integridad y restauración, y bloquean migraciones/regeneraciones si falla la copia. El flujo desatendido se detiene antes de copiar o ingerir archivos si falla el respaldo. La limpieza conserva como mínimo 3 copias verificadas y excluye copias manuales, alteradas y de otras bases. |
| 20. Orden del código y Git | Inventariados los consumidores en [Orden del código y Git](orden-codigo-y-git.md). La interfaz deja de precargar módulos sin uso directo y `excel_generator.py` carga `excel_writer.py` sólo para la regeneración histórica. Se creó una copia recuperable, se reescribió `main` y la rama alternativa y se retiraron las ramas remotas obsoletas. Siguen pendientes la migración del flujo heredado y la purga por GitHub Support de los refs internos de PR y vistas cacheadas. | Retirar una pieza sólo tras migrar sus consumidores y pasar la suite. Confirmar con GitHub Support que los objetos antiguos dejaron de estar accesibles. Renovar los clones de trabajo que aún conserven referencias antiguas, respetando sus ramas locales. |

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
