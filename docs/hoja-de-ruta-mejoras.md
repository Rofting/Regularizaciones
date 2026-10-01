# Hoja de ruta de las 20 mejoras

Fecha: 1 de octubre de 2026. Revisión de partida: `a60926d`.

Este repositorio es la instalación personal del usuario. Los CUPS y la
configuración de sus comunidades son reales y se conservan. La aceptación
automática usa documentos y bases temporales sintéticos; la aceptación con
documentos reales se ejecuta sobre una copia verificada de la base.

El alcance de esta entrega es corregir el clon limpio, añadir CI y preparar
esta hoja de ruta. Las demás filas describen trabajo pendiente; no representan
funciones ya entregadas.

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
- `tests/test_dependencies.py` comprueba también `rapidfuzz`, ya declarado en
  los requisitos. Evita confundir una instalación incompleta con un fallo del
  reconocimiento aproximado de proveedores.

La versión de LibreOffice instalada por los gestores de paquetes de CI todavía
puede cambiar. Fijar la misma versión estable y validar todo el recorrido en
Windows forma parte de la mejora 12. Un workflow añadido no equivale a una
ejecución de Windows aprobada: hay que comprobar el resultado de Actions.

## Orden de ejecución

| Entrega | Mejoras | Motivo y resultado |
| --- | --- | --- |
| A. Base reproducible | 11 y CI de 20 | Detectar regresiones en un clon limpio. Es el alcance de esta entrega. |
| B. Protección del cálculo | 19, 8, 10, 9, 12 | Crear copias verificadas, detectar importes heredados y comprobar el cálculo antes de ampliar automatismos. |
| C. Entrada de documentos | 6, 5, 2, 3, 4, 1, 7 | Aprender identidades confirmadas, repartir por comunidad y mejorar extracción conservando trazabilidad. |
| D. Trabajo diario | 14, 15, 13 | Revisar documentos dentro de la aplicación, ver todos los expedientes y completar altas en lote. |
| E. Salidas y distribución | 17, 16, 18 | Obtener PDF, preparar/envíar cartas con control por destinatario y construir el instalador Windows. |
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
| 8. Plantillas antiguas con gastos heredados | El modelo común ya tiene los gastos fijos vacíos y `fixed_costs.py` escribe importes del período. Una plantilla antigua puede conservarlos si no se introduce otro valor. Añadir una revisión previa de las filas conocidas de `OTROS GASTOS`. | Si el libro contiene un gasto sin valor confirmado para esa comunidad/período, mostrar celda, importe y acción para confirmar o corregir. `993,44 €` es un indicio que revisar, nunca una razón para borrar un gasto legítimo automáticamente. |
| 9. Comparación con el ejercicio previo | La base mantiene períodos y resultados históricos; falta un informe de cambios anómalos. Comparar vivienda/suministro/unidad y períodos de duración comparable, con umbrales configurables. | Un consumo o importe multiplicado por tres genera un aviso explicable con ambos valores. Sin período comparable no se inventa una referencia; rectificaciones y cambios de titular quedan identificados. |
| 10. Coherencia antes del Excel | Ya existen controles de fuentes, importes obligatorios, lecturas y conciliación monetaria. Añadir comprobación de consumo facturado frente a lecturas cuando suministro, unidad e intervalo sean comparables; suma de coeficientes y solapes de facturas. | Una diferencia relevante o un solape del mismo suministro/CUPS se puede revisar antes de publicar. No comparar kWh con m³ directamente. Los coeficientes expresados como porcentajes deben sumar 100 % dentro de tolerancia; los pesos relativos admitidos por el reparto se normalizan de forma explícita. |
| 11. Pruebas independientes de la 658 | Corregido en esta entrega: proyecto y workbook temporales sintéticos en las pruebas del flujo guiado. | La suite pasa desde un checkout que nunca ha tenido `plantillas/comunidades/658`. Las pruebas no dejan archivos privados ni salidas en el árbol de trabajo. |
| 12. Recorrido real con LibreOffice | Ya existe una prueba de exportación real con verificación de diseño, fórmulas y valores cacheados. `office_recalculation.py` conserva el diseño original y copia los resultados calculados. CI ejecuta esa comprobación en Linux y Windows; falta un ensayo completo alta → factura → Excel → reparto → cartas en ambos sistemas y acordar la versión estable de referencia. | Con dos propietarios y documentos sintéticos, el mismo recorrido conserva impresión/diseño, cuadra en céntimos y genera las cartas esperadas. Registrar y fijar la versión estable tras comprobar Windows. Una prueba con un motor simulado no sustituye a esta aceptación. |

## Uso diario e interfaz

| Nº | Estado comprobado y trabajo pendiente | Criterio de aceptación |
| --- | --- | --- |
| 13. Alta masiva | Ya hay detección de comunidades en lote y un alta guiada desde una carpeta. Falta conectar ambos para completar propietarios, lecturas, perfil y expediente por cada subcarpeta. | Un resumen permite aprobar los datos de cada comunidad; repetir el lote no duplica altas. El fallo de una comunidad conserva su diagnóstico y permite continuar con las demás. |
| 14. Vista previa de incidencias | Las extracciones guardan fragmentos/localizadores; la revisión todavía exige abrir fuentes aparte. Añadir un visor de página, navegación y resaltado cuando se conozca la caja del dato. | Al seleccionar una incidencia se ve su página y evidencia. Si sólo existe un fragmento textual, se muestra como tal sin inventar una posición en el PDF. |
| 15. Panel de todas las comunidades | La pantalla principal trabaja con la comunidad activa. Añadir una consulta global de expedientes con estado, incidencias y última salida, filtros y apertura del caso elegido. | El panel muestra pendientes, bloqueados, calculados y cartas listas, se actualiza tras una acción y abre el expediente correcto sin cambiar de período por error. |
| 16. Correo de cartas | Existe un plan detallado en `docs/superpowers/plans/2026-09-29-correo-cartas.md`; no hay servicio de envío implementado. Empezar por borradores EML auditables, después transporte SMTP con credenciales del sistema y confirmación visible del lote. | Falta de email, adjunto cambiado, error y envío exitoso quedan registrados por propietario. Reintentar no reenvía los éxitos. Las pruebas usan un transporte simulado y no envían correos reales. |
| 17. Cartas en PDF | El servicio actual publica Word y registra su identidad. Añadir conversión con LibreOffice Writer y vincular DOCX/PDF al mismo propietario y lote. | Dos propietarios producen dos DOCX y dos PDF con sus importes correctos. Se conserva el diseño, se comprueba el número de páginas y un fallo de conversión impide presentar ese PDF como listo. |

## Mantenimiento e instalación

| Nº | Estado comprobado y trabajo pendiente | Criterio de aceptación |
| --- | --- | --- |
| 18. Instalador Windows | Hay plan portable y `scripts/exportar_producto.py`, pero no un ejecutable construido. Resolver rutas de datos escribibles, empaquetar con PyInstaller, modelos OCR y herramientas documentales, y generar el instalador en Windows. | Instalar y arrancar en Windows sin Python ni descargas en tiempo de ejecución. Probar rutas con espacios/acentos, primera base vacía, importación y cartas. Actualizar conserva los datos y permite recuperar una copia. |
| 19. Copias automáticas | `database_reset.py` ya crea una copia SQLite verificada al reiniciar la base; faltan copias al abrir, migrar y regenerar. Crear un servicio común con `sqlite3.backup`, `integrity_check`, retención configurable y registro. | Las copias contienen una base íntegra incluso con WAL; una copia fallida no autoriza una migración destructiva. La limpieza sólo afecta a copias antiguas verificadas y conserva el mínimo acordado; demostrar restauración. |
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
