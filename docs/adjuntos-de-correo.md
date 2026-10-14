# Importar documentos desde correos guardados

En un expediente, abre **Fuentes → Añadir fuentes → Añadir archivos** y
selecciona correos `.eml` o `.msg`, junto con los PDF, Excel y CSV habituales.
**Añadir carpeta** también encuentra correos en las subcarpetas visibles.
No se requiere Outlook instalado para leer los MSG.

El correo se archiva sin cambios en las fuentes del expediente. Sus adjuntos
PDF, XLSX, XLS y CSV se conservan y pasan por el mismo detector de comunidad,
extracción, revisión y aplicación que los documentos seleccionados directamente.
El cuerpo del correo no se interpreta como factura. El correo original queda
como evidencia informativa y no participa en el cálculo ni en la reevaluación
de documentos; se reevaluarán sus adjuntos incorporados.

La pantalla informa de correos archivados, adjuntos incorporados, duplicados,
formatos omitidos y errores. Un correo sin adjuntos compatibles, un MSG dañado
o un adjunto ilegible no detiene los demás documentos del lote. Los adjuntos
que indican otra comunidad se apartan, se conservan y no se aplican al
expediente seleccionado. Para incorporarlos, abre el expediente de su comunidad
y añade el correo allí. Esta entrega integra correos en **Añadir fuentes**;
la Bandeja global y las altas desde carpetas siguen recibiendo documentos
PDF, Excel y CSV directamente.

En **Ver en la fuente**, la revisión muestra también el correo de origen,
su asunto y remitente. SQLite conserva la relación de cada adjunto con el
correo, su nombre original, ruta archivada, huella SHA-256 y documento aplicado.
Dos adjuntos con el mismo nombre conservan sus contenidos separados. Repetir
un correo, recibir la misma factura en otro correo o haberla incorporado suelta
no duplica el documento ni la factura dentro del expediente. Los reintentos
no reemplazan las decisiones ya confirmadas. Si el correo o un adjunto
archivado han cambiado, se informa del error y no se reutilizan silenciosamente.

Sólo se extraen adjuntos con nombre y formato compatible; no se ejecutan archivos
ni se descomprimen ZIP. Se admiten hasta 100 adjuntos por correo, 25 MB por
adjunto y 50 MB por correo o conjunto de adjuntos. Los nombres se normalizan
para Windows y Linux y no pueden escribir fuera de su carpeta de archivo.

La migración 23 añade `email_imports` y `email_attachments`; se aplica con la
copia de seguridad previa habitual. Los MSG se leen con `extract-msg`, incluido
en la distribución Windows. El workflow portable abre un EML y un MSG sintéticos
con el ejecutable construido para comprobar que contiene el lector y sus codecs.

Aceptación sintética: `tests/test_email_ingestion.py` usa MIME real y genera MSG
OLE/MAPI que lee la biblioteca real. Verifica archivado, aplicación y deduplicación
de facturas, nombres repetidos, rutas, formatos, tamaños, comunidad, errores y
reevaluación. `tests/test_expedient_ui.py` recorre el botón de selección y el
resumen con un correo válido y otro dañado. Queda pendiente comprobar correos
reales del despacho, que no se han proporcionado en esta entrega.
