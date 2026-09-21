# Manual de operación

## Dos formas de empezar

### Expediente conocido

1. Selecciona la comunidad y el período de trabajo.
2. Crea o selecciona el expediente.
3. Pulsa **Añadir fuentes** y elige archivos o una carpeta. Se admiten PDF,
   XLS, XLSX y CSV; los originales se archivan sin modificar. La barra inferior
   muestra las fases **Huella**, **Leyendo texto**, **Clasificando**,
   **Identificando proveedor**, **Extrayendo campos** y **Finalizado**.
4. El sistema elimina duplicados por contenido, reutiliza el texto/OCR ya leído
   y separa facturas, abonos, lecturas, propietarios, presupuestos, albaranes,
   justificantes e informes. Un archivo ilegible no detiene el resto del lote.
5. Las fuentes completas y seguras se aplican sin confirmación individual.
   **Validar** muestra sólo decisiones raíz pendientes. Si falta el emisor se
   pregunta una vez por el proveedor y después se reanalizan sus campos.
6. Al pulsar **Abrir archivo**, contrasta el dato con el original y confirma
   únicamente el valor que solicita la pantalla.
7. Cuando no haya incidencias, pulsa **Confirmar fuentes**. Este paso aplica
   los datos revisados al expediente.
8. Pulsa **Generar Excel oficial**, después **Calcular reparto** y, una vez
   conciliado, **Generar cartas**.

El botón lateral **Reparto** nunca salta los pasos anteriores: abre la acción
pendiente que muestre el panel central.

### Carpeta mixta de correo

1. Pulsa **Bandeja global** sin seleccionar comunidad ni período.
2. Elige la carpeta que contiene las facturas, lecturas o listados.
3. La aplicación agrupa por el código al comienzo del nombre, por ejemplo
   `658 - Factura feb 2026.pdf`. También muestra el mes/año y el tipo orientativo
   detectados en el nombre.
4. Las comunidades nuevas aparecen marcadas para crearse. Confirma sólo las
   filas correctas. Los documentos sin código o en `Sin_Comunidad` quedan sin
   asignar: nunca se envían a la comunidad que esté abierta en pantalla.
5. Crea o selecciona después el expediente de cada grupo para incorporar sus
   fuentes y completar la validación. El mes del nombre es una ayuda de
   clasificación, no sustituye las fechas exactas que se extraen de la factura.

## PDF escaneados

Al encontrar un PDF sin texto, la aplicación intenta OCR de la primera página
para evitar esperas largas. Si no puede leerlo, la incidencia explica si el
problema fue la imagen, el motor local o el lector PDF. Abre el original y
clasifícalo o introduce el campo solicitado; no es necesario seguir enlaces de
instalación desde la incidencia.

El texto se guarda en una caché local por huella SHA-256 y versión del lector.
Al reanalizar una fuente idéntica no se vuelve a ejecutar OCR. Todo el proceso
es local: el PDF y su texto no se envían a servicios externos.

## Reconocimiento y aplicación

- Reconocer una factura no significa aplicarla automáticamente. Deben coincidir
  comunidad, período y módulo activo de la comunidad.
- Una factura reconocida de un servicio no activo se conserva en el histórico,
  pero no altera el Excel, el reparto ni las cartas.
- Un valor de confianza baja queda como evidencia para revisión; nunca pasa a
  las tablas canónicas.
- Después de corregir proveedor o tipo documental, pulsa **Reanalizar fuentes**.
  Se mantienen las correcciones manuales válidas y se recalcula lo pendiente.

## Auditor privado del catálogo

El auditor sirve para medir el reconocimiento sobre una carpeta real sin
guardar en el repositorio nombres, texto de facturas ni datos personales. Desde
PowerShell, en la carpeta del proyecto:

```powershell
$env:REGULARIZACIONES_SOURCE_ROOT = "C:\ruta\de\las\fuentes"
.\.venv-fase1\Scripts\python.exe scripts\audit_provider_catalog.py `
  --root "$env:REGULARIZACIONES_SOURCE_ROOT" `
  --database data\gestion.db `
  --output .private-audit\provider-catalog.json
```

El JSON queda en `.private-audit/`, carpeta ignorada por Git. Los indicadores
principales son:

- `text_invoices_recognised_pct`: porcentaje de facturas con proveedor resuelto;
  el objetivo operativo es al menos 90 %.
- `provider_decisions_remaining`: documentos que aún requieren decidir emisor;
  el objetivo del corpus de referencia es menos de 60.
- `non_invoice_false_positives`: documentos no facturables tratados como factura;
  debe permanecer en 0.

## Reglas de seguridad

- Las copias con nombre `(2)` se tratan como posible duplicado, pero la huella
  del archivo decide si son realmente la misma fuente.
- Una comunidad nueva se crea sólo cuando el código es inequívoco. CIF distintos
  en el mismo grupo bloquean su alta automática.
- El período oficial se basa en las fechas verificadas del documento. Nunca se
  calcula un reparto sólo a partir del mes incluido en el nombre del correo.
