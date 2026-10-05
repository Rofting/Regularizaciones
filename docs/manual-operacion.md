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

Cada propietario recibe un Word y un PDF en la misma carpeta del lote. La
aplicación registra ambos archivos y el número de páginas del PDF. Si falla la
conversión, ese propietario figura como fallido y no se publican archivos
parciales; revisa el error y vuelve a generar las cartas.

Tras generar todas las cartas, pulsa **Preparar correos**. La aplicación crea
un borrador EML con el PDF de cada propietario que tenga correo válido, e
informa de direcciones ausentes o repetidas. Abre la carpeta y revisa los
destinatarios y adjuntos. Cambiar una carta o un destinatario crea un lote
distinto al preparar otra vez.

**Enviar correos** es opcional. Requiere confirmar el lote en pantalla e
indicar los propietarios que quieres excluir con su motivo, si los hay, e
introducir servidor SMTP con STARTTLS, puerto, usuario y contraseña. La
contraseña se entrega al almacén de credenciales del sistema; no se guarda en
la base ni en los borradores. El resultado de cada propietario queda registrado.
Al reintentar, los envíos ya completados no se repiten. Las pruebas nunca
establecen conexiones SMTP reales.

El botón lateral **Reparto** nunca salta los pasos anteriores: abre la acción
pendiente que muestre el panel central.

## Qué significa «listo»

Que no queden incidencias manuales no significa por sí solo que el expediente
pueda calcularse. El panel central comprueba también los datos estructurales y
muestra una única acción inmediata:

- **Añadir fuentes** si todavía no hay documentos útiles.
- **Gestionar períodos** si las fuentes no están ligadas al intervalo correcto.
- **Importar propietarios** si faltan las unidades activas de la comunidad.
- **Completar lecturas** sólo cuando un concepto activo se reparte por consumo.
- **Revisar coeficientes** cuando un concepto activo los necesita y alguno no
  tiene un valor positivo guardado.
- **Preparar o revalidar el perfil Excel** si falta, hay más de uno activo o su
  huella ya no coincide con el registro de la base de datos.
- **Generar Excel oficial** antes del reparto y **Calcular reparto** antes de las
  cartas. Un Excel o un reparto antiguo no autoriza a seguir si cambiaron sus
  entradas.

Después de analizar una carpeta, la aplicación vuelve a seleccionar el mismo
expediente y actualiza este diagnóstico. Si falta algo, no es necesario buscar
un botón oculto: usa la acción principal que aparece en el centro.

### Carpeta mixta de correo

1. Pulsa **Bandeja global** sin seleccionar comunidad ni período.
2. Elige la carpeta que contiene las facturas, lecturas o listados.
3. Cada archivo se asigna a una comunidad sólo si todas sus evidencias
   coinciden: código en el nombre o la carpeta (`658 - Factura feb 2026.pdf`),
   CIF de una comunidad registrada en el texto, CUPS ya confirmado o el mismo
   archivo ya archivado en un expediente (aunque se haya renombrado). Cada
   grupo muestra cuántos archivos se reconocieron por cada evidencia y el
   expediente abierto al que irán.
4. Los archivos con evidencias contradictorias (por ejemplo, el nombre dice 701
   y el CIF es de 658) o sin ninguna quedan sin asignar con su motivo. Elige su
   comunidad en el desplegable si la conoces; nunca se envían a la comunidad
   que esté abierta en pantalla.
5. **Repartir en expedientes** añade cada grupo como fuentes del expediente
   abierto más reciente de su comunidad, con el mismo análisis que «Añadir
   fuentes». Las comunidades sin alta o sin expediente abierto se indican y no
   se reparten. Repetir el reparto no duplica documentos.
6. Las comunidades nuevas siguen pudiéndose crear con **Crear comunidades
   seleccionadas**. El mes del nombre es una ayuda de clasificación, no
   sustituye las fechas exactas que se extraen de la factura.

Al añadir fuentes a un expediente también se apartan los archivos cuyo
contenido indica otra comunidad (CIF, CUPS o archivo ya archivado en ella).

## PDF escaneados

Al encontrar un PDF sin texto, la aplicación intenta OCR de la primera página
para evitar esperas largas. Si no puede leerlo, la incidencia explica si el
problema fue la imagen, el motor local o el lector PDF. Abre el original y
clasifícalo o introduce el campo solicitado; no es necesario seguir enlaces de
instalación desde la incidencia.

Si la portada escaneada está torcida, tenue o con puntos de polvo, se prepara
una copia de la imagen (enderezado de hasta ±5°, reducción de ruido y
contraste) y se lee también. La aplicación se queda con la lectura que trae más
fechas, importes y CIF, así que una página que ya se leía bien no empeora. El
original nunca se modifica y el tratamiento aplicado queda en el diagnóstico de
la fuente. Para comprobarlo con documentos reales:
`python scripts/comparar_ocr.py CARPETA` muestra, por archivo, los problemas
detectados y los datos leídos con y sin preparación, sin tocar nada.

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
