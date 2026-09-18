# Manual de operación

## Dos formas de empezar

### Expediente conocido

1. Selecciona la comunidad y el período de trabajo.
2. Crea o selecciona el expediente.
3. Pulsa **Añadir fuentes** y elige archivos o una carpeta. Se admiten PDF,
   XLS, XLSX y CSV; los originales se archivan sin modificar.
4. Revisa las incidencias. Al pulsar **Abrir archivo**, contrasta el dato con
   el original y confirma únicamente el valor que solicita la pantalla.
5. Cuando no haya incidencias, pulsa **Confirmar fuentes**. Este paso aplica
   los datos revisados al expediente.
6. Pulsa **Generar Excel oficial**, después **Calcular reparto** y, una vez
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

## Reglas de seguridad

- Las copias con nombre `(2)` se tratan como posible duplicado, pero la huella
  del archivo decide si son realmente la misma fuente.
- Una comunidad nueva se crea sólo cuando el código es inequívoco. CIF distintos
  en el mismo grupo bloquean su alta automática.
- El período oficial se basa en las fechas verificadas del documento. Nunca se
  calcula un reparto sólo a partir del mes incluido en el nombre del correo.
