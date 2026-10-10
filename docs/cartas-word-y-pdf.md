# Cartas Word y PDF

La mejora 17 está implementada en el flujo guiado del expediente. **Generar
cartas** produce un Word editable y un PDF por propietario, en el mismo lote
de `salidas/cartas`. LibreOffice Writer convierte el Word; el servicio abre
el PDF y registra sus páginas antes de presentar la carta como generada.

Un error de conversión o de registro deja ese propietario como fallido, retira
sus archivos parciales y conserva el diagnóstico. El resto del lote puede
terminar; los fallos deben resolverse antes de presentar todas las cartas como
listas. Los lotes anteriores se conservan en sus propias carpetas.

## Revisión del 10 de octubre de 2026

La implementación existente generaba correctamente dos Word y dos PDF con
importes individuales, pero sólo comprobaba la existencia de los archivos
para reutilizar un lote. Una carta sustituida por la de otro propietario,
incluso un PDF válido, seguía considerándose correcta.

Ahora cada carta registra también las huellas SHA-256 del Word y del PDF
publicados. Al repetir la generación, se reutiliza un lote sólo si ambos
archivos siguen coincidiendo con lo registrado. Si se modificaron, se
eliminaron o no son legibles, se genera un lote nuevo.

El correo comprueba la huella del PDF desde su generación, además de sus
controles existentes sobre el borrador y el adjunto. Un PDF sustituido no se
acepta como una nueva versión válida sólo por volver a preparar el correo.
Los ensayos de envío usan un transporte simulado, sin enviar correos reales.

## Cartas anteriores y edición manual

La migración 22 añade `output_sha256` y `pdf_sha256` a `generated_letters`.
No rellena esas huellas a partir de archivos antiguos que podrían haber sido
editados. Conserva sus rutas, páginas y estados; la copia automática protege
la base antes de migrarla.

Si un lote antiguo no tiene huellas, vuelve a **Generar cartas** desde el
expediente antes de preparar su correo. Se crea otro lote y se conserva el
anterior. Una carta editada manualmente también requiere regeneración para
que las salidas del expediente vuelvan a coincidir con sus datos confirmados.

## Validación

Las pruebas sintéticas comprueban la conversión con Writer real, propietario,
importe y número de páginas del PDF; fallos de conversión o registro; sustitución
de Word/PDF por los de otro vecino; reutilización de un lote intacto;
regeneración de lotes sin huellas y rechazo del correo con un PDF sustituido
o sin huella de generación. La migración conserva los documentos históricos
sin inventar una certificación de sus archivos.
