# Validación de expedientes e incidencias

Esta guía describe el flujo actual de apertura de un expediente y revisión manual de sus fuentes. El documento de origen se archiva como copia para su revisión; los originales permanecen intactos. La extracción automática llegará en una entrega posterior: por ahora, el documento se presenta y se revisa manualmente.

## Pasos operativos

1. **Abrir la app y seleccionar comunidad.**
2. **Pulsar Crear expediente y elegir fecha inicial/final.** Puede usarse cualquier rango, incluido uno arbitrario de seis meses.
3. **Pulsar Añadir fuentes; los originales permanecen intactos.**
4. **Elegir el tipo de fuente y seleccionar archivos.**
5. **Pulsar Resolver incidencias; cada incidencia abre su copia archivada.**
6. **Confirmar valor y motivo; el sistema conserva la auditoría.**
7. **Esperar el estado Listo para cálculo antes de usar las salidas posteriores.**

## Comprobaciones de aceptación

- Un rango arbitrario de seis meses se puede abrir y conservar como expediente.
- Al añadir una factura PDF se crean para confirmar los tres campos `fecha_inicio`, `fecha_fin` e `importe_total`.
- Cancelar o dejar en blanco una confirmación no escribe cambios.
- Volver a añadir una fuente no duplica la fuente ni sobrescribe una corrección ya confirmada.
- Cambiar de comunidad impide utilizar un expediente antiguo.
- Una incidencia sin resolver bloquea el estado **Listo para cálculo**.
- Los originales no se modifican: durante la resolución se abre únicamente la copia archivada.

## Alcance actual

La fuente se archiva y se revisa manualmente durante esta entrega. Este flujo no parsea automáticamente PDFs, no calcula ni prorratea costes, no produce el Excel final, no genera una lógica nueva de cartas ni envía correos. La extracción automática se incorporará en una entrega posterior.
