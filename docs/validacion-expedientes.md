# Validación de expedientes e incidencias

Esta guía describe el flujo actual de apertura de un expediente y revisión manual de sus fuentes. El documento de origen se archiva como copia para su revisión; los originales permanecen intactos. La extracción automática llegará en una entrega posterior: por ahora, el documento se presenta y se revisa manualmente.

## Pasos operativos

1. **Abrir la app y seleccionar comunidad.** Si todavía no existe, pulsa **+
   Comunidad** y elige **Registro rápido** o **Alta guiada desde fuentes**.
2. **Pulsar Crear expediente y elegir fecha inicial/final.** Puede usarse cualquier rango, incluido uno arbitrario de seis meses.
3. **Pulsar Añadir fuentes; los originales permanecen intactos.**
4. **Elegir el tipo de fuente y seleccionar archivos.**
5. **Pulsar Resolver incidencias; cada incidencia abre su copia archivada.**
6. **Confirmar valor y motivo; el sistema conserva la auditoría.**
7. **Esperar el estado Listo para cálculo antes de usar las salidas posteriores.**

## Alta guiada desde fuentes

El asistente permite crear una comunidad sin disponer de un Excel manual
perfecto. Presenta cinco etapas consecutivas y visibles:

1. **Identidad:** código, nombre y, opcionalmente, datos del primer período.
2. **Fuentes:** lista de propietarios CSV y una o más lecturas Excel o PDF.
   Ambos grupos son obligatorios; las facturas PDF son opcionales.
3. **Detección:** servicios y ambigüedades encontrados durante un análisis que
   todavía no escribe datos.
4. **Confirmaciones:** selección de servicios y respuestas obligatorias para
   cualquier dato incierto. El asistente no inventa respuestas.
5. **Resumen:** revisión final antes de crear la comunidad, su perfil, su
   plantilla y, si se indicó un período, el expediente inicial.

La publicación se ejecuta en segundo plano y la barra de estado muestra la
operación en curso. Al terminar se actualizan los selectores de comunidad,
período y expediente. Las fuentes archivadas, la base de datos, los perfiles
generados y las plantillas instaladas son datos locales y privados: no se
incluyen en el control de versiones.

## Comprobaciones de aceptación

- Un rango arbitrario de seis meses se puede abrir y conservar como expediente.
- Al añadir una factura PDF se crean para confirmar los tres campos `fecha_inicio`, `fecha_fin` e `importe_total`.
- Cancelar o dejar en blanco una confirmación no escribe cambios.
- Volver a añadir una fuente no duplica la fuente ni sobrescribe una corrección ya confirmada.
- Cambiar de comunidad impide utilizar un expediente antiguo.
- Una incidencia sin resolver bloquea el estado **Listo para cálculo**.
- Los originales no se modifican: durante la resolución se abre únicamente la copia archivada.
- El resumen del alta guiada no está disponible mientras falten propietarios,
  lecturas o respuestas obligatorias.

## Alcance actual

La revisión ordinaria de expedientes conserva el flujo manual descrito. El
alta guiada sí analiza las fuentes para proponer un perfil, pero no calcula ni
prorratea costes, no genera una lógica nueva de cartas ni envía correos. Un
Excel manual previo es opcional y no bloquea el alta.
