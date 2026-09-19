# Incidencias, OCR y continuidad de lecturas

## Objetivo

Eliminar el bloqueo al cerrar incidencias repetidas, hacer resolubles los
reinicios de contador agrupados desde la aplicación y aumentar la detección de
proveedores y documentos escaneados. Las lecturas recibidas se conservarán por
fecha y fuente; un valor `0` podrá reutilizar la última lectura fiable del
mismo contador sin borrar nunca el valor recibido.

## Límites de seguridad

- Una corrección manual no puede perder su auditoría ni sobrescribir una
  observación original.
- Un valor `0` no se interpreta como arrastre si no existe una lectura anterior
  del mismo propietario y suministro, o si hay evidencia de sustitución de
  contador; en esos casos se crea una incidencia.
- El valor original `0` queda almacenado aunque se derive una lectura efectiva
  arrastrada.
- La detección de proveedor requiere evidencia positiva: nombre de archivo,
  firma normalizada del documento o ambas. No se inventa un proveedor ni un
  tipo de suministro.
- RapidOCR se usa como lector local autocontenido. Tesseract continúa como
  respaldo opcional para instalaciones que ya lo tengan disponible.

## Datos y migración

`lecturas_vecino` seguirá siendo la tabla canónica de valores efectivos que
consume el reparto: una lectura por propietario, suministro y fecha. La tabla
ya está indexada por fecha y no se sustituye para evitar romper cálculos y
expedientes existentes.

Se añadirá `reading_observations` como registro inmutable de cada lectura
recibida. Contendrá el propietario, suministro, fecha, valor observado,
documento/fuente de origen, estado de aceptación, fecha de registro y, cuando
proceda, la lectura efectiva a la que dio lugar. Admitirá varias observaciones
para una misma fecha procedentes de documentos distintos.

Una observación de valor cero con lectura anterior fiable generará una lectura
efectiva de estado `estimado`, método `carry_forward_zero` y notas que enlacen
la observación y la lectura previa. La fecha de la observación se mantiene. El
registro anterior no se modifica. Si en esa fecha ya existe un valor real
diferente, se registra conflicto y no se altera la lectura canónica.

La migración de incidencias reconstruirá `review_issues` dentro de una única
transacción para eliminar la restricción que hace único el estado histórico
`(id_document, code, field_name, status)`: SQLite no permite retirar esa
restricción directamente. Después creará un índice parcial que solo impide dos
incidencias `open` equivalentes. Así se conservan resoluciones previas y puede
aparecer una nueva incidencia abierta si una relectura realmente vuelve a
requerir decisión.

## Flujo de incidencias

Al pulsar `Resolver incidencias`, todos los grupos se muestran como decisiones
accionables. Para un grupo `COUNTER_RESET` habrá siempre una acción primaria
visible: conservar lecturas previas temporalmente o abrir la estimación
individual cuando el grupo tenga un único contador. Abrir el archivo seguirá
siendo una ayuda secundaria, nunca la única acción.

Al resolver una incidencia corriente, la operación queda en una transacción:
se inserta la corrección auditada, se actualiza el candidato manual, se cierra
la incidencia y se reaplica el documento. Un choque histórico no podrá abortar
el cierre. Si la reaplicación encuentra un conflicto canónico, se revierte todo
y se enseña una incidencia comprensible, sin guardar una corrección parcial.

## OCR y proveedores

La extracción se hará en capas:

1. Capa de texto nativa de PDF o de celdas en hojas de cálculo.
2. RapidOCR sobre la primera página/imágen cuando el texto sea insuficiente,
   con preprocesado de rotación, contraste y escalado.
3. Tesseract como respaldo si RapidOCR no puede producir texto útil y el motor
   ya está presente.
4. Normalización de mayúsculas, acentos, espacios, símbolos de marca y errores
   comunes de OCR antes de comparar firmas.
5. Puntuación de proveedores: coincidencias de nombre, firma, CIF, CUPS y
   términos característicos. Se acepta solo una mejor coincidencia por encima
   de un umbral y con margen frente a la segunda; los empates se revisan.

Los perfiles de `proveedores.json` seguirán siendo editables. Se ampliarán sus
alias y firmas solo con documentos reales o pruebas representativas. Cada
resultado expondrá las señales que condujeron a la detección para facilitar la
revisión.

## Interfaz y errores

Las vistas agrupadas indicarán cantidad, archivo, regla aplicada y botón de
acción. Para un arrastre automático por cero se mostrará en el historial y en
la ficha de la lectura como `Lectura 0 recibida; se conserva temporalmente la
lectura de <fecha>`.

Los fallos de OCR se traducirán a estados breves: imagen ilegible, lector no
disponible, sin texto útil o revisión de proveedor. No se mostrarán trazas ni
se solicitará al usuario final instalar programas.

## Pruebas y aceptación

- Resolver una nueva incidencia con el mismo documento/código/campo que una
  histórica ya resuelta funciona y conserva ambas filas auditables.
- Las incidencias abiertas duplicadas siguen bloqueadas.
- Un grupo de reinicios ofrece una acción resolutiva desde la bandeja y desde
  `Resolver incidencias`.
- Una lectura `0` con anterior real crea una observación inmutable y una
  efectiva arrastrada; sin anterior, con sustitución o con conflicto no altera
  el valor efectivo y crea revisión.
- RapidOCR se invoca para imágenes/PDF sin texto, normaliza una firma con OCR
  imperfecto y no clasifica proveedores ambiguos.
- Las suites de revisión documental, UI, ingestión, análisis de fuentes y
  migraciones pasan; también pasan la compilación y la comprobación de espacios.
