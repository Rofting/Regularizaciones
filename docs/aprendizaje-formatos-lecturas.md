# Aprendizaje de formatos de lecturas

La mejora 1 conserva la distribución de las columnas de un informe confirmado
para reconocer próximos informes de la misma empresa. Funciona en el análisis
de tablas de PDF, Excel y CSV. Si una tabla no se reconoce, siguen disponibles
los lectores de formatos anteriores.

## Uso

1. Importa el primer informe de una empresa de lecturas. El formato todavía no
   aprendido queda pendiente de confirmación de la fuente.
2. Comprueba las viviendas, servicio, fechas y valores y confirma el documento.
   El formato sólo se aprende si la aplicación de las lecturas termina validada.
3. Los siguientes informes con las mismas cabeceras consultan el mapeo guardado,
   incluso en otra comunidad. Si los datos son completos y coherentes se pueden
   aplicar automáticamente por el flujo habitual.
4. Un orden de columnas diferente o un consumo que no cuadra abre «Revisar
   formato y cuadre de lecturas». Comprueba el original y resuelve las incidencias
   antes de confirmar la fuente. Al confirmar un formato diferente se conserva
   como otro formato de esa empresa, sin borrar los anteriores.

## Datos y límites

La migración 21 añade `learned_reading_formats`. Guarda empresa, firma de
cabeceras, versión del extractor, índices y funciones de las columnas, fuente,
persona que confirmó y fecha del aprendizaje. La firma conserva el orden de las
cabeceras y normaliza sus fechas; el mapeo no almacena comunidad, propietarios,
valores ni fechas del período. Cada informe aporta sus datos actuales.

No se aprende de la aplicación automática ni de una empresa sin identificar.
Se ignoran versiones anteriores del extractor y registros de mapeo dañados.
El aprendizaje no sustituye las comprobaciones de servicio, consumo, período,
viviendas, ceros o reinicios de contador.

La aceptación sintética comprueba confirmación y trazabilidad, reutilización en
otra comunidad/período, importación PDF/Excel/CSV, revisión por cambio de columnas,
bloqueo previo a confirmar y rechazo de un consumo incoherente entre veinte filas.
Queda pendiente repetir la aceptación con informes reales sobre una copia
verificada de la base del despacho.
