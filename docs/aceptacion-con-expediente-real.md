# Aceptación con un expediente real

Se ha pasado por la aplicación la carpeta completa de una comunidad real.
Primero se importó el estudio del ejercicio anterior como «modelo inicial».
Después se añadieron las facturas PDF, las lecturas y los listados del
ejercicio siguiente. El resultado se ha comparado con el estudio hecho a mano.

La suite sintética pasaba entera, pero el recorrido real se bloqueaba en
varios puntos. Las correcciones de esta entrega tienen su prueba en
`tests/test_aceptacion_real_regresiones.py`.

## Bloqueos corregidos

| Problema con datos reales | Corrección |
| --- | --- |
| La «Fecha fra.» del gas venía arrastrada del ejercicio anterior. Al corregirla en la incidencia sólo se guardaba el valor: la factura nunca se importaba. El expediente quedaba «listo» y el Excel se negaba con «Faltan facturas normalizadas de GAS». | Las correcciones confirmadas de una celda del maestro se aplican al momento. Las facturas se vuelven a leer del maestro archivado, sin duplicar las ya importadas. |
| «Cerrar: no corresponde al período» sobre una fila marcaba el maestro entero como no aplicable. | En un maestro se excluye sólo esa fila. El nuevo botón «Excluir fila» sirve para previsiones («sin recibir aún») y filas incompletas. |
| Las filas preparadas (fecha y fórmulas a cero) creaban incidencias. | Una fila sin ningún importe no es una factura y se ignora. |
| Un maestro de otro año tenía el gas con una columna menos. El perfil leía la «Lectura inicial» como fecha de inicio. | Si el rótulo de la columna contradice su significado, se busca la única columna del encabezado que sí lo cumple. |
| Locales y garajes pedían lecturas de ACS (13 incidencias falsas). En el informe de lecturas aparecían como «lectura cero». | Sólo las viviendas participan en el reparto por contador, igual que en el resto de la aplicación. |
| Faltaba una lectura final en el informe. Al resolver la incidencia no se guardaba nada y el Excel fallaba con «Faltan lecturas». | El valor confirmado se guarda como estimación manual aprobada. Si faltan las dos, se indican como «inicial;final». |
| Las correcciones de celdas del maestro no se validaban: se aceptaba cualquier texto como fecha. | Se validan como fecha o importe. |
| El listado de propietarios en Excel no se reconocía al importar el modelo. Los correos impresos en la línea siguiente se perdían. | El listado se convierte al CSV canónico, como en el alta. Se conserva el correo de la línea siguiente. |
| Con importes de más de dos decimales (fórmulas del maestro), la conciliación fallaba por un céntimo. | La conciliación se hace céntimo a céntimo por factura, igual que se escribe la hoja. |
| Sin cuotas anotadas, la hoja de cobros recibía el coste real en vez de lo cobrado. La diferencia salía siempre a cero. | Se usa lo cobrado del período. Si no existe, se mantiene el comportamiento anterior. |
| El término fijo y el variable se escribían sin impuestos y el total con impuestos. Los % de fijo y variable del análisis no sumaban 100 % y una parte del coste no se repartía. | Los impuestos se reparten entre ambos términos en proporción a su importe, como en el estudio manual. Si la diferencia supera los impuestos habituales, «Coherencia» lo avisa. |
| `GAS!Q31` mostraba `#DIV/0!`: el precio medio por m³ de un gas facturado en kWh. | Los precios medios de la fila de sumas quedan vacíos si el divisor es cero. |
| Cada factura de mantenimiento de un solo día pedía una aceptación en «Coherencia». | Sólo se avisa en los suministros con consumo. |
| Las facturas de Gómez Group Metering y algunos importes de Mantenimientos Zaragoza no se reconocían. | Se han ajustado sus reglas en el catálogo de proveedores. |

## Decisiones del despacho, ya aplicadas

- **Criterio de fecha.** Cuenta la factura emitida dentro del ejercicio, aunque
  el consumo sea del mes anterior. Sólo si la factura no trae fecha de emisión
  se juzga por el período facturado.
- **Mantenimiento y lecturas de contadores.** Se deducen de las facturas del
  ejercicio y se escriben en las filas de «OTROS GASTOS» como importe mensual
  medio, de modo que el anual es exactamente la suma de las facturas. Lo que se
  anote en «Gastos fijos» manda sobre lo deducido.
- **Reparaciones.** Entran por una quinta parte en el ejercicio de la factura y
  en los cuatro siguientes (también las anotadas en `gastos_extra`). La
  descripción de la fila indica «1/5 de …». Una factura de mantenimiento es
  reparación si su texto (o su nota) habla de reparación, avería o sustitución.
- **Previsiones («sin recibir aún»).** No cuentan; se excluyen con «Excluir fila».

## Pendiente

- **Lecturas del ejercicio.** Un informe tiene todos los contadores a cero.
  Otro muestra que los contadores se cambiaron (la lectura final es menor que
  la inicial). Hace falta decidir la estimación de esos contadores.
- **Placas solares.** No se distinguen del resto del mantenimiento: todo el
  mantenimiento va a «sala de calderas» salvo que se anote a mano.
