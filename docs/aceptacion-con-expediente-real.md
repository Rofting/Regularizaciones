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
  reparación si su texto, su nota o el nombre del archivo hablan de reparación,
  avería, fuga o sustitución.
- **Previsiones («sin recibir aún»).** No cuentan; se excluyen con «Excluir fila».

- **Lecturas.** No se adivinan. Si una lectura falta, viene a cero o baja, se
  conserva la anterior; cuando llega la lectura real, sustituye a la conservada.
  La lectura que abre el ejercicio se busca primero entre las del propio
  período, aunque el cierre del anterior esté fechado unos días después.
- **Factura, abono y refactura.** Una factura y su abono íntegro (mismo
  suministro, mismo tramo, importe opuesto) se anulan entre sí y no cuentan
  como solape; queda la factura rectificada. Un abono parcial sí pide revisión.
- **TotalEnergies.** El término fijo incluye fijo, peajes («Otros»), cuota
  GTS/CNMC y alquiler; el variable es el resto. Ambos suman el total cobrado.

## Comprobado con el ejercicio 2025-2026

Con la carpeta real el expediente llega hasta el final: Excel sin errores de
fórmula, reparto y 50 cartas. Lo que hace falta hacer a mano:

1. Apartar las facturas de otro ejercicio (botón «Apartar las N facturas de
   otro ejercicio» en la bandeja).
2. Anotar las cuotas cobradas (400 €/mes desde agosto, 300 €/mes desde enero y
   las liquidaciones por consumo).
3. Aceptar en «Coherencia» las variaciones de consumo por vivienda.

Diferencias con el estudio manual, todas por documentos que no están en la carpeta:

| Concepto | App | Estudio | Motivo |
|---|---|---|---|
| Gas | 11.553,36 € | 27.789,64 € | La carpeta tiene 8 de las 21 facturas del estudio; las presentes coinciden al céntimo. |
| Agua | 1.516,57 € | 2.569,46 € | 7 de las 12 facturas; coinciden al céntimo. |
| Electricidad | 2.554,85 € | 2.554,85 € | Igual. |
| Lecturas de contadores | 993,44 € | 993,44 € | Igual. |
| Consumo ACS | 358 m³ | 769 m³ | Ver «Pendiente»: cambio de contadores. |

## Pendiente

- **Cambio de contadores (decisión del despacho).** En 2026 se sustituyeron los
  contadores: el informe de enero a marzo viene entero a cero y el anual acaba
  en lecturas pequeñas (contador nuevo). Con la regla actual se conserva la
  última lectura del contador viejo, así que el consumo del ejercicio se queda
  en 358 m³ y no cuenta lo medido por los contadores nuevos (236 m³). El
  estudio manual usa 769 m³. Falta decidir si, ante un cambio de contador, se
  suma la lectura del contador nuevo.
- **Placas solares.** No se distinguen del resto del mantenimiento: todo el
  mantenimiento va a «sala de calderas» salvo que se anote a mano.
