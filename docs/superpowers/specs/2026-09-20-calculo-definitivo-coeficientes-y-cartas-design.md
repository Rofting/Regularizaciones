# Cálculo definitivo, coeficientes y cartas

## Objetivo

Completar el flujo de regularización para que una comunidad pueda pasar de
facturas, listados de propietarios y lecturas a un Excel reproducible, un
reparto trazable y cartas correctas. El sistema debe tolerar ceros temporales
de contador sin perder las observaciones originales, importar una sola vez los
enteros de participación, derivar los precios de las facturas y permitir
repetir pasos sin borrar el historial.

## Evidencia de los ficheros reales

El informe `658 Lecturas consumo 07 2025 a 07 2026.xls` contiene 62 unidades:
50 cierres inferiores a su apertura y 12 cierres cero. El total acumulado baja
de 11.069 a 236. El informe `01 2026 a 03 2026` contiene ceros en ambas
lecturas para las 62 unidades. Los informes intermedios hasta enero de 2026 sí
contienen lecturas crecientes, salvo un reinicio individual.

El listado de propietarios de la 658 incluye 62 coeficientes que suman
100,0002: las 50 viviendas suman 80,0002, los locales 9 y los garajes 11. El
fichero `644 CAL-ACS 2025.xlsx` incluye 120 propiedades y una columna `Enteros
Participación` que suma 89,044. Estos datos demuestran que el coeficiente se
puede importar de la fuente y que el porcentaje aplicado depende de las
unidades que participen en cada concepto.

## Lecturas observadas y lecturas efectivas

Toda cifra recibida se guarda primero en `reading_observations`, incluida una
lectura cero o inferior a la anterior. Nunca se sustituye ni elimina el valor
de Meditrade.

Para cada propietario y servicio, las fuentes se procesan en orden
cronológico:

1. Una lectura positiva igual o superior a la última lectura fiable se publica
   como lectura efectiva real.
2. Una lectura cero conserva como lectura efectiva la última lectura fiable.
   La observación queda marcada como `carried_forward` y la lectura efectiva
   como estimada mediante `carry_forward_zero`.
3. Una lectura positiva inferior a la última lectura fiable se trata igual que
   un reinicio temporal: se conserva la última lectura fiable y se registra la
   disminución. No se calcula consumo negativo.
4. Cuando llega una lectura posterior igual o superior a la última lectura
   fiable, el contador recupera automáticamente la continuidad y el consumo se
   calcula desde aquella lectura fiable.
5. Si no existe una lectura anterior, no se inventa cero ni consumo. Se crea
   una incidencia que identifica propiedad, servicio, fecha, archivo y el
   valor inicial que debe introducir el usuario.
6. Una sustitución real de contador requiere confirmar el cambio y sus valores
   de retirada/instalación. Esa decisión queda auditada y permite iniciar una
   nueva serie sin esperar a que el contador alcance el acumulado anterior.

El Excel, el reparto y las cartas utilizan exclusivamente lecturas efectivas.
La interfaz muestra cuándo una lectura es real, arrastrada o confirmada por
sustitución. Un lote de ceros o disminuciones del mismo documento se resuelve
como una decisión agrupada, no como una pregunta por vivienda.

## Enteros de participación

`propietarios.coeficiente` conserva el entero registral original. El
analizador aceptará cabeceras normalizadas como `Coeficiente`, `Enteros
Participación`, `Participación` y variantes de acentos, espacios y
mayúsculas.

La primera fuente fiable asigna el valor y conserva su procedencia. Las
siguientes fuentes con el mismo valor no generan trabajo. Si falta un valor,
se solicita únicamente para esa propiedad. Si difiere del almacenado por más
de 0,0001, se crea una incidencia agrupada y nunca se sobrescribe en silencio.

El porcentaje aplicado se calcula por concepto y período:

`coeficiente aplicado = entero registral / suma de enteros elegibles`

Garajes, locales o unidades inactivas quedan fuera cuando el perfil del
concepto así lo establece. El resultado histórico guarda el entero original,
la suma elegible y el porcentaje aplicado para que un reparto pasado pueda
reconstruirse aunque cambie la configuración futura.

La carta mostrará ambos valores cuando exista reparto por coeficiente:

- `Participación registral: 2,026 %`
- `Coeficiente aplicado al reparto: 2,533 %`

## Conceptos y cálculo de costes

El catálogo técnico conserva conceptos atómicos para no perder detalle:

- ACS variable.
- ACS fijo.
- Calefacción variable.
- Calefacción fija.
- Gastos extraordinarios y ajustes cuando la comunidad los active.

La carta agrupa los cuatro conceptos ordinarios en tres líneas comprensibles:
`ACS`, `Calefacción` y `Cuota fija`. `Cuota fija` suma las partes fijas activas
de ACS y calefacción. Cada comunidad decide qué conceptos están activos en su
perfil; una comunidad sin calefacción no muestra una fila vacía.

La liquidación por consumo no pide un importe manual. Se deriva de las
facturas comprendidas o prorrateadas dentro del intervalo:

1. Separar base variable, base fija, impuestos e IVA mediante los componentes
   extraídos de cada factura.
2. Asignar al servicio la parte definida por el perfil cuando una factura
   conjunta cubra ACS y calefacción.
3. Calcular el precio neto unitario dividiendo el coste variable neto por el
   consumo efectivo total del servicio.
4. Aplicar al propietario su consumo por el precio neto y añadir los impuestos
   e IVA correspondientes con una reconciliación al céntimo contra las
   facturas.
5. Repartir las bases fijas con el método configurado para la comunidad y
   presentarlas en `Cuota fija`.

Si faltan componentes necesarios, el sistema indica qué factura y qué importe
debe confirmarse. No acepta un precio unitario derivado de un consumo total
cero.

## Repetición de pasos e historial

Los pasos anteriores siguen disponibles después de generar Excel, reparto o
cartas. Repetir un paso no borra salidas anteriores:

- Reanalizar fuentes conserva observaciones y decisiones manuales.
- Regenerar Excel crea una nueva ejecución y marca la anterior como sustituida.
- Recalcular reparto crea una versión nueva y obliga a regenerar cartas.
- Regenerar cartas crea un lote nuevo y conserva el anterior en el historial.

Cada cambio invalida únicamente las salidas posteriores. La interfaz explica
qué se recalculará antes de ejecutar la acción.

## Gráficas de las cartas

La carta conserva una sola página y dos gráficos compactos:

- Histórico de consumo del propietario por períodos comparables.
- Comparación con la comunidad mediante cinco bandas: `Muy bajo`, `Bajo`,
  `Medio`, `Alto` y `Muy alto`.

Los límites son parte del perfil del servicio para que tengan unidades y
significado estables. Para ACS se usan inicialmente `0–10`, `10–20`, `20–30`,
`30–40` y `más de 40 m³`, tal como requiere el modelo de la 658. Calefacción
debe declarar sus cuatro límites en kWh antes de publicar cartas; no reutiliza
los límites de ACS ni inventa una escala. Cada banda muestra cuántas viviendas
comparables contiene y destaca al propietario. Si no hay suficientes datos
válidos, se muestra `Comparación no disponible`. Valores provisionales se
distinguen de los reales.

## Componentes afectados

- Análisis tabular y alta de propietarios.
- Persistencia y migraciones de lecturas, coeficientes y resultados.
- Aplicación cronológica de fuentes de lectura.
- Motor de reparto, prorrateo de facturas y reconciliación.
- Estado del expediente y repetición de pasos.
- Datos de carta, plantilla y gráficas.
- Configuración de perfiles por comunidad.

Las modificaciones respetarán los cambios locales actuales. No se restaurarán
archivos, base de datos, plantillas ni perfiles que no pertenezcan al bloque en
ejecución.

## Pruebas de aceptación

- La secuencia 144, 0, 0, 160 conserva 144 durante los ceros y termina con 16
  de consumo acumulado desde la lectura fiable; todas las observaciones quedan
  auditadas.
- Un cero sin lectura anterior crea una sola incidencia manual con propiedad,
  servicio, fecha y archivo.
- Una disminución masiva del mismo informe se agrupa y nunca produce consumo
  negativo.
- La 658 importa 62 coeficientes y separa 50 viviendas de locales y garajes.
- La 644 reconoce `Enteros Participación`, importa 120 coeficientes y conserva
  la suma 89,044.
- Un coeficiente diferente no sobrescribe el maestro sin confirmación.
- El reparto guarda el coeficiente aplicado y reconcilia cada concepto al
  céntimo.
- La carta muestra participación registral y coeficiente aplicado, tres grupos
  de conceptos y ambas gráficas en una sola página.
- Repetir Excel, reparto o cartas mantiene las ejecuciones anteriores y sólo
  invalida lo posterior.
- Una liquidación deriva su precio de facturas y consumo; con consumo cero o
  componentes ausentes queda bloqueada con una explicación concreta.
