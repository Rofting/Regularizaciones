# Comparación con el ejercicio anterior

La mejora 9 se integra en **Coherencia** y en el control previo al Excel oficial.
Los avisos muestran los dos valores, sus fechas, la vivienda o factura y el
sentido del cambio. Se revisan antes de generar el Excel y las cartas.

## Uso y umbrales

Abre **Coherencia** en el expediente. El umbral inicial es **×3**, tanto para
subidas como para bajadas. Puedes cambiar el factor a cualquier valor mayor
que 1. También puedes ajustar la diferencia máxima de duración entre los dos
intervalos; por defecto es el 10 % de la duración anterior. Guarda con
**Guardar configuración y coeficientes**.

Un aviso se corrige revisando la fuente o se acepta con un motivo, por ejemplo
un cambio comprobado de tarifa u ocupación. La aceptación conserva la huella
de los datos actuales y anteriores: modificar cualquiera de ellos vuelve a
abrir la revisión. Los consumos que pasan de cero a positivos, o de positivos
a cero, generan aviso aunque no se pueda calcular un cociente.

## Referencia e intervalos

- Se usa el último período de la **misma comunidad** que haya terminado al
  comenzar el expediente. Si varios períodos cierran ese día, la referencia
  se considera ambigua. No se busca arbitrariamente otro ejercicio para
  sustituir un período reciente incompleto o incompatible.
- Los intervalos deben tener duración similar dentro de la tolerancia y
  situarse aproximadamente un año después: ambos límites pueden diferir como
  máximo 31 días de su aniversario. Se contempla el 29 de febrero.
- No se prorratean valores. Las fechas reales se muestran para que puedas
  comprobar las diferencias de duración y temporada.

## Facturas

Se compara el importe total en euros de cada factura con una factura del
ejercicio anterior del mismo suministro, CUPS/referencia y unidad de consumo.
Entre los intervalos compatibles se elige el más próximo al aniversario de
ambos límites. Un empate queda sin comparar: puede haber originales y
rectificaciones. El número de factura puede cambiar de un año a otro.

Los importes negativos o nulos se identifican y excluyen de la comparación
ordinaria. Las anotaciones que contienen «rectific…» se señalan para revisión.
No se deduce una rectificación positiva sin evidencia. Sin punto identificado,
unidad o factura compatible se explica por qué no hay comparación.

## Viviendas y titulares

Se compara **consumo del intervalo**, no el valor acumulado del contador.
Se utilizan las mismas lecturas efectivas y aprobaciones que el reparto,
incluidas las conservaciones automáticas de la última lectura fiable. Los
intervalos reales de lecturas también deben ser comparables. Una vivienda
sin lecturas anteriores válidas queda sin comparar, con su motivo visible.

ACS usa la unidad m³ del modelo existente. En calefacción debes confirmar
la misma unidad de ambos ejercicios: kWh o unidades de repartidor. Puedes
seleccionar **Sin confirmar** para evitar comparar contadores cuya unidad
haya cambiado o no esté comprobada. Si la unidad guardada en un reparto
anterior difiere, esa vivienda no se compara.

La migración 17 añade nombre del titular y código de vivienda a las copias
históricas de los futuros repartos. Si el titular cambia, el panel muestra
ambos nombres y mantiene la comparación de la vivienda. Un cambio del código
histórico deja su identidad pendiente de revisión y evita compararla.

Los repartos ya existentes conservan sus importes y dejan esos campos vacíos:
no se copia el titular actual como si fuera el de entonces. El panel indica
que en esos históricos no puede comprobarse el cambio de titular.

## Salidas y pruebas

Las fuentes históricas forman parte de la identidad de las entradas del
expediente. Modificar el período de referencia, una factura, lectura o reparto
histórico hace que el Excel anterior necesite regenerarse antes de repartir o
generar cartas. Las decisiones sobre los nuevos avisos utilizan las copias
automáticas y el registro de revisión existentes. La migración se protege con
la copia verificada que ya crea el sistema antes de actualizar una base.

Las pruebas usan comunidades, propietarios y documentos sintéticos. Cubren
incrementos y caídas, límites y configuración, ceros, meses contiguos,
rectificaciones, unidades/puntos diferentes, referencias ambiguas, lecturas
ausentes o sin aprobar, invalidación de salidas y conservación del titular
anterior al recalcular. También verifican la actualización desde la versión
16 sin inventar datos en los repartos existentes.
