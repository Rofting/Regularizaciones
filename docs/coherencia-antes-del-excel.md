# Coherencia antes del Excel

La mejora 10 añade controles a la preparación del expediente y a la generación
del Excel oficial. Una llamada directa al exportador también los ejecuta.
El mismo panel incluye la [comparación con el ejercicio anterior](comparacion-con-ejercicio-anterior.md)
de la mejora 9, con avisos y umbrales propios.

## Uso

1. Selecciona una comunidad y un expediente con período enlazado.
2. Abre **Coherencia** en las herramientas. Si hay un bloqueo, el expediente
   ofrece **Revisar coherencia**.
3. Revisa los coeficientes. Selecciona **Porcentajes** si representan porcentajes
   de participación; deben sumar 100 %, con tolerancia de 0,01 puntos. El modo
   inicial es **Pesos relativos**, compatible con el reparto existente: usa
   valores positivos y los normaliza por su suma. Puedes corregirlos y guardar.
4. Declara las comparaciones de consumo aplicables a esa comunidad: suministro,
   ACS/calefacción, unidades y, si hay varios puntos, CUPS o referencia. Guarda
   una tolerancia de consumo; por defecto es el 10 % del consumo facturado.
5. Corrige los errores. Un solape o una diferencia de consumo pueden aceptarse
   con un motivo, por ejemplo un abono comprobado o un consumo de zonas comunes.
   Los coeficientes inválidos, porcentajes incorrectos, fechas invertidas y
   lecturas que necesitan aprobación requieren corrección.

Los coeficientes pertenecen al listado de propietarios de la comunidad:
corregirlos afecta también a otros expedientes que usen ese listado.

## Qué se comprueba

- **Coeficientes:** viviendas activas del reparto que utiliza coeficientes;
  distingue porcentajes de pesos relativos. Un concepto opcional sin importe
  no obliga a introducir coeficientes que el cálculo no usa.
- **Solapes:** intervalos de facturas del mismo suministro y punto. Dos puntos
  identificados diferentes no generan un aviso. Si falta identificar alguno,
  se indica «Posible solape». Compartir sólo la fecha de límite no cuenta como
  solape. Se muestran ambas facturas y las fechas para revisar rectificaciones.
- **Consumo:** utiliza las mismas lecturas efectivas, estimaciones y requisitos
  de aprobación que el reparto. Sólo compara las reglas declaradas, con unidades
  iguales, un único punto identificado y facturas que cubran exactamente el
  intervalo de lecturas, sin huecos ni solapes.

No se asume que el agua general equivalga al ACS ni que la energía de gas sea
el consumo de los repartidores. No se convierten unidades ni se prorratean
facturas incompletas. Los intervalos se comparan por sus fechas exactas: facturas
que terminan un día y comienzan al siguiente no se consideran continuas sin
evidencia de cómo interpreta el proveedor sus límites.

Si no se ha declarado una comparación, o faltan esas condiciones, el panel y
la actividad al generar el Excel explican que el consumo no se ha comparado.
Ese mensaje no certifica un cuadre de consumos.

## Registro y protección de datos

Configuración y decisiones se guardan por expediente en `period_parameters`.
Cada aceptación conserva motivo, usuario local, fecha UTC y huella de los datos
revisados. Si cambian facturas, coeficientes, lecturas, perfil o configuración,
la aceptación deja de servir para los datos nuevos. **Actualizar revisión**
recarga el panel; tampoco se permite aceptar datos cambiados desde su apertura.

Antes de guardar configuración, coeficientes o decisiones se crea una copia
verificada de la base. Estos cambios entran en la identidad de las entradas
del expediente y hacen que las salidas anteriores necesiten regenerarse.

## Comprobación

`tests/test_case_coherence.py` usa una base y un perfil sintéticos para verificar
los controles y el bloqueo del exportador. `tests/test_expedient_ui.py` ejecuta
las acciones del panel con widgets simulados: corregir porcentajes y aceptar
un solape con motivo. No modifica comunidades reales ni sustituye una revisión
visual de la pantalla en el equipo del usuario.
