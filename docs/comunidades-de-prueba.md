# Comunidades de prueba

`Regularizaciones_prueba.bat` abre la aplicación sobre una instalación aparte,
en `%LOCALAPPDATA%\Regularizaciones-prueba`. Tiene su propia base de datos:
la del despacho no se abre ni se modifica. La primera vez la crea
`scripts/crear_comunidades_prueba.py` (alrededor de un minuto). Con
`Regularizaciones_prueba.bat nueva` se borra y se vuelve a crear desde cero.

En el título de la ventana aparece «PRUEBA · datos inventados». Todos los
nombres, CIF, CUPS y correos son inventados; los correos usan el dominio
reservado `example.invalid`, así que nunca llegan a nadie.

| Comunidad | Para qué sirve |
| --- | --- |
| **901 · Lista para calcular** | Facturas de gas, luz y agua, mantenimiento, servicio de lecturas, propietarios, lecturas de ACS, cuotas cobradas y gastos fijos ya incorporados con «Añadir fuentes» y confirmados. Pulsa «Generar Excel», «Calcular reparto» y «Generar cartas». |
| **902 · Con incidencias** | Los mismos documentos y tres problemas habituales: una factura del ejercicio anterior, una reparación cuyo importe no se pudo leer y una factura de agua de un proveedor que el catálogo no conoce. Además lleva una copia repetida de una factura, que se descarta sola. |
| **903 · Bandeja de fuentes** | Expediente vacío. En `%LOCALAPPDATA%\Regularizaciones-prueba\fuentes_903` hay 56 documentos para probar «Añadir fuentes» de principio a fin. Después queda lista para calcular. |

Los PDF son de texto (no escaneados) e imitan los formatos de Naturgy,
Endesa, Mantenimientos Zaragoza y Gómez Group Metering del catálogo. El
proveedor de agua «Aguas de Prueba» se declara en
`config/proveedores_despacho.json` de la instalación de prueba, igual que el
despacho añadiría uno propio.

Al construir estas comunidades aparecieron fallos reales que también afectan
a las comunidades del despacho. Están corregidos:

- Una comunidad nueva, sin perfil Excel todavía, aceptaba facturas de otro
  ejercicio sin comprobar su período. Ahora se usan los módulos del modelo
  común.
- Una factura con total 0,00 € (importe ilegible) se aplicaba sin avisar.
  Ahora abre una incidencia.
- «Consumo Total 980,00 kWh» se tomaba como importe total de la factura.
- Con cuotas cobradas que incluyen el consumo de cada liquidación, el m³ del
  período se contaba dos veces en «LECTURAS ACS M3».

`tests/test_comunidades_prueba.py` crea la instalación en una carpeta
temporal y comprueba el estado de cada comunidad.
