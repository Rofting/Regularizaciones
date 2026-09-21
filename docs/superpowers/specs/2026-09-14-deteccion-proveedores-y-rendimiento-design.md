# Detección de proveedores y rendimiento de fuentes

## Objetivo

Reducir las revisiones manuales de fuentes del expediente sin convertir datos
financieros dudosos en datos canónicos. El caso de referencia es la comunidad
658: once lecturas se identifican correctamente y cincuenta y cinco PDFs de
facturas quedan como tipo desconocido porque el catálogo actual no reconoce sus
proveedores.

El flujo debe seguir siendo portable: una factura nueva de un proveedor no
catalogado ha de clasificarse al menos como factura, y un proveedor se mejora
una sola vez en el catálogo para que sirva a todas las comunidades posteriores.

## Diseño

### Catálogo reutilizable de proveedores

`config/proveedores.json` seguirá siendo el único catálogo editable de
proveedores. Cada perfil contiene firmas textuales, tipo de suministro y reglas
de extracción; no se duplicarán perfiles por comunidad. Los datos específicos
de cada comunidad (como CUPS permitidos) permanecen en el bloque de la
comunidad dentro del perfil.

Se incorporarán perfiles para los formatos que aparecen en 658 y que puedan
reutilizarse: Mantenimientos Zaragoza, Naturgy Clientes y los demás formatos
confirmados durante la muestra. Cada perfil deberá declarar evidencia textual
inequívoca, no sólo una coincidencia de nombre de archivo.

### Clasificación escalonada

El análisis de PDF conservará la salida detallada del lector (`motivo`,
`detalle`, proveedor y localizador) en vez de degradarla directamente a
`unknown`.

1. Se intenta un perfil del catálogo y sus extracciones.
2. Si no hay perfil, un clasificador genérico busca evidencia conjunta de
   factura: una cabecera de factura, identificador o fecha y un total o bloque
   de importes. Entonces crea una fuente `invoice` con confianza media y sólo
   propone los campos realmente presentes.
3. Si sólo hay evidencia de lectura o de propietarios se mantiene su clasificador
   especializado.
4. Sin evidencia suficiente, permanece `unknown` y se crea una única incidencia
   de clasificación, con el motivo técnico y el fragmento del documento.

La clasificación genérica nunca asigna automáticamente tipo de suministro,
conceptos de reparto, consumo ni importe cuando el texto no lo contiene de
forma inequívoca. Esos campos seguirán solicitándose de forma concreta.

### Rendimiento e interfaz

La configuración de proveedores se cargará una vez por lote de análisis, no una
vez por PDF. El análisis seguirá fuera del hilo de interfaz y publicará avances
por documento.

La pantalla de incidencias mostrará primero un resumen agrupado y una página de
un máximo de diez elementos. Incluirá navegación de páginas y filtros por tipo
de incidencia. Así se evita construir decenas de tarjetas CustomTkinter a la
vez y se mantiene visible el número total pendiente.

`Reanalizar fuentes` volverá a calcular sólo las incidencias automáticas. Las
clasificaciones y correcciones confirmadas por el usuario permanecen intactas.

## Flujo resultante

```text
PDF/XLS/CSV
  -> extracción de texto o tabla
  -> perfil global de proveedor
       -> datos candidatos + revisión puntual
  -> sin perfil: clasificador genérico seguro
       -> factura reconocida + sólo campos que falten
  -> sin evidencia: incidencia de clasificación con contexto
  -> datos confirmados -> tablas canónicas -> Excel -> reparto -> cartas
```

## Errores y seguridad

- Un PDF corrupto, escaneado sin texto o ambiguo no se acepta como factura.
- Un archivo ya archivado se recupera sólo si coincide su hash, conforme al
  comportamiento de reintento ya implantado.
- Las incidencias conservan ruta archivada, fragmento y localizador cuando estén
  disponibles.
- Una factura fuera del periodo o con CUPS incompatible se conserva para
  revisión; no se aplica al reparto.

## Pruebas

- Pruebas unitarias de perfiles y del clasificador genérico: factura conocida,
  factura de proveedor nuevo, lectura, propietarios y documento ambiguo.
- Prueba de caché: un lote de N PDFs no vuelve a leer el catálogo N veces.
- Prueba de reanálisis: elimina incidencias automáticas obsoletas y conserva
  una resolución manual.
- Prueba de paginación: con 55 incidencias, la primera vista contiene sólo diez
  tarjetas y el total sigue siendo visible.
- Prueba de regresión de comunidad 658 con muestras anonimizadas por patrones,
  sin incorporar archivos privados al repositorio.
