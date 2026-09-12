# Ingesta unificada de fuentes y reinicio seguro de base de datos

## Objetivo

Hacer que un expediente pueda iniciarse desde una base vacía con una carpeta
mixta de facturas, lecturas de contador y archivos auxiliares. Cada archivo se
clasificará y extraerá antes de crear incidencias. Los datos confirmados
alimentarán las tablas canónicas que generan el Excel oficial, el reparto y las
cartas.

La comunidad 658 será la primera validación, pero no tendrá reglas codificadas
en el producto: la solución será reutilizable para cualquier comunidad y
despacho.

## Problema confirmado

La carga guiada actual registra cada archivo con el tipo elegido por el usuario
sin ejecutar el extractor. Cuando se elige `invoice` para una carpeta, todos
los documentos reciben los requisitos de factura (`fecha_inicio`, `fecha_fin`,
`importe_total`). En la base actual había 68 documentos marcados como facturas,
por lo que se produjeron exactamente 204 incidencias repetidas, incluso para
lecturas XLS/XLSX/CSV y PDF.

El parser que reconoce facturas y lecturas existe en un flujo heredado distinto
del expediente. La corrección debe unificar ambas rutas y no limitarse a ocultar
incidencias.

## Seguridad y base limpia

La sustitución de la base será una operación explícita y segura:

1. Cerrar las conexiones activas a `data/gestion.db`.
2. Crear `data/backups/gestion-YYYYMMDD-HHMMSS.db` como copia íntegra.
3. Verificar que la copia SQLite se puede abrir y contiene el esquema esperado.
4. Inicializar una nueva `data/gestion.db` mediante las migraciones vigentes.
5. Registrar la fecha y ruta de la copia en la actividad de la aplicación.
6. Si una fase falla, mantener la base original intacta y explicar el error.

Los archivos fuente no se eliminan: cada expediente conserva un archivo propio
en `data/expedientes/<caso>/fuentes`. La nueva base puede empezar sin
comunidades, facturas ni lecturas; las fuentes archivadas y la copia de la base
anterior permiten recuperar y comparar información.

## Arquitectura

Se incorporará un orquestador de ingesta de expediente, separado de la UI y de
la persistencia:

```text
archivo → clasificador → extractor específico → candidatos + contexto
        → revisión, si hace falta → tablas canónicas → salidas
```

### Clasificador

Recibe una ruta archivada y devuelve una categoría y una confianza:

- `invoice`: factura de proveedor PDF con extracción normalizada.
- `reading`: lectura de contadores PDF, XLS, XLSX o CSV.
- `owners`: listado de propietarios o coeficientes.
- `other`: documento auxiliar que puede conservarse sin bloquear.
- `unknown`: documento que no puede clasificarse con seguridad.

Los PDF reutilizan el lector existente de proveedores y lecturas. Los formatos
tabulares inspeccionan sus cabeceras y estructura antes de clasificarse. La
elección manual seguirá existiendo como corrección, pero no será el origen
normal de la clasificación.

### Extracción y persistencia

Cada extractor devuelve datos normalizados, campos que son obligatorios para
esa categoría y contexto de origen. La aplicación:

- registra los candidatos junto con el documento;
- crea incidencias únicamente para campos obligatorios de la categoría
  detectada que estén realmente ausentes;
- incorpora datos `confirmado` a las tablas de facturas, lecturas, propietarios
  o conceptos;
- deja datos dudosos en `pendiente_revision`, sin sustituirlos por cero ni
  inventar consumos;
- permite `other` sin bloquear cuando no afecta a módulos activos;
- crea para `unknown` una sola incidencia de clasificación por documento.

Los contadores que retroceden se mantienen como una incidencia de reinicio:
se conserva la lectura previa hasta que exista una lectura posterior válida o
el usuario confirme una estimación documentada.

### Reanálisis idempotente

Cada expediente tendrá una acción **Reanalizar fuentes**. Recorre sus archivos
archivados, recalcula clasificación y candidatos y reemplaza únicamente las
incidencias generadas automáticamente que sigan abiertas. Nunca sobrescribe
una corrección manual ni datos ya confirmados. Permitirá recuperar el expediente
658: las 204 incidencias genéricas se sustituirán por resultados reales.

## Experiencia de usuario

Después de añadir archivos o carpeta, la pantalla presenta un resumen:

```text
12 facturas detectadas · 4 lecturas · 1 listado de propietarios
2 documentos requieren revisión
```

La bandeja de incidencias se agrupa por documento y prioridad. Una incidencia
describe el dato concreto, el motivo y el impacto en el Excel/reparto. No se
repiten tres campos de factura en documentos que no son facturas.

Cada incidencia ofrece:

- **Ver contexto** cuando el extractor conoce una página, hoja, celda o
  fragmento relevante.
- **Abrir archivo** como alternativa universal.
- **Resolver** con validación del tipo de dato y motivo de corrección.

Para PDF se almacenará página y fragmento cuando el parser los conozca. La UI
mostrará esa referencia; abrir directamente una página dependerá del visor de
Windows, por lo que no será requisito bloqueante. En Excel se mostrará, cuando
exista, hoja y celda/rango.

El Excel oficial solo se habilita cuando las fuentes requeridas por los módulos
activos están confirmadas. Reparto y cartas permanecen detrás de ese control.

## Pruebas y criterios de aceptación

Pruebas automatizadas cubrirán:

1. Clasificación de factura PDF, lectura PDF, lectura tabular, auxiliar y
   documento desconocido.
2. Una lectura nunca recibe los requisitos de una factura.
3. Un documento desconocido crea una sola incidencia de clasificación.
4. Reanálisis de 68 fuentes no genera 204 incidencias genéricas.
5. Reanálisis no borra una corrección manual confirmada.
6. Reinicio seguro crea una copia verificable antes de crear la nueva base.
7. Contexto de origen se conserva en una incidencia cuando está disponible.

La aceptación manual será una base nueva, una comunidad creada desde la
aplicación y una carpeta mixta de fuentes de la 658. El resultado debe mostrar
un resumen por tipo; solo los documentos realmente ambiguos deben requerir
intervención. Tras resolverlos, el programa debe generar el Excel oficial, el
reparto final y las cartas sin depender de un Excel manual previo.

## Fuera de alcance inmediato

- OCR avanzado para PDFs puramente escaneados sin texto.
- Garantizar navegación automática a una página concreta en todos los visores
  de PDF de Windows.
- Migrar automáticamente datos dudosos desde la base antigua: la copia queda
  disponible para comparación y recuperación controlada.
