# Detección global fiable de proveedores y documentos

**Fecha:** 2026-09-21

**Estado:** aprobado para planificación

**Ámbito:** fuentes PDF de todas las comunidades y despachos

## 1. Objetivo

Reducir las incidencias manuales identificando de forma fiable el tipo de
documento, el proveedor y los campos de facturación antes de incorporar datos
al expediente. El detector será global: una mejora para un proveedor se aplica
a todas las comunidades, mientras que CUPS, referencias de cliente y otros
datos particulares permanecen asociados a su comunidad.

La automatización no debe aumentar el riesgo financiero. Un valor dudoso se
mantiene como candidato o incidencia; nunca se inventa ni se publica como dato
canónico.

## 2. Línea base auditada

Se revisaron las carpetas reales de correo, versiones anteriores, comunidad
658 y archivos archivados por la aplicación:

| Resultado | Cantidad |
|---|---:|
| Rutas PDF examinadas | 822 |
| PDF únicos por SHA-256 | 661 |
| Reconocidos por los perfiles actuales | 206 |
| Facturas claras sin perfil suficiente | 377 |
| PDF escaneados sin capa de texto | 46 |
| Presupuestos, informes, albaranes u otros | 32 |

Los principales huecos son Tiersan, variantes de Gómez Group Metering,
Echeman, Fenie Energía, Cerrajera Moncasi, Limpiezas Cote, Schindler, JPG
Reparaciones Eléctricas, Limpiezas Utebo, Orona, Moeve, Tritermia y Teleser.
También hay grupos menores recurrentes de limpieza, ascensores, energía,
calderas, cerrajería, protección contra incendios y mantenimiento.

La auditoría demostró además que el nombre mostrado primero en el PDF puede ser
el cliente, Meditrade, una distribuidora o un banco. Por tanto, el nombre de la
primera línea no es una evidencia suficiente del emisor.

## 3. Principios

1. **Identidad fiscal antes que nombre.** El CIF/NIF confirmado del emisor es
   la señal de mayor peso.
2. **Tipo documental antes que extracción.** Un presupuesto o justificante no
   puede convertirse en factura por contener importes.
3. **Identidad y formato son independientes.** Un proveedor puede emitir
   varios formatos históricos sin duplicarse en el catálogo.
4. **Confianza por campo.** Identificar el proveedor no implica que fecha,
   período, importe o servicio sean fiables.
5. **Una lectura por archivo.** Texto y OCR se reutilizan por SHA-256 y versión
   del extractor.
6. **Datos privados locales.** No se enviarán documentos ni texto a servicios
   externos.
7. **Revisión mínima y concreta.** Solo se solicita el campo o decisión que no
   supera las reglas de confianza.

## 4. Arquitectura

### 4.1 Extracción y caché

Una nueva capa `document_text_service` será la única responsable de obtener
texto. Recibirá la ruta y la huella SHA-256 y devolverá:

- texto normalizado;
- método: capa PDF, OCR RapidOCR, OCR alternativo o sin texto;
- páginas procesadas;
- diagnóstico estructurado;
- versión del extractor;
- duración.

La base guardará una caché local por `(sha256, extractor_version)`. La primera
pasada usa como máximo las dos primeras páginas. El OCR se ejecuta únicamente
si no hay texto suficiente o faltan señales esenciales. Una huella ya cacheada
no vuelve a renderizarse mientras no cambie la versión del extractor.

Cada documento tendrá un límite de tiempo y un resultado recuperable. Un PDF
problemático no puede detener el lote completo.

### 4.2 Clasificador de tipo documental

Antes de resolver el proveedor se asigna uno de estos tipos:

- `invoice`;
- `credit_note`;
- `reading`;
- `owners`;
- `quote`;
- `delivery_note`;
- `bank_receipt`;
- `report`;
- `other`;
- `unknown`.

El clasificador utiliza señales positivas y exclusiones. Por ejemplo, un
presupuesto con total no es factura; una nota de entrega de combustible no se
aplica al reparto; un justificante bancario se archiva como soporte; y una
tabla de contadores se envía al importador de lecturas.

Solo `invoice` y `credit_note` pasan al extractor económico. Los demás tipos se
archivan, importan por su flujo especializado o generan una única decisión de
clasificación cuando son ambiguos.

### 4.3 Registro global de proveedores

`config/proveedores.json` seguirá siendo el catálogo único, con un esquema
compatible que admite campos nuevos opcionales:

```json
{
  "tax_ids": ["B00000000"],
  "aliases": ["PROVEEDOR EJEMPLO, S.L."],
  "document_types": ["invoice", "credit_note"],
  "service_family": "MANTENIMIENTO",
  "extractor_family": "standard_spanish_invoice",
  "required_signatures": [],
  "excluded_signatures": [],
  "formats": []
}
```

Los perfiles actuales continúan siendo válidos. La migración del catálogo se
hará de forma incremental y sus validadores rechazarán CIF duplicados,
expresiones inválidas o familias inexistentes.

Los datos propios de una comunidad —CUPS, número de cliente, referencia o CIF
de la comunidad— no definen un proveedor y se mantienen fuera de su identidad
global.

### 4.4 Resolución de proveedor

El resolutor recopila todos los CIF/NIF y nombres candidatos y aplica esta
prioridad:

1. CIF/NIF exacto de un proveedor global y evidencia de documento compatible;
2. firma fuerte requerida y alias normalizado;
3. combinación inequívoca de alias, estructura y familia;
4. proveedor desconocido.

La presencia de un banco, distribuidora, administrador o CIF de comunidad no
debe competir como emisor. Si dos proveedores globales obtienen la misma
evidencia, el resultado es ambiguo y se solicita una única confirmación.

El resultado incluye proveedor, confianza, señales que justifican la decisión
y fragmento/localizador. Las correcciones manuales confirmadas crean un alias
o propuesta revisable; nunca modifican silenciosamente el catálogo.

### 4.5 Familias de extracción

Las reglas compartidas se concentran en extractores reutilizables:

- factura española estándar;
- electricidad;
- gas y combustible;
- limpieza y mantenimiento periódico;
- ascensores;
- calderas y climatización;
- contadores y gestión energética;
- agua y tasas públicas.

Cada familia extrae número, fecha de factura, inicio y fin del servicio, base,
IVA, total, consumo, unidad, CUPS o referencia cuando estén presentes. Un
perfil de proveedor puede declarar formatos concretos que sobrescriben solo
los campos diferentes.

Los importes deben reconciliar base, impuestos y total con tolerancia de
céntimos. Las fechas se validan semánticamente y el período no se deduce de un
número sin etiqueta. Para mantenimiento sin intervalo explícito podrá usarse
la fecha de factura únicamente cuando el perfil lo declare.

### 4.6 Confianza por campo

Cada candidato guardará:

- valor normalizado;
- confianza `high`, `medium` o `low`;
- fuente: texto, OCR, nombre de archivo o corrección manual;
- página, fragmento o celda;
- regla y versión que lo produjo.

Un documento se aplica automáticamente solo si el tipo, proveedor y todos los
campos obligatorios tienen confianza alta y superan sus validaciones. La
confianza media reduce la revisión al campo concreto. La confianza baja no
alimenta tablas canónicas.

### 4.7 Elegibilidad para el reparto

Reconocer una factura y su proveedor no autoriza por sí solo a incluirla en
una regularización. Después de extraerla se aplica una barrera independiente
basada en el perfil de la comunidad, el período del expediente y sus conceptos
activos.

Para entrar en el reparto, el documento debe cumplir conjuntamente:

- comunidad identificada con confianza alta;
- período de servicio compatible con el expediente;
- familia de servicio vinculada a un concepto activo de esa comunidad;
- estado documental apto (`invoice` o `credit_note`);
- campos económicos obligatorios validados.

Una factura válida pero ajena a los conceptos activos se conserva en el
histórico documental con el motivo `not_applicable_to_case`; no genera gasto,
consumo ni carta. Si la comunidad o el concepto no son inequívocos, se crea una
sola decisión agrupada y el documento queda sin aplicar hasta resolverla.

### 4.8 Incidencias y aprendizaje controlado

Las incidencias automáticas se agrupan por documento y decisión. No se crearán
tres incidencias de fechas e importe cuando la decisión real sea “identificar
el proveedor”. Después de resolverla se reanaliza el documento y se muestran
solo los campos que continúen faltando.

Una corrección de proveedor queda auditada con el CIF, alias, documento y
persona responsable. La interfaz podrá ofrecer “proponer para el catálogo
global”, pero la incorporación exige validación y pruebas; una respuesta
manual aislada no modifica reglas para todas las comunidades.

## 5. Flujo de datos

```text
archivo
  -> SHA-256 y deduplicación
  -> texto cacheado o extracción/OCR
  -> clasificación documental
  -> resolución global del proveedor
  -> familia + formato de extracción
  -> validación y confianza por campo
  -> elegibilidad: comunidad + período + concepto activo
  -> aplicación automática segura
       o incidencia concreta y agrupada
  -> datos canónicos -> Excel -> reparto -> cartas
```

La bandeja global resuelve primero comunidad y período. Un documento que
declara otra comunidad se aparta antes de llegar al expediente seleccionado.

## 6. Rendimiento y experiencia de uso

- Deduplicación antes de abrir el PDF.
- Catálogo cargado una vez por lote.
- Texto y OCR cacheados por huella.
- Extracción inicial limitada a portada/dos primeras páginas.
- OCR selectivo, cancelable y con límite por archivo.
- Progreso por fases: huella, texto, tipo, proveedor y campos.
- Procesamiento de archivos independientes en paralelo con un límite seguro.
- La interfaz pagina y agrupa incidencias; no crea cientos de controles a la
  vez.

## 7. Despliegue por etapas

### Etapa 1 — Infraestructura segura

- esquema ampliado del catálogo;
- resolución por CIF/NIF;
- clasificación documental previa;
- caché de texto/OCR y métricas;
- compatibilidad con los 16 perfiles actuales.

### Etapa 2 — Mayor reducción de incidencias

Perfiles y pruebas para Tiersan, Gómez Group Metering, Echeman, Fenie Energía,
Cerrajera Moncasi, Limpiezas Cote, Schindler, JPG Reparaciones Eléctricas,
Limpiezas Utebo y Orona.

### Etapa 3 — Familias recurrentes

Moeve, Tritermia, Teleser, ISS, Laboil, ista, Ascensores Sales y los grupos de
limpieza, mantenimiento, protección contra incendios y cerrajería que aparecen
entre tres y seis veces.

### Etapa 4 — Cola residual

- proveedores de una o dos muestras;
- escaneos difíciles;
- documentos sin emisor visible;
- revisión de falsos positivos y tiempos.

## 8. Pruebas y corpus

Los documentos privados no se incorporarán al repositorio. Las pruebas usarán
fragmentos sintéticos anonimizados que conserven estructura, etiquetas y
variantes OCR relevantes.

Habrá pruebas de:

- CIF del proveedor frente a CIF del cliente;
- banco y distribuidora mencionados sin convertirse en emisor;
- varios formatos del mismo proveedor;
- factura, abono, presupuesto, albarán, justificante, informe y lectura;
- importes españoles, negativos, puntos de millar y OCR defectuoso;
- períodos explícitos, fecha única y fechas ambiguas;
- caché por huella y cambio de versión;
- tiempo máximo y continuidad del lote;
- reanálisis sin perder correcciones manuales;
- separación de comunidades;
- aplicación automática solo con confianza alta.

Un verificador privado, ejecutado localmente, medirá el catálogo contra la
muestra real sin guardar texto ni datos personales en Git.

## 9. Criterios de aceptación

1. Reconocer proveedor y tipo en al menos el 90 % de las facturas con texto de
   la muestra actual.
2. Reducir las 377 facturas sin perfil suficiente a menos de 60 decisiones de
   proveedor.
3. Cero presupuestos, albaranes, informes o justificantes aplicados como
   factura en el conjunto revisado.
4. Cero documentos de otra comunidad aplicados al expediente seleccionado.
5. Un archivo cacheado no repite OCR con la misma versión del extractor.
6. Un PDF problemático no bloquea el resto del lote.
7. Ningún valor de baja confianza entra en facturas, lecturas, reparto o
   cartas.
8. Toda corrección automática o humana conserva trazabilidad.
9. Una factura reconocida pero ajena a los conceptos activos se archiva sin
   modificar el reparto.
10. Las pruebas existentes y las nuevas pruebas del detector pasan juntas.

## 10. Fuera de alcance

- envío automático de cartas por correo;
- OCR o clasificación mediante servicios externos;
- aprendizaje autónomo que cambie reglas globales sin aprobación;
- extracción perfecta de documentos ilegibles o sin datos obligatorios;
- migración masiva de otros despachos en esta entrega.

## 11. Riesgos y mitigaciones

- **CIF del cliente confundido con emisor:** resolver únicamente contra CIF
  fiscales registrados y evidencia de tipo compatible.
- **Proveedor cambia de formato:** mantener formatos versionados bajo la misma
  identidad.
- **OCR incorrecto:** confianza por campo, reconciliación económica y revisión
  cuando falle.
- **Catálogo demasiado amplio:** firmas fuertes, exclusiones y pruebas de
  colisión.
- **Base de datos creciente:** guardar una sola extracción por huella y versión,
  con política futura de mantenimiento sin borrar auditoría.
- **Datos privados en pruebas:** corpus sintético en Git y verificador real
  exclusivamente local.
