# Núcleo configurable de regularizaciones

Fecha: 30 de agosto de 2026

## 1. Objetivo y relación con el diseño anterior

Este documento evoluciona el diseño de `2026-08-29-regularizacion-facturas-design.md`. Conserva su objetivo —datos trazables, cálculo al céntimo, Excel y cartas— y define el núcleo operativo que faltaba: un expediente de regularización configurable, con revisión humana obligatoria de documentos inciertos y períodos de fechas libres.

La comunidad 658 es el primer caso de aceptación. No es una excepción de código ni una plantilla rígida. La misma arquitectura debe admitir después 644 y el resto de comunidades, y convertirse en una instalación portable para otros despachos.

El primer envío de este producto termina en Word revisable; no envía correo electrónico automáticamente.

## 2. Principios no negociables

- La base de datos es la fuente de verdad; los PDF, Excel y Word se guardan como evidencia o salida.
- Un dato ausente, ambiguo, incoherente o de baja confianza crea una incidencia. No se usa para calcular mientras no sea confirmado.
- Resolver una incidencia exige ver el archivo de origen y confirmar el campo concreto en la app. No existe un valor silencioso por defecto.
- Cada corrección mantiene valor original, valor confirmado, origen, motivo, instante y usuario local responsable.
- Una comunidad configura servicios, componentes y reglas. No se codifican ramas por código de comunidad.
- Un período siempre tiene fecha inicial y final exactas, sin asumir un año natural ni una duración fija.
- Facturas, gastos, lecturas, resultados, Excel y cartas pertenecen a un expediente y a su período.
- Un descuadre económico bloquea la generación de Excel y cartas.
- Marca, datos de contacto, plantillas, rutas y proveedores son configurables por instalación; ninguna salida depende de Meditrade.

## 3. Expediente de regularización

Un expediente representa una regularización de una comunidad para un rango de fechas. Su ciclo de vida es:

```text
Borrador → Reuniendo fuentes → En revisión → Listo para cálculo
        → Cuadrado → Entregas generadas → Cerrado
```

El sistema no permite saltar etapas. Volver a importar, editar una corrección o cambiar el rango devuelve el expediente a la etapa que requiera revalidación.

Cada expediente almacena:

- comunidad, nombre opcional y fecha inicial/final;
- estado y fechas de creación, cálculo y cierre;
- configuración de comunidad congelada al calcular, para poder reproducir el resultado;
- documentos fuente y sus huellas SHA-256;
- extracciones candidatas, incidencias y correcciones manuales;
- lotes de cálculo, conciliaciones, Excel y cartas generados.

La interfaz trabaja sobre un expediente seleccionado; no procesa una carpeta global sin contexto de comunidad y fechas.

## 4. Configuración de despacho, comunidad y conceptos

### 4.1 Despacho

Una instalación contiene un perfil de despacho: nombre mostrado, logo opcional, dirección, correo/teléfono, firma de cartas, paleta, plantilla Word, plantilla Excel, carpetas y política de copia de seguridad. Si falta un dato, se usa una identidad neutra de «Regularización de facturas», nunca una marca de otro despacho.

### 4.2 Comunidad

Una comunidad define código, denominación, CIF, dirección, propiedades, suministros, identificadores como CUPS y su catálogo de componentes. La app tendrá un asistente de alta y una pantalla de edición; ambas guardan configuración, no constantes Python.

### 4.3 Servicio y componente

Un **servicio** describe el suministro: ACS, calefacción, electricidad, agua u otro nombre configurado. Un **componente** es una partida repartible de ese servicio. No se limita a una sola cuota por servicio.

Ejemplos válidos:

```text
ACS / cuota fija / por vivienda
ACS / cuota variable / por consumo m³
Calefacción / cuota fija / por vivienda
Calefacción / cuota variable / por consumo o coeficiente
Calefacción / gasto extraordinario / por coeficiente
Electricidad / gasto común / por coeficiente
```

Cada componente configurable contiene, como mínimo: clave estable, etiqueta de carta, servicio, clase (`fijo`, `variable`, `extraordinario`, `ajuste`, `abono` u `otro`), unidad opcional, regla de reparto, orden de presentación, si requiere lecturas y si aparece por defecto en cartas. Las reglas iniciales son por vivienda, partes iguales, coeficiente, consumo y porcentaje fijo. Nuevas reglas se añaden como estrategias del motor, no como excepciones de plantilla.

## 5. Fechas y prorrateo

La persona usuaria elige dos fechas y un nombre opcional. La app calcula duración, pero no restringe el rango: admite dos, seis, nueve, doce meses u otro intervalo.

Para cada documento económico:

1. se usa su intervalo de suministro/gasto si está disponible;
2. si el intervalo no solapa el expediente, el documento queda fuera;
3. si solapa parcialmente, se distribuye de forma proporcional por días naturales: `importe del expediente = importe × días solapados / días del documento`;
4. se conserva el importe original, el intervalo fuente, días totales, días incluidos y el importe prorrateado;
5. si faltan fechas, el documento crea una incidencia que pide fecha inicial y final o la confirmación de que su importe completo corresponde al expediente.

Para componentes por consumo, el período requiere una lectura inicial y otra final asociadas al servicio. Si no son exactas, la app crea una incidencia donde se confirma la lectura o la regla de estimación. Un descenso aislado de contador conserva la lectura anterior como estimada y deja la propiedad pendiente de lectura posterior. Un reinicio global se registra como tal y se procesa con la regla explícita del servicio.

## 6. Documentos, extracción e incidencias

Cada archivo añadido se conserva en una carpeta del expediente o se referencia de forma inmutable, con nombre original, hash, tipo y fecha. El lector entrega candidatos con valor, confianza, evidencia y página/celda cuando se conoce.

Una incidencia posee: código, severidad, estado, documento, campo, valor detectado, explicación, formulario de resolución y auditoría. Las causas iniciales cubren proveedor desconocido, PDF sin texto, fechas de suministro faltantes, importe incoherente, comunidad/CUPS no coincidente, propiedad sin correspondencia, lectura anómala y regla de reparto sin datos.

El flujo de revisión es obligatorio:

1. El usuario abre la incidencia.
2. La app abre o previsualiza el archivo fuente; para PDF muestra la página relevante cuando exista esa referencia.
3. Muestra solo los campos requeridos, con el valor detectado y la explicación.
4. El usuario confirma, corrige o marca el documento como no aplicable.
5. La app guarda una corrección auditable, ejecuta las validaciones afectadas y actualiza la cola.

Las correcciones repetidas pueden producir una **sugerencia** de regla para un proveedor/formato. Ninguna sugerencia se activa sola: hay que aceptarla en configuración y se conserva su versión.

## 7. Pantalla única del expediente

La pantalla principal pasa a tener seis etapas visibles y permanentes:

```text
1. Configurar comunidad y fechas
2. Añadir fuentes
3. Resolver incidencias
4. Revisar Excel y reparto
5. Generar cartas
6. Cerrar expediente
```

La primera tarjeta muestra comunidad, fechas, duración, número de propiedades y estado. «Nueva comunidad» y «Nuevo período» se sustituyen por acciones claras: «Crear comunidad» y «Crear expediente».

La segunda muestra archivos añadidos, tipo detectado, servicio propuesto, resultado y progreso. La tercera es una bandeja filtrable de incidencias. Cada fila abre su fuente y formulario de corrección; el contador de pendientes impide avanzar.

La cuarta ofrece resumen de costes originales/prorrateados, resultados por componente y conciliación al céntimo. El Excel es una vista previa y un archivo exportable. La quinta permite seleccionar componentes para las cartas y previsualizar una carta antes de crear todas. La sexta registra la salida y protege el expediente contra cambios accidentales; reabrirlo exige confirmar y deja versión de la entrega anterior.

## 8. Cálculo, Excel y cartas

El motor toma únicamente documentos validados del expediente. Agrupa importes por componente, aplica prorrateo cuando proceda, distribuye por su regla y guarda por propiedad: cobrado, coste real, diferencia, consumo y referencias de origen.

El sistema concilia por componente y total. Acepta una diferencia máxima configurable de un céntimo por redondeo; cualquier exceso bloquea salida.

El generador Excel recibe el expediente calculado y un adaptador de plantilla. El primer adaptador reproduce el modelo validado de la 658 y conserva formato, fórmulas de presentación y hojas aplicables. Los datos no dependen de celdas fijas: el adaptador ubica filas y componentes mediante etiquetas/configuración. Futuras comunidades podrán elegir otro adaptador sin cambiar el motor.

Las cartas Word continúan siendo de una página. El usuario elige los componentes; los que no tengan resultado o no estén seleccionados se omiten. La carta usa la identidad del despacho, una tabla dinámica de componentes y las dos gráficas aprobadas: franja de consumo frente a vecinos e histórico real de la propiedad. Si no hay datos suficientes para una gráfica, se adapta el espacio sin inventarlos. Una generación que no cabe en una página se bloquea y explica qué debe revisarse.

## 9. Portabilidad y operación

La distribución separa código, configuración de instalación, datos del despacho, base de datos, plantillas y archivos de trabajo. La primera ejecución crea una configuración neutra y permite elegir carpetas. Las copias de seguridad se guardan antes de regenerar Excel o sobrescribir una entrega.

No se incluye todavía envío de correo, conexión a un proveedor externo ni migración masiva de todos los despachos. La salida Word queda revisable y lista para el envío manual ya acordado.

## 10. Entregas y orden de implementación

### Entrega A — Expediente e incidencias (primera entrega utilizable)

- Migraciones para expediente, documento fuente, candidatos, incidencias y correcciones.
- Servicio de validación y cola de incidencias sin interfaz dependiente.
- Importación de archivo con hash, deduplicación y auditoría.
- Pantalla de expediente con fechas libres y bandeja de incidencias.
- Formulario manual que persiste una corrección y revalida.

**Aceptación:** una factura o lectura con un campo requerido ausente bloquea el cálculo, se abre desde la app, se corrige manualmente y pasa a validada sin perder su original.

### Entrega B — Componentes y períodos configurables

- Catálogo de servicios/componentes por comunidad, incluidas cuota fija y variable de calefacción.
- Motor de reglas por vivienda, iguales, coeficiente y consumo.
- Selección de fechas, solape y prorrateo por días auditables.
- Cálculo y conciliación dinámicos.

**Aceptación:** la 658 funciona como ACS fijo/variable y una configuración de prueba incluye calefacción fija/variable y gasto extraordinario sin condicional de comunidad.

### Entrega C — Excel de salida automático

- Adaptador de la plantilla 658 desde datos validados.
- Comparación de los campos de control con los Excel de referencia.
- Copia de seguridad y salida por expediente.

**Aceptación:** con fuentes validadas de la 658, se produce un Excel revisable que cuadra al céntimo sin rellenarlo manualmente.

### Entrega D — Cartas, cierre y perfil portable

- Cartas configuradas por despacho/comunidad, selección de componentes y previsualización.
- Cierre/versionado de expediente.
- Perfil de despacho y asistente de alta de comunidad.

**Aceptación:** otra comunidad puede configurarse sin que aparezca Meditrade ni se modifique el código del motor.

### Entrega E — Ampliación documental y validación 644

- Extractores y reglas por proveedores adicionales; OCR y cuarentena para escaneados.
- Adaptador/validación con la comunidad 644 y sus conceptos reales.
- Pruebas de regresión contra 658 y 644.

## 11. Estrategia de pruebas

- Cada servicio de dominio se desarrolla con una prueba que falla antes de su implementación.
- Las pruebas usan bases temporales y documentos sintéticos; no versionan datos personales ni facturas reales.
- Se cubren fechas sin solape, solape parcial, límites de día, importes negativos, componentes sin pesos, correcciones manuales y reintentos idempotentes.
- Los casos de aceptación reales de 658 y 644 se ejecutan fuera de Git y solo conservan conteos, diferencias y resultados de QA sin datos personales.
- Las cartas se renderizan con LibreOffice y se verifica una única página.

## 12. Límites explícitos de este ciclo

Este ciclo implementa primero la Entrega A y planifica las Entregas B–E. No se añade envío automático de correo ni se inicia una migración masiva hasta que los expedientes de 658 y 644 pasen sus validaciones.
