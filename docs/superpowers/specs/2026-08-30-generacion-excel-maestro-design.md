# Diseño: Excel maestro, reparto final y cartas

**Fecha:** 2026-08-30  
**Estado:** aprobado para especificación; pendiente de revisión del documento  
**Primer caso real:** comunidad 658, periodo 2025-2026

## Objetivo

Generar el Excel oficial de una comunidad únicamente a partir de datos validados en la base de datos, conservando la estructura, fórmulas, estilo, impresión y hojas del modelo que el despacho ya considera correcto. A continuación, calcular el reparto final por propietario y generar una carta Word individual para cada propietario.

El primer resultado debe reproducir el libro de referencia de la comunidad 658, conciliar los importes repartidos y producir sus cartas. El diseño debe permitir que una segunda comunidad con el mismo patrón, como la 644, active módulos adicionales sin duplicar el motor.

## Material de referencia

- Modelo oficial de 658: <private-658-master-workbook>.
- Modelo de comparación de 644: <private-644-master-workbook>.

La 658 contiene las hojas DATOS, GAS, ELECTRICIDAD, AGUA, OTROS GASTOS, LECTURAS ACS M3 y ANALISIS. La 644 comparte ese núcleo y añade LECTURAS CALEF KWH, varios análisis anuales y RESUMEN.

## Decisiones

### 1. Plantillas maestras inmutables y versionadas

Cada comunidad tendrá una plantilla maestra en el proyecto, identificada por código de comunidad, perfil y hash SHA-256. La plantilla nunca se modifica al generar un resultado.

Una exportación:

1. verifica que la plantilla registrada coincide con su hash;
2. crea una copia temporal;
3. rellena únicamente los rangos de entrada definidos para el perfil;
4. conserva fórmulas, estilos, bordes, anchos, áreas de impresión y hojas;
5. valida el resultado; y
6. lo publica de forma atómica, con copia del Excel oficial anterior en Excels_Maestros/backups/.

La primera plantilla registrada será una copia preservada del Excel real de la 658. Las futuras modificaciones conscientes de diseño se registrarán como una nueva versión, nunca sobrescribiendo la anterior.

### 2. Perfil de Excel por comunidad, no lógica repetida

Un perfil describe cómo la base de datos se representa en un libro. Contendrá:

- plantilla y versión;
- módulos activos;
- hojas obligatorias;
- rangos y anclas semánticas;
- campos obligatorios y opcionales;
- reglas de formato y de fórmulas que debe conservar el libro;
- comprobaciones de conciliación.

El perfil 658_acs_v1 activa GAS, ELECTRICIDAD, AGUA, otros gastos y ACS. No activa calefacción. El perfil de la 644 activa, además, calefacción y sus análisis/resumen.

Los perfiles vivirán como configuración versionada del proyecto y se registrarán en la base de datos junto con el hash de la plantilla usada por cada salida. Así el motor sigue siendo único y otras comunidades o despachos añaden configuración y plantillas, no código específico.

### 3. Datos normalizados y datos de plantilla

Las tablas existentes ya cubren comunidad, periodo, facturas, lecturas, gastos, configuraciones y reparto. Se ampliarán únicamente donde la plantilla requiere desglose que una factura resumida no puede expresar:

- componentes de factura: por ejemplo, agua variable/fija de cada proveedor, conceptos de gas, impuestos y abonos;
- parámetros del periodo: cuota fija/variable, tarifas, reparto ACS/calefacción y otras reglas activas;
- ejecución de exportación: perfil, plantilla, caso/periodo, hash de las entradas, estado, ruta y errores;
- vínculo de las celdas de origen con el dato normalizado para auditoría.

Una celda no se rellenará con un valor inventado. Si un campo obligatorio de un módulo activo no está en la base, se creará una incidencia de revisión y la exportación oficial se bloqueará. Los módulos inactivos conservarán su estructura de plantilla sin forzar importes a cero.

### 4. Arranque de la 658 desde su Excel real

La base instalada no contiene aún las facturas de la 658, mientras que el Excel 2025-2026 es el dato real más reciente. Por ello se añadirá una importación inicial de Excel maestro que:

- registra el archivo como fuente auditada;
- lee los bloques de GAS, ELECTRICIDAD, AGUA, otros gastos y lecturas ACS por etiquetas/anclas, no por posiciones ciegas;
- guarda facturas, componentes, parámetros, gastos y lecturas en las tablas normalizadas;
- conserva hoja y celda de origen para cada dato;
- detecta datos ambiguos, vacíos o incompatibles y abre incidencias manuales en vez de interpretarlos silenciosamente;
- es idempotente por hash: volver a importar el mismo libro no duplica datos.

Esta importación es de transición. En el flujo final, los PDFs y las lecturas alimentarán las mismas tablas normalizadas y el generador no dependerá del Excel original.

El estudio de la 658 contiene costes y parámetros agregados, pero no las lecturas y la identificación de cada propietario. Para poder completar el reparto y las cartas de esta primera ejecución, el mismo arranque registrará también el listado de propietarios y los ficheros de lecturas asociados al periodo. Todos quedan vinculados como fuentes del expediente. Si falta una lectura de inicio o cierre, o no coincide con el rango del expediente, se abrirá una incidencia y no habrá reparto ni cartas.

### 5. Generación y controles previos

La acción de interfaz será Generar Excel oficial. Sólo estará disponible para un expediente de la comunidad que esté ready_for_calculation y con todos sus datos necesarios validados.

Antes de publicar, el sistema comprobará:

- coincidencia de comunidad, periodo y perfil;
- todas las incidencias cerradas;
- presencia de datos exigidos por los módulos activos;
- que las sumas de las filas de entrada coinciden con los importes normalizados;
- que las fórmulas, hojas y áreas de impresión requeridas siguen presentes;
- que los totales de gasto, ingreso y diferencia concilian al céntimo según la regla aplicable.

La interfaz mostrará progreso por etapa: validar expediente, preparar plantilla, escribir cada hoja, recalcular, conciliar y publicar. Si falla, conservará el temporal para diagnóstico sólo fuera de la carpeta oficial y no sustituirá el último Excel válido.

El resultado oficial se ubicará en Excels_Maestros/Comunidad_<codigo>.xlsx; cada ejecución también quedará registrada y tendrá el backup correspondiente.

### 6. Reparto final por propietario

Una vez publicado el Excel y conciliados sus datos de entrada, el sistema calculará el resultado final de cada propietario. El reparto se hará por conceptos declarados por el perfil, no por una lista fija de columnas:

- cuota fija de ACS y calefacción: partes iguales o el método configurado;
- consumo variable de ACS y calefacción: lecturas validadas o la estimación auditada aprobada;
- agua, mantenimiento y gastos extraordinarios: el método configurado para el concepto;
- abonos y ajustes: importes negativos trazables;
- conceptos adicionales: la regla declarada para esa comunidad.

Cada importe se distribuye en céntimos con un método determinista de mayor resto. La suma de propietarios por concepto debe coincidir exactamente con el total conciliado del concepto. Los resultados canónicos se guardarán por propietario, periodo y concepto en la tabla de resultados trazables; la tabla histórica de repartos sólo se mantendrá como proyección compatible mientras siga siendo usada por funciones existentes.

No se calcularán repartos ni cartas si falta una lectura obligatoria, si un contador está reiniciado sin una estimación revisada, si no hay pesos válidos de reparto o si cualquier conciliación queda descuadrada. Los errores se convierten en incidencias accionables, no en repartos a cero.

### 7. Cartas individuales

Cuando todos los conceptos y el total general concilien, la aplicación generará una carta Word por propietario. La carta toma exclusivamente los resultados finales de la base de datos y el perfil de carta de la comunidad.

La carta:

- muestra sólo los conceptos activos y relevantes del propietario;
- diferencia importes cobrados, coste real, ajuste y total;
- incorpora cuota fija y variable cuando corresponda;
- muestra consumo propio frente a vecinos y el histórico disponible;
- conserva el diseño corporativo, sombras, separadores y una sola página aprobados para el despacho;
- no muestra una marca fija de un despacho: nombre, logotipo y firma proceden de la configuración;
- se guarda en una carpeta por comunidad y periodo, sin enviar correos automáticamente.

La generación de cartas registra una ejecución y sus errores por propietario. Si una carta falla, el resto puede generarse, pero la ejecución queda marcada como incompleta y no se presenta como lista para envío.

### 8. Fórmulas y recalculado

El generador conservará las fórmulas que formen parte del modelo. Después de escribir los datos se recalculará el libro con un motor compatible de escritorio antes de la conciliación final. La aplicación detectará si dicho motor no está disponible y explicará cómo instalarlo; no marcará un libro como validado basándose sólo en valores de fórmula obsoletos.

Los datos de entrada se escribirán como fechas y números reales, nunca como textos formateados. Las fórmulas y sus referencias se comprobarán antes de publicar para evitar referencias rotas, divisiones por cero o errores de cálculo.

### 9. Diseño de interfaz

No se rediseña la aplicación. En la tarjeta de expediente actual se añadirá:

- perfil de Excel que se usará;
- comprobación resumida de datos antes de exportar;
- acción principal Generar Excel oficial;
- progreso por etapa y enlace para abrir el resultado o su registro de validación.

Tras la conciliación aparecerán las acciones Calcular reparto final y Generar cartas. Cada una mostrará su propio resumen: conceptos calculados, propietarios incluidos, diferencias conciliadas y cartas correctas o con error.

El lenguaje será directo: “Faltan 2 datos para generar el Excel”, “Conciliación correcta” o “Excel oficial publicado”. El sistema no llamará “perfecto” a una salida que no haya pasado los controles.

## Flujo completo

PDFs / lecturas / gastos → expediente e incidencias manuales → base de datos normalizada → perfil y copia de plantilla maestra → hojas y fórmulas rellenadas/recalculadas → conciliación y registro de exportación → Excel oficial + backup → reparto final por propietario → conciliación al céntimo → cartas Word individuales.

Para la primera ejecución de la 658, la entrada será el Excel real como fuente de arranque. En las siguientes, los PDFs y lecturas usarán el mismo contrato de datos.

## Verificación

La implementación incluirá:

- pruebas de migración y de perfiles;
- pruebas de importación idempotente de una plantilla sanitizada;
- pruebas de bloqueo por campos obligatorios o incidencias abiertas;
- pruebas de generación atómica y backup;
- pruebas de conservación de hojas, fórmulas, estilos críticos y áreas de impresión;
- conciliación al céntimo de las sumas relevantes;
- pruebas de reparto de cuota fija, consumo variable, abonos y conceptos configurables;
- pruebas de redondeo que conservan exactamente el total por concepto;
- pruebas de bloqueo por contador reiniciado sin estimación aprobada;
- pruebas de generación de cartas con conceptos opcionales, datos sin identificar y errores aislados;
- una prueba de comparación privada contra el libro real de la 658, sin incorporar datos personales al repositorio;
- comprobación visual manual en Windows del Excel generado, de su recalculado y de una carta de cada perfil activo.

## Fuera de alcance de esta entrega

- extracción automática de los PDFs de la 658;
- envío de correo;
- conversión automática de todas las comunidades existentes.

La entrega deja operativo el flujo completo de la 658: importación inicial del modelo real, Excel oficial, reparto conciliado y cartas individuales.
