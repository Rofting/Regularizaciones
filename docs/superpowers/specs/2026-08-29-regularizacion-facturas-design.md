# Diseño del sistema de regularización de facturas y consumos

Fecha: 29 de agosto de 2026

## 1. Objetivo

Construir una aplicación que transforme facturas, lecturas y datos de propietarios en una regularización anual completa, trazable y repetible:

1. leer documentos de entrada;
2. solicitar intervención cuando falte un dato o la extracción no sea fiable;
3. guardar únicamente datos validados en la base de datos;
4. calcular el reparto por propiedad;
5. generar el Excel anual con la estructura de los libros de referencia;
6. generar cartas Word individuales de una página, con detalle económico y gráficas;
7. permitir una revisión final antes de cualquier envío.

La solución se validará primero con la comunidad 658 y después con la 644, pero su modelo, reglas e interfaz deben servir para todas las comunidades del despacho. La arquitectura debe permitir convertirla posteriormente en una aplicación de marca blanca para otros despachos.

## 2. Principios de diseño

- La base de datos es la única fuente de verdad operativa.
- Los Excel perfectos existentes son patrones de validación durante el desarrollo. En el flujo definitivo serán salidas del programa, no entradas necesarias.
- Ninguna regla de cálculo puede estar codificada exclusivamente para la comunidad 658 o la 644.
- Los datos dudosos o incompletos no se guardan como definitivos.
- Todo importe debe poder rastrearse hasta su factura, lectura, corrección manual o regla de reparto.
- Las diferencias económicas se comprueban al céntimo antes de generar las cartas.
- Las configuraciones propias de una comunidad se guardan como datos, no como cambios de código.
- El nombre, logotipo, colores, firma y datos de contacto del despacho son configurables por instalación.

## 3. Alcance por fases

### Fase 1: validación con la comunidad 658

- Importar a la base de datos los Excel perfectos de 2024-2025 y 2025-2026.
- Mantener el archivo y las celdas de origen como información de auditoría.
- Comprobar propiedades, lecturas, conceptos, importes y totales.
- Generar cartas Word desde la base de datos.
- Comparar los resultados con ambos Excel de referencia.
- Generar las dos gráficas aprobadas usando datos reales.

Esta importación desde Excel es una herramienta de validación y migración inicial; no será el flujo operativo final.

### Fase 2: validación cruzada con la comunidad 644

- Ejecutar el mismo proceso con una comunidad que contiene calefacción, ACS, cuotas y un histórico más amplio.
- Detectar supuestos que solo funcionen para la 658.
- Confirmar que los conceptos dinámicos y las gráficas funcionan con otra estructura de reparto.

### Fase 3: incorporación del resto del despacho

- Dar de alta progresivamente todas las comunidades.
- Configurar suministros, identificadores, conceptos, periodos y propietarios sin modificar código.
- Mostrar en la app si una comunidad está lista o qué configuración le falta.

### Fase 4: flujo documental completo

- Leer facturas PDF, PDFs de lecturas y listados de propietarios.
- Resolver incidencias mediante revisión asistida.
- Calcular la regularización.
- Generar automáticamente el Excel anual y las cartas.

### Fase 5: producto portable

- Separar por instalación la base de datos, la marca, las plantillas y la configuración.
- Crear un proceso de instalación y alta inicial para otros despachos.
- Mantener el mismo motor de extracción, validación, cálculo y generación.

## 4. Flujos de datos

### 4.1. Flujo temporal de validación

```text
Excel perfecto 658
    -> importador con trazabilidad
    -> validaciones y cuadre
    -> base de datos
    -> selección de conceptos en la app
    -> vista previa
    -> cartas Word
```

### 4.2. Flujo operativo definitivo

```text
Facturas PDF + PDFs de lecturas + propietarios
    -> identificación de comunidad, proveedor, servicio y ejercicio
    -> extracción de campos
    -> validación automática
    -> revisión manual de incidencias
    -> base de datos validada
    -> motor de reparto
    -> control de totales
    -> Excel anual
    -> selección de conceptos y vista previa
    -> cartas Word
    -> revisión final
    -> envío posterior
```

El envío por correo no forma parte de la primera entrega. Permanecerá separado de la generación para impedir envíos accidentales sin revisión.

## 5. Modelo de información

### 5.1. Despacho e instalación

- Nombre comercial opcional.
- Logotipo opcional.
- Colores de marca.
- Dirección y datos de contacto.
- Firma utilizada en las cartas.
- Rutas de entrada, salida, plantillas y copias de seguridad.

Si no se configura una marca, la interfaz mostrará un título neutro, por ejemplo, «Gestor de regularizaciones».

### 5.2. Comunidad

- Código interno.
- Nombre y CIF.
- Dirección.
- Número de propiedades.
- Periodo o ciclo habitual.
- Servicios aplicables.
- CUPS, pólizas u otros identificadores de suministro.
- Conceptos habituales de sus cartas.
- Estado de preparación y lista de datos pendientes.

### 5.3. Propiedad y propietario

- Identificador de vivienda o local.
- Titular.
- Correo electrónico.
- Coeficientes aplicables.
- Datos de contadores y servicios.
- Vigencia temporal para conservar cambios de titularidad.

### 5.4. Ejercicio

- Comunidad.
- Nombre del ejercicio.
- Fecha inicial y final exactas.
- Estado: preparación, revisión, validado o cerrado.
- Archivos de origen y salidas generadas.

### 5.5. Documentos y extracción

- Archivo, huella digital y fecha de importación.
- Tipo documental y proveedor.
- Comunidad y ejercicio detectados.
- Campos extraídos, valor, confianza y evidencia de origen.
- Estado de revisión.
- Correcciones manuales y usuario que las realizó.

### 5.6. Lecturas y facturas

- Servicio.
- Fechas inicial y final.
- Lecturas inicial y final.
- Consumo.
- Unidades.
- Importes fijos, variables, impuestos y total.
- Documento de origen.

### 5.7. Conceptos y regularización

Los conceptos se modelan como registros dinámicos. Como mínimo se admitirán:

- cuota fija de calefacción;
- cuota fija de ACS;
- consumo de calefacción;
- consumo de ACS;
- gastos extraordinarios;
- ajustes;
- abonos;
- conceptos configurables adicionales.

Cada resultado por propiedad guardará concepto, importe cobrado, coste real y diferencia. Las cartas mostrarán filas, no columnas nuevas, para los conceptos adicionales. Las filas sin datos o no seleccionadas se omitirán.

### 5.8. Auditoría

- Lote de importación.
- Documento o Excel de origen.
- Regla aplicada.
- Valor extraído y valor corregido.
- Fecha y responsable de la corrección.
- Versiones del cálculo, Excel y carta generados.

## 6. Extracción y revisión de incidencias

1. El sistema extrae los campos y valida los obligatorios.
2. Un campo ausente, ambiguo, incoherente o de baja confianza crea una incidencia.
3. El documento queda pendiente y no alimenta el cálculo definitivo.
4. La app abre el archivo problemático, en la página correspondiente cuando sea posible.
5. Se presenta un formulario únicamente con los campos que necesitan intervención.
6. Tras la corrección se repiten las validaciones de fechas, importes, comunidad, servicio y totales.
7. Solo una operación completamente válida se confirma en la base de datos.

Después de una corrección, la app ofrecerá «Usar esta corrección en futuras facturas similares». Esta acción creará una propuesta de regla limitada al proveedor y formato detectados. La regla no se activará de forma silenciosa: deberá confirmarse y probarse contra documentos históricos. Se conservará su versión y podrá desactivarse.

Los documentos no resolubles podrán dejarse pendientes o enviarse a cuarentena con un motivo explícito.

## 7. Motor de cálculo y validación

- El motor trabaja exclusivamente con datos validados.
- Las reglas fijas y variables se configuran por comunidad, servicio, concepto y ejercicio.
- La asignación verano/invierno se deriva de fechas y reglas configuradas, no de posiciones fijas en un Excel.
- Los abonos y rectificaciones conservan el signo y el vínculo con la factura corregida.
- Los cálculos se guardan con suficiente precisión y se redondean de forma explícita al presentar importes.
- Antes de cerrar un ejercicio se concilian totales de facturas, costes distribuidos, importes cobrados y diferencias.
- Una diferencia de conciliación bloquea el Excel y las cartas hasta su resolución.

## 8. Generación del Excel

El generador construirá el libro anual desde la base de datos y respetará la estructura validada de los modelos:

- DATOS;
- GAS;
- ELECTRICIDAD;
- AGUA;
- OTROS GASTOS;
- lecturas de ACS;
- lecturas de calefacción cuando correspondan;
- análisis del ejercicio;
- resumen e históricos cuando correspondan.

Las hojas no aplicables a una comunidad podrán ocultarse u omitirse según su plantilla. Las fórmulas, totales, formatos y referencias deben ser auditables. El sistema mantendrá una copia de seguridad antes de sustituir una salida existente.

En la fase inicial, la comparación automática con los Excel perfectos verificará celdas de control, totales y resultados por propiedad.

## 9. Cartas Word

### 9.1. Preparación

Antes de generar cartas, la app mostrará los conceptos disponibles mediante casillas. Recuperará como propuesta la última configuración utilizada en esa comunidad, pero permitirá modificarla. Se mostrará un resumen de propietarios afectados e importes antes de continuar.

### 9.2. Composición aprobada

- Una sola página.
- Respeto de la estructura, sombras suaves, relieve, líneas finas y filas alternas en gris claro de la carta corporativa validada.
- Cabecera y datos de marca configurables.
- Resultado de la regularización destacado.
- Datos de propiedad y ejercicio.
- Detalle de lecturas, consumos e importes.
- Resumen con columnas estables: concepto, cobrado, coste real y diferencia.
- Filas dinámicas para cuotas, consumos, gastos extraordinarios, ajustes y abonos.
- Texto informativo y firma.

### 9.3. Gráficas

Las dos gráficas aparecen en paralelo dentro de la misma página:

1. Distribución de vecinos por franjas de consumo, por ejemplo 0-10, 10-20, 20-30, 30-40 y sucesivas. Se destaca la franja de la propiedad destinataria.
2. Evolución del consumo anual de esa propiedad frente a sus ejercicios anteriores disponibles.

No se mostrará una gráfica de consumo por periodos. No se inventarán datos mensuales ni históricos. Si no existe información suficiente para una gráfica, la composición redistribuirá el espacio sin incluir una visualización incompleta.

El generador ajustará espacios, filas vacías y bloques opcionales para mantener una página sin reducir el texto por debajo del tamaño legible definido en la plantilla. Si una selección excepcional de conceptos no cabe, la app bloqueará esa generación y pedirá revisar o agrupar conceptos; nunca creará silenciosamente una segunda página.

## 10. Interfaz de usuario

### 10.1. Dirección visual aprobada

- Estructura cercana y profesional basada en la opción visual 3.
- Tipografía Trebuchet MS o una alternativa portable de métrica y personalidad equivalentes si las pruebas de empaquetado lo requieren.
- Fondo blanco y gris muy suave.
- Verde petróleo para navegación y acciones.
- Verde menta para estados correctos.
- Colores de aviso y error usados únicamente con significado funcional.
- Marca blanca configurable; ninguna referencia fija a Meditrade.

### 10.2. Pantalla principal

- No se abre un diálogo obligatorio de periodo al arrancar.
- Comunidad y ejercicio/ciclo se seleccionan en una tarjeta visible.
- «Crear comunidad» aparece junto al selector de comunidad.
- «Crear ejercicio» aparece junto al selector de ejercicio.
- Se muestran fechas, número de propiedades y estado de preparación.
- El flujo se organiza en cinco fases visibles: inicio, archivos, revisión, cálculo y entrega.
- La pantalla resume documentos leídos, incidencias y propiedades.

### 10.3. Alta de comunidad

El alta es un asistente de tres pasos:

1. Datos generales: código, CIF, nombre, dirección, propiedades y periodo habitual.
2. Servicios y reglas: ACS, calefacción, cuotas, gastos, identificadores de suministro y conceptos habituales.
3. Importación inicial: propietarios, correos, coeficientes, contadores y datos históricos disponibles.

El asistente informa de lo que falta y deja la comunidad en estado «lista» o «pendiente de configuración».

### 10.4. Flujo guiado

1. Seleccionar comunidad y ejercicio.
2. Cargar documentos.
3. Resolver incidencias.
4. Validar y calcular.
5. Generar Excel.
6. Seleccionar conceptos de cartas.
7. Revisar una vista previa representativa.
8. Generar todos los Word.

## 11. Gestión de errores

- Un error en un documento no detiene el procesamiento de los demás.
- Cada incidencia muestra archivo, causa, campo afectado y acción requerida.
- No se ocultan excepciones que puedan cambiar importes.
- Las operaciones que afectan a varias tablas se realizan de forma transaccional.
- Los reintentos no duplican facturas, lecturas, repartos ni cartas.
- Los archivos procesados, pendientes y en cuarentena quedan separados.
- Cada ejecución produce un resumen legible y un registro técnico.

## 12. Pruebas y criterios de aceptación

### 12.1. Comunidad 658

- Ambos ejercicios se importan sin duplicados.
- Propiedades, lecturas, conceptos e importes coinciden con los Excel de referencia.
- Los totales cuadran al céntimo por concepto, propiedad y comunidad.
- Las cartas usan los conceptos seleccionados y omiten los restantes.
- Las cartas caben en una página con las dos gráficas cuando existen datos suficientes.
- La franja de consumo de cada propiedad es correcta.
- El histórico refleja únicamente ejercicios existentes.

### 12.2. Comunidad 644

- El mismo modelo admite calefacción, ACS, cuotas y un histórico distinto.
- No se introduce ninguna condición específica en el código para resolver diferencias de estructura.
- Los resultados históricos existentes continúan cuadrando.

### 12.3. Incidencias

- Un documento incompleto queda pendiente y no contamina la base de datos.
- La corrección manual permite completar y revalidar el documento.
- Las propuestas de reglas futuras requieren confirmación.
- La auditoría conserva valores originales y corregidos.

### 12.4. Regresión

Los Excel perfectos de la 658 y los datos históricos validados de la 644 permanecen como casos automáticos. Cualquier cambio que altere lecturas, importes, reparto o cartas deberá detectarse antes de publicar una versión.

## 13. Fuera de la primera entrega

- Envío automático de correos.
- Integraciones con WhatsApp u otros canales.
- Gestión comercial o facturación del despacho.
- Migración completa de todas las comunidades en un único lote.
- Instalador comercial para terceros.

Estos elementos quedan previstos por la arquitectura, pero se abordarán después de validar el núcleo con la 658 y la 644.

## 14. Resultado esperado

Al finalizar el proyecto, un usuario podrá dar de alta una comunidad, cargar sus documentos, resolver únicamente los campos dudosos, validar el cálculo y obtener un Excel anual y cartas Word consistentes. La incorporación de una comunidad o despacho nuevos requerirá configuración y datos, no cambios en el código del motor.

## 15. Componentes y límites

- **Persistencia:** esquema, migraciones, transacciones y consultas. No interpreta PDFs ni decide el diseño de las salidas.
- **Importación de referencias:** convierte los Excel perfectos en registros normalizados con trazabilidad. Se usa para validación y migración inicial.
- **Lectores documentales:** identifican y extraen facturas, lecturas y listados; producen datos candidatos con evidencia y confianza.
- **Servicio de revisión:** crea incidencias, presenta los campos dudosos, aplica correcciones y propone reglas reutilizables.
- **Configuración:** gestiona despacho, comunidades, servicios, periodos, conceptos, proveedores y reglas.
- **Motor de reparto:** calcula resultados por propiedad a partir de datos validados y reglas configuradas.
- **Conciliación:** verifica totales y decide si un ejercicio puede pasar a generación.
- **Generador de Excel:** produce el libro anual desde la base de datos.
- **Generador de cartas:** compone la carta dinámica, las gráficas y la salida Word de una página.
- **Interfaz:** coordina el flujo y muestra estados; no contiene fórmulas de negocio.
- **Auditoría:** registra orígenes, decisiones, versiones y resultados de cada ejecución.

Cada componente expone entradas y resultados explícitos para poder probarlo sin abrir la interfaz y cambiar su implementación sin romper los demás.

## 16. Límite del primer plan de implementación

Esta especificación describe la arquitectura completa, pero no se ejecutará como un único cambio. El primer plan abarcará solamente:

1. ampliaciones mínimas del modelo de datos necesarias para la validación;
2. importación trazable de los dos Excel perfectos de la 658;
3. conciliación al céntimo;
4. selección y persistencia de conceptos de carta;
5. generación Word de una página con las dos gráficas aprobadas;
6. pruebas automáticas de la 658.

La 644, la extracción documental completa, la migración del resto de comunidades y el empaquetado para terceros tendrán planes posteriores, usando los mismos límites arquitectónicos.
