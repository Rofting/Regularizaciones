# Alta guiada de comunidades — Diseño

## Objetivo

Dar de alta una comunidad nueva sin depender de un Excel perfecto previo. El
usuario aporta propietarios, lecturas y facturas; la aplicación conserva las
fuentes, detecta su estructura, solicita sólo confirmaciones necesarias y
guarda un perfil versionado con el que podrá generar el Excel oficial y las
cartas en los ejercicios posteriores.

## Alcance de esta fase

- Añadir un asistente de alta de comunidad accesible desde la interfaz v3.
- Aceptar como fuentes: listado de propietarios, lecturas en Excel o PDF y
  facturas PDF.
- Analizar de forma no destructiva las fuentes seleccionadas y presentar una
  propuesta de conceptos, tipos de lectura y campos pendientes.
- Permitir confirmar o corregir cada propuesta de baja confianza antes de
  crear el perfil.
- Persistir un perfil declarativo y una plantilla de Excel generada para la
  comunidad, sin incorporar los documentos privados al repositorio.
- Crear un expediente inicial sólo cuando el usuario indique el período.

No se incluye todavía el envío de correo ni se asume que un proveedor concreto
de facturas sea universal. Un Excel manual existente es opcional y únicamente
sirve como referencia de comparación durante una futura validación.

## Experiencia de usuario

El botón **+ Comunidad** abre dos opciones: **Registro rápido** (el formulario
actual) y **Alta guiada desde fuentes**. La opción guiada se organiza en cinco
pasos visibles y guarda un borrador entre pasos:

1. **Identidad**: código, nombre, CIF opcional y número estimado de viviendas.
2. **Fuentes**: propietarios (CSV/XLSX), lecturas (XLSX/XLS/PDF) y una o más
   facturas PDF. Se muestra qué archivos se archivarán y cuáles faltan.
3. **Propuesta detectada**: servicios/conceptos encontrados, contador asociado
   y períodos detectados. Cada propuesta lleva estado `confirmado`,
   `revisar` o `no detectado`.
4. **Confirmaciones**: sólo se editan campos `revisar` o `no detectado`;
   además se puede desactivar un concepto que no aplique. No hay valores
   preseleccionados que inventen importes, consumos o períodos.
5. **Resumen y creación**: muestra el perfil que se guardará, las fuentes
   archivadas y las incidencias que seguirán abiertas. El botón final crea la
   comunidad, su perfil y el expediente si se ha indicado período.

La pantalla mantiene la paleta y tipografía v3 actual. Los textos describen
la decisión del usuario: “Confirmar servicio”, “Indicar columna de lectura” y
“Crear comunidad”, no nombres internos de tablas o perfiles.

## Arquitectura

### Análisis de fuentes

Se crea un servicio de dominio `community_onboarding` sin controles Tk. Recibe
rutas explícitas, valida extensiones y construye un `OnboardingDraft` con:

- identidad propuesta;
- fuentes y sus huellas SHA-256;
- hojas, encabezados y períodos detectados de las lecturas Excel;
- resultados existentes del lector de PDF para facturas y lecturas PDF;
- conceptos candidatos y su nivel de confianza;
- preguntas pendientes, cada una con un identificador estable.

El análisis no escribe en la base de datos ni instala plantillas. Archiva las
fuentes únicamente al confirmar el paso final.

### Perfil declarativo

Tras las confirmaciones, el servicio construye un `ExcelProfile` declarativo
con una clave segura basada en código de comunidad y versión `1`. Los conceptos
activos se derivan exclusivamente de confirmaciones explícitas. El perfil se
guarda como JSON local en `config/excel_profiles/` y debe superar la misma
validación que el perfil 658.

En esta fase la plantilla se crea con el generador oficial de libro base y se
instala en la ruta privada de la comunidad. Las comunidades complejas pueden
seguir aportando después un modelo de Excel de referencia; esa comparación no
altera el perfil sin una nueva revisión/versionado.

### Persistencia y trazabilidad

La confirmación final se ejecuta de forma transaccional:

1. crea o actualiza la comunidad;
2. guarda el perfil validado;
3. archiva las fuentes y registra sus huellas;
4. crea el período/expediente inicial si se proporcionaron fechas;
5. registra incidencias abiertas para datos que no se hayan confirmado.

Si falla cualquier paso, no se crea un perfil activo ni una plantilla parcial.
Las fuentes originales nunca se modifican. Los archivos privados, bases de
datos y plantillas instaladas permanecen ignorados por Git.

## Manejo de incertidumbre

Un valor se marca `confirmado` sólo cuando procede de una regla inequívoca o
de una confirmación humana. Los casos ambiguos (por ejemplo, varias columnas
posibles de contador, fecha ilegible de una factura o concepto no reconocido)
se convierten en preguntas del borrador. La aplicación abre el archivo
archivado cuando el usuario lo solicite y no permite finalizar mientras haya
preguntas obligatorias sin resolver.

La ausencia legítima de un concepto se confirma como “no aplica”; no se crea
una cuota fija, variable, ACS, calefacción o gasto extraordinario de forma
automática.

## Integración con el flujo existente

- `expedient_service` sigue siendo la autoridad para comunidades, períodos y
  expedientes.
- `excel_profiles` valida y carga perfiles creados por el asistente.
- `excel_bootstrap_importer` sigue siendo útil para comunidades históricas
  que ya tienen un Excel maestro; el nuevo asistente no lo requiere.
- `case_workflow_actions` recibe el expediente creado y conserva las mismas
  etapas: revisión → Excel oficial → reparto → cartas.

## Pruebas y criterios de aceptación

- Un borrador con propietarios, lecturas Excel y facturas PDF detecta fuentes
  y no escribe archivos ni filas de base de datos durante el análisis.
- Una columna de lecturas ambigua produce una pregunta obligatoria; confirmar
  una columna crea un perfil válido y trazable.
- Una comunidad sin ACS, calefacción o gasto extraordinario no incluye esos
  conceptos en el perfil.
- Lecturas PDF y Excel se tratan como fuentes alternativas del mismo dato.
- Un fallo en la confirmación no deja perfil, comunidad, plantilla ni fuentes
  parciales.
- La interfaz enruta los cinco pasos sin iniciar ventanas en pruebas unitarias.
- La suite completa, compilación y comprobación de diferencias se ejecutan en
  Windows con el entorno virtual del proyecto.

## Decisiones

- El Excel perfecto manual es opcional y no bloquea el alta.
- Propietarios y lecturas son las fuentes mínimas esperadas de una comunidad
  nueva; las lecturas pueden llegar en Excel o PDF.
- Las facturas PDF son la fuente para proponer conceptos y períodos, pero todo
  dato incierto se confirma antes de convertirse en configuración.
- La portabilidad se basa en perfiles locales declarativos y no en nombres,
  colores o reglas fijas de un despacho concreto.
