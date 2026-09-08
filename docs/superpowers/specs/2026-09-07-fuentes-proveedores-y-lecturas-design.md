# Facturas de proveedores y lecturas de contador — diseño

## Objetivo

Completar el alta guiada para que una comunidad se configure a partir de dos
fuentes complementarias: facturas de proveedores para conceptos, importes y
períodos; y lecturas de contadores para consumo. El programa debe generar un
perfil, un Excel canónico y un reparto compatibles sin requerir un Excel
manual previo.

## Principios no negociables

- Una factura de proveedor nunca se utiliza como lectura de contador.
- Una lectura de contador nunca inventa un importe o un concepto facturable.
- Todo campo incierto, ausente o contradictorio se formula como una pregunta
  manual obligatoria. Incluye proveedor, fecha/período, importe, concepto,
  contador, suministro, columna, fecha de lectura y valor de lectura.
- La persona usuaria puede declarar explícitamente “no aplica” sólo cuando
  el dato no sea necesario para el módulo confirmado. Esa decisión queda
  trazada; no existe un perfil vacío implícito.
- Un módulo ACS y un módulo de calefacción pueden usar columnas/contadores
  distintos. No se reutiliza una selección de lectura entre módulos sin una
  confirmación explícita.
- El perfil generado usa los nombres canónicos del motor: `acs_fixed`,
  `acs_variable`, `heating_fixed` y `heating_variable`, con plantilla y
  bindings que permitan exportación y reparto.
- Fuentes, base de datos, perfiles runtime y plantillas instaladas siguen
  siendo locales y no se versionan.

## Modelo de propuesta y confirmación

El análisis produce evidencia separada:

1. `invoice_evidence`: proveedor, período, importes y candidatos de concepto
   extraídos de facturas PDF.
2. `reading_evidence`: contador, columna, fechas, lecturas y candidatos de
   suministro extraídos de Excel o PDF de lecturas.

La evidencia sólo propone. Una normalización única de respuestas confirmadas
construye la configuración publicable. El resumen usa esa misma configuración
para mostrar, por módulo, factura/concepto confirmado, contador/columna y
período. La publicación no puede reconstruir una variante distinta.

Para ACS y calefacción simultáneos, la confirmación contiene una decisión de
lectura por módulo. Una columna puede compartirse únicamente si la persona lo
elige para ambos módulos y la evidencia lo permite.

## Flujo

1. Seleccionar lista de propietarios, lecturas Excel/PDF y facturas PDF
   opcionales.
2. Analizar ambas fuentes sin escribir datos.
3. Mostrar propuesta separada de facturas y lecturas, con preguntas manuales
   para toda incertidumbre o ausencia relevante.
4. Normalizar decisiones y mostrar un resumen verificable de cada módulo.
5. Crear perfil, plantilla, expediente y archivo de fuentes de manera segura.
6. Importar el maestro, generar Excel canónico, calcular reparto y generar
   cartas mediante los flujos existentes.

## Validación del maestro inicial

La excepción de plantilla vacía se limita a la plantilla recién creada y
reconocible mediante una huella/estado de bootstrap específico. Un maestro
posterior vacío o sin parámetros obligatorios se informa como incompleto: no
puede pasar revisión sólo porque el perfil proceda de alta guiada.

## Pruebas de aceptación

- Una factura de calefacción propone calefacción sin aportar una lectura.
- Una lectura ambigua o incompleta genera una pregunta manual y bloquea la
  publicación hasta responderla.
- Una factura con importe/período incierto genera confirmaciones manuales.
- ACS y calefacción pueden confirmar lecturas distintas y producen conceptos,
  bindings, reparto y exportación canónicos.
- El resumen refleja las decisiones de factura y lectura que publicará.
- Un libro posterior completamente vacío falla validación aunque use un
  perfil runtime de onboarding.
