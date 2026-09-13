# Validación de expedientes e incidencias

Esta guía describe el flujo de alta, apertura de un expediente y revisión
manual de sus fuentes. El documento de origen se archiva como copia para su
revisión y los originales permanecen intactos. El análisis automático propone
valores; toda ausencia, ambigüedad o contradicción relevante sigue necesitando
confirmación humana.

## Arranque

Desde PowerShell, abre la carpeta del proyecto, activa su entorno virtual y
ejecuta:

```powershell
python core/app.py
```

Antes de trabajar, comprueba en **Configurar rutas** la base de datos, el
archivo de expedientes, las salidas y la plantilla de cartas. Esas ubicaciones,
las fuentes, los perfiles runtime y las plantillas instaladas son locales.

## Pasos operativos

1. **Abrir la app y seleccionar comunidad.** Si todavía no existe, pulsa **+
   Comunidad** y elige **Registro rápido** o **Alta guiada desde fuentes**.
2. **Pulsar Crear expediente y elegir fecha inicial/final.** Puede usarse cualquier rango, incluido uno arbitrario de seis meses.
3. **Pulsar Añadir fuentes; los originales permanecen intactos.**
4. **Elegir el tipo de fuente y seleccionar archivos.**
5. **Pulsar Resolver incidencias; cada incidencia abre su copia archivada.**
6. **Confirmar valor y motivo; el sistema conserva la auditoría.**
7. **Esperar el estado Listo para cálculo antes de usar las salidas posteriores.**

## Flujo unificado de fuentes

Sigue este recorrido para un expediente nuevo. Todos los originales se conservan
sin cambios; la aplicación trabaja con sus copias archivadas dentro del
expediente.

1. Crea o selecciona la comunidad con **+ Comunidad**. Para una comunidad que
   aún no existe, elige **Registro rápido** o **Alta guiada desde fuentes** y
   completa su código y nombre.
2. Selecciona la comunidad y usa **Crear expediente** para indicar el nombre y
   las fechas inicial y final del período que vas a regularizar.
3. En **Añadir fuentes**, pulsa **Añadir carpeta** para incorporar todos los
   PDF, XLSX, XLS y CSV visibles de una carpeta, o **Añadir archivos** para
   escoger documentos concretos. La carpeta puede contener una mezcla de
   facturas, lecturas y listados de propietarios; no hace falta separarlos
   antes.
4. Lee el cuadro **Fuentes analizadas** al finalizar: agrupa las copias nuevas
   y duplicadas como *facturas detectadas*, *lecturas*, *listados de
   propietarios*, *otros documentos* y *documentos por revisar*. Una lectura
   o un listado no recibe los campos obligatorios de una factura. Cada
   *documento por revisar* abre una sola incidencia de clasificación; no
   confirmes que es una factura salvo que el original lo demuestre.
5. Abre **Resolver incidencias**. Antes de confirmar, usa **Ver contexto** para
   consultar página/fragmento de PDF u hoja/celda de Excel cuando estén
   disponibles, y **Abrir archivo** para ver la copia archivada. Escribe el
   valor confirmado y el motivo o fuente que lo respalda. Para una clasificación
   desconocida, selecciona el tipo correcto y explica el motivo; la app vuelve
   a analizar la copia para crear sólo los datos que requiere ese tipo.
6. Si cambian los originales, se corrige el extractor o quieres actualizar las
   propuestas automáticas, usa **Reanalizar fuentes**. Esta acción reutiliza las
   copias archivadas del expediente: no reimporta ni modifica los originales y
   conserva las correcciones confirmadas manualmente. Revisa de nuevo su resumen
   y sus incidencias.
7. Con cero incidencias abiertas y el estado **Listo para cálculo**, pulsa
   **Generar Excel oficial**, revisa el libro publicado y después **Calcular
   reparto final**. Continúa con **Generar cartas**, revisa conceptos,
   destinatarios y una muestra de los archivos antes de enviarlos manualmente.
   Excel, reparto y cartas siguen bloqueados mientras quede una incidencia
   abierta.

## Empezar con una base vacía de forma segura

Para una aceptación manual que requiera una base limpia, primero termina el
trabajo en curso y cierra la aplicación para que no quede otra sesión usando la
base. No borres, renombres ni sustituyas `data/gestion.db` desde el explorador.
Al volver a abrir la aplicación, entra en **Configurar rutas** y selecciona
**Nueva base segura**. Confirma el diálogo sólo si quieres iniciar de cero.

La acción crea y verifica una copia de seguridad de la base anterior en la
carpeta `backups` situada junto a la base configurada antes de publicar una
base vacía. La ruta de esa copia aparece al terminar. La base anterior nunca se
elimina silenciosamente: si la copia o la nueva base no superan la verificación,
el reinicio se cancela y la base anterior se conserva. Conserva la ruta indicada
para poder recuperar el contexto de la prueba anterior.

## Alta guiada desde fuentes

El asistente permite crear una comunidad sin disponer de un Excel manual
perfecto. Presenta cinco etapas consecutivas y visibles:

1. **Identidad:** código, nombre y, opcionalmente, datos del primer período.
2. **Fuentes:** lista de propietarios CSV y una o más lecturas Excel, XLS o
   PDF. Ambos grupos son obligatorios. Las facturas PDF son opcionales y se
   pueden seleccionar varias.
3. **Detección:** propuestas separadas de lecturas y facturas, sin escrituras.
   Las facturas sólo aportan proveedor, período, importe y concepto; las
   lecturas sólo aportan suministro/módulo, contador, columna, fecha y valor.
4. **Confirmaciones:** selección de servicio y respuestas obligatorias para
   cualquier campo ausente, ambiguo o contradictorio. Para ACS y calefacción,
   confirma una lectura por módulo: no se comparte contador o columna salvo
   que se elija expresamente para ambos. **NO_APLICA** sólo es válido para un
   módulo inactivo y queda trazado.
5. **Resumen:** revisión de columna, contador, fecha y valor por módulo, además
   de proveedor, período, importe y concepto de cada factura. La vista usa la
   configuración ya normalizada; no vuelve a interpretar las fuentes. La
   publicación permanece bloqueada mientras falte una confirmación.

La publicación se ejecuta en segundo plano y la barra de estado muestra la
operación en curso. Al terminar se actualizan los selectores de comunidad,
período y expediente. Las fuentes archivadas, la base de datos, los perfiles
generados y las plantillas instaladas son datos locales y privados: no se
incluyen en el control de versiones.

Si se omite el período, el asistente publica únicamente comunidad, perfil y
plantilla local. Si se rellenan nombre, fecha inicial y fecha final, crea
también el expediente y archiva copias verificadas de las fuentes. No se admite
completar sólo una parte de esas fechas.

## Del alta a Excel, reparto y cartas

1. Selecciona la comunidad y el expediente inicial, o crea uno si el alta se
   hizo sin período.
2. Usa **Importar modelo inicial** una sola vez para incorporar el maestro
   histórico, si existe. En un alta guiada se utiliza la plantilla canónica
   instalada para esa comunidad. Puedes acompañarla con propietarios y
   lecturas. El original no se modifica.
3. La plantilla canónica recién creada tiene una excepción de arranque
   reconocible. Esa excepción no se aplica a un maestro posterior vacío: si
   faltan parámetros obligatorios, se generan incidencias
   `MISSING_REQUIRED_FIELD`.
4. Añade las facturas y lecturas del período que no se incorporaron durante el
   alta o la importación inicial.
5. Abre **Resolver incidencias** y confirma valor y motivo hasta alcanzar
   **Listo para cálculo**.
6. Pulsa **Generar Excel oficial**. Revisa el libro publicado; se crea desde la
   base de datos validada, no desde cambios manuales en una salida anterior.
7. Pulsa **Calcular reparto final**. Comprueba que los conceptos activos y el
   total conciliado cuadran al céntimo.
8. Pulsa **Generar cartas**, revisa conceptos y destinatarios y abre una muestra
   antes de usar el lote. El envío por correo es siempre manual.

## Comprobaciones de aceptación

- Un rango arbitrario de seis meses se puede abrir y conservar como expediente.
- Al añadir una factura PDF se crean para confirmar los tres campos `fecha_inicio`, `fecha_fin` e `importe_total`.
- Cancelar o dejar en blanco una confirmación no escribe cambios.
- Volver a añadir una fuente no duplica la fuente ni sobrescribe una corrección ya confirmada.
- Cambiar de comunidad impide utilizar un expediente antiguo.
- Una incidencia sin resolver bloquea el estado **Listo para cálculo**.
- Los originales no se modifican: durante la resolución se abre únicamente la copia archivada.
- El resumen del alta guiada no está disponible mientras falten propietarios,
  lecturas o respuestas obligatorias.
- Una factura de calefacción puede proponer ese módulo, pero nunca se convierte
  en una lectura de contador.
- Una configuración combinada puede publicar columnas, contadores y valores
  distintos para ACS y calefacción.
- El resumen final muestra las mismas decisiones de lectura y factura que
  consumen perfil, plantilla y publicación.

## Alcance actual

La revisión ordinaria de expedientes conserva el flujo manual descrito. El
alta guiada analiza fuentes y publica el perfil y la plantilla canónicos; el
cálculo, reparto y cartas se ejecutan después mediante las acciones ordinarias
del expediente. Un Excel manual previo es opcional y no bloquea el alta.
