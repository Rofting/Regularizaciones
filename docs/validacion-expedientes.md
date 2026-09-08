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
