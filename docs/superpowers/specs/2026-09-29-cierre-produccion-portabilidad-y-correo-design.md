# Cierre de producción, portabilidad y correo

Fecha: 29 de septiembre de 2026

## 1. Objetivo

Completar el producto para que un despacho pueda partir de una base vacía,
incorporar una carpeta con documentos de varias comunidades, resolver sólo las
decisiones realmente dudosas y obtener para cada expediente un Excel validado,
un reparto conciliado, cartas de una página y correos preparados o enviados con
confirmación explícita.

La entrega debe cerrar primero los casos reales 658 y 644, conservar la base de
datos como fuente de verdad y separar el programa de los datos privados de cada
instalación. No se consideran terminados los componentes que sólo funcionan en
pruebas sintéticas: cada salida debe superar también una aceptación privada con
documentos reales.

## 2. Estado de partida

La rama `feature/global-provider-detection` contiene el flujo más avanzado. En
la base local examinada:

- la 658 tiene propietarios, lecturas, facturas y dos exportaciones Excel
  validadas, pero todavía no tiene un reparto ni un lote de cartas completados;
- la 644 tiene facturas y cero incidencias abiertas, pero carece de propietarios,
  lecturas, perfil registrado, exportación y reparto;
- el estado `ready_for_calculation` puede alcanzarse sin esos datos estructurales,
  por lo que cero incidencias no equivale todavía a expediente completo;
- la rama de trabajo está diecisiete commits por delante de `master`.

Estos hechos determinan el orden de trabajo: primero integridad, después
aceptación real, luego distribución portable y correo, y por último publicación
estable.

## 3. Principios

- Una etapa sólo aparece completa cuando cumple todos sus requisitos, no cuando
  simplemente carece de incidencias abiertas.
- Un dato ausente se presenta como una acción agrupada y comprensible. No se
  generan decenas de preguntas equivalentes por vivienda.
- Ningún dato económico, lectura, coeficiente o dirección de correo se inventa.
- Los documentos originales y la base habitual no se modifican durante una
  validación privada.
- Las salidas se versionan. Repetir Excel, reparto, cartas o correo no borra el
  historial anterior.
- Ningún correo se envía durante pruebas ni sin una confirmación final visible.
- El programa no contiene una marca, ruta, comunidad o credencial de un despacho
  concreto.
- Los secretos no se guardan en Git, JSON, SQLite ni archivos de registro.

## 4. Entregas independientes

El trabajo se divide en cinco subproyectos. Cada uno produce software utilizable,
pruebas y un commit verificable antes de iniciar el siguiente.

### 4.1 Integridad y preparación del expediente

Se añadirá un servicio puro de requisitos que produzca un
`CaseReadinessReport`. El informe contendrá bloqueos agrupados, avisos y la
acción que resuelve cada problema. La interfaz y los servicios de Excel,
reparto y cartas consumirán el mismo informe para evitar criterios diferentes.

Los requisitos se evalúan por etapa:

#### Fuentes confirmadas

- no hay incidencias abiertas;
- no quedan documentos reconocidos pendientes de confirmar o aplicar;
- toda factura elegible tiene comunidad, período, suministro, fechas e importe
  reconciliable;
- todo documento excluido conserva un motivo auditable.

#### Excel oficial

- existe al menos una propiedad activa cuando el perfil requiere reparto;
- el perfil Excel activo se puede resolver o crear desde el modelo común;
- los conceptos obligatorios tienen fuentes económicas válidas;
- las lecturas requeridas identifican propiedad, servicio y fecha;
- LibreOffice está disponible antes de iniciar una exportación que necesite
  recálculo y validación.

#### Reparto

- existe una exportación Excel validada para la huella actual;
- cada concepto tiene una base económica y un método de reparto;
- los conceptos por coeficiente tienen coeficiente registral en todas las
  unidades elegibles y una suma elegible positiva;
- los conceptos por consumo tienen lectura inicial y final efectiva para todas
  las unidades participantes, o una estimación expresamente aprobada;
- las cuotas fijas conocen el número de unidades participantes;
- el total puede conciliarse al céntimo.

#### Cartas y correo

- existe un reparto conciliado vigente;
- cada destinatario tiene una identidad de propiedad y una carta generable;
- la ausencia de correo no bloquea la carta, pero sí su inclusión en un envío;
- una carta de más de una página o que no pueda renderizarse queda bloqueada con
  una explicación concreta.

El estado del expediente se derivará del informe y de la última ejecución
vigente. `ready_for_calculation` no podrá mantenerse cuando falten propietarios,
coeficientes o lecturas exigidas. La pantalla mostrará una tarjeta resumida,
por ejemplo: `Faltan propietarios`, `Faltan lecturas ACS: 62 propiedades` o
`Falta confirmar el período de 3 facturas`.

## 5. Aceptación privada de 658 y 644

El runner privado actual de la 658 se generalizará para aceptar un manifiesto
local por comunidad. El manifiesto indicará código, período, fuentes, Excel de
referencia opcional y controles esperados, pero permanecerá fuera de Git.

Cada ejecución:

1. crea una carpeta aislada y una SQLite nueva;
2. copia únicamente configuración pública y fuentes indicadas;
3. ejecuta el mismo análisis, confirmación, Excel, reparto y cartas que la app;
4. compara conteos y totales con controles independientes;
5. renderiza el Excel afectado y una muestra de cartas;
6. genera un informe saneado sin nombres, correos, texto documental ni rutas de
   origen.

### 5.1 Aceptación 658

- conservar las 62 unidades y distinguir viviendas, locales y garajes;
- utilizar lecturas efectivas, nunca consumos negativos;
- aplicar arrastre de la última lectura fiable a ceros o disminuciones;
- reconciliar ACS fijo y variable al céntimo;
- validar Excel, reparto completo y cartas de una página;
- mostrar histórico y comparación vecinal sólo con datos válidos.

### 5.2 Aceptación 644

- importar el listado completo de propietarios y `Enteros Participación`;
- conservar la suma registral observada y calcular la suma elegible por concepto;
- importar lecturas ACS y calefacción con límites de período correctos;
- configurar ACS fijo/variable, calefacción fija/variable y gastos
  extraordinarios presentes;
- producir Excel, reparto y cartas sin condicionales de código `644`.

Una diferencia entre la referencia manual y el programa se clasificará como
fuente incompleta, regla de negocio no configurada o defecto. Sólo el tercer
caso se corrige en código; nunca se ajusta un resultado para que coincida sin
trazabilidad.

## 6. Auditoría del reconocimiento documental

El auditor privado se ejecutará sobre las carpetas reales y conservará sólo
métricas agregadas y huellas. Sus umbrales de publicación son:

- al menos 90 % de facturas con texto y proveedor reconocido;
- menos de 60 decisiones de proveedor en el corpus de referencia;
- cero presupuestos, albaranes, informes o justificantes aplicados como factura;
- cero documentos asignados a otra comunidad;
- un PDF defectuoso no detiene el lote;
- una fuente cacheada no repite OCR con la misma versión.

Los nuevos proveedores se incorporarán al catálogo global por CIF/NIF y firmas
fuertes de factura. Los nombres de cliente, bancos, distribuidoras y textos
genéricos se mantendrán como exclusiones. Un formato nuevo de un proveedor
existente será otra familia de extracción, no otro proveedor.

## 7. Instalación portable

Se publicará primero una distribución Windows `onedir`. Un único ejecutable
autoextraíble no se usará porque complica los modelos OCR, las plantillas y el
diagnóstico de LibreOffice.

### 7.1 Separación de código y datos

El código y los recursos públicos serán de solo lectura. Una instalación tendrá
un directorio de datos seleccionado en el primer arranque. Si el usuario no lo
elige, se usará `%LOCALAPPDATA%\Regularizaciones`. La variable
`REGULARIZACIONES_HOME` permitirá una instalación portable controlada.

Dentro del directorio de datos existirán:

```text
config/      identidad, rutas y perfiles locales
data/        SQLite y caché de texto/OCR
fuentes/     copias archivadas por expediente
salidas/     Excel, cartas y borradores de correo
backups/     copias verificadas de SQLite y configuración
logs/        diagnósticos sin contenido documental
```

### 7.2 Primer arranque

El asistente solicitará nombre del despacho, logo opcional, firma, datos de
contacto, directorio de trabajo y ruta de LibreOffice. Después creará una base
verificada, configuración neutra y una copia inicial. Ningún dato de Meditrade
u otro despacho se heredará por defecto.

La configuración podrá exportarse sin base, fuentes ni credenciales para
instalar el programa en otro despacho. Las actualizaciones aplicarán migraciones
de base y conservarán una copia previa recuperable.

## 8. Preparación y envío de correo

El correo será un subsistema posterior a las cartas, con dos transportes:

1. **Borrador EML**, activado por defecto y sin credenciales. Genera un archivo
   estándar por destinatario con asunto, cuerpo y carta adjunta.
2. **SMTP**, opcional. Sólo se habilita cuando el despacho configura servidor,
   puerto, cifrado y remitente. La contraseña se guarda mediante el almacén de
   credenciales de Windows; si no está disponible se solicita por sesión.

Outlook, Gmail y otros clientes pueden abrir los EML. No se añadirá automatismo
COM ni OAuth específico mientras el transporte estándar cubra el flujo.

### 8.1 Modelo y estados

Una migración añadirá lotes y entregas de correo con:

- expediente, lote de cartas y propietario;
- destinatario normalizado;
- huellas de carta, asunto y cuerpo;
- estado `draft`, `ready`, `sending`, `sent`, `failed` o `skipped`;
- identificador de mensaje, error saneado e instantes;
- usuario que confirmó el envío.

La huella impedirá enviar dos veces la misma carta al mismo destinatario. Un
nuevo lote de cartas crea borradores nuevos y deja los anteriores en el
histórico.

### 8.2 Interfaz segura

La pantalla mostrará destinatarios válidos, ausentes y duplicados. El usuario
podrá abrir una muestra, excluir filas y generar todos los borradores. El botón
`Enviar correos` requerirá transporte configurado, una confirmación final con
conteo y que no exista otra ejecución activa. Los fallos se aislarán por
destinatario y podrán reintentarse sin reenviar los ya completados.

Las pruebas usarán un transporte falso. Ninguna prueba de unidad, aceptación o
empaquetado establecerá conexiones SMTP reales.

## 9. Integración y publicación

La rama funcional se integrará después de completar los bloques anteriores.
Antes del merge se preservarán `data/gestion.db`, preferencias, fuentes,
salidas y perfiles locales. Sólo se eliminarán archivos del sistema y worktrees
tras verificar su ruta exacta y que no contengan trabajo sin integrar.

La publicación seguirá este orden:

1. pruebas unitarias y de integración completas;
2. aceptación privada 658 y 644;
3. auditoría del catálogo dentro de umbrales;
4. construcción portable en una carpeta limpia;
5. prueba desde una base vacía y desde una copia migrada;
6. revisión visual de Excel y cartas;
7. merge a `master`, nueva ejecución completa y etiqueta de versión;
8. manual de instalación, operación, copia y recuperación.

No se forzará ni sobrescribirá `master`. Si la rama remota y la local difieren,
se reconciliarán mediante merge o rebase revisable, conservando el historial.

## 10. Manejo de errores y seguridad

- Toda escritura de salida usa un temporal y publicación atómica.
- Las operaciones de base relacionadas se ejecutan en una transacción.
- Los originales se archivan por huella y nunca se reescriben.
- Las rutas se resuelven dentro del directorio de instalación configurado.
- Los informes y logs no incluyen nombres, correos, texto OCR ni contenido de
  facturas.
- Las credenciales se referencian por identificador seguro y nunca se exportan.
- Un fallo de OCR, LibreOffice, una carta o un correo no invalida resultados
  anteriores correctos.
- Las copias de seguridad se verifican antes de sustituir una base o publicar
  una actualización.

## 11. Estrategia de pruebas

Cada comportamiento nuevo comenzará con una prueba que falle. La cobertura se
divide en:

- pruebas puras del informe de requisitos y transiciones;
- bases SQLite temporales con combinaciones de datos ausentes;
- pruebas del flujo Excel, reparto y cartas con perfiles sintéticos;
- pruebas del runner privado sin datos personales en los resultados;
- pruebas del catálogo con ejemplos saneados y auditoría agregada;
- pruebas de configuración portable y migración entre directorios;
- pruebas del generador EML, deduplicación y transporte SMTP falso;
- prueba completa desde base vacía hasta borradores de correo;
- renderizado visual de todos los Excel afectados y cartas representativas.

Los casos de borde incluyen cero frente a vacío, propiedades duplicadas,
coeficiente cero, falta de lectura inicial o final, períodos parciales, abonos,
reinicios de contador, direcciones de correo inválidas, reintentos y cierre de
la aplicación durante una operación.

## 12. Criterios finales de aceptación

La entrega se considerará completa cuando:

1. una base vacía pueda descubrir y crear comunidades desde una carpeta mixta;
2. el programa explique en una sola vista todos los requisitos que faltan;
3. la 658 produzca Excel, reparto y cartas que cuadren con sus controles reales;
4. la 644 complete el mismo flujo con ACS, calefacción y coeficientes propios;
5. el catálogo cumpla sus umbrales sobre el corpus privado;
6. otra instalación pueda configurar despacho y comunidades sin cambiar código;
7. las cartas puedan convertirse en borradores EML y enviarse por SMTP sólo con
   confirmación explícita;
8. todas las pruebas pasen desde `master` y la distribución portable arranque
   en un directorio limpio;
9. el manual permita instalar, operar, actualizar y recuperar la aplicación;
10. ninguna fuente privada, base, credencial o identidad de despacho quede
    incluida en Git o en el paquete público.

## 13. Exclusiones

No se implementan en este ciclo servicios externos de OCR, aprendizaje que
modifique reglas sin aprobación, envío por API propietaria de Gmail/Outlook,
facturación comercial del despacho ni migración automática de todos los datos
de un tercero. Esas capacidades podrán añadirse sobre las interfaces de texto,
proveedor y transporte ya definidas.
