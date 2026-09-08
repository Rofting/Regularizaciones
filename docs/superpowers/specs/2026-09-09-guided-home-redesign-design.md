# Rediseño de la pantalla principal: flujo guiado

## Objetivo

Sustituir la pantalla principal actual por una interfaz que haga evidente el
trabajo siguiente de un expediente de regularización. Debe servir a cualquier
despacho y comunidad, no llevar identidad de un despacho concreto, y conservar
el flujo y los datos ya implementados: fuentes, revisión humana, reparto,
Excel oficial y cartas.

La dirección aprobada es **A: flujo guiado, limpio y enfocado**. El usuario
trabaja sobre un único expediente cada vez y no tiene que interpretar varios
paneles equivalentes ni decidir qué botón es el correcto.

## Principios visuales

1. Una acción principal por estado. La zona central muestra el siguiente paso
   necesario, su contexto y un único botón primario.
2. Contexto compacto. Comunidad, período y expediente se agrupan en una
   cabecera estrecha; no ocupan una tarjeta dominante.
3. Progreso persistente y sobrio. Una navegación lateral de cuatro pasos
   muestra Fuentes, Validar, Reparto y Cartas. No se usan cinco tarjetas verdes
   para explicar el mismo flujo.
4. Lenguaje neutral. Título e identidad visual de "Regularizaciones", sin
   referencia a Meditrade u otro despacho.
5. Accesibilidad funcional. Todas las acciones conservan texto explícito,
   bordes visibles para controles secundarios y contraste suficiente en claro
   y oscuro.

## Sistema de color y controles

La paleta existente conserva su turquesa, pero se usa con semántica estricta:

| Elemento | Tratamiento |
| --- | --- |
| Acción para avanzar o crear | Turquesa sólido |
| Guardar, confirmar o terminar correctamente | Verde sólido |
| Acción secundaria | Fondo transparente, borde estable y texto neutro |
| Incidencia pendiente | Ámbar, solo como aviso |
| Error bloqueante o eliminación | Rojo, solo cuando corresponda |
| Información pasiva | Gris suave o turquesa muy claro |

La tipografía se normaliza a una familia de sistema legible y títulos más
contenidos. No habrá cajas de color saturado con texto descriptivo cuando una
etiqueta y un divisor basten.

## Estructura de la pantalla

### Cabecera de contexto

Contiene, en una única fila adaptable:

- selector de comunidad y botón secundario "Nueva comunidad";
- selector de período y acceso a "Gestionar períodos";
- selector de expediente o botón primario "Crear expediente";
- acceso discreto a Ajustes y al cambio de tema.

En tamaños estrechos los selectores se apilan, pero las acciones siguen siendo
visibles y no se recortan.

### Navegación del expediente

La columna lateral utiliza cuatro pasos numerados:

1. Fuentes
2. Validar
3. Reparto
4. Cartas

Cada paso muestra estado: pendiente, activo, bloqueado, listo o terminado. Un
paso no permitido sigue pudiéndose leer, pero explica la condición que falta;
no aparenta ser una acción disponible.

### Área de trabajo

La zona principal cambia según el paso activo:

| Paso | Contenido | Acción principal |
| --- | --- | --- |
| Fuentes | Lista de documentos, recuento por clase, progreso de carga y errores por archivo | Añadir carpeta / Añadir archivos |
| Validar | Incidencias abiertas agrupadas por fuente, explicación entendible y acceso al original | Resolver siguiente incidencia |
| Reparto | Resumen de facturas, lecturas y comprobaciones aprobadas | Calcular reparto |
| Cartas | Resumen final por propietario y resultado de generación | Generar cartas |

En todos los pasos se ve un bloque de estado breve: comunidad, período,
expediente, fuentes incorporadas, incidencias abiertas y siguiente acción. El
registro técnico queda disponible en un desplegable o panel secundario; no
compite con la tarea diaria.

## Comportamientos conservados y expuestos

- **Alta de comunidad:** desde la cabecera abre el formulario existente y, al
  guardar, confirma qué registro se escribió en la base de datos activa.
- **Períodos:** sigue admitiendo intervalos no anuales. Todas las validaciones
  y cálculos se limitan a la fecha inicial y final elegidas.
- **Fuentes:** se conservan "Añadir archivos" y "Añadir carpeta"; esta última
  recorre subcarpetas, acepta PDF/XLS/XLSX/CSV y omite archivos ocultos,
  temporales o no compatibles. Durante la ingestión muestra `X de Y`.
- **Incidencias:** el diálogo mantiene "Abrir archivo" y explica el campo que
  falta, dónde buscarlo, formato, ejemplo y motivo. Un reinicio de contador
  continúa exigiendo estimación aprobada y trazable.
- **Gates de seguridad:** Excel, reparto y cartas no se habilitan mientras
  falten validaciones, lecturas obligatorias o una estimación aprobada.
- **Base de datos y rutas:** se concentran en Ajustes, separados del trabajo
  normal. Se puede elegir una base de prueba antes de crear una comunidad.

## Arquitectura de implementación

No cambia el contrato de base de datos ni los servicios de negocio. La capa
visual sigue siendo CustomTkinter y se reorganiza alrededor de una vista de
estado del expediente.

- `core/app.py`: reemplaza la composición de `_crear_ui_v3` por cabecera de
  contexto, navegación de pasos y área de trabajo; reutiliza los comandos y
  refrescos existentes.
- `core/expedient_ui.py`: conserva los diálogos de fuentes e incidencias y
  expone los resúmenes que necesita cada paso, sin duplicar reglas de negocio.
- `core/ui_moderna.py`: formaliza tokens de estado y componentes reutilizables
  para botón secundario, estado de paso y tarjeta informativa.
- Servicios de ingestión, revisión, reparto, exportación Excel y cartas:
  permanecen como fuente única de verdad; la interfaz solo los invoca y
  representa sus resultados.

## Estados y errores

- Sin comunidad, período o expediente: se presenta solo la acción de crear o
  seleccionar lo que falte.
- Operación en curso: el botón principal se desactiva y muestra avance con
  nombre de archivo o etapa, sin bloquear el registro de actividad.
- Error de una fuente: el resto continúa; el resultado enumera qué archivos
  fallaron y permite volver a abrir o reintentar.
- Gate bloqueado: muestra una frase directa y un enlace/botón al paso que lo
  resuelve, en vez de un botón que falla tarde.

## Compatibilidad

Se conservan la comunidad 658, las demás comunidades configuradas, períodos
existentes, bases de datos de prueba y preferencias de tema. El rediseño no
migra ni reescribe datos. Las rutas de configuración actuales seguirán
funcionando.

## Verificación

1. Pruebas unitarias para derivar paso activo, estado y acción siguiente desde
   un expediente con y sin bloqueos.
2. Pruebas de interfaz para asegurar la presencia de acciones principales y
   secundarias en cada estado.
3. Pruebas existentes de ingestión de carpeta, guías de incidencia, alta de
   comunidad, reparto, exportación y cartas.
4. Compilación Python y suite completa antes de integrar.
5. Revisión visual manual en modo claro y oscuro, a resolución de escritorio
   y tamaño reducido, verificando que los controles no pierdan borde ni se
   recorten.
