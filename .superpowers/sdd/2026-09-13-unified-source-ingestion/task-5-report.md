# Task 5 — Acciones y contexto en la interfaz guiada

## Resultado

La ventana de fuentes analiza cada documento antes de incorporarlo mediante
`analyse_source` y `add_analysed_document_to_case`. Ya no hay un selector global
que convierta accidentalmente lecturas o propietarios en facturas. Al terminar
se muestra un resumen agrupado por clasificación, junto con nuevas fuentes,
duplicadas y errores. Se conserva la selección de archivos o carpetas y el
filtrado de ocultos, temporales y extensiones incompatibles.

Las fuentes desconocidas ofrecen un selector manual únicamente en su incidencia
de clasificación. La decisión exige un motivo auditable y se guarda mediante
`resolve_issue`. A continuación se reanalizan las fuentes para aplicar la
clasificación efectiva y crear los campos requeridos antes de comprobar la
preparación del expediente. Una transacción exterior mantiene juntas la decisión,
el reanálisis y esa comprobación: un fallo no deja una corrección parcial ni una
incidencia cerrada prematuramente.

El botón secundario **Reanalizar fuentes** se incorpora a la navegación del
expediente de la pantalla guiada. Comprueba el expediente activo y su pertenencia
a la comunidad, trabaja sobre las copias archivadas y presenta el resumen de
clasificaciones persistidas. No se recupera ninguna pantalla heredada.

En **Ajustes**, **Nueva base segura** requiere confirmación explícita con «No»
como opción predeterminada e impide iniciar el reinicio mientras hay trabajo en
curso. Usa `reset_database`, crea los respaldos en `data/backups` junto a la base
seleccionada y solo limpia las selecciones y devuelve el flujo al inicio después
de recibir el resultado correcto. Informa de la ruta exacta de la copia tanto en
el registro como en el diálogo final. Los errores mantienen el contexto de la UI.
La inicialización ejecuta las tablas, índices y migraciones de `gestor_bd` y
cierra explícitamente la conexión SQLite antes de publicar el reemplazo, un
requisito comprobado en Windows.

Las incidencias tienen **Ver contexto** además de **Abrir archivo**. Se consulta
`extraction_candidates.source_context`, dando preferencia al campo de la
incidencia y usando otro fragmento del mismo documento cuando un campo ausente
no tiene candidato propio. Se muestran página y fragmento PDF, hoja/celda de
Excel o texto histórico. Si no hay metadatos útiles se indica cómo consultar el
original. No se altera el modelo público `ReviewIssue`.

## TDD y verificación

Primera ejecución dirigida, antes de la implementación: 36 pruebas, ocho fallos
esperados. La ingesta anterior clasificaba los tres documentos de la muestra
como facturas; faltaban el resumen, el contexto, la clasificación manual, el
reanálisis y el reinicio de UI.

Una segunda ronda añadió regresiones de persistencia. Fallaron la clasificación
manual interrumpida (dejaba la incidencia resuelta) y el reinicio completo usando
`gestor_bd.crear_bd` (su conexión abierta impedía reemplazar la base en Windows).
Se corrigieron la transacción exterior y el cierre explícito del inicializador.

Verificación final:

```text
..\..\.venv-fase1\Scripts\python.exe -m unittest tests.test_expedient_ui -v
39 pruebas — OK

..\..\.venv-fase1\Scripts\python.exe -m unittest tests.test_expedient_ui tests.test_source_analysis tests.test_document_review tests.test_expedient_flow tests.test_database_reset -v
93 pruebas — OK

git diff --check
Sin errores de espacios; solo avisos habituales de conversión LF/CRLF.
```

Las nuevas pruebas ejecutan callbacks reales de la UI con sustitutos de dibujo
Tk. Ingesta, registro, revisión y consultas se ejecutan contra SQLite temporal;
solo se sustituye la extracción PDF de la muestra. El reinicio tiene pruebas de
cancelación, fallo, éxito y una integración real que comprueba base vacía y copia
con los datos anteriores. No se ha reiniciado ninguna base real del proyecto.

## Límites y observaciones

- El contexto está disponible como texto legible, sin depender de un visor PDF o
  de Excel. **Abrir archivo** mantiene la asociación nativa; no garantiza saltar
  a la página o celda. No se ha realizado una captura visual con una ventana Tk
  nativa durante esta tarea.
- La clasificación manual reutiliza el reanálisis del expediente completo. Es
  síncrona dentro del diálogo de corrección, como las demás acciones de guardar;
  el botón general de reanálisis y la ingesta se ejecutan en segundo plano.
- Solo se han modificado `core/expedient_ui.py`, `core/app.py`,
  `tests/test_expedient_ui.py` y este informe. No hay código específico de una
  comunidad ni cambios en las APIs de las tareas anteriores.
