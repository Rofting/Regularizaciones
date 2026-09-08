# Informe — Tarea 4: asistente visual v3 y guía

## Cambios realizados

- `core/expedient_ui.py`
  - añade la función pura `onboarding_step_route(state)`;
  - añade `open_community_onboarding_dialog(app)` con las cinco etapas
    identidad, fuentes, detección, confirmaciones y resumen;
  - ofrece selectores separados para propietarios CSV, una o más lecturas
    Excel/PDF y facturas PDF opcionales;
  - delega el análisis exclusivamente en `analyse_sources` y la publicación
    exclusivamente en `confirm_onboarding`;
  - ejecuta análisis y publicación con `_en_hilo` y `_estado`, el mecanismo de
    trabajo y progreso ya existente;
  - actualiza comunidad, período y expediente al finalizar el alta.
- `core/app.py`
  - el botón `+ Comunidad` abre un selector entre `Alta guiada desde fuentes`
    y `Registro rápido`;
  - el formulario anterior se conserva sin cambios funcionales bajo
    `_registro_rapido_comunidad`.
- `tests/test_expedient_ui.py`
  - cubre el enrutado sin crear controles Tk;
  - impide saltar identidad o fuentes y bloquea el resumen mientras falten
    respuestas obligatorias.
- `GUIA_PROCESAR_TODO.txt` y `docs/validacion-expedientes.md`
  - documentan las cinco etapas, propietarios + lecturas como mínimo,
    lecturas Excel/PDF alternativas, facturas y Excel manual opcionales, y el
    carácter local/no versionado de fuentes, base, perfiles y plantillas.
- Dependencia autorizada de la tarea: se generalizaron en `task-3-report.md`
  cinco referencias al intérprete local mediante
  `$env:REGULARIZACION_PYTHON`. No se alteró ningún dato funcional ni prueba.

## Decisiones

- Se reutilizaron paleta, tipografía, radios, divisores y jerarquía v3: panel
  sobrio, verde petróleo como único acento y carril numerado porque el alta es
  una secuencia real.
- Se mantuvo un único panel de contenido por etapa para evitar un rediseño de
  la aplicación o una colección de tarjetas genéricas. No se añadió marca ni
  referencia a un despacho concreto.
- La lista de propietarios se limita a CSV porque es el contrato que valida
  actualmente `analyse_sources`; la UI no implementa lectores alternativos.
- Los módulos detectados se presentan como confirmaciones explícitas que el
  usuario puede desactivar. Las preguntas ambiguas empiezan vacías y las
  obligatorias bloquean el resumen.
- El período inicial es opcional, pero si se empieza a rellenar deben estar
  presentes nombre, fecha inicial y fecha final.
- Los valores de Tk se capturan en el hilo principal antes de ejecutar el
  worker. Cambiar identidad o fuentes invalida el borrador para impedir
  publicar un análisis obsoleto.

## TDD y pruebas reales

- RED: `-m unittest tests.test_expedient_ui -q` produjo 5 errores por ausencia
  de `onboarding_step_route` antes de implementar producción.
- GREEN UI/onboarding: `-m unittest tests.test_expedient_ui tests.test_community_onboarding -q`
  ejecutó 22 pruebas, todas correctas.
- Primera suite completa: `-m unittest discover -s tests -t . -q` ejecutó 178
  pruebas y detectó un único fallo de privacidad en el informe de Tarea 3 por
  una ruta de usuario versionada.
- RED/GREEN de privacidad: la guardia
  `Private658ValidationGuardsTest.test_tracked_documents_do_not_contain_known_private_identifiers_or_local_paths`
  falló antes de generalizar las rutas y pasó después (1 prueba correcta).
- Suite completa tras la corrección: 178 pruebas correctas en 28,485 s.
- `-m compileall -q core tests`: código de salida 0.
- `git diff --check`: código de salida 0; sólo avisos informativos de futura
  normalización LF/CRLF de Git.

Las ejecuciones se realizaron con el intérprete del entorno
`..\..\.venv-fase1` y `PYTHONUTF8=1`.

## Auto-revisión

- Se comprobó que la nueva UI sólo llama a `community_onboarding.analyse_sources`
  y `community_onboarding.confirm_onboarding` para el flujo de alta.
- No hay lectura de Excel/PDF ni generación, archivado o escritura de perfiles
  reimplementados en los controles.
- Se revisaron rutas hacia atrás: cualquier edición posterior a un análisis
  fuerza uno nuevo y la función pura redirige un resumen incompleto a la etapa
  que falta.
- El registro rápido conserva la misma implementación; sólo cambió su punto de
  entrada para poder presentar las dos opciones.
- `.gitignore` ya excluye `data/`, plantillas de comunidades y perfiles
  runtime.

## Preocupaciones

- No se realizó una sesión visual interactiva con ventanas Tk en esta ejecución;
  el aspecto se revisó contra los componentes y tokens v3 existentes. El
  enrutado, el servicio y la compilación sí quedan cubiertos por pruebas
  automatizadas.

## Ronda 1 de corrección — revisión sobre `18b94ee`

### Hallazgos y cambios

- P1: la devolución del worker establecía `detected`, pero `render("detection")`
  lo sobrescribía. La etapa real es ahora `detection` y el éxito se representa
  con `analysis_complete`. El enrutado acepta también el alias `detected`.
  El botón real de detección alcanza confirmaciones, exige las respuestas
  obligatorias, llega al resumen y permite publicar.
- P2: `has_sources` indica solamente que están elegidos propietarios y lecturas.
  Un análisis completado se distingue mediante `analysis_complete`. Volver a
  identidad y continuar conduce a fuentes, donde puede iniciarse/reintentarse
  el análisis; una pantalla de detección sin worker ni borrador válido vuelve
  a fuentes. Los errores no dejan una pantalla de espera infinita.
- P2: cada análisis invalida borrador, respuestas y variables anteriores antes
  de mostrar progreso. Una generación identifica la petición; los cambios de
  identidad (incluido el período) la invalidan y un callback desfasado no instala
  resultados. Mientras trabaja, los callbacks de selección, navegación,
  confirmación y publicación están bloqueados. La publicación captura borrador
  y respuestas antes de lanzar el worker, sin leer variables Tk desde él.
- P2: el resumen y el progreso explican que sin período inicial no se archivan
  fuentes. El mensaje de éxito usa el número real de documentos devueltos por
  el servicio. El contrato existente se llama `OnboardingResult.source_document_ids`,
  no `documents_registered`; se utiliza su longitud sin modificar el servicio.
- P3: `Quitar facturas` vacía la selección opcional, actualiza su etiqueta e
  invalida el análisis y las respuestas previas.
- Se conservan el registro rápido, el mecanismo `_en_hilo`/`_estado`, los
  servicios de dominio y los componentes y tokens visuales v3.

### Pruebas y comandos reales

Las ejecuciones usan `PYTHONUTF8=1` y el intérprete
`../../.venv-fase1/Scripts/python.exe`, desde este worktree.

- RED: `-m unittest tests.test_expedient_ui -q` ejecutó 12 pruebas y produjo
  8 fallos (incluidos subcasos): transición real bloqueada, retorno tras error,
  resultado desfasado, navegación del borrador anterior, fuentes editables,
  ausencia de retirada de facturas y mensajes de archivado.
- Tras corregir la transición, la misma orden dejó 6 fallos que confirmaban
  específicamente los demás problemas, incluidos resumen y mensaje de éxito.
- Se añadieron 10 pruebas de callbacks del diálogo y una guardia pura. Los
  dobles sustituyen sólo controles/variables Tk, selectores, servicios y la
  planificación del worker. Se ejecutan `open_community_onboarding_dialog`,
  los comandos entregados a sus botones y los callbacks `after` reales.
  Ninguna prueba crea Tk. Se cubren también reanálisis exitoso con respuestas
  nuevas, publicación antigua bloqueada, ausencia de preguntas y cambio de
  período entre la finalización del worker y la devolución a la interfaz.
- GREEN focal: `-m unittest tests.test_expedient_ui tests.test_community_onboarding -q`
  ejecutó 33 pruebas correctas en 1,856 s.
- Suite completa: `-m unittest discover -s tests -t . -q` ejecutó 189 pruebas
  correctas en 27,066 s.
- `-m compileall -q core tests`: salida 0.
- `git diff --check`: salida 0; avisos informativos de normalización LF/CRLF.

### Preocupaciones restantes

- No se abrió una sesión gráfica Tk; los callbacks y el contenido se verifican
  sin ventanas, y los servicios tienen su batería de integración independiente.
- El archivado requiere período inicial por contrato del servicio existente;
  ahora queda explicado antes de crear y el resultado informa lo registrado.
