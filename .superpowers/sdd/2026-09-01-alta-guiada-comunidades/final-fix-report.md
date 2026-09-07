# Corrección final — alta guiada de comunidades

Fecha: 2026-09-07

Rama: `feature/community-onboarding`
HEAD de partida: `1befd2f67c734348ecca979600fb0756a87354f0`

## Diagnóstico

El análisis detectaba servicios únicamente en las fuentes clasificadas como
lecturas. Después, la UI mantenía dos canales de decisión incompatibles: la
respuesta `service` y checkboxes `module:*`. Finalmente,
`build_profile_payload` ignoraba `service` y `reading_column`, y generaba un
único concepto `acs` o `calefaccion`; el reparto sólo reconoce conceptos de
consumo `acs_*` y `heating_*`, y el exportador sólo escribía lecturas ACS.

## RED

Las regresiones se escribieron antes de modificar producción y se ejecutaron
contra `1befd2f`:

- `python -m unittest tests.test_community_onboarding -q`: 22 pruebas,
  6 fallos y 1 error. Reprodujo factura ignorada para detección, ausencia sin
  preguntas obligatorias, respuestas sin efecto en el payload y los dos flujos
  de reparto vacíos/no canónicos.
- `python -m unittest tests.test_expedient_ui -q`: 17 pruebas, 1 fallo y
  1 error. Reprodujo los checkboxes paralelos y que publicación recibía el
  diccionario crudo en vez de la configuración normalizada.
- La revisión de bindings añadió una segunda regresión previa al ajuste:
  las 2 pruebas E2E ACS/calefacción fallaron porque los cuatro
  `parameter_cells` canónicos de cada servicio estaban ausentes.
- La corrección de una detección inequívoca se probó también antes del cambio:
  2 pruebas dieron 1 error y 1 fallo porque la lista sólo permitía aceptar el
  servicio propuesto o descartarlo, y no elegir el otro servicio canónico.

## Decisiones de implementación

- `resolve_onboarding_configuration` es la única normalización de las
  respuestas. Resumen, payload y publicación consumen el mismo objeto
  inmutable.
- Toda detección solicita una decisión explícita de servicio; cuando no hay
  evidencia ofrece ACS, calefacción, ambos o `NO_APLICA`. La ausencia o
  ambigüedad de columna crea una pregunta obligatoria.
- Las facturas participan en la evidencia de servicio, pero las columnas se
  extraen exclusivamente de fuentes `meter_reading_*`.
- La configuración persistida conserva decisión, módulos activos, columna,
  bindings por módulo y huellas/nombres relativos de las fuentes. El cargador
  valida el contrato; el perfil público 658 puede seguir omitiéndolo.
- ACS publica `acs_fixed` (partes iguales) y `acs_variable` (consumo).
  Calefacción publica `heating_fixed` y `heating_variable` con los mismos
  contratos de fuentes `period_parameters.*_actual/*_billed` que reparto.
- La plantilla canónica y el exportador escriben y concilian lecturas e
  importes de ambos tipos. El resultado de calefacción conserva unidad `kWh`.
- Una plantilla de salida recién creada, con todas sus celdas económicas
  vacías, no se interpreta como un maestro histórico incompleto durante el
  bootstrap; si contiene cualquier importe, mantiene el parser y sus
  incidencias obligatorias normales.
- Se eliminaron los checkboxes `module:*`; `service` es la única decisión que
  determina módulos activos.

## GREEN y verificación

- Focales de dominio/UI/perfiles/exportación/reparto:
  `python -m unittest tests.test_community_onboarding tests.test_expedient_ui tests.test_excel_profiles tests.test_excel_bootstrap_importer tests.test_excel_export_service tests.test_case_distribution -q`
  → 105 pruebas, OK.
- Extremo a extremo ACS y calefacción:
  `python -m unittest tests.test_community_onboarding.CommunityOnboardingTest.test_acs_onboarding_profile_reaches_canonical_export_and_distribution tests.test_community_onboarding.CommunityOnboardingTest.test_heating_onboarding_profile_reaches_canonical_export_and_distribution -v`
  → 2 pruebas, OK.
- Batería completa:
  `python -m unittest discover -s tests -t . -q`
  → 196 pruebas, OK.
- `python -m compileall -q core tests` → exit 0.
- `git diff --check` → exit 0; sólo avisos de conversión LF/CRLF de Git.

## Alcance y preocupaciones

No se modificó el JSON público `658_acs_v1.json`, ni la ruta de registro
rápido. Las pruebas de UI usan dobles y no crean Tk. No se añadieron fuentes,
bases de datos, plantillas privadas ni perfiles runtime al control de
versiones. La configuración guiada canónica de contador queda deliberadamente
limitada a ACS y calefacción; otros suministros conservan los flujos ya
existentes y requieren una ampliación de contrato separada.
