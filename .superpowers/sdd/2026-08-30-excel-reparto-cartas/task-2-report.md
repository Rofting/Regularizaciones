# Informe — Tarea 2: bootstrap auditable del modelo 658

## Archivos

- `core/excel_bootstrap_importer.py`: importación maestra y de fuentes complementarias.
- `core/document_review.py`: alta idempotente de incidencias genéricas.
- `tests/test_excel_bootstrap_importer.py`: ocho pruebas con fuentes temporales saneadas.

## Contratos añadidos

- `BootstrapImportResult` e `import_master_excel(...)`.
- `CompanionImportResult` e `import_companion_sources(...)`.
- Archivo previo de cada fuente con SHA-256 y lotes idempotentes por hash.
- Trazabilidad por hoja/celda para facturas, componentes, parámetros, propietarios y lecturas.
- Facturas normalizadas con identidad estable; un libro modificado solo añade facturas/componentes nuevos y no sobrescribe los ya confirmados.
- Incidencias por campo requerido vacío, cabecera desconocida, fecha incompatible, total descuadrado, propietario sin correspondencia, tramo de lectura incompleto y reinicio de contador.
- Una lectura final inferior a la inicial conserva el valor informado, usa `contador_averiado` y bloquea el expediente hasta revisión; nunca se convierte en consumo cero.
- Las importaciones conservan propietarios existentes, correcciones manuales e incidencias resueltas.

## Esquema y compatibilidad

No se añadió migración: `import_batches`, `source_values`, `invoice_components` y `period_parameters` ya estaban presentes en las migraciones aprobadas de las tareas anteriores.

El lector previo de propietarios actualizaba filas existentes y el lector previo de lecturas convertía determinados descensos en reinicio anual o arrastre estimado. Ninguno era compatible con el requisito vinculante de preservar datos confirmados y bloquear todo descenso hasta estimación aprobada. Por ello se reutilizan sus formatos, pero la persistencia se realiza en el nuevo importador con semántica conservadora. Se admiten lecturas `.xls` y `.xlsx`; el listado de propietarios de este contrato es CSV con separador `;`.

Cada lote nuevo guarda una fotografía completa de sus celdas para auditoría. La regla «solo añade valores nuevos» se aplica a las entidades normalizadas: no duplica ni sobrescribe facturas, componentes, propietarios o lecturas existentes.

## TDD y verificación

- RED inicial: `ModuleNotFoundError: No module named 'excel_bootstrap_importer'`.
- RED adicional: gas sin kWh se etiquetaba incorrectamente como kWh pese a disponer de m³.
- GREEN focalizado: 8 pruebas del importador, todas correctas.
- Suite completa: 62 pruebas, todas correctas.
- Compilación: `excel_bootstrap_importer.py`, `document_review.py` y su prueba compilan correctamente.
- `git diff --check`: correcto; solo aparece el aviso esperado LF/CRLF de Windows.
- Fuentes: exclusivamente libros y listados temporales con identidades inventadas; no se abrió ni versionó ninguna fuente privada.

## Commit

Esta entrega queda contenida en un único commit con el mensaje `Importa fuentes maestras de Excel de forma auditable`.

## Corrección de revisión — layout obligatorio en onboarding

Se cerró el hueco que permitía a un perfil con `onboarding_configuration`
omitir `workbook_layout`: el perfil quedaba validado con un mapeo vacío y el
importador posterior podía informar `MISSING_REQUIRED_FIELD` contra un maestro
sin las celdas canónicas. Ahora esos perfiles deben aportar un layout no vacío;
después se validan los `parameter_cells` canónicos y el marcador
`bootstrap_template` con estado `fresh_onboarding` y SHA-256. Los perfiles
públicos que no tienen onboarding siguen pudiendo no declarar dicha
configuración.

### RED/GREEN y verificación

- RED: `test_profile_validation_requires_nonempty_layout_for_onboarding`
  demostró que sin layout no se levantaba `ValueError`; con `{}` el rechazo era
  solo incidental por una clave estructural ausente.
- GREEN focalizado: 2 pruebas correctas, incluida la carga del perfil público
  `658_acs_v1` sin `onboarding_configuration`.
- Suite completa: 215 pruebas correctas.
- Compilación: `core/excel_profiles.py`,
  `tests/test_community_onboarding.py` y `tests/test_excel_profiles.py`
  compilan correctamente.
