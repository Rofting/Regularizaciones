# Informe — corrección del bootstrap por perfil

## Resultado

- La importación de facturas prioriza ahora las filas y columnas declaradas
  por `workbook_layout` cuando contiene filas sustantivas, sin depender de
  alias frágiles de cabeceras de Office. Los perfiles sin layout conservan el
  lector genérico como alternativa.
- El perfil declarativo 658 incorpora los extremos de suministro confirmados:
  inicio GAS `G` e inicio AGUA `F`; no se usan duraciones ni ratios para
  inventar fechas.
- Cuando el bloque económico genérico no puede interpretarse, se leen los
  `parameter_cells` declarados. Se conserva tanto el valor calculado como el
  valor/fórmula original y un cero explícito sigue siendo un importe válido.
- Una lectura inicial puede provenir exclusivamente del mes anterior si el
  expediente empieza el día 1, la columna es única y el cierre coincide de
  forma exacta. La celda y cabecera quedan auditadas. Los reinicios de
  contador no se modifican: permanecen como incidencias pendientes de
  estimación aprobada.

## Pruebas

- RED observado: el libro saneado con cabeceras descriptivas insertaba `0/3`
  facturas y las lecturas con base en el mes anterior insertaban `0/4`.
- GREEN focal: `tests.test_excel_bootstrap_importer` — 24 pruebas correctas.
- GREEN global: 131 pruebas correctas.
- Compilación: `core/excel_bootstrap_importer.py` y `core/excel_profiles.py`.
- `git diff --check`: correcto (solo avisos no bloqueantes LF/CRLF de Windows).

## Privacidad

No se abrió, copió ni registró ningún archivo real durante esta corrección.
Las nuevas pruebas crean únicamente libros y propietarios sintéticos en
directorios temporales.
