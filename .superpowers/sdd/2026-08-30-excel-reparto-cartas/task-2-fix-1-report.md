# Informe — Tarea 2 / corrección de revisión 1

Base revisada: `5dbd68f`.

## Correcciones

1. Las cabeceras de lecturas se interpretan como mes/año. Solo se usan las columnas que coinciden con el inicio y fin del expediente. Un rango ajeno crea `INCOMPATIBLE_DATE`; un extremo ausente crea `MISSING_READING_RANGE`; en ambos casos no se persisten lecturas engañosas.
2. Facturas, consumos, términos y coeficientes distinguen `0` explícito de valor ausente o ilegible. Una factura incompleta no se confirma y un propietario sin coeficiente no se crea. El cero explícito continúa siendo válido.
3. La migración 4 añade `case_import_batches`, que enlaza un lote físico con cada expediente, período y clase de fuente. El mismo listado por hash se deduplica en `import_batches`, pero ambos expedientes conservan su relación lógica y trazabilidad.
4. Tras registrar una fuente se verifica el SHA-256 de la copia archivada. El maestro, CSV de propietarios y libro de lecturas se leen exclusivamente desde esa copia. También se corrigió el saneado del nombre archivado para conservar la extensión; los bytes verificados del maestro se abren sin depender del original.

## Evidencia TDD

- RED: una lectura 2020–2021 se persistía dentro del caso 2025–2026 (`4 != 0`).
- RED: una factura sin término fijo seguía insertándose (`3 != 2`).
- RED: un coeficiente vacío creaba propietario con `0.0`.
- RED: faltaba `case_import_batches` al reutilizar el listado en dos períodos.
- RED: mutar el original tras archivarlo cambiaba el total persistido (`999 != 100`).
- RED: una copia archivada manipulada no era rechazada.
- GREEN focalizado: 14 pruebas del importador correctas.
- GREEN global fresco: 68 pruebas correctas en 7.277 s.
- `py_compile`: cinco módulos afectados correctos.
- `git diff --check`: correcto; únicamente avisos LF/CRLF de Windows.

No se abrió ni copió ninguna fuente privada y no se inició Tk. Las pruebas usan nombres, correos y libros temporales inventados.

## Entrega

La corrección queda en el único commit que contiene este informe, con mensaje `Corrige riesgos de importación auditable`. El SHA efectivo se comunica junto con la entrega, ya que un commit no puede incluir de forma autorreferencial su propio SHA.
