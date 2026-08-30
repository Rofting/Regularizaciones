# Informe — Tarea 2 / corrección de revisión 2

Base revisada: `f1a6695`.

## Correcciones

1. El lector de facturas distingue filas realmente vacías y filas de suma de aquellas que contienen datos. Si una fila de datos tiene fecha de factura vacía se crea `MISSING_REQUIRED_FIELD`; si es ilegible o queda fuera del expediente se crea `INCOMPATIBLE_DATE`. Los extremos inicial y final de suministro son obligatorios, deben ser fechas válidas y mantener orden cronológico. La factura afectada no se confirma.
2. El atajo idempotente comprueba ahora si el lote ya está enlazado al mismo expediente, período y clase. Un reintento del mismo caso continúa devolviendo inmediatamente el lote existente. Si el mismo hash se usa en otro caso, el maestro revalida su período, propietarios vuelve a validar el listado y lecturas reinterpreta sus cabeceras para el nuevo rango. El documento nuevo queda `validated` o `under_review` según sus propias incidencias y todos los resultados devuelven el período del caso actual.

## Evidencia TDD

- RED: una fila GAS con fecha ilegible se omitía sin incidencia.
- RED: una fila GAS sin inicio y con fin ilegible se insertaba igualmente.
- RED: reutilizar el maestro en otro caso devolvía el período del lote anterior (`1 == 1`) y no creaba incidencia temporal.
- RED: documentos complementarios reutilizados quedaban `registered` y las lecturas no se revalidaban para el período nuevo.
- GREEN focalizado: 18 pruebas del importador correctas.
- GREEN global fresco: 72 pruebas correctas en 8.059 s.
- `py_compile`: importador y pruebas correctos.
- `git diff --check`: correcto; solo avisos LF/CRLF de Windows.

No se abrió ninguna fuente privada y no se inició Tk. Todas las fuentes de prueba son temporales y usan identidades inventadas.

## Entrega

La corrección queda en el único commit que contiene este informe, con mensaje `Revalida fuentes reutilizadas por periodo`. El SHA efectivo se comunica junto con la entrega porque un commit no puede contener autorreferencialmente su propio SHA.
