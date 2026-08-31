# Informe — Tarea 3: exportación oficial atómica

## Resultado

- Añadidos recalculador LibreOffice con timeout/error comprensible y doble determinista para pruebas.
- Añadidos validación de hojas, áreas de impresión, fórmulas, errores, totales y huella estructural.
- Añadido exportador por expediente: valida condiciones, registra plantilla/hash/ejecución, usa copia temporal, escribe solo el layout declarativo, recalcula, valida y publica atómicamente con copia de seguridad.
- El generador heredado conserva su ruta anterior y delega en el exportador al recibir `id_case` elegible.
- El perfil 658 incorpora únicamente el layout declarativo necesario para los cinco módulos activos.

## Corrección adicional de alcance

Un gasto extraordinario configurado como concepto opcional puede estar ausente en un período. Se añadió prueba RED y la conciliación lo compara contra cero, preservando la sección vacía; los parámetros exigidos siguen bloqueando la exportación.

## Evidencia

- RED heredado de la tarea: el módulo de exportación no existía.
- RED adicional: ausencia de `extraordinary_expense_actual` producía `ExportBlockedError` aunque el concepto era opcional.
- GREEN focalizado: `9` pruebas de exportación correctas.
- GREEN global: `81` pruebas correctas.
- Compilación: módulos de exportación/perfil y prueba correctos.
- `git diff --check`: correcto; solo avisos no bloqueantes LF/CRLF de Windows.

## Privacidad y operación

Las pruebas crean un libro 658 sintético en un directorio temporal. No se abrió, copió ni versionó la plantilla privada real. El servicio de producción exige LibreOffice al ejecutar una exportación oficial y falla de forma segura si no está disponible.

## Continuidad

El agente inicial agotó su cuota antes de preparar commit/informe. El coordinador inspeccionó los cambios pendientes, añadió el caso opcional en TDD y ejecutó la verificación completa antes de este cierre.
