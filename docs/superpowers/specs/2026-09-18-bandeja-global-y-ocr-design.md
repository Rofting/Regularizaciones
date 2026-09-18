# Bandeja global, OCR y continuidad de reparto

## Objetivo

Permitir que un despacho deje una carpeta con facturas, lecturas y listados de
distintas comunidades y periodos sin seleccionar previamente comunidad ni
periodo. El sistema debe clasificar, proponer altas seguras, conservar los
documentos originales y guiar cada expediente hasta reparto y cartas.

## Límites de seguridad

- Nunca se asignará un archivo a una comunidad solo por un número de factura.
- La comunidad se identifica por código inicial del nombre, etiqueta explícita,
  carpeta numérica, CIF del PDF o coincidencia inequívoca posterior.
- Si dos documentos del mismo código tienen CIF distintos, se bloquea el alta
  automática y se solicita revisión.
- Un archivo sin comunidad o sin fechas suficientes se conserva sin mover y se
  muestra en una bandeja de revisión; no se fuerza a una comunidad activa.
- La clasificación desde el nombre es evidencia inicial, no una sustitución de
  los datos extraídos del documento.

## Bandeja de entrada global

La acción principal de fuentes dejará de exigir una comunidad o expediente. Al
seleccionar archivos o una carpeta, se realizará un preanálisis por lotes:

1. Se deduplica por huella SHA-256 y se agrupa por comunidad detectada.
2. Se infiere una fecha/periodo desde el PDF cuando esté disponible; como
   alternativa se entiende el mes/año del nombre del correo.
3. Se conserva la categoría de entrada: `Habituales`, `Extraordinarias` o
   `Sin_Comunidad`.
4. Se presenta una revisión agrupada: comunidad, nombre/CIF, periodo propuesto,
   tipo, cantidad de archivos y cualquier bloqueo.
5. Al aceptar, se crean las comunidades nuevas con sus datos extraídos y se
   crean o reutilizan expedientes solo cuando el intervalo sea completo y
   verificable. Los casos sin intervalo quedan en la bandeja, no en una fecha
   inventada.
6. Tras confirmar, los documentos pasan al mismo sistema de fuentes,
   incidencias, Excel, reparto y cartas ya existente.

## Patrón de correo admitido

Los documentos observados en `flujo mail` siguen formas como:

- `658 - Factura feb 2026.pdf`
- `644 - Limpieza mar 2026.pdf`
- `643 - Electricidad feb.pdf`
- `625 - Mant. preventivo mar 2026.pdf`
- duplicados legítimos con sufijo ` (2)`.

El analizador reconocerá el código inicial y normalizará pistas de nombre:
factura, electricidad, agua, gas, limpieza, ascensor, mantenimiento,
reparación y extraordinaria. Estas pistas reducen incidencias de clasificación,
pero una factura solo recibe importe, periodo y datos canónicos tras extracción
del documento o revisión humana.

## OCR de PDFs escaneados

El OCR se ejecutará automáticamente para un PDF sin capa de texto. La aplicación
detectará Tesseract y Poppler instalados, seleccionará el idioma disponible
(`spa` si existe; de lo contrario `eng`) y probará primero la portada. Si no
obtiene texto utilizable, conservará el diagnóstico estructurado de la causa:
motor ausente, conversor ausente, imagen ilegible o texto insuficiente.

La interfaz no mostrará instrucciones de instalación a un usuario final. En su
lugar indicará que se intentó OCR, qué resultado tuvo y qué campo concreto debe
revisarse. Abrir archivo seguirá disponible.

## Continuidad hacia reparto

El botón lateral `Reparto` no podrá saltarse pasos. Consultará el estado real
del expediente y hará una de estas acciones:

- fuentes sin aplicar: abrir `Confirmar fuentes`;
- incidencias: abrir la primera incidencia;
- listo para cálculo: generar el Excel oficial;
- Excel generado: calcular reparto;
- reparto conciliado: abrir la generación de cartas.

Además, cualquier bloqueo se mostrará en un cuadro visible con la causa y la
acción siguiente, además del registro de actividad.

## Reutilización de versiones anteriores

`Flujo_Regularizacion` y `flujo mail` son fuentes de reglas y de documentos
reales, no ramas que se vayan a copiar completas. Se compararán perfiles de
proveedores, patrones de nombre, formatos de lectura y reglas de importación.
Solo se integrarán reglas que sean compatibles con la base, el flujo guiado y
la validación actuales. No se recuperará la interfaz antigua ni se sobrescribirá
la estructura de la base actual.

## Criterios de aceptación

- Una carpeta mixta puede analizarse sin comunidad ni periodo seleccionados.
- Cada propuesta muestra el motivo de comunidad y periodo, o una razón clara
  por la que requiere revisión.
- Los nombres habituales de correo se clasifican sin confundir el código con
  números de factura o duplicados.
- Un PDF escaneado intenta OCR automáticamente y la incidencia no recomienda
  descargar software que ya está integrado.
- El flujo de reparto siempre informa de qué acción falta y nunca falla de
  forma silenciosa.
