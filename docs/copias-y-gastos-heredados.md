# Copias automáticas y gastos de plantillas antiguas

Primera entrega de protección del cálculo: mejoras 19 y 8 de la hoja de ruta.

## Copias de seguridad

La aplicación crea una copia de la base:

- Al abrirla, antes de cargar comunidades o aplicar migraciones.
- Antes de aplicar un lote de migraciones a una base que ya contiene datos.
- Antes de generar o regenerar Excel, reparto o cartas desde el flujo guiado.
- Antes de reevaluar las fuentes desde la interfaz.
- Antes de iniciar el flujo desatendido `core/pipeline.py`, incluidos la ingesta,
  el reparto y la regeneración de Excel y cartas. Una base nueva todavía vacía
  no requiere copia.

Una base nueva vacía y las conexiones en memoria no generan copias de
migración. La copia usa la API de SQLite, incluye los cambios confirmados en
WAL y pasa `PRAGMA integrity_check`. No copia a ciegas sólo el archivo `.db`.

Las copias están en la carpeta `backups` junto a la base, normalmente
`data/backups`. Cada copia automática tiene un manifiesto `.db.json` con
fecha, motivo, base de origen y SHA-256. El arranque y las acciones muestran
la copia creada en el registro de la aplicación.

Si no se puede crear o verificar la copia, la operación se detiene antes de
regenerar o migrar. El flujo desatendido también se detiene antes de copiar
archivos de una carpeta extra o ingerirlos. Si falla únicamente la limpieza
de copias antiguas, la copia nueva se conserva y se registra el problema.

### Retención

Por defecto se conservan las **20 últimas copias automáticas verificadas de
cada base**. Para cambiarlo, crea `backup_settings.json` junto a la base:

```json
{"max_backups": 30}
```

El mínimo permitido es 3. Una configuración inválida se informa y detiene
la operación; no autoriza borrar copias. La limpieza conserva las copias
manuales, las del reinicio seguro, las de otras bases y los archivos alterados
o que no se pueden verificar. Por ello la carpeta puede contener más archivos
que el límite configurado.

### Recuperación

Para comprobar una copia sin modificarla:

```bash
.venv/bin/python -c "import sys; from pathlib import Path; sys.path.insert(0, 'core'); from database_backup import verify_backup; verify_backup(Path('data/backups/NOMBRE_DE_LA_COPIA.db'))"
```

Para recuperar datos, cierra la aplicación y todas las conexiones a la base.
Conserva primero la base actual y sus archivos `-wal`/`-shm`, si existen, en
otra carpeta. Copia la copia verificada al nombre y ubicación de la base
activa, sin dejar allí los archivos WAL/SHM de la base anterior. Abre la
aplicación y comprueba comunidades y expedientes. La suite verifica una
restauración a otra base usando únicamente el archivo de la copia.

Estas copias protegen la base de datos. Los documentos archivados, plantillas
y cartas siguen necesitando su propia copia para recuperar una instalación
completa. No se ha añadido todavía un asistente de restauración.

## Gastos heredados

La revisión comprueba las cuatro filas conocidas de `OTROS GASTOS`, tanto
sus importes mensuales `B6:B9` como sus totales anuales `C6:C9`. Si hay un
importe o expresión sin confirmar para la comunidad y período actuales, la
pantalla ofrece **Revisar gastos fijos** y bloquea el Excel. El aviso muestra
celda, servicio e importe, distinguiendo euros al mes y al año.

En el formulario, guarda el importe mensual correcto. Dejar un campo vacío
confirma explícitamente cero. Guardar también sustituye, en el Excel generado,
los totales anuales de esas filas por la fórmula mensual × 12 del modelo
común: así no reaparece un `993,44` escrito a mano en una plantilla antigua.
Las fórmulas no estándar de esas celdas también se muestran para revisión.

La plantilla original se conserva. Sólo se modifican estas celdas en las
filas cuyo rótulo coincide con el modelo; el resto del diseño y sus estilos
siguen sujetos a la validación del Excel. Un gasto legítimo puede confirmarse:
el importe `993,44` no se usa como una regla para borrarlo automáticamente.

## Validación

Las pruebas usan bases y libros sintéticos. Cubren recuperación desde WAL,
retención, fallos de copia, bloqueo previo a migraciones/regeneraciones,
arranque, navegación al formulario y confirmación de cero. La prueba con
LibreOffice real comprueba que el total anual heredado desaparece del libro
recalculado conservando el diseño y los resultados de las facturas.
