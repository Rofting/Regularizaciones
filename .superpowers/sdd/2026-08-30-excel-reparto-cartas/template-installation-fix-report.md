# Corrección — instalación de plantilla privada tras bootstrap

## Causa

El importador archivaba y verificaba el Excel maestro, pero no publicaba una
copia en la ruta relativa de plantilla del perfil ni registraba su SHA-256.
Por ello, el primer botón de Excel no encontraba la plantilla registrada.

## Corrección

- La ruta de aplicación pasa `project_root` al bootstrap.
- Desde la copia archivada y verificada se crea un archivo temporal en el
  directorio de plantillas privadas y se publica con enlace atómico; el SHA-256
  se comprueba antes y después.
- Sólo tras verificar la copia se registra el perfil activo y su hash. Un
  reintento de los mismos bytes es idempotente.
- Una plantilla o registro diferente no se reemplaza: deja una incidencia
  `TEMPLATE_VERSION_CONFLICT` que indica registrar una nueva versión.
- `plantillas/comunidades/` queda ignorado por Git. No se introdujeron fuentes
  reales en pruebas, commits ni este informe.
- La interfaz anuncia “Plantilla instalada y verificada” únicamente cuando el
  resultado contiene una copia comprobada; si hay conflicto, refresca la
  bandeja de incidencias.

## TDD y comprobación

- RED observado: `import_master_excel` no aceptaba `project_root`, no existían
  errores de instalación y el directorio privado no estaba ignorado.
- Pruebas sintéticas añadidas: instalación/registro con huella, reintento,
  conflicto sin sobreescritura, ruta fuera de proyecto y regla de Git ignore.
- Focales bootstrap/controlador: 33 OK.
- Suite completa: 123 OK.
- `py_compile` de módulos modificados y `git diff --check` local: correctos.
