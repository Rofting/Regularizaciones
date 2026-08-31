# Validación privada aislada del flujo

Esta comprobación demuestra el flujo completo con las fuentes reales de una
comunidad sin modificar las fuentes, la base de datos habitual ni el proyecto.
No forma parte de la aplicación gráfica ni se ejecuta en la suite ordinaria.

## Preparación

Use una carpeta de salida nueva o vacía situada fuera del repositorio y fuera
de las carpetas que contienen las fuentes. Defina explícitamente estas cuatro
variables, apuntando las tres primeras a archivos y la última al directorio de
salida:

```powershell
$env:REGULARIZACION_658_MASTER = "<archivo maestro>"
$env:REGULARIZACION_658_PROPIETARIOS = "<listado propietarios>"
$env:REGULARIZACION_658_LECTURAS = "<lecturas>"
$env:REGULARIZACION_658_VALIDATION_ROOT = "<carpeta nueva y vacía>"
```

Si LibreOffice no está en `PATH`, defina también `LIBREOFFICE_PATH` con la ruta
del ejecutable `soffice`. El proceso no descarga ni instala programas.

## Ejecución

Desde la carpeta del proyecto, con el entorno virtual activado:

```powershell
python core/private_658_validation.py
```

El runner crea una subcarpeta única dentro de la raíz indicada. Copia sólo la
configuración y la plantilla Word públicas. Archiva el maestro y registra la
plantilla Excel privada dentro de esa ejecución aislada.

## Resultado esperado

1. Se crea una base SQLite nueva y un expediente del período estructural del
   maestro.
2. Se importan maestro, propietarios y lecturas mediante los servicios
   ordinarios.
3. Si hay incidencias abiertas, termina con código `2` y deja un informe con
   el número de incidencias por tipo. Hay que resolverlas manualmente: nunca
   se completan valores automáticamente.
4. Si no hay incidencias, regenera y valida el Excel oficial, calcula el
   reparto al céntimo, genera las cartas y renderiza una carta para exigir una
   única página.
5. El informe final contiene estado, conteos y una ruta sólo dentro de la
   raíz de validación. No contiene datos de propietarios, importes
   individuales, nombres de las fuentes ni sus rutas.

Tras un resultado correcto, abra en Windows el Excel publicado y la carta
renderizada para una revisión visual. El envío de correo sigue siendo manual.

## Seguridad y limpieza

El runner rechaza la raíz si coincide con el repositorio, lo contiene, está
dentro de él, es una carpeta de fuentes o no está vacía. No borra nada. Para
retirar el material de validación, elimine manualmente sólo la subcarpeta de
ejecución confirmada dentro de la raíz privada.
