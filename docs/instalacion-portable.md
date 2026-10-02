# Distribución Windows portable

La distribución `Regularizaciones-Windows-onedir` se construye en GitHub Actions
con `scripts/build_windows.ps1`. Contiene el ejecutable, Python y librerías,
modelos OCR, plantillas públicas y Poppler. No contiene una base, fuentes,
cartas, correos ni preferencias del despacho.

## Instalar y abrir

1. Descarga el artefacto `Regularizaciones-Windows-onedir` del workflow
   **Distribución Windows portable** y descomprímelo en una carpeta propia.
2. Instala LibreOffice 24.2 o posterior si el equipo no lo tiene. Writer
   convierte las cartas a PDF y Calc recalcula el Excel oficial.
3. Abre `Regularizaciones.exe`. No se necesita instalar Python ni descargar
   paquetes al arrancar. En el primer arranque se crea una base vacía y la
   aplicación pide los datos del despacho.

Los datos se guardan por defecto en
`%LOCALAPPDATA%\Regularizaciones`. Para usar una carpeta portable, define
`REGULARIZACIONES_HOME` antes de abrir el programa. La ruta puede contener
espacios y acentos. Si no es escribible, el arranque indica cómo elegir otra.

```text
config/       preferencias y ajustes propios
data/         base y expedientes
data/backups/ copias verificadas
salidas/      Excel, cartas y borradores EML
logs/         diagnósticos
```

## Actualizar y recuperar

Cierra el programa, sustituye solo la carpeta del ejecutable y vuelve a abrir.
La base y las salidas permanecen en el hogar de datos. Antes de migrar una
base existente, la aplicación crea una copia SQLite verificada. Para recuperar
una copia, cierra el programa y usa la función de restauración de base del
manual de operación.

Si colocas el ejecutable en la carpeta de una instalación antigua y el hogar
nuevo aún no tiene base, se importa una copia verificada de `data/gestion.db`
y se copian los perfiles, plantillas de comunidades, expedientes y salidas sin
sobrescribir archivos ya presentes.

Para comprobar una instalación sin abrir ventanas, ejecuta
`Regularizaciones.exe --self-test`. Este comando crea o abre la base en el hogar
configurado y comprueba su integridad.

## Construir desde el código

En Windows 10/11, con Python 3.12 y LibreOffice disponibles, instala Poppler y
ejecuta `powershell -ExecutionPolicy Bypass -File scripts/build_windows.ps1`.
El resultado queda en `dist\Regularizaciones\`. El build instala dependencias
de `requirements-build.txt` en el entorno de construcción; el ejecutable no
las instala al arrancar.
