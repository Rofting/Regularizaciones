# Final fix report

## Scope

Corrección final del flujo de alta guiada para lecturas ACS/calefacción:

- La UI no obliga a responder preguntas de módulos inactivos.
- La importación complementaria usa los `reading_bindings` confirmados por módulo, columna y `sha256`.
- La evidencia de lecturas Excel conserva relaciones de fila/celda reales, no solo encabezados.
- La plantilla canónica escribe cabeceras desde los bindings de layout.
- El export de perfiles onboarding materializa el consumo agregado para que el libro sea verificable con `data_only=True` sin depender de caché de fórmulas del motor externo.

## RED

Comando:

```powershell
<python-runtime> -X utf8 -m unittest tests.test_onboarding_final_flow -v
```

Resultado inicial:

```text
FAILED (failures=3)
AssertionError: 30 != None
AssertionError: 300 != None
```

La importación ya dejaba las lecturas en la base de datos, pero el Excel oficial generado para perfiles onboarding dejaba la celda de consumo `G8` sin valor visible al abrir con `data_only=True`, porque solo escribía fórmula y el `DeterministicRecalculator` no genera cachés.

## Fix

`core/excel_export_service.py` ahora materializa `final - initial` en la columna de consumo únicamente cuando el perfil tiene `onboarding_configuration`. Los perfiles históricos mantienen la fórmula `=D8-F8`, protegidos por `tests.test_excel_export_service`.

## GREEN

Comandos:

```powershell
<python-runtime> -X utf8 -m unittest tests.test_onboarding_final_flow tests.test_excel_export_service -v
<python-runtime> -X utf8 -m unittest discover -s tests -v
```

Resultados:

```text
Ran 22 tests in 5.280s
OK

Ran 222 tests in 29.000s
OK
```

## Notes

Dependencias faltantes instaladas en el Python empaquetado para poder ejecutar la suite real: `xlrd`, `xlwt`, `matplotlib`, `customtkinter`.
