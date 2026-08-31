# Informe — resolución manual de reinicios de contador

## Resultado

- Se añade `approve_counter_reset_estimate`, una operación transaccional que
  sólo admite incidencias abiertas `COUNTER_RESET` ligadas a una lectura
  canónica inicial/final del expediente.
- El consumo confirmado debe ser positivo; se transforma en una lectura final
  virtual (`lectura inicial + consumo`) con estado `estimado`, método explícito
  `counter_reset_manual`, aprobador, fecha y notas de trazabilidad.
- El valor de origen se conserva tanto en `manual_corrections` como en las
  notas; la corrección actualiza además el candidato revisado, sin modificar
  las lecturas de otros propietarios.
- Se añade el cierre auditado `dismiss_invoice_outside_period`, restringido al
  código `INVOICE_OUTSIDE_PERIOD`: marca sólo la incidencia como descartada y
  no inserta ni modifica facturas.
- El importador clasifica una fecha de factura válida pero fuera del intervalo
  como `INVOICE_OUTSIDE_PERIOD`; fechas malformadas y extremos inválidos siguen
  siendo `INCOMPATIBLE_DATE`.
- El diálogo v3 escoge la ruta específica sin cambiar paleta ni branding:
  muestra una explicación humana, el consumo estimado y el motivo, o el cierre
  seguro de una factura fuera de período. El enrutado se prueba sin crear una
  ventana Tk.

## Verificación

Ejecutado sin fuentes reales ni envío de correo:

```text
python -m unittest discover -s tests -t . -q
Ran 138 tests in 20.249s
OK

python -m py_compile core/document_review.py core/expedient_ui.py core/excel_bootstrap_importer.py
git diff --check
```

Las pruebas cubren estimación aprobada, entrada inválida, aislamiento entre
propietarios, rollback de auditoría al fallar la lectura, desbloqueo del reparto
al aprobar todas las incidencias y clasificación/cierre de factura fuera de
período.
