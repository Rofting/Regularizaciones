"""Script temporal para inspeccionar la estructura del Excel."""
import openpyxl
import re

ruta = r"C:\Users\Jose\Proyectos\Flujo_Regularizacion\Excels_Maestros\Comunidad_644.xlsx"
wb = openpyxl.load_workbook(ruta, data_only=False)

print("=== HOJAS ===")
for name in wb.sheetnames:
    print(f"  {name!r}")

print()

# Para cada hoja de datos, mostrar primeras 30 filas relevantes
hojas_datos = ["GAS", "ELECTRICIDAD", "AGUA", "OTROS GASTOS", "LECTURAS ACS M3", "LECTURAS CALEF KWH"]

for hoja in hojas_datos:
    if hoja not in wb.sheetnames:
        print(f"HOJA NO ENCONTRADA: {hoja!r}")
        continue
    ws = wb[hoja]
    print(f"\n=== HOJA: {hoja} (max_row={ws.max_row}) ===")
    patron = re.compile(r'^\d{4}-\d{4}$')
    for i, row in enumerate(ws.iter_rows(min_row=1, max_row=min(ws.max_row, 80), values_only=True), 1):
        colA = str(row[0]).strip() if row[0] is not None else ""
        colB = str(row[1]).strip() if row[1] is not None else ""
        colF = str(row[5]).strip() if len(row) > 5 and row[5] is not None else ""
        if patron.match(colA) or colB == "Suma" or colF == "Suma" or patron.match(colB):
            print(f"  fila {i:3d}: A={colA!r:20s} B={colB!r:30s} F={colF!r}")
        elif colA or colB:
            # Mostrar solo filas con algo en A o B
            print(f"  fila {i:3d}: A={colA!r:20s} B={colB!r:30s}")
