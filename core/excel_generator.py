"""
excel_generator.py
==================
Regenera el Excel Maestro COMPLETO de una comunidad desde la BD SQLite.

FILOSOFÍA: la BD es la única fuente de verdad. El Excel es un informe de
salida que se reconstruye entero cada vez. Así desaparecen los problemas de
corrupción y fórmulas descuadradas de la escritura incremental:
  - Nunca se insertan filas en un libro con datos previos.
  - Los bloques de año se escriben en orden cronológico, append-only.
  - Las hojas ANALISIS/RESUMEN se crean al final, cuando las filas de suma
    ya están en su posición definitiva.
  - La escritura es atómica: se construye en un temporal y solo si todo va
    bien sustituye al archivo real (el anterior queda en backups/).

USO:
    python excel_generator.py --comunidad 644
    python excel_generator.py --comunidad 644 --bd ../data/gestion.db

Desde código:
    from excel_generator import regenerar_excel_comunidad
    resultado = regenerar_excel_comunidad("644")
"""

import os
import sys
import shutil
import sqlite3
import argparse
from copy import deepcopy
from datetime import datetime
from pathlib import Path

# En consola cp1252 (Windows en español, fuera de una terminal ya en UTF-8)
# los emoji de _log_defecto (✅, ⚠️...) revientan con UnicodeEncodeError.
# Ver el mismo arreglo, con más detalle, en pipeline.py.
for _stream in (sys.stdout, sys.stderr):
    if hasattr(_stream, "reconfigure"):
        try:
            _stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass

sys.path.insert(0, str(Path(__file__).parent))
import excel_writer

BASE_DIR    = Path(__file__).parent.parent
RUTA_BD     = BASE_DIR / "data" / "gestion.db"
RUTA_EXCELS = BASE_DIR / "Excels_Maestros"
MAX_BACKUPS = 5


_CANONICAL_TABLES = {
    "GAS": {
        "sheet": "GAS", "start_row": 10, "end_row": 30,
        "input_columns": {
            "invoice_date": "B", "days": "C", "start_date": "D",
            "end_date": "E", "consumption": "F", "fixed": "G",
            "variable": "H", "provider": "J",
        },
        "derived_columns": {"total": "I"},
    },
    "ELECTRICIDAD": {
        "sheet": "ELECTRICIDAD", "start_row": 10, "end_row": 30,
        "input_columns": {
            "invoice_date": "B", "days": "C", "start_date": "D",
            "end_date": "E", "consumption": "F", "fixed": "G",
            "variable": "H", "provider": "J",
        },
        "derived_columns": {"total": "I"},
    },
    "AGUA": {
        "sheet": "AGUA", "start_row": 10, "end_row": 30,
        "input_columns": {
            "invoice_date": "B", "days": "C", "start_date": "D",
            "end_date": "E", "consumption": "F", "fixed": "G",
            "variable": "H", "provider": "J",
        },
        "derived_columns": {"total": "I"},
    },
    "OTROS_GASTOS": {
        "sheet": "OTROS GASTOS", "start_row": 10, "end_row": 60,
        "columns": {"date": "B", "description": "C", "amount": "D"},
    },
    "ACS": {
        "sheet": "LECTURAS ACS M3", "start_row": 8, "end_row": 40,
        "input_columns": {
            "charge_date": "B", "final_date": "C", "final": "D",
            "initial_date": "E", "initial": "F", "variable_fee": "H",
            "fixed_fee": "I",
        },
        "derived_columns": {
            "consumption": "G", "total": "J", "variable_unit": "L",
            "fixed_unit": "M",
        },
    },
    "CALEFACCION": {
        "sheet": "LECTURAS CALEF KWH", "start_row": 8, "end_row": 40,
        "input_columns": {
            "charge_date": "B", "final_date": "C", "final": "D",
            "initial_date": "E", "initial": "F", "variable_fee": "H",
            "fixed_fee": "I",
        },
        "derived_columns": {
            "consumption": "G", "total": "J", "variable_unit": "L",
            "fixed_unit": "M",
        },
    },
}


def canonical_workbook_layout(active_modules: tuple[str, ...] | list[str]) -> dict:
    """Return the declarative layout used by generated community templates."""
    tables = {
        module: deepcopy(_CANONICAL_TABLES[module]) for module in active_modules
    }
    total_checks: dict[str, list[str]] = {}
    for module, table in tables.items():
        sheet = table["sheet"]
        start = table["start_row"]
        end = table["end_row"]
        if module in {"GAS", "ELECTRICIDAD", "AGUA"}:
            total_checks.update({
                f"invoice_total:{module}": [sheet, f"I{start}:I{end}"],
                f"invoice_component:{module}:fixed": [sheet, f"G{start}:G{end}"],
                f"invoice_component:{module}:variable": [sheet, f"H{start}:H{end}"],
            })
        elif module == "OTROS_GASTOS":
            total_checks["parameter:extraordinary_expense_actual"] = [
                sheet, f"D{start}:D{end}"
            ]
        elif module == "ACS":
            total_checks["reading_total:ACS"] = [sheet, f"G{start}:G{end}"]
    return {
        "metadata_cells": {
            "community_name": ["DATOS", "A1"],
            "period_label": ["DATOS", "A3"],
            "owner_count": ["DATOS", "D4"],
        },
        "tables": tables,
        "parameter_cells": {},
        "total_checks": total_checks,
    }


def create_canonical_community_template(
    destination: str | Path,
    *,
    community_code: str,
    community_name: str,
    active_modules: tuple[str, ...] | list[str],
) -> dict:
    """Create a portable workbook containing only confirmed module sheets."""
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill

    from excel_profiles import MODULE_REQUIRED_SHEETS

    modules = tuple(active_modules)
    layout = canonical_workbook_layout(modules)
    workbook = Workbook()
    data_sheet = workbook.active
    data_sheet.title = "DATOS"
    data_sheet["A1"] = community_name
    data_sheet["A2"] = f"CÓDIGO: {community_code}"
    data_sheet["A3"] = "PERIODO: pendiente de expediente"
    data_sheet["D4"] = 0

    sheet_names = ["DATOS"]
    for module in modules:
        for sheet_name in MODULE_REQUIRED_SHEETS[module]:
            if sheet_name not in sheet_names:
                workbook.create_sheet(sheet_name)
                sheet_names.append(sheet_name)

    title_fill = PatternFill("solid", fgColor="1A3A5C")
    header_fill = PatternFill("solid", fgColor="D6EAF8")
    for sheet in workbook.worksheets:
        sheet.print_area = "A1:M60"
        sheet.freeze_panes = "B8" if sheet.title != "DATOS" else "A5"
        sheet["A1"].font = Font(bold=True, color="1A3A5C")

    headers = {
        "GAS": ("Fecha factura", "Días", "Inicio", "Fin", "Consumo", "Fijo", "Variable", "Total", "Proveedor"),
        "ELECTRICIDAD": ("Fecha factura", "Días", "Inicio", "Fin", "Consumo", "Fijo", "Variable", "Total", "Proveedor"),
        "AGUA": ("Fecha factura", "Días", "Inicio", "Fin", "Consumo", "Fijo", "Variable", "Total", "Proveedor"),
        "OTROS_GASTOS": ("Fecha", "Descripción", "Importe"),
        "ACS": ("Fecha cargo", "Fecha final", "Lectura final", "Fecha inicial", "Lectura inicial", "Consumo", "Variable", "Fijo", "Total"),
        "CALEFACCION": ("Fecha cargo", "Fecha final", "Lectura final", "Fecha inicial", "Lectura inicial", "Consumo", "Variable", "Fijo", "Total"),
    }
    for module, table in layout["tables"].items():
        sheet = workbook[table["sheet"]]
        sheet["A2"] = f"{module} — {community_name}"
        sheet["A2"].font = Font(bold=True, color="FFFFFF")
        sheet["A2"].fill = title_fill
        header_row = table["start_row"] - 1
        columns = list(table.get("input_columns", table.get("columns", {})).values())
        columns += list(table.get("derived_columns", {}).values())
        for column, label in zip(columns, headers[module]):
            cell = sheet[f"{column}{header_row}"]
            cell.value = label
            cell.font = Font(bold=True, color="1A3A5C")
            cell.fill = header_fill

    destination = Path(destination)
    destination.parent.mkdir(parents=True, exist_ok=True)
    try:
        workbook.save(destination)
    finally:
        workbook.close()
    return layout


def _log_defecto(mensaje: str, tipo: str = "neutro"):
    print(mensaje)


def _rotar_backups(carpeta: Path, codigo: str):
    """Mantiene solo los MAX_BACKUPS backups más recientes de una comunidad."""
    backups = sorted(carpeta.glob(f"Comunidad_{codigo}_*.xlsx"),
                     key=lambda p: p.stat().st_mtime, reverse=True)
    for viejo in backups[MAX_BACKUPS:]:
        try:
            viejo.unlink()
        except OSError:
            pass


def _personalizar_datos(ruta_xlsx: str, com, ruta_bd: str):
    """Escribe nombre, código y nº de viviendas de la comunidad en la hoja DATOS."""
    import openpyxl
    n_viv = 0
    try:
        con = sqlite3.connect(ruta_bd)
        n_viv = con.execute(
            "SELECT COUNT(*) FROM propietarios WHERE id_comunidad=? AND activo=1",
            (com["id_comunidad"],)
        ).fetchone()[0] or 0
        con.close()
    except sqlite3.Error:
        pass

    wb = openpyxl.load_workbook(ruta_xlsx)
    if "DATOS" in wb.sheetnames:
        ws = wb["DATOS"]
        ws["A1"] = com["nombre"] or f"COMUNIDAD {com['codigo']}"
        if n_viv:
            # D7 es la celda que referencian las hojas de lecturas (=DATOS!D7)
            ws["D7"] = n_viv
        wb.save(ruta_xlsx)


def regenerar_excel_comunidad(codigo: str,
                              ruta_bd: str = None,
                              ruta_excels: str = None,
                              log=None,
                              *,
                              id_case: int | None = None,
                              project_root: str | Path | None = None,
                              recalculator=None) -> dict:
    """
    Reconstruye Comunidad_{codigo}.xlsx entero desde la BD.

    Returns:
        dict con: ok, archivo, periodos_volcados (list), filas_escritas,
                  backup, errores (list)
    """
    log         = log or _log_defecto
    ruta_bd     = str(ruta_bd or RUTA_BD)
    ruta_excels = Path(ruta_excels or RUTA_EXCELS)
    ruta_excels.mkdir(parents=True, exist_ok=True)

    resultado = {"ok": False, "archivo": None, "periodos_volcados": [],
                 "filas_escritas": 0, "backup": None, "errores": []}

    # El nuevo flujo por expediente queda aislado del regenerador histórico.
    # Los callers antiguos, que no pasan id_case, conservan exactamente el
    # comportamiento previo basado en comunidad y todos sus periodos.
    if id_case is not None:
        from excel_export_service import generate_official_excel

        connection = sqlite3.connect(ruta_bd)
        connection.row_factory = sqlite3.Row
        connection.execute("PRAGMA foreign_keys = ON")
        try:
            output_root = (
                ruta_excels.parent
                if ruta_excels.name.lower() == "excels_maestros"
                else ruta_excels
            )
            exported = generate_official_excel(
                connection,
                id_case=id_case,
                project_root=Path(project_root or BASE_DIR),
                output_root=output_root,
                recalculator=recalculator,
                progress=lambda stage, payload: log(f"    · {stage}"),
            )
            period = connection.execute(
                """SELECT p.nombre FROM regularization_cases c
                   JOIN periodos p ON p.id_periodo=c.id_periodo
                   WHERE c.id_case=?""",
                (id_case,),
            ).fetchone()
            resultado.update({
                "ok": True,
                "archivo": str(exported.output_path),
                "periodos_volcados": [period[0]] if period else [],
                "backup": str(exported.backup_path) if exported.backup_path else None,
            })
        except Exception as error:
            resultado["errores"].append(f"{type(error).__name__}: {error}")
            log(f"  ❌ {resultado['errores'][-1]}", "error")
        finally:
            connection.close()
        return resultado

    # ── Datos de la comunidad y sus periodos ─────────────────────────────
    con = sqlite3.connect(ruta_bd)
    con.row_factory = sqlite3.Row
    com = con.execute(
        "SELECT id_comunidad, codigo, nombre FROM comunidades WHERE codigo=?",
        (str(codigo),)
    ).fetchone()
    if not com:
        con.close()
        resultado["errores"].append(f"Comunidad '{codigo}' no existe en la BD")
        return resultado

    periodos = con.execute(
        "SELECT id_periodo, nombre, fecha_inicio FROM periodos "
        "WHERE id_comunidad=? ORDER BY fecha_inicio ASC, nombre ASC",
        (com["id_comunidad"],)
    ).fetchall()
    con.close()

    if not periodos:
        resultado["errores"].append(
            f"La comunidad {codigo} no tiene periodos en la BD; nada que volcar")
        return resultado

    destino  = ruta_excels / f"Comunidad_{codigo}.xlsx"
    temporal = ruta_excels / f"~Comunidad_{codigo}_generando.xlsx"

    # Plantilla rica (con hoja ANALISIS y fórmulas) si existe
    ruta_plantilla = BASE_DIR / "plantillas" / "Comunidad_PLANTILLA.xlsx"

    try:
        # ── 1. Libro nuevo desde plantilla limpia ────────────────────────
        if temporal.exists():
            temporal.unlink()
        if ruta_plantilla.exists():
            shutil.copy2(str(ruta_plantilla), str(temporal))
            _personalizar_datos(str(temporal), com, ruta_bd)
        else:
            log("  ⚠️ plantillas/Comunidad_PLANTILLA.xlsx no encontrada — "
                "se usa plantilla básica (sin hojas ANALISIS)")
            excel_writer.crear_plantilla_excel(str(temporal),
                                               codigo=com["codigo"],
                                               nombre=com["nombre"] or "")

        # ── 2. Volcar cada periodo en orden cronológico ──────────────────
        for per in periodos:
            r = excel_writer.actualizar_excel_maestro(
                ruta_excel=str(temporal),
                ruta_bd=ruta_bd,
                id_comunidad=com["id_comunidad"],
                id_periodo=per["id_periodo"],
                hacer_backup=False,
            )
            resultado["filas_escritas"] += r.get("filas_escritas", 0)
            resultado["periodos_volcados"].append(per["nombre"])
            for err in r.get("errores", []):
                resultado["errores"].append(f"{per['nombre']}: {err}")
            log(f"    · Periodo {per['nombre']}: "
                f"{r.get('filas_escritas', 0)} filas "
                f"({', '.join(r.get('pestañas_actualizadas', [])) or 'sin datos'})")

        # ── 3. Backup del actual y sustitución atómica ───────────────────
        if destino.exists():
            carpeta_bak = ruta_excels / "backups"
            carpeta_bak.mkdir(exist_ok=True)
            nombre_bak = f"Comunidad_{codigo}_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
            ruta_bak = carpeta_bak / nombre_bak
            shutil.copy2(str(destino), str(ruta_bak))
            resultado["backup"] = str(ruta_bak)
            _rotar_backups(carpeta_bak, codigo)

        os.replace(str(temporal), str(destino))
        resultado["archivo"] = str(destino)
        resultado["ok"] = True
        log(f"  ✅ Excel regenerado: {destino.name} "
            f"({len(periodos)} periodos, {resultado['filas_escritas']} filas)")

    except PermissionError:
        resultado["errores"].append(
            f"No se pudo escribir {destino.name}: el archivo está abierto en Excel. "
            "Ciérralo y vuelve a ejecutar.")
    except Exception as e:
        resultado["errores"].append(f"{type(e).__name__}: {e}")
    finally:
        if temporal.exists():
            try:
                temporal.unlink()
            except OSError:
                pass

    if resultado["errores"] and not resultado["ok"]:
        for err in resultado["errores"]:
            log(f"  ❌ {err}", "error")
    return resultado


# ---------------------------------------------------------------------------
if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Regenera el Excel Maestro de una comunidad desde la BD")
    parser.add_argument("--comunidad", required=True, help="Código (ej: 644)")
    parser.add_argument("--bd", default=None, help="Ruta a gestion.db")
    parser.add_argument("--excels", default=None, help="Carpeta Excels_Maestros")
    args = parser.parse_args()

    r = regenerar_excel_comunidad(args.comunidad, args.bd, args.excels)
    sys.exit(0 if r["ok"] else 1)
