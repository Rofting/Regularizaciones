"""Validación estructural y económica previa a publicar un Excel oficial."""

from __future__ import annotations

from dataclasses import dataclass
from decimal import Decimal, ROUND_HALF_UP
import hashlib
from pathlib import Path
from typing import Mapping
from zipfile import ZipFile

from openpyxl import load_workbook
from openpyxl.utils.cell import range_boundaries
from openpyxl.utils.cell import coordinate_from_string

from excel_profiles import ExcelProfile


FORMULA_ERRORS = frozenset({
    "#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A",
})


class WorkbookValidationError(ValueError):
    """El libro no conserva el contrato necesario para ser oficial."""


@dataclass(frozen=True)
class WorkbookFingerprint:
    sheet_names: tuple[str, ...]
    print_areas: tuple[tuple[str, str], ...]
    required_formulas: tuple[tuple[str, str, str], ...]
    merges: tuple[tuple[str, tuple[str, ...]], ...]
    dimensions: tuple[tuple[str, tuple, tuple], ...]
    view_and_print_settings: tuple[tuple[str, str, str, str, str], ...]
    drawings: tuple[tuple[str, int, tuple[str, ...], int, tuple[str, ...]], ...]
    immutable_cells: tuple[tuple[str, str, object, int, str], ...]
    ooxml_design_parts: tuple[tuple[str, str], ...]


def _print_area(sheet) -> str:
    value = sheet.print_area
    return str(value) if value is not None else ""


def _input_columns(table) -> Mapping[str, str]:
    return table.get("input_columns", table.get("columns", {}))


def _mutable_addresses(profile: ExcelProfile) -> dict[str, set[str]]:
    result: dict[str, set[str]] = {}
    layout = profile.workbook_layout
    for sheet_name, address in layout.get("metadata_cells", {}).values():
        result.setdefault(sheet_name, set()).add(address)
    for sheet_name, address in layout.get("parameter_cells", {}).values():
        result.setdefault(sheet_name, set()).add(address)
    for table in layout.get("tables", {}).values():
        sheet_name = table["sheet"]
        for row in range(int(table["start_row"]), int(table["end_row"]) + 1):
            for column in (*_input_columns(table).values(), *table.get("derived_columns", {}).values()):
                result.setdefault(sheet_name, set()).add(f"{column}{row}")
    return result


def _dimension_signature(dimensions) -> tuple:
    return tuple(
        sorted(
            (key, getattr(dimension, "height", None), getattr(dimension, "width", None),
             getattr(dimension, "hidden", None), getattr(dimension, "outlineLevel", None),
             getattr(dimension, "collapsed", None), getattr(dimension, "style_id", None))
            for key, dimension in dimensions.items()
        )
    )


def _anchor_signature(anchor) -> str:
    if isinstance(anchor, str):
        return anchor
    start = getattr(anchor, "_from", None)
    end = getattr(anchor, "to", None)
    def marker(value):
        if value is None:
            return ""
        return ":".join(str(getattr(value, attr, "")) for attr in ("col", "row", "colOff", "rowOff"))
    return f"{type(anchor).__name__}:{marker(start)}:{marker(end)}"


def _design_part_hashes(path: Path) -> tuple[tuple[str, str], ...]:
    prefixes = ("xl/charts/", "xl/drawings/", "xl/media/", "customXml/", "xl/theme/", "xl/printerSettings/")
    with ZipFile(path, "r") as archive:
        names = [
            name for name in archive.namelist()
            if name.startswith(prefixes) or name.endswith(".rels")
            or name in {"xl/styles.xml", "[Content_Types].xml", "_rels/.rels"}
        ]
        return tuple(
            (name, hashlib.sha256(archive.read(name)).hexdigest())
            for name in sorted(set(names))
        )


def workbook_fingerprint(path: Path, profile: ExcelProfile) -> WorkbookFingerprint:
    workbook = load_workbook(path, data_only=False, read_only=False)
    try:
        formulas = []
        for sheet_name, address in profile.required_formula_cells:
            value = workbook[sheet_name][address].value if sheet_name in workbook.sheetnames else None
            formulas.append((sheet_name, address, value if isinstance(value, str) else ""))
        mutable = _mutable_addresses(profile)
        immutable_cells = []
        merges = []
        dimensions = []
        views = []
        drawings = []
        for sheet in workbook.worksheets:
            merges.append((sheet.title, tuple(sorted(str(value) for value in sheet.merged_cells.ranges))))
            dimensions.append((
                sheet.title, _dimension_signature(sheet.row_dimensions),
                _dimension_signature(sheet.column_dimensions),
            ))
            views.append((
                sheet.title, str(sheet.freeze_panes or ""), _print_area(sheet),
                str(sheet.page_setup), str(sheet.page_margins),
            ))
            drawings.append((
                sheet.title, len(sheet._charts),
                tuple(_anchor_signature(chart.anchor) for chart in sheet._charts),
                len(sheet._images),
                tuple(_anchor_signature(image.anchor) for image in sheet._images),
            ))
            for cell in sheet._cells.values():
                if cell.coordinate in mutable.get(sheet.title, set()):
                    continue
                if cell.value is not None or cell.has_style:
                    immutable_cells.append((
                        sheet.title, cell.coordinate, cell.value, cell.style_id, cell.data_type,
                    ))
        return WorkbookFingerprint(
            sheet_names=tuple(workbook.sheetnames),
            print_areas=tuple(
                (name, _print_area(workbook[name]))
                for name in profile.required_sheets
                if name in workbook.sheetnames
            ),
            required_formulas=tuple(formulas),
            merges=tuple(merges),
            dimensions=tuple(dimensions),
            view_and_print_settings=tuple(views),
            drawings=tuple(drawings),
            immutable_cells=tuple(sorted(immutable_cells)),
            ooxml_design_parts=_design_part_hashes(path),
        )
    finally:
        workbook.close()


def _to_cents(value) -> int:
    if value is None or isinstance(value, bool):
        return 0
    try:
        amount = Decimal(str(value)).quantize(Decimal("0.01"), ROUND_HALF_UP)
    except Exception as error:
        raise WorkbookValidationError(f"El total contiene un valor no numérico: {value!r}") from error
    return int(amount * 100)


def _sum_reference(values_sheet, formula_sheet, reference: str) -> int:
    if ":" not in reference:
        return _to_cents(values_sheet[reference].value)
    minimum_column, minimum_row, maximum_column, maximum_row = range_boundaries(reference)
    total = 0
    for row in values_sheet.iter_rows(
        min_row=minimum_row,
        max_row=maximum_row,
        min_col=minimum_column,
        max_col=maximum_column,
    ):
        for cell in row:
            formula_value = formula_sheet[cell.coordinate].value
            if isinstance(formula_value, str) and formula_value.startswith("="):
                evaluated = _simple_formula_cents(values_sheet, formula_value)
                if evaluated is not None:
                    total += evaluated
                continue
            total += _to_cents(cell.value)
    return total


def _simple_formula_cents(values_sheet, formula: str) -> int | None:
    """Evalúa sólo la suma/resta de dos celdas usada por filas derivadas.

    LibreOffice aporta el caché en producción. Esta ruta hace que el doble de
    recálculo de las pruebas pueda validar exactamente las fórmulas de fila,
    sin convertir este módulo en otro motor de hojas de cálculo.
    """
    expression = formula[1:].replace("$", "").strip()
    for operator in ("+", "-"):
        if expression.count(operator) != 1:
            continue
        left, right = (part.strip() for part in expression.split(operator))
        try:
            coordinate_from_string(left)
            coordinate_from_string(right)
        except ValueError:
            continue
        left_cents = _to_cents(values_sheet[left].value)
        right_cents = _to_cents(values_sheet[right].value)
        return left_cents + right_cents if operator == "+" else left_cents - right_cents
    return None


def _formula_errors(workbook) -> list[str]:
    errors: list[str] = []
    for sheet in workbook.worksheets:
        for row in sheet.iter_rows():
            for cell in row:
                value = cell.value
                if cell.data_type == "e" or (
                    isinstance(value, str) and value.upper() in FORMULA_ERRORS
                ):
                    errors.append(f"{sheet.title}!{cell.coordinate}={value}")
    return errors


def validate_workbook(
    path: Path,
    profile: ExcelProfile,
    expected_totals_cents: Mapping[str, int],
    *,
    expected_fingerprint: WorkbookFingerprint | None = None,
) -> None:
    """Comprueba estructura, fórmulas, impresión, errores y conciliación."""
    path = Path(path)
    if not path.is_file():
        raise WorkbookValidationError(f"No existe el libro a validar: {path}")
    formula_book = load_workbook(path, data_only=False, read_only=False)
    values_book = load_workbook(path, data_only=True, read_only=False)
    try:
        missing = [name for name in profile.required_sheets if name not in formula_book.sheetnames]
        if missing:
            raise WorkbookValidationError(
                "Faltan hojas obligatorias: " + ", ".join(missing)
            )
        for sheet_name in profile.required_sheets:
            if not _print_area(formula_book[sheet_name]):
                raise WorkbookValidationError(
                    f"La hoja {sheet_name} no conserva un área de impresión"
                )
        for sheet_name, address in profile.required_formula_cells:
            value = formula_book[sheet_name][address].value
            if not isinstance(value, str) or not value.startswith("="):
                raise WorkbookValidationError(
                    f"Falta la fórmula obligatoria {sheet_name}!{address}"
                )
        errors = _formula_errors(formula_book) + _formula_errors(values_book)
        if errors:
            raise WorkbookValidationError(
                "El libro contiene un error de fórmula: " + ", ".join(sorted(set(errors))[:5])
            )

        total_checks = profile.workbook_layout.get("total_checks", {})
        absent_checks = sorted(set(expected_totals_cents).difference(total_checks))
        if absent_checks:
            raise WorkbookValidationError(
                "El perfil no declara comprobaciones para: " + ", ".join(absent_checks)
            )
        for key, expected in expected_totals_cents.items():
            sheet_name, reference = total_checks[key]
            if sheet_name not in values_book.sheetnames:
                raise WorkbookValidationError(f"No existe la hoja del total {key}: {sheet_name}")
            actual = _sum_reference(
                values_book[sheet_name], formula_book[sheet_name], reference
            )
            if actual != int(expected):
                raise WorkbookValidationError(
                    f"El total {key} no concilia: esperado {expected} céntimos, "
                    f"obtenido {actual} céntimos"
                )
    finally:
        values_book.close()
        formula_book.close()

    if expected_fingerprint is not None:
        actual_fingerprint = workbook_fingerprint(path, profile)
        if actual_fingerprint.sheet_names != expected_fingerprint.sheet_names:
            raise WorkbookValidationError("La estructura u orden de hojas cambió")
        if actual_fingerprint.print_areas != expected_fingerprint.print_areas:
            raise WorkbookValidationError("Las áreas de impresión cambiaron")
        if actual_fingerprint.required_formulas != expected_fingerprint.required_formulas:
            raise WorkbookValidationError("Las fórmulas declaradas cambiaron")
        if actual_fingerprint.merges != expected_fingerprint.merges:
            raise WorkbookValidationError("Las celdas combinadas cambiaron")
        if actual_fingerprint.dimensions != expected_fingerprint.dimensions:
            raise WorkbookValidationError("Las dimensiones de filas o columnas cambiaron")
        if actual_fingerprint.view_and_print_settings != expected_fingerprint.view_and_print_settings:
            raise WorkbookValidationError("La vista o configuración de impresión cambió")
        if actual_fingerprint.drawings != expected_fingerprint.drawings:
            raise WorkbookValidationError("Los gráficos, imágenes o sus anclas cambiaron")
        if actual_fingerprint.immutable_cells != expected_fingerprint.immutable_cells:
            raise WorkbookValidationError("Una celda o estilo fuera de las entradas cambió")
        if actual_fingerprint.ooxml_design_parts != expected_fingerprint.ooxml_design_parts:
            raise WorkbookValidationError("Las partes OOXML de diseño cambiaron")
