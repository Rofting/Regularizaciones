"""Validación estructural y económica previa a publicar un Excel oficial."""

from __future__ import annotations

from dataclasses import dataclass
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
from typing import Mapping

from openpyxl import load_workbook
from openpyxl.utils.cell import range_boundaries

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


def _print_area(sheet) -> str:
    value = sheet.print_area
    return str(value) if value is not None else ""


def workbook_fingerprint(path: Path, profile: ExcelProfile) -> WorkbookFingerprint:
    workbook = load_workbook(path, data_only=False, read_only=False)
    try:
        formulas = []
        for sheet_name, address in profile.required_formula_cells:
            value = workbook[sheet_name][address].value if sheet_name in workbook.sheetnames else None
            formulas.append((sheet_name, address, value if isinstance(value, str) else ""))
        return WorkbookFingerprint(
            sheet_names=tuple(workbook.sheetnames),
            print_areas=tuple(
                (name, _print_area(workbook[name]))
                for name in profile.required_sheets
                if name in workbook.sheetnames
            ),
            required_formulas=tuple(formulas),
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
                continue
            total += _to_cents(cell.value)
    return total


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
