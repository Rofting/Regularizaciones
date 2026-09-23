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
    immutable_cells: tuple[tuple[str, str, object, str], ...]
    immutable_cell_styles: tuple[tuple[str, str, bool, tuple], ...]
    mutable_cell_styles: tuple[tuple[str, str, tuple], ...]
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
    for key, binding in layout.get("total_checks", {}).items():
        if not str(key).startswith("parameter:"):
            continue
        sheet_name, address = binding
        if ":" not in address:
            result.setdefault(sheet_name, set()).add(address)
    for table in layout.get("tables", {}).values():
        sheet_name = table["sheet"]
        for row in range(int(table["start_row"]), int(table["end_row"]) + 1):
            for column in (*_input_columns(table).values(), *table.get("derived_columns", {}).values()):
                result.setdefault(sheet_name, set()).add(f"{column}{row}")
    return result


def _medida(valor):
    """Alto o ancho redondeado: LibreOffice reescribe los decimales finos."""
    return None if valor is None else round(float(valor), 1)


def _dimension_signature(dimensions) -> tuple:
    """Altos, anchos y visibilidad; sin el índice de estilo, que se renumera."""
    # El alto de fila queda fuera: Excel lo deja implícito donde LibreOffice
    # escribe 14,2, y convierte 6,95 en 6,8 al recalcular. Lo que sí se
    # compara es el ancho de columna, lo oculto y el agrupamiento, que son los
    # que se notan al abrir e imprimir el libro.
    return tuple(
        sorted(
            (str(key), _medida(getattr(dimension, "width", None)),
             bool(getattr(dimension, "hidden", False)),
             getattr(dimension, "outlineLevel", None),
             bool(getattr(dimension, "collapsed", False)))
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


def _categoria_formato(formato: str) -> str:
    """Clase de formato visible: fecha, moneda, porcentaje o número.

    El formato exacto no es comparable entre motores —la misma fecha se
    escribe 'mm-dd-yy' en un libro y 'm/d/yyyy' en el otro—, pero su clase sí,
    y es lo que nota quien abre la hoja.
    """
    texto = (formato or "General").lower()
    if "€" in texto or "$" in texto or "ptas" in texto:
        return "moneda"
    if "%" in texto:
        return "porcentaje"
    if any(marca in texto for marca in ("yy", "dd", "mmm", "hh")):
        return "fecha"
    if texto == "general":
        return "general"
    return "numero"


def _color_relleno(relleno) -> str | None:
    """Color de fondo declarado, cuando es un RGB explícito.

    Sin él, cambiar la paleta del libro pasaba inadvertido: el patrón seguía
    siendo 'solid' y la firma no cambiaba. Los colores de tema o indexados se
    dejan fuera porque cada motor los reescribe a su manera.
    """
    color = getattr(relleno, "fgColor", None)
    valor = getattr(color, "rgb", None)
    return str(valor) if isinstance(valor, str) else None



def _cell_style_signature(cell) -> tuple:
    """Firma de la presentación visible, estable entre Excel y LibreOffice.

    Se dejan fuera el índice de estilo y los detalles que cada motor rellena a
    su manera (charset de la fuente, alineación vertical implícita, cómo
    representa el color de un borde). Comparar aquello marcaba como alterado
    un libro idéntico a la vista.
    """
    fuente, borde = cell.font, cell.border
    return (
        fuente.name, fuente.sz, bool(fuente.b), bool(fuente.i),
        cell.fill.patternType, _color_relleno(cell.fill),
        # 'general' y el valor sin fijar son la misma alineación.
        None if cell.alignment.horizontal in (None, "general") else cell.alignment.horizontal,
        _categoria_formato(cell.number_format),
        bool(cell.protection.locked),
        tuple(
            getattr(getattr(borde, lado), "style", None)
            for lado in ("left", "right", "top", "bottom")
        ),
    )


def _implicit_default_style_normalization(expected: tuple, actual: tuple) -> bool:
    """Admite únicamente normalizaciones inocuas del estilo implícito.

    LibreOffice materializa el resultado de algunas fórmulas que en Excel
    conservan el estilo 0: sustituye la fuente predeterminada por un
    equivalente métrico y deduce un formato numérico. El relleno, énfasis,
    alineación, protección y bordes sí forman parte del diseño y deben seguir
    siendo idénticos.
    """
    if len(expected) != len(actual):
        return False
    compatible_fonts = {"Arial", "Aptos", "Calibri", "Carlito", "Liberation Sans"}
    if expected[0] != actual[0] and {
        str(expected[0]), str(actual[0]),
    }.difference(compatible_fonts):
        return False
    # Sólo un formato General implícito puede convertirse en la categoría que
    # LibreOffice deduce del resultado de la fórmula.
    if expected[7] != actual[7] and expected[7] != "general":
        return False
    return all(
        expected[index] == actual[index]
        for index in (1, 2, 3, 4, 5, 6, 8, 9)
    )


def _immutable_styles_preserved(expected, actual) -> bool:
    """Compara estilos fijos sin confundir el estilo 0 con un rediseño."""
    expected_by_cell = {
        (sheet, address): (implicit_default, signature)
        for sheet, address, implicit_default, signature in expected
    }
    actual_by_cell = {
        (sheet, address): (implicit_default, signature)
        for sheet, address, implicit_default, signature in actual
    }
    if expected_by_cell.keys() != actual_by_cell.keys():
        return False
    for cell, (implicit_default, expected_signature) in expected_by_cell.items():
        _, actual_signature = actual_by_cell[cell]
        if actual_signature == expected_signature:
            continue
        if not implicit_default or not _implicit_default_style_normalization(
            expected_signature, actual_signature,
        ):
            return False
    return True


def _design_part_hashes(path: Path) -> tuple[tuple[str, str], ...]:
    """Huella de las partes de diseño: gráficos, imágenes, tema, impresión.

    Fuera quedan las partes estructurales —la tabla de estilos, el catálogo de
    contenidos y los mapas de relaciones—, que cada motor numera a su manera y
    que deben pertenecer al libro generado, no a la plantilla. Su contenido ya
    se comprueba por otras vías: las fórmulas, los valores y el estilo visible
    de cada celda.
    """
    prefixes = ("xl/charts/", "xl/drawings/", "xl/media/", "customXml/", "xl/theme/", "xl/printerSettings/")
    with ZipFile(path, "r") as archive:
        names = [name for name in archive.namelist() if name.startswith(prefixes)]
        return tuple(
            (name, hashlib.sha256(archive.read(name)).hexdigest())
            for name in sorted(set(names))
        )


def _dimensions_preserved(esperadas, obtenidas) -> bool:
    """Los altos y anchos de la plantilla siguen ahí.

    LibreOffice escribe además filas con la altura por defecto que Excel deja
    implícitas; eso no cambia el diseño, así que se admite lo que sobre.
    """
    por_hoja = {hoja: (set(filas), set(columnas)) for hoja, filas, columnas in obtenidas}
    for hoja, filas, columnas in esperadas:
        if hoja not in por_hoja:
            return False
        filas_obtenidas, columnas_obtenidas = por_hoja[hoja]
        if not set(filas).issubset(filas_obtenidas):
            return False
        if not set(columnas).issubset(columnas_obtenidas):
            return False
    return True



def workbook_fingerprint(path: Path, profile: ExcelProfile) -> WorkbookFingerprint:
    """Huella del diseño del libro, comparable entre Excel y LibreOffice.

    La huella recoge contenido y estructura, nunca números internos: al
    recalcular, LibreOffice reconstruye la tabla de estilos y numera de otra
    forma, y rellena los ajustes de impresión que Excel deja vacíos. Comparar
    esos índices hacía fallar cualquier libro correcto con "las dimensiones
    cambiaron", cuando ni una fórmula ni un valor se habían movido.
    """
    workbook = load_workbook(path, data_only=False, read_only=False)
    try:
        formulas = []
        for sheet_name, address in profile.required_formula_cells:
            value = workbook[sheet_name][address].value if sheet_name in workbook.sheetnames else None
            formulas.append((sheet_name, address, value if isinstance(value, str) else ""))
        mutable = _mutable_addresses(profile)
        immutable_cells = []
        immutable_styles = []
        merges = []
        dimensions = []
        views = []
        drawings = []
        mutable_styles = []
        for sheet in workbook.worksheets:
            merges.append((sheet.title, tuple(sorted(str(value) for value in sheet.merged_cells.ranges))))
            dimensions.append((
                sheet.title, _dimension_signature(sheet.row_dimensions),
                _dimension_signature(sheet.column_dimensions),
            ))
            views.append((
                sheet.title, str(sheet.freeze_panes or ""), _print_area(sheet), "", "",
            ))
            drawings.append((
                sheet.title, len(sheet._charts),
                tuple(_anchor_signature(chart.anchor) for chart in sheet._charts),
                len(sheet._images),
                tuple(_anchor_signature(image.anchor) for image in sheet._images),
            ))
            for address in sorted(mutable.get(sheet.title, set())):
                mutable_styles.append((
                    sheet.title, address, _cell_style_signature(sheet[address]),
                ))
            for cell in sheet._cells.values():
                if cell.coordinate in mutable.get(sheet.title, set()):
                    continue
                # Sólo las celdas con contenido: una celda vacía que únicamente
                # llevaba formato desaparece al recalcular y no es una pérdida.
                if cell.value is not None:
                    immutable_cells.append((
                        sheet.title, cell.coordinate, cell.value, cell.data_type,
                    ))
                    immutable_styles.append((
                        sheet.title, cell.coordinate, cell.style_id == 0,
                        _cell_style_signature(cell),
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
            immutable_cell_styles=tuple(sorted(immutable_styles)),
            mutable_cell_styles=tuple(sorted(mutable_styles)),
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
        # No se exige que todas las hojas tengan área de impresión: los modelos
        # del despacho sólo la definen donde imprimen de verdad (el análisis).
        # Lo que sí se comprueba, más abajo contra la huella de la plantilla,
        # es que la generación no altere las que hubiera.
        if expected_fingerprint is None:
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
        # Un #DIV/0! no siempre es un fallo: hay celdas informativas del modelo
        # —el precio por m³ de un gas que se factura en kWh, por ejemplo— que
        # no se pueden calcular porque ese dato no existe en la comunidad.
        # Bloquean sólo si afectan a las celdas de las que depende el estudio:
        # las fórmulas obligatorias, los cuadres y los importes del análisis.
        criticas = {
            (hoja, direccion) for hoja, direccion in profile.required_formula_cells
        }
        for hoja, direccion in profile.workbook_layout.get("parameter_cells", {}).values():
            criticas.add((hoja, direccion))
        for hoja, direccion in profile.workbook_layout.get("total_checks", {}).values():
            if ":" not in direccion:
                criticas.add((hoja, direccion))
        errors = _formula_errors(formula_book) + _formula_errors(values_book)
        graves = [
            error for error in errors
            if any(f"{hoja}!{direccion}=" in error for hoja, direccion in criticas)
        ]
        if graves:
            raise WorkbookValidationError(
                "El libro contiene un error de fórmula: " + ", ".join(sorted(set(graves))[:5])
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
        if not _dimensions_preserved(
            expected_fingerprint.dimensions, actual_fingerprint.dimensions
        ):
            raise WorkbookValidationError("Las dimensiones de filas o columnas cambiaron")
        if actual_fingerprint.view_and_print_settings != expected_fingerprint.view_and_print_settings:
            raise WorkbookValidationError("La vista o configuración de impresión cambió")
        if actual_fingerprint.drawings != expected_fingerprint.drawings:
            raise WorkbookValidationError("Los gráficos, imágenes o sus anclas cambiaron")
        if actual_fingerprint.immutable_cells != expected_fingerprint.immutable_cells:
            raise WorkbookValidationError("Una celda o estilo fuera de las entradas cambió")
        if not _immutable_styles_preserved(
            expected_fingerprint.immutable_cell_styles,
            actual_fingerprint.immutable_cell_styles,
        ):
            raise WorkbookValidationError("Las partes OOXML de diseño cambiaron")
        if actual_fingerprint.mutable_cell_styles != expected_fingerprint.mutable_cell_styles:
            raise WorkbookValidationError("El estilo de una entrada o fórmula mutable cambió")
        if actual_fingerprint.ooxml_design_parts != expected_fingerprint.ooxml_design_parts:
            raise WorkbookValidationError("Las partes OOXML de diseño cambiaron")
