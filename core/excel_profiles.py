"""Perfiles declarativos para reproducir libros Excel por comunidad."""

from __future__ import annotations

import json
import re
from dataclasses import dataclass, field
from pathlib import Path, PurePosixPath, PureWindowsPath
from types import MappingProxyType
from typing import Any, Mapping


SUPPORTED_ALLOCATION_METHODS = frozenset({
    "equal",
    "coefficient",
    "consumption",
    "direct",
})

MODULE_REQUIRED_SHEETS = {
    "GAS": ("GAS",),
    "ELECTRICIDAD": ("ELECTRICIDAD",),
    "AGUA": ("AGUA",),
    "OTROS_GASTOS": ("OTROS GASTOS",),
    "ACS": ("LECTURAS ACS M3", "ANALISIS"),
    "CALEFACCION": ("LECTURAS CALEF KWH",),
}


@dataclass(frozen=True)
class ConceptRule:
    key: str
    allocation_method: str
    actual_source: str
    billed_source: str | None
    required: bool


@dataclass(frozen=True)
class ExcelProfile:
    key: str
    version: str
    community_code: str
    template_relative_path: str
    active_modules: tuple[str, ...]
    required_sheets: tuple[str, ...]
    required_formula_cells: tuple[tuple[str, str], ...]
    concepts: tuple[ConceptRule, ...]
    workbook_layout: Mapping[str, Any] = field(
        default_factory=lambda: MappingProxyType({})
    )


_PROFILE_FIELDS = {
    "key",
    "version",
    "community_code",
    "template_relative_path",
    "active_modules",
    "required_sheets",
    "required_formula_cells",
    "concepts",
}
_CONCEPT_FIELDS = {
    "key",
    "allocation_method",
    "actual_source",
    "billed_source",
    "required",
}


def _require_fields(data: dict[str, Any], required: set[str], context: str) -> None:
    missing = sorted(required.difference(data))
    if missing:
        raise ValueError(f"Faltan campos en {context}: {', '.join(missing)}")


def _validate_relative_template_path(raw_path: Any, project_root: Path) -> str:
    if not isinstance(raw_path, str) or not raw_path.strip():
        raise ValueError("template_relative_path debe ser una ruta relativa no vacía")
    windows_path = PureWindowsPath(raw_path)
    posix_path = PurePosixPath(raw_path)
    if (
        windows_path.drive
        or windows_path.root
        or windows_path.anchor
        or posix_path.drive
        or posix_path.root
        or posix_path.anchor
        or ".." in windows_path.parts
        or ".." in posix_path.parts
    ):
        raise ValueError("template_relative_path debe ser una ruta relativa del proyecto")
    root = Path(project_root).resolve()
    resolved = (root / raw_path).resolve()
    try:
        resolved.relative_to(root)
    except ValueError:
        raise ValueError(
            "template_relative_path debe ser una ruta relativa del proyecto"
        ) from None
    return raw_path


def _string_tuple(value: Any, field: str) -> tuple[str, ...]:
    if not isinstance(value, list) or not all(
        isinstance(item, str) and item for item in value
    ):
        raise ValueError(f"{field} debe ser una lista de textos no vacíos")
    if len(set(value)) != len(value):
        raise ValueError(f"{field} contiene valores duplicados")
    return tuple(value)


_A1_CELL_REFERENCE = re.compile(r"\$?([A-Za-z]{1,3})\$?([1-9][0-9]*)\Z")


def _is_a1_cell_reference(value: str) -> bool:
    match = _A1_CELL_REFERENCE.fullmatch(value)
    if match is None:
        return False
    column_letters, row_text = match.groups()
    column = 0
    for letter in column_letters.upper():
        column = column * 26 + ord(letter) - ord("A") + 1
    return column <= 16_384 and int(row_text) <= 1_048_576


def _formula_cells(value: Any) -> tuple[tuple[str, str], ...]:
    if not isinstance(value, list):
        raise ValueError("required_formula_cells debe ser una lista")
    result: list[tuple[str, str]] = []
    for item in value:
        if (
            not isinstance(item, list)
            or len(item) != 2
            or not isinstance(item[0], str)
            or not item[0]
            or not isinstance(item[1], str)
            or not _is_a1_cell_reference(item[1])
        ):
            raise ValueError(
                "Cada fórmula obligatoria debe indicar una hoja y una única celda A1"
            )
        result.append((item[0], item[1]))
    if len(set(result)) != len(result):
        raise ValueError("required_formula_cells contiene valores duplicados")
    return tuple(result)


def _concepts(value: Any) -> tuple[ConceptRule, ...]:
    if not isinstance(value, list):
        raise ValueError("concepts debe ser una lista")
    concepts: list[ConceptRule] = []
    keys: set[str] = set()
    for index, item in enumerate(value):
        if not isinstance(item, dict):
            raise ValueError(f"El concepto {index} no es un objeto")
        _require_fields(item, _CONCEPT_FIELDS, f"concepts[{index}]")
        key = item["key"]
        method = item["allocation_method"]
        if not isinstance(key, str) or not key:
            raise ValueError("Todo concepto debe tener una clave no vacía")
        if key in keys:
            raise ValueError(f"Clave de concepto duplicada: {key}")
        if method not in SUPPORTED_ALLOCATION_METHODS:
            raise ValueError(f"Método de reparto no soportado: {method}")
        if not isinstance(item["actual_source"], str) or not item["actual_source"]:
            raise ValueError(f"actual_source no válido para {key}")
        billed_source = item["billed_source"]
        if billed_source is not None and (
            not isinstance(billed_source, str) or not billed_source
        ):
            raise ValueError(f"billed_source no válido para {key}")
        if not isinstance(item["required"], bool):
            raise ValueError(f"required debe ser booleano para {key}")
        keys.add(key)
        concepts.append(ConceptRule(
            key=key,
            allocation_method=method,
            actual_source=item["actual_source"],
            billed_source=billed_source,
            required=item["required"],
        ))
    return tuple(concepts)


def _freeze_json(value: Any) -> Any:
    """Convierte la configuración JSON en una estructura de solo lectura."""
    if isinstance(value, dict):
        return MappingProxyType({key: _freeze_json(item) for key, item in value.items()})
    if isinstance(value, list):
        return tuple(_freeze_json(item) for item in value)
    return value


def _workbook_layout(value: Any) -> Mapping[str, Any]:
    if value is None:
        return MappingProxyType({})
    if not isinstance(value, dict):
        raise ValueError("workbook_layout debe ser un objeto")
    for required in ("metadata_cells", "tables", "parameter_cells", "total_checks"):
        if required not in value or not isinstance(value[required], dict):
            raise ValueError(f"workbook_layout.{required} debe ser un objeto")
    for module, table in value["tables"].items():
        if not isinstance(table, dict):
            raise ValueError(f"La tabla {module} debe ser un objeto")
        for key in ("sheet", "start_row", "end_row"):
            if key not in table:
                raise ValueError(f"Falta workbook_layout.tables.{module}.{key}")
        if (
            not isinstance(table["start_row"], int)
            or not isinstance(table["end_row"], int)
            or table["start_row"] < 1
            or table["end_row"] < table["start_row"]
        ):
            raise ValueError(f"Rango de filas no válido para {module}")
        legacy_columns = table.get("columns")
        input_columns = table.get("input_columns")
        derived_columns = table.get("derived_columns")
        if legacy_columns is not None:
            if input_columns is not None or derived_columns is not None:
                raise ValueError(
                    f"La tabla {module} no puede mezclar columnas antiguas y separadas"
                )
            if not isinstance(legacy_columns, dict) or not legacy_columns:
                raise ValueError(f"La tabla {module} no declara columnas")
        else:
            if not isinstance(input_columns, dict) or not input_columns:
                raise ValueError(f"La tabla {module} no declara columnas de entrada")
            if not isinstance(derived_columns, dict):
                raise ValueError(f"La tabla {module} no declara columnas derivadas")
            overlap = set(input_columns).intersection(derived_columns)
            if overlap:
                raise ValueError(
                    f"La tabla {module} repite columnas mutables y derivadas: "
                    + ", ".join(sorted(overlap))
                )
    return _freeze_json(value)


def load_profile(profile_key: str, project_root: Path) -> ExcelProfile:
    """Carga y valida un perfil versionado sin salir del directorio de configuración."""
    if not profile_key or Path(profile_key).name != profile_key:
        raise ValueError("Clave de perfil no válida")
    path = Path(project_root) / "config" / "excel_profiles" / f"{profile_key}.json"
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except FileNotFoundError:
        raise LookupError(f"Perfil Excel no encontrado: {profile_key}") from None
    except json.JSONDecodeError as error:
        raise ValueError(f"JSON de perfil no válido: {error.msg}") from error
    if not isinstance(data, dict):
        raise ValueError("El perfil Excel debe ser un objeto JSON")
    _require_fields(data, _PROFILE_FIELDS, "perfil")
    if data["key"] != profile_key:
        raise ValueError("La clave interna del perfil no coincide con el archivo")

    active_modules = _string_tuple(data["active_modules"], "active_modules")
    unsupported_modules = sorted(set(active_modules).difference(MODULE_REQUIRED_SHEETS))
    if unsupported_modules:
        raise ValueError(f"Módulos no soportados: {', '.join(unsupported_modules)}")
    required_sheets = _string_tuple(data["required_sheets"], "required_sheets")
    missing_sheets = sorted({
        sheet
        for module in active_modules
        for sheet in MODULE_REQUIRED_SHEETS[module]
        if sheet not in required_sheets
    })
    if missing_sheets:
        raise ValueError(
            f"Faltan hojas obligatorias para los módulos activos: {', '.join(missing_sheets)}"
        )
    formula_cells = _formula_cells(data["required_formula_cells"])
    formula_sheets = {sheet for sheet, _ in formula_cells}
    absent_formula_sheets = sorted(formula_sheets.difference(required_sheets))
    if absent_formula_sheets:
        raise ValueError(
            "Las fórmulas obligatorias usan hojas no declaradas: "
            + ", ".join(absent_formula_sheets)
        )
    for field in ("version", "community_code"):
        if not isinstance(data[field], str) or not data[field]:
            raise ValueError(f"{field} debe ser un texto no vacío")

    return ExcelProfile(
        key=profile_key,
        version=data["version"],
        community_code=data["community_code"],
        template_relative_path=_validate_relative_template_path(
            data["template_relative_path"], project_root
        ),
        active_modules=active_modules,
        required_sheets=required_sheets,
        required_formula_cells=formula_cells,
        concepts=_concepts(data["concepts"]),
        workbook_layout=_workbook_layout(data.get("workbook_layout")),
    )
