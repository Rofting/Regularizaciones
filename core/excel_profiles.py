"""Perfiles declarativos para reproducir libros Excel por comunidad."""

from __future__ import annotations

import json
from dataclasses import dataclass
from pathlib import Path, PurePath
from typing import Any


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


def _validate_relative_template_path(raw_path: Any) -> str:
    if not isinstance(raw_path, str) or not raw_path.strip():
        raise ValueError("template_relative_path debe ser una ruta relativa no vacía")
    path = PurePath(raw_path)
    if path.is_absolute() or ".." in path.parts:
        raise ValueError("template_relative_path debe permanecer dentro del proyecto")
    return raw_path


def _string_tuple(value: Any, field: str) -> tuple[str, ...]:
    if not isinstance(value, list) or not all(
        isinstance(item, str) and item for item in value
    ):
        raise ValueError(f"{field} debe ser una lista de textos no vacíos")
    if len(set(value)) != len(value):
        raise ValueError(f"{field} contiene valores duplicados")
    return tuple(value)


def _formula_cells(value: Any) -> tuple[tuple[str, str], ...]:
    if not isinstance(value, list):
        raise ValueError("required_formula_cells debe ser una lista")
    result: list[tuple[str, str]] = []
    for item in value:
        if (
            not isinstance(item, list)
            or len(item) != 2
            or not all(isinstance(part, str) and part for part in item)
        ):
            raise ValueError("Cada fórmula obligatoria debe indicar hoja y celda")
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
            data["template_relative_path"]
        ),
        active_modules=active_modules,
        required_sheets=required_sheets,
        required_formula_cells=formula_cells,
        concepts=_concepts(data["concepts"]),
    )
