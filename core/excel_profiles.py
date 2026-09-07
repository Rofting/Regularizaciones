"""Perfiles declarativos para reproducir libros Excel por comunidad."""

from __future__ import annotations

import json
import hashlib
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
_ONBOARDING_SERVICE_MODULES = {
    "ACS": ("ACS",),
    "CALEFACCION": ("CALEFACCION",),
    "ACS+CALEFACCION": ("ACS", "CALEFACCION"),
    "NO_APLICA": (),
}
MODULE_CONCEPTS = {
    "ACS": ("acs_fixed", "acs_variable"),
    "CALEFACCION": ("heating_fixed", "heating_variable"),
}
_ONBOARDING_CONCEPT_RULES = {
    "acs_fixed": (
        "equal", "period_parameters.acs_fixed_actual",
        "period_parameters.acs_fixed_billed",
    ),
    "acs_variable": (
        "consumption", "period_parameters.acs_variable_actual",
        "period_parameters.acs_variable_billed",
    ),
    "heating_fixed": (
        "equal", "period_parameters.heating_fixed_actual",
        "period_parameters.heating_fixed_billed",
    ),
    "heating_variable": (
        "consumption", "period_parameters.heating_variable_actual",
        "period_parameters.heating_variable_billed",
    ),
}
_ONBOARDING_PARAMETER_CELLS = {
    "ACS": {
        "acs_variable_actual": ("LECTURAS ACS M3", "H8"),
        "acs_fixed_actual": ("LECTURAS ACS M3", "I9"),
        "acs_variable_billed": ("LECTURAS ACS M3", "H4"),
        "acs_fixed_billed": ("LECTURAS ACS M3", "I4"),
    },
    "CALEFACCION": {
        "heating_variable_actual": ("LECTURAS CALEF KWH", "H8"),
        "heating_fixed_actual": ("LECTURAS CALEF KWH", "I9"),
        "heating_variable_billed": ("LECTURAS CALEF KWH", "H4"),
        "heating_fixed_billed": ("LECTURAS CALEF KWH", "I4"),
    },
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
    source_sha256: str = ""
    onboarding_configuration: Mapping[str, Any] = field(
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


def canonical_concepts_for_module(module: str) -> tuple[dict[str, object], ...]:
    """Devuelve las reglas de reparto canónicas de un módulo de onboarding."""
    try:
        keys = MODULE_CONCEPTS[module]
    except KeyError:
        raise ValueError(f"Módulo de onboarding no soportado: {module}") from None
    return tuple(
        {
            "key": key,
            "allocation_method": _ONBOARDING_CONCEPT_RULES[key][0],
            "actual_source": _ONBOARDING_CONCEPT_RULES[key][1],
            "billed_source": _ONBOARDING_CONCEPT_RULES[key][2],
            "required": True,
        }
        for key in keys
    )


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


def _onboarding_configuration(
    value: Any,
    active_modules: tuple[str, ...],
) -> Mapping[str, Any]:
    if value is None:
        return MappingProxyType({})
    if not isinstance(value, dict):
        raise ValueError("onboarding_configuration debe ser un objeto")
    _require_fields(
        value,
        {"schema_version", "invoice_decisions", "not_applicable_modules"},
        "onboarding_configuration actual",
    )
    if value["schema_version"] != 2:
        raise ValueError("onboarding_configuration.schema_version no es válido")
    _require_fields(
        value,
        {
            "service_decision", "active_modules", "reading_column",
            "reading_bindings", "sources",
        },
        "onboarding_configuration",
    )
    configured_modules = _string_tuple(
        value["active_modules"], "onboarding_configuration.active_modules"
    )
    if configured_modules != active_modules:
        raise ValueError(
            "onboarding_configuration.active_modules no coincide con el perfil"
        )
    service_decision = value["service_decision"]
    if not isinstance(service_decision, str) or not service_decision:
        raise ValueError("onboarding_configuration.service_decision no es válido")
    if _ONBOARDING_SERVICE_MODULES.get(service_decision) != configured_modules:
        raise ValueError(
            "onboarding_configuration.service_decision contradice los módulos activos"
        )
    not_applicable_modules = _string_tuple(
        value.get("not_applicable_modules", []),
        "onboarding_configuration.not_applicable_modules",
    )
    if (
        set(not_applicable_modules).difference({"ACS", "CALEFACCION"})
        or set(not_applicable_modules).intersection(configured_modules)
    ):
        raise ValueError(
            "onboarding_configuration.not_applicable_modules contiene módulos activos"
        )
    if not isinstance(value["reading_column"], str):
        raise ValueError("onboarding_configuration.reading_column no es válido")
    bindings = value["reading_bindings"]
    if not isinstance(bindings, list):
        raise ValueError("onboarding_configuration.reading_bindings debe ser una lista")
    bound_modules: list[str] = []
    for index, binding in enumerate(bindings):
        if not isinstance(binding, dict):
            raise ValueError(f"El binding de lectura {index} no es un objeto")
        _require_fields(
            binding,
            {"module", "column", "source_sha256s"},
            f"onboarding_configuration.reading_bindings[{index}]",
        )
        _require_fields(
            binding,
            {"meter", "date", "value"},
            f"onboarding_configuration.reading_bindings[{index}] actual",
        )
        if (
            not isinstance(binding["module"], str)
            or not isinstance(binding["column"], str)
            or not binding["column"]
            or not isinstance(binding.get("meter", ""), str)
            or not isinstance(binding.get("date", ""), str)
            or not isinstance(binding.get("value", ""), str)
            or not all(binding[field] for field in ("meter", "date", "value"))
            or not isinstance(binding["source_sha256s"], list)
            or not binding["source_sha256s"]
            or not all(
                isinstance(item, str) and len(item) == 64
                for item in binding["source_sha256s"]
            )
        ):
            raise ValueError(f"Binding de lectura no válido para {binding.get('module')}")
        bound_modules.append(binding["module"])
    if tuple(bound_modules) != active_modules:
        raise ValueError("Cada módulo activo debe tener un binding de lectura")
    invoice_decisions = value.get("invoice_decisions", [])
    if not isinstance(invoice_decisions, list):
        raise ValueError("onboarding_configuration.invoice_decisions debe ser una lista")
    for index, decision in enumerate(invoice_decisions):
        if not isinstance(decision, dict):
            raise ValueError(f"La decisión de factura {index} no es un objeto")
        _require_fields(
            decision,
            {"source_sha256", "provider", "period", "amount", "concept"},
            f"onboarding_configuration.invoice_decisions[{index}]",
        )
        if (
            not isinstance(decision["source_sha256"], str)
            or len(decision["source_sha256"]) != 64
            or not all(
                isinstance(decision[field], str) and decision[field]
                for field in ("provider", "period", "amount", "concept")
            )
        ):
            raise ValueError(f"Decisión de factura no válida en posición {index}")
    sources = value["sources"]
    if not isinstance(sources, list) or not all(isinstance(item, dict) for item in sources):
        raise ValueError("onboarding_configuration.sources debe ser una lista de objetos")
    for index, source in enumerate(sources):
        _require_fields(source, {"kind", "name", "sha256"}, f"fuente {index}")
        if (
            not isinstance(source["kind"], str)
            or not source["kind"]
            or not isinstance(source["name"], str)
            or not source["name"]
            or "/" in source["name"]
            or "\\" in source["name"]
            or not isinstance(source["sha256"], str)
            or len(source["sha256"]) != 64
        ):
            raise ValueError(f"Traza de fuente no válida en posición {index}")
    invoice_source_sha256s = sorted(
        source["sha256"] for source in sources
        if source["kind"] == "invoice_pdf"
    )
    invoice_decision_sha256s = sorted(
        decision["source_sha256"] for decision in invoice_decisions
    )
    if invoice_source_sha256s != invoice_decision_sha256s:
        raise ValueError("Cada factura debe tener una decisión completa")
    return _freeze_json(value)


def _validate_onboarding_concepts(
    onboarding_configuration: Mapping[str, Any],
    active_modules: tuple[str, ...],
    concepts: tuple[ConceptRule, ...],
) -> None:
    if not onboarding_configuration:
        return
    expected_keys = tuple(
        key for module in active_modules for key in MODULE_CONCEPTS[module]
    )
    if tuple(concept.key for concept in concepts) != expected_keys:
        raise ValueError("Los conceptos del onboarding no son los canónicos")
    for concept in concepts:
        expected_method, expected_actual, expected_billed = (
            _ONBOARDING_CONCEPT_RULES[concept.key]
        )
        if (
            concept.allocation_method != expected_method
            or concept.actual_source != expected_actual
            or concept.billed_source != expected_billed
            or not concept.required
        ):
            raise ValueError(
                f"El concepto canónico {concept.key} no tiene sus fuentes y reparto esperados"
            )


def _validate_onboarding_parameter_cells(
    onboarding_configuration: Mapping[str, Any],
    active_modules: tuple[str, ...],
    workbook_layout: Mapping[str, Any],
) -> None:
    if not onboarding_configuration or not workbook_layout:
        return
    expected = {
        key: binding
        for module in active_modules
        for key, binding in _ONBOARDING_PARAMETER_CELLS[module].items()
    }
    actual = workbook_layout["parameter_cells"]
    if set(actual) != set(expected):
        raise ValueError(
            "Las celdas de parámetros del onboarding no son las canónicas"
        )
    for key, expected_binding in expected.items():
        if tuple(actual[key]) != expected_binding:
            raise ValueError(
                f"El binding canónico {key} no apunta a su celda esperada"
            )


def _validate_onboarding_bootstrap_marker(
    onboarding_configuration: Mapping[str, Any],
    workbook_layout: Mapping[str, Any],
) -> None:
    if not onboarding_configuration or not workbook_layout:
        return
    marker = workbook_layout.get("bootstrap_template")
    if (
        not isinstance(marker, Mapping)
        or marker.get("state") != "fresh_onboarding"
        or not isinstance(marker.get("sha256"), str)
        or re.fullmatch(r"[0-9a-f]{64}", marker["sha256"]) is None
    ):
        raise ValueError(
            "workbook_layout.bootstrap_template no identifica la plantilla inicial"
        )


def validate_profile_payload(payload: Mapping[str, Any], project_root: Path) -> ExcelProfile:
    """Validate an in-memory profile payload without creating any files."""
    if not isinstance(payload, Mapping):
        raise ValueError("El perfil Excel debe ser un objeto JSON")
    data = dict(payload)
    profile_key = data.get("key")
    if not isinstance(profile_key, str) or not profile_key or Path(profile_key).name != profile_key:
        raise ValueError("Clave de perfil no válida")
    _require_fields(data, _PROFILE_FIELDS, "perfil")

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

    concepts = _concepts(data["concepts"])
    workbook_layout = _workbook_layout(data.get("workbook_layout"))
    onboarding_configuration = _onboarding_configuration(
        data.get("onboarding_configuration"), active_modules
    )
    _validate_onboarding_concepts(
        onboarding_configuration, active_modules, concepts
    )
    _validate_onboarding_parameter_cells(
        onboarding_configuration, active_modules, workbook_layout
    )
    _validate_onboarding_bootstrap_marker(
        onboarding_configuration, workbook_layout
    )
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
        concepts=concepts,
        workbook_layout=workbook_layout,
        source_sha256="",
        onboarding_configuration=onboarding_configuration,
    )


def runtime_profile_path(profile_key: str, project_root: Path) -> Path:
    """Return the ignored local path reserved for onboarding runtime profiles."""
    if not profile_key or Path(profile_key).name != profile_key:
        raise ValueError("Clave de perfil no válida")
    return (
        Path(project_root) / "config" / "excel_profiles" / "runtime"
        / f"{profile_key}.json"
    )


def configured_profile_paths(project_root: Path) -> tuple[Path, ...]:
    """List public profiles plus non-shadowed local runtime profiles."""
    directory = Path(project_root) / "config" / "excel_profiles"
    public = sorted(directory.glob("*.json"))
    public_keys = {path.stem for path in public}
    runtime = [
        path for path in sorted((directory / "runtime").glob("*.json"))
        if path.stem not in public_keys
    ]
    return tuple((*public, *runtime))


def _profile_source_path(profile_key: str, project_root: Path) -> Path:
    public = Path(project_root) / "config" / "excel_profiles" / f"{profile_key}.json"
    if public.is_file():
        return public
    runtime = runtime_profile_path(profile_key, project_root)
    if runtime.is_file():
        return runtime
    return public


def load_profile(profile_key: str, project_root: Path) -> ExcelProfile:
    """Carga y valida un perfil versionado sin salir del directorio de configuración."""
    if not profile_key or Path(profile_key).name != profile_key:
        raise ValueError("Clave de perfil no válida")
    path = _profile_source_path(profile_key, project_root)
    try:
        source_bytes = path.read_bytes()
        data = json.loads(source_bytes.decode("utf-8"))
    except FileNotFoundError:
        raise LookupError(f"Perfil Excel no encontrado: {profile_key}") from None
    except json.JSONDecodeError as error:
        raise ValueError(f"JSON de perfil no válido: {error.msg}") from error
    if not isinstance(data, dict):
        raise ValueError("El perfil Excel debe ser un objeto JSON")
    if data.get("key") != profile_key:
        raise ValueError("La clave interna del perfil no coincide con el archivo")
    profile = validate_profile_payload(data, project_root)
    return ExcelProfile(
        key=profile.key,
        version=profile.version,
        community_code=profile.community_code,
        template_relative_path=profile.template_relative_path,
        active_modules=profile.active_modules,
        required_sheets=profile.required_sheets,
        required_formula_cells=profile.required_formula_cells,
        concepts=profile.concepts,
        workbook_layout=profile.workbook_layout,
        source_sha256=hashlib.sha256(source_bytes).hexdigest(),
        onboarding_configuration=profile.onboarding_configuration,
    )


def calculate_profile_sha256(profile: ExcelProfile, project_root: Path) -> str:
    """Calcula la huella actual de los bytes del JSON que define el perfil."""
    path = _profile_source_path(profile.key, Path(project_root).resolve())
    if path.is_file():
        return hashlib.sha256(path.read_bytes()).hexdigest()
    if profile.source_sha256:
        return profile.source_sha256
    raise LookupError(f"Perfil Excel no encontrado: {profile.key}")
