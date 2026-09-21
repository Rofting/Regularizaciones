"""Global, community-independent provider identity registry."""

from __future__ import annotations

import json
import re
import unicodedata
from dataclasses import dataclass
from pathlib import Path
from types import MappingProxyType
from typing import Mapping


_TAX_ID_PATTERN = re.compile(
    r"(?:[ABCDEFGHJNPQRSUVW]\d{7}[0-9A-J]|\d{8}[A-Z]|[XYZ]\d{7}[A-Z])"
)


@dataclass(frozen=True)
class ProviderProfile:
    key: str
    display_name: str
    tax_ids: tuple[str, ...]
    aliases: tuple[str, ...]
    document_types: tuple[str, ...]
    service_family: str
    extractor_family: str
    required_signatures: tuple[str, ...]
    excluded_signatures: tuple[str, ...]
    legacy: Mapping[str, object]


@dataclass(frozen=True)
class ProviderMatch:
    provider_key: str
    confidence: str
    rule_id: str
    evidence: tuple[str, ...]


def _normalize_words(value: str) -> str:
    decomposed = unicodedata.normalize("NFKD", value or "")
    without_accents = "".join(
        char for char in decomposed if not unicodedata.combining(char)
    )
    return re.sub(r"\s+", " ", without_accents.upper()).strip()


def normalize_tax_id(value: object) -> str:
    return re.sub(r"[^A-Z0-9]", "", str(value or "").upper())


def _valid_tax_id(value: str) -> bool:
    return bool(_TAX_ID_PATTERN.fullmatch(value))


def _as_strings(value: object) -> tuple[str, ...]:
    if value is None:
        return ()
    if isinstance(value, str):
        return (value,) if value.strip() else ()
    if not isinstance(value, (list, tuple)):
        raise ValueError(f"Se esperaba una lista de textos, no {type(value).__name__}")
    return tuple(str(item).strip() for item in value if str(item).strip())


def _default_document_types(raw: Mapping[str, object]) -> tuple[str, ...]:
    supply = str(raw.get("tipo_suministro") or "").upper()
    if bool(raw.get("es_justificante")):
        return ("bank_receipt",)
    if supply.startswith("LECTURA") or supply in {"CONTADORES", "LECTURAS"}:
        return ("reading",)
    return ("invoice", "credit_note")


def _default_service_family(raw: Mapping[str, object]) -> str:
    supply = str(raw.get("tipo_suministro") or "OTROS").strip().upper()
    if bool(raw.get("es_justificante")):
        return "PAGO"
    if supply.startswith("LECTURA"):
        return "CONTADORES"
    return supply or "OTROS"


def _default_extractor_family(service_family: str) -> str:
    if service_family == "ELECTRICIDAD":
        return "electricity"
    if service_family in {"GAS", "GASOLEO", "COMBUSTIBLE"}:
        return "gas_fuel"
    if service_family in {"AGUA", "TASAS", "CANON"}:
        return "water_public_fees"
    if service_family in {"CONTADORES", "LECTURAS"}:
        return "metering_management"
    if service_family in {"MANTENIMIENTO", "LIMPIEZA"}:
        return "periodic_maintenance"
    return "standard_spanish_invoice"


def provider_registry_from_payload(
    payload: Mapping[str, object],
) -> Mapping[str, ProviderProfile]:
    raw_profiles = payload.get("proveedores", {})
    if not isinstance(raw_profiles, Mapping):
        raise ValueError("El catálogo debe contener un objeto 'proveedores'")

    registry: dict[str, ProviderProfile] = {}
    tax_id_owners: dict[str, str] = {}
    for raw_key, value in raw_profiles.items():
        key = str(raw_key).strip().upper()
        if not key or not isinstance(value, Mapping):
            raise ValueError("Cada proveedor necesita una clave y un objeto de configuración")

        tax_ids = tuple(normalize_tax_id(item) for item in _as_strings(value.get("tax_ids")))
        for tax_id in tax_ids:
            if not _valid_tax_id(tax_id):
                raise ValueError(f"Identificador fiscal no válido en {key}: {tax_id}")
            previous = tax_id_owners.get(tax_id)
            if previous and previous != key:
                raise ValueError(
                    f"Identificador fiscal duplicado {tax_id}: {previous} y {key}"
                )
            tax_id_owners[tax_id] = key

        aliases = _as_strings(value.get("aliases"))
        if not aliases:
            aliases = _as_strings(value.get("firmas_identificacion"))

        document_types = _as_strings(value.get("document_types"))
        if not document_types:
            document_types = _default_document_types(value)
        document_types = tuple(item.lower() for item in document_types)

        service_family = str(
            value.get("service_family") or _default_service_family(value)
        ).strip().upper()
        extractor_family = str(
            value.get("extractor_family") or _default_extractor_family(service_family)
        ).strip()

        required = _as_strings(value.get("required_signatures"))
        if not required:
            required = _as_strings(value.get("firmas_requeridas"))
        excluded = _as_strings(value.get("excluded_signatures"))
        if not excluded:
            excluded = _as_strings(value.get("firma_exclusion"))

        registry[key] = ProviderProfile(
            key=key,
            display_name=str(value.get("display_name") or value.get("nombre_display") or key),
            tax_ids=tax_ids,
            aliases=aliases,
            document_types=document_types,
            service_family=service_family,
            extractor_family=extractor_family,
            required_signatures=required,
            excluded_signatures=excluded,
            legacy=MappingProxyType(dict(value)),
        )
    return MappingProxyType(registry)


def load_provider_registry(path: str | Path) -> Mapping[str, ProviderProfile]:
    with Path(path).open("r", encoding="utf-8") as handle:
        payload = json.load(handle)
    return provider_registry_from_payload(payload)


def _signature_matches(pattern: str, raw_text: str, normalized_text: str) -> bool:
    try:
        if re.search(pattern, raw_text, re.IGNORECASE):
            return True
    except re.error:
        pass
    normalized_pattern = _normalize_words(pattern)
    return bool(normalized_pattern and normalized_pattern in normalized_text)


def resolve_provider(
    registry: Mapping[str, ProviderProfile],
    text: str,
    filename: str,
    document_kind: str,
) -> ProviderMatch | None:
    raw_document = f"{filename or ''}\n{text or ''}"
    normalized_document = _normalize_words(raw_document)
    compact_document = re.sub(r"[^A-Z0-9]", "", raw_document.upper())
    kind = str(document_kind or "").lower()

    compatible = [
        profile for profile in registry.values() if kind in profile.document_types
    ]

    fiscal_matches: list[tuple[ProviderProfile, str]] = []
    for profile in compatible:
        for tax_id in profile.tax_ids:
            if tax_id in compact_document:
                fiscal_matches.append((profile, tax_id))
    if len(fiscal_matches) == 1:
        profile, tax_id = fiscal_matches[0]
        return ProviderMatch(
            profile.key,
            "high",
            "provider:tax-id:v1",
            (tax_id,),
        )
    if len({profile.key for profile, _ in fiscal_matches}) > 1:
        return None

    candidates: list[tuple[int, ProviderProfile, tuple[str, ...], bool]] = []
    for profile in compatible:
        if any(
            _signature_matches(pattern, raw_document, normalized_document)
            for pattern in profile.excluded_signatures
        ):
            continue

        aliases = tuple(
            alias
            for alias in profile.aliases
            if _signature_matches(alias, raw_document, normalized_document)
        )
        if not aliases:
            continue

        required_ok = all(
            _signature_matches(pattern, raw_document, normalized_document)
            for pattern in profile.required_signatures
        )
        if profile.required_signatures and not required_ok:
            continue

        score = 100 + (20 * len(aliases)) + (30 if required_ok and profile.required_signatures else 0)
        candidates.append((score, profile, aliases, required_ok))

    if not candidates:
        return None
    candidates.sort(key=lambda item: item[0], reverse=True)
    best_score = candidates[0][0]
    winners = [item for item in candidates if item[0] == best_score]
    if len(winners) != 1:
        return None

    _, profile, aliases, required_ok = winners[0]
    confidence = "high" if required_ok and profile.required_signatures else "medium"
    return ProviderMatch(
        profile.key,
        confidence,
        "provider:signature:v1",
        tuple(aliases),
    )


__all__ = [
    "ProviderMatch",
    "ProviderProfile",
    "load_provider_registry",
    "normalize_tax_id",
    "provider_registry_from_payload",
    "resolve_provider",
]
