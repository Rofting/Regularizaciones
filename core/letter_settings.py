"""Configuración de conceptos que se muestran en las cartas de regularización."""

from __future__ import annotations

import json
import sqlite3
from dataclasses import dataclass
from pathlib import Path


@dataclass(frozen=True)
class LetterConcept:
    key: str
    label: str
    service: str | None
    unit: str | None
    display_order: int


@dataclass(frozen=True)
class LetterIdentity:
    """Identidad de despacho configurable por comunidad, sin marca implícita."""

    office_name: str
    footer: str
    signature: str
    city: str
    logo_path: Path | None = None


def load_community_letter_identity(project_root: Path, community_code: str) -> LetterIdentity:
    """Carga una identidad opcional y segura para las cartas de una comunidad.

    El archivo no contiene datos de reparto y puede instalarse en cada despacho:
    ``config/letter_identities.json``. Un logo sólo se acepta si queda dentro
    del proyecto para impedir que una configuración abra rutas arbitrarias.
    """
    root = Path(project_root).resolve()
    defaults = {
        "office_name": "Administración de fincas",
        "footer": "Atención de la comunidad",
        "signature": "La Administración",
        "city": "",
    }
    path = root / "config" / "letter_identities.json"
    if path.is_file():
        try:
            payload = json.loads(path.read_text(encoding="utf-8"))
        except json.JSONDecodeError as error:
            raise ValueError("La configuración de identidad de cartas no es válida") from error
        if not isinstance(payload, dict):
            raise ValueError("La configuración de identidad de cartas no es válida")
        communities = payload.get("communities", {})
        if not isinstance(communities, dict):
            raise ValueError("communities debe ser un objeto en la identidad de cartas")
        configured = communities.get(str(community_code), {})
        if not isinstance(configured, dict):
            raise ValueError("La identidad de la comunidad no es válida")
        for key in defaults:
            value = configured.get(key)
            if value is not None:
                if not isinstance(value, str) or not value.strip():
                    raise ValueError(f"{key} debe ser un texto no vacío")
                defaults[key] = value.strip()
        logo_raw = configured.get("logo_path")
    else:
        logo_raw = None

    logo_path = None
    if logo_raw is not None:
        if not isinstance(logo_raw, str) or not logo_raw.strip():
            raise ValueError("logo_path debe ser una ruta relativa no vacía")
        candidate = (root / logo_raw).resolve()
        try:
            candidate.relative_to(root)
        except ValueError:
            raise ValueError("logo_path debe quedar dentro del proyecto") from None
        if candidate.is_file():
            logo_path = candidate
    return LetterIdentity(logo_path=logo_path, **defaults)


def available_concepts(connection: sqlite3.Connection) -> tuple[LetterConcept, ...]:
    rows = connection.execute(
        """
        SELECT concept_key, label, service, unit, display_order
        FROM regularization_concepts
        WHERE active=1 ORDER BY display_order, concept_key
        """
    ).fetchall()
    return tuple(LetterConcept(*row) for row in rows)


def _default_keys(connection: sqlite3.Connection, id_comunidad: int) -> tuple[str, ...]:
    """Selecciona por defecto solo conceptos con resultados en la comunidad."""
    rows = connection.execute(
        """
        SELECT DISTINCT r.concept_key
        FROM owner_concept_results r
        JOIN propietarios p ON p.id_propietario=r.id_propietario
        WHERE p.id_comunidad=?
        ORDER BY r.concept_key
        """,
        (id_comunidad,),
    ).fetchall()
    keys = {row[0] for row in rows}
    ordered = [concept.key for concept in available_concepts(connection) if concept.key in keys]
    return tuple(ordered)


def load_selected_concepts(
    connection: sqlite3.Connection, id_comunidad: int
) -> tuple[str, ...]:
    row = connection.execute(
        "SELECT selected_concepts_json FROM community_letter_settings WHERE id_comunidad=?",
        (id_comunidad,),
    ).fetchone()
    if row:
        try:
            selected = tuple(str(value) for value in json.loads(row[0]))
        except (TypeError, ValueError, json.JSONDecodeError) as exc:
            raise ValueError("La configuración de conceptos está dañada") from exc
        known = {concept.key for concept in available_concepts(connection)}
        unknown = sorted(set(selected) - known)
        if unknown:
            raise ValueError("Conceptos no disponibles: " + ", ".join(unknown))
        return tuple(key for key in (concept.key for concept in available_concepts(connection)) if key in selected)
    return _default_keys(connection, id_comunidad)


def save_selected_concepts(
    connection: sqlite3.Connection, id_comunidad: int, selected: tuple[str, ...] | list[str]
) -> tuple[str, ...]:
    available = available_concepts(connection)
    order = [concept.key for concept in available]
    selected_set = {str(key) for key in selected}
    unknown = sorted(selected_set - set(order))
    if unknown:
        raise ValueError("Conceptos no disponibles: " + ", ".join(unknown))
    if not selected_set:
        raise ValueError("Debe seleccionarse al menos un concepto")
    normalized = tuple(key for key in order if key in selected_set)
    connection.execute(
        """
        INSERT INTO community_letter_settings(id_comunidad, selected_concepts_json, updated_at)
        VALUES (?, ?, datetime('now'))
        ON CONFLICT(id_comunidad) DO UPDATE SET
            selected_concepts_json=excluded.selected_concepts_json,
            updated_at=datetime('now')
        """,
        (id_comunidad, json.dumps(normalized, ensure_ascii=False)),
    )
    connection.commit()
    return normalized
