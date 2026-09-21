"""Alta automática de la plantilla Excel de una comunidad.

El despacho usa el mismo modelo de estudio para todas las comunidades: cambian
los datos y los módulos activos, no la estructura. Antes, cada comunidad exigía
que alguien eligiera a mano su Excel maestro antes de poder generar nada, y sin
ese paso el expediente se quedaba bloqueado aunque tuviera todas sus fuentes.

Aquí se crea sola a partir del modelo canónico. La comunidad que ya tenga su
libro propio puede seguir instalándolo con «Importar modelo inicial»: esto sólo
actúa cuando no hay ninguno.
"""

from __future__ import annotations

import json
import shutil
import sqlite3
from dataclasses import dataclass
from pathlib import Path

from excel_profiles import load_profile, runtime_profile_path, validate_profile_payload

MODELO_RELATIVO = "plantillas/modelo/modelo_acs_v1.xlsx"
PERFIL_BASE_RELATIVO = "config/modelo_excel/perfil_base.json"


class PlantillaNoDisponible(RuntimeError):
    """Falta el modelo canónico del que partir."""


@dataclass(frozen=True)
class PlantillaCreada:
    clave: str
    perfil: Path
    plantilla: Path
    reutilizada: bool


def _clave(codigo: str) -> str:
    return f"{codigo}_acs_v1"


def _perfil_existente(project_root: Path, codigo: str) -> str | None:
    """Clave del perfil ya configurado para esa comunidad, si lo hay."""
    from excel_profiles import configured_profile_paths

    for ruta in configured_profile_paths(project_root):
        try:
            perfil = load_profile(ruta.stem, project_root)
        except (LookupError, ValueError):
            continue
        if perfil.community_code == str(codigo):
            return perfil.key
    return None


def asegurar_plantilla(
    connection: sqlite3.Connection,
    *,
    community_id: int,
    project_root: Path,
) -> PlantillaCreada | None:
    """Garantiza que la comunidad tenga perfil y plantilla con los que generar.

    Devuelve ``None`` cuando ya había una plantilla registrada y utilizable, de
    modo que llamar a esto siempre es seguro.
    """
    root = Path(project_root).resolve()
    fila = connection.execute(
        "SELECT codigo FROM comunidades WHERE id_comunidad=?", (community_id,)
    ).fetchone()
    if fila is None:
        raise LookupError("La comunidad no existe")
    codigo = str(fila["codigo"] if hasattr(fila, "keys") else fila[0]).strip()

    registrado = connection.execute(
        """SELECT profile_key FROM excel_template_profiles
           WHERE id_comunidad=? AND status='active' LIMIT 1""",
        (community_id,),
    ).fetchone()
    if registrado is not None:
        return None

    clave = _perfil_existente(root, codigo)
    if clave is not None:
        perfil = load_profile(clave, root)
        plantilla = (root / perfil.template_relative_path).resolve()
        if plantilla.is_file():
            # Perfil y libro ya preparados: generar los registrará.
            return PlantillaCreada(clave, Path(), plantilla, reutilizada=True)
        _copiar_modelo(root, plantilla)
        return PlantillaCreada(clave, Path(), plantilla, reutilizada=False)

    clave = _clave(codigo)
    destino_perfil = runtime_profile_path(clave, root)
    payload = _perfil_para(root, codigo, clave)
    plantilla = (root / payload["template_relative_path"]).resolve()
    _copiar_modelo(root, plantilla)
    validate_profile_payload(payload, root)
    destino_perfil.parent.mkdir(parents=True, exist_ok=True)
    destino_perfil.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8",
    )
    return PlantillaCreada(clave, destino_perfil, plantilla, reutilizada=False)


def _perfil_para(root: Path, codigo: str, clave: str) -> dict:
    origen = root / PERFIL_BASE_RELATIVO
    if not origen.is_file():
        raise PlantillaNoDisponible(
            f"Falta el perfil base del modelo canónico: {origen}"
        )
    texto = origen.read_text(encoding="utf-8")
    payload = json.loads(texto.replace("{codigo}", codigo).replace("{clave}", clave))
    payload.pop("_descripcion", None)
    return payload


def _copiar_modelo(root: Path, destino: Path) -> None:
    modelo = root / MODELO_RELATIVO
    if not modelo.is_file():
        raise PlantillaNoDisponible(
            f"Falta el modelo canónico del que copiar la plantilla: {modelo}"
        )
    try:
        destino.relative_to(root)
    except ValueError:
        raise PlantillaNoDisponible("La plantilla quedaría fuera del proyecto") from None
    destino.parent.mkdir(parents=True, exist_ok=True)
    if not destino.is_file():
        shutil.copy2(modelo, destino)
