"""Rutas del producto empaquetado y hogar escribible del despacho."""

from __future__ import annotations

import os
import shutil
import sqlite3
import sys
from contextlib import closing
from dataclasses import dataclass
from pathlib import Path
from typing import Mapping
from uuid import uuid4


PUBLIC_RESOURCES = (
    "config/proveedores.json",
    "config/palabras_clave.json",
    "config/proveedores_despacho.ejemplo.json",
    "config/rutas.ejemplo.json",
    "config/modelo_excel",
    "plantillas/Plantilla_Cartas.docx",
    "plantillas/modelo",
)
LEGACY_DATA = (
    "config/excel_profiles/runtime",
    "config/letter_identities.json",
    "config/proveedores_despacho.json",
    "plantillas/comunidades",
    "data/expedientes",
    "salidas",
)


@dataclass(frozen=True)
class ApplicationPaths:
    resources: Path
    home: Path
    executable: Path

    @classmethod
    def resolve(cls, environ: Mapping[str, str] | None = None,
                executable: Path | None = None, local_app_data: Path | None = None,
                resources: Path | None = None, frozen: bool | None = None) -> "ApplicationPaths":
        values = os.environ if environ is None else environ
        is_frozen = bool(getattr(sys, "frozen", False)) if frozen is None else frozen
        binary = Path(executable or sys.executable).resolve()
        source = Path(resources or getattr(sys, "_MEIPASS", Path(__file__).resolve().parents[1])).resolve()
        override = values.get("REGULARIZACIONES_HOME", "").strip()
        if override:
            home = Path(override).expanduser().resolve()
        elif is_frozen:
            local = local_app_data or values.get("LOCALAPPDATA")
            if not local:
                raise RuntimeError("No se encuentra LOCALAPPDATA; define REGULARIZACIONES_HOME")
            home = (Path(local) / "Regularizaciones").resolve()
        else:
            home = source
        return cls(source, home, binary)

    @property
    def database(self) -> Path:
        return self.home / "data" / "gestion.db"

    @property
    def sources(self) -> Path:
        return self.home / "data" / "expedientes"

    @property
    def outputs(self) -> Path:
        return self.home / "salidas"

    @property
    def backups(self) -> Path:
        return self.home / "data" / "backups"

    @property
    def logs(self) -> Path:
        return self.home / "logs"

    def prepare(self) -> None:
        """Comprueba escritura y siembra sólo recursos públicos que aún faltan."""
        try:
            self.home.mkdir(parents=True, exist_ok=True)
            probe = self.home / f".regularizaciones-write-test-{uuid4().hex}"
            with probe.open("x", encoding="utf-8") as stream:
                stream.write("ok")
            probe.unlink()
        except OSError as error:
            raise RuntimeError(
                f"No se puede escribir en {self.home}. Elige otra carpeta con REGULARIZACIONES_HOME: {error}"
            ) from error
        for folder in (self.database.parent, self.sources, self.outputs, self.backups, self.logs):
            folder.mkdir(parents=True, exist_ok=True)
        if self.home == self.resources:
            return
        for relative in PUBLIC_RESOURCES:
            source = self.resources / relative
            if not source.exists():
                raise RuntimeError(f"Falta un recurso del programa: {relative}")
            targets = (source.rglob("*") if source.is_dir() else (source,))
            for item in targets:
                if not item.is_file():
                    continue
                target = self.home / item.relative_to(self.resources)
                target.parent.mkdir(parents=True, exist_ok=True)
                temporary = target.with_name(f".{target.name}.{uuid4().hex}.tmp")
                try:
                    shutil.copy2(item, temporary)
                    os.replace(temporary, target)
                finally:
                    temporary.unlink(missing_ok=True)
        self._import_legacy_installation()

    def _import_legacy_installation(self) -> None:
        """Migra una instalación fuente contigua al EXE sin sobrescribir datos."""
        legacy = self.executable.parent
        old_database = legacy / "data" / "gestion.db"
        if self.database.exists() or not old_database.is_file() or legacy == self.home:
            return
        temporary = self.database.with_suffix(".importando")
        try:
            with closing(sqlite3.connect(old_database.resolve().as_uri() + "?mode=ro", uri=True)) as source, \
                 closing(sqlite3.connect(temporary)) as target:
                source.backup(target)
                if target.execute("PRAGMA integrity_check").fetchone()[0] != "ok":
                    raise RuntimeError("La base antigua no supera la comprobación de integridad")
            os.replace(temporary, self.database)
            for relative in LEGACY_DATA:
                source = legacy / relative
                target = self.home / relative
                if source.is_dir() and not target.exists():
                    shutil.copytree(source, target)
                elif source.is_file() and not target.exists():
                    target.parent.mkdir(parents=True, exist_ok=True)
                    shutil.copy2(source, target)
        finally:
            temporary.unlink(missing_ok=True)
