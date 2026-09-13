"""Reinicio seguro de bases de datos SQLite con copia de seguridad verificada."""

from __future__ import annotations

import os
import sqlite3
import tempfile
from collections.abc import Callable
from contextlib import closing
from dataclasses import dataclass
from datetime import UTC, datetime
from pathlib import Path
from uuid import uuid4


class DatabaseResetError(RuntimeError):
    """El reinicio no se completó y la base de datos original se conserva."""


@dataclass(frozen=True)
class ResetResult:
    backup_path: Path
    new_path: Path


def reset_database(
    database_path: Path,
    *,
    backup_root: Path,
    initialise: Callable[[Path], None],
) -> ResetResult:
    """Sustituye una base SQLite tras respaldarla y validar ambas copias."""
    database_path = Path(database_path)
    backup_root = Path(backup_root)

    if not database_path.is_file():
        raise DatabaseResetError(f"La base de datos no existe: {database_path}")

    try:
        backup_root.mkdir(parents=True, exist_ok=True)
        backup_path = _backup_path(database_path, backup_root)
        _copy_database(database_path, backup_path)
        _verify_database(backup_path)
    except DatabaseResetError:
        raise
    except (OSError, sqlite3.Error) as error:
        raise DatabaseResetError(
            f"No se pudo crear o verificar la copia de seguridad de {database_path}: {error}"
        ) from error

    replacement_path = _temporary_database_path(database_path)
    try:
        initialise(replacement_path)
        _verify_database(replacement_path)
        os.replace(replacement_path, database_path)
    except DatabaseResetError:
        _remove_temporary_database(replacement_path)
        raise
    except Exception as error:
        _remove_temporary_database(replacement_path)
        raise DatabaseResetError(
            f"No se pudo inicializar o validar la nueva base de datos: {error}"
        ) from error

    return ResetResult(backup_path=backup_path, new_path=database_path)


def _backup_path(database_path: Path, backup_root: Path) -> Path:
    timestamp = datetime.now(UTC).strftime("%Y%m%dT%H%M%S%fZ")
    suffix = database_path.suffix or ".db"
    return backup_root / f"{database_path.stem}-{timestamp}-{uuid4().hex}{suffix}"


def _copy_database(source_path: Path, destination_path: Path) -> None:
    source_uri = source_path.resolve().as_uri() + "?mode=ro"
    with closing(sqlite3.connect(source_uri, uri=True)) as source:
        with closing(sqlite3.connect(destination_path)) as destination:
            source.backup(destination)


def _verify_database(database_path: Path) -> None:
    database_uri = database_path.resolve().as_uri() + "?mode=ro"
    with closing(sqlite3.connect(database_uri, uri=True)) as connection:
        result = connection.execute("PRAGMA integrity_check").fetchall()
    if result != [("ok",)]:
        raise DatabaseResetError(
            f"La comprobación de integridad falló para {database_path}: {result}"
        )


def _temporary_database_path(database_path: Path) -> Path:
    descriptor, temporary_name = tempfile.mkstemp(
        dir=database_path.parent,
        prefix=f".{database_path.stem}-replacement-",
        suffix=database_path.suffix or ".db",
    )
    os.close(descriptor)
    return Path(temporary_name)


def _remove_temporary_database(path: Path) -> None:
    try:
        path.unlink(missing_ok=True)
    except OSError:
        pass
