"""Copias SQLite automáticas verificadas y retención de copias propias."""

from __future__ import annotations

import hashlib
import json
import sqlite3
import time
from contextlib import closing
from dataclasses import dataclass
from datetime import UTC, datetime
from pathlib import Path
from uuid import uuid4


class DatabaseBackupError(RuntimeError):
    """No se pudo proteger la base antes de modificarla."""


@dataclass(frozen=True)
class BackupResult:
    backup_path: Path
    removed_paths: tuple[Path, ...] = ()
    cleanup_warnings: tuple[str, ...] = ()


def _retention(database_path: Path) -> int:
    settings = database_path.parent / "backup_settings.json"
    if not settings.exists():
        return 20
    payload = json.loads(settings.read_text(encoding="utf-8"))
    if not isinstance(payload, dict):
        raise ValueError("backup_settings.json debe contener un objeto JSON")
    count = payload.get("max_backups", 20)
    if isinstance(count, bool) or not isinstance(count, int) or count < 3:
        raise ValueError("max_backups debe ser un entero de al menos 3")
    return count


def _digest(path: Path) -> str:
    with path.open("rb") as stream:
        return hashlib.file_digest(stream, "sha256").hexdigest()


def verify_backup(path: Path) -> None:
    """Comprueba integridad y presencia de tablas sin migrar ni escribir."""
    with closing(sqlite3.connect(path.resolve().as_uri() + "?mode=ro", uri=True)) as con:
        if con.execute("PRAGMA integrity_check").fetchall() != [("ok",)]:
            raise DatabaseBackupError(f"La copia no supera la comprobación de integridad: {path}")
        if not con.execute(
            "SELECT 1 FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%'"
        ).fetchone():
            raise DatabaseBackupError(f"La copia no contiene tablas: {path}")


def _prune(root: Path, source: Path, prefix: str, keep: int) -> tuple[tuple[Path, ...], tuple[str, ...]]:
    verified = []
    warnings = []
    for manifest in root.glob(prefix + "*.db.json"):
        path = manifest.with_suffix("")
        try:
            if path.is_symlink() or manifest.is_symlink():
                continue
            record = json.loads(manifest.read_text(encoding="utf-8"))
            if not isinstance(record, dict):
                continue
            if record.get("source") != str(source) or record.get("file") != path.name:
                continue
            if record.get("sha256") != _digest(path):
                warnings.append(f"Se conserva una copia alterada: {path.name}")
                continue
            verify_backup(path)
            verified.append((path, manifest))
        except (OSError, ValueError, sqlite3.Error, DatabaseBackupError) as error:
            warnings.append(f"Se conserva {path.name}: {error}")
    verified.sort(key=lambda item: item[0].name, reverse=True)
    removed = []
    for path, manifest in verified[keep:]:
        try:
            path.unlink()
            removed.append(path)
            manifest.unlink()
        except OSError as error:
            warnings.append(f"No se pudo limpiar {path.name}: {error}")
    return tuple(removed), tuple(warnings)


def backup_database(database_path: str | Path, *, reason: str) -> BackupResult:
    """Respalda datos confirmados, incluido WAL, antes de iniciar una operación.

    Cada copia tiene integridad comprobada y un manifiesto con su SHA-256.
    La limpieza afecta sólo a copias automáticas de esta misma base que siguen
    verificadas; las copias manuales o del reinicio seguro no se eliminan.
    """
    source = Path(database_path).resolve()
    destination = None
    manifest = None
    manifest_created = False
    try:
        if not source.is_file():
            raise DatabaseBackupError(f"No existe la base que se quiere respaldar: {source}")
        keep = _retention(source)
        root = source.parent / "backups"
        root.mkdir(parents=True, exist_ok=True)
        identity = hashlib.sha256(str(source).encode()).hexdigest()[:12]
        prefix = f"automatic-{identity}-"
        timestamp = datetime.now(UTC).strftime("%Y%m%dT%H%M%S%fZ")
        while destination is None:
            candidate = root / f"{prefix}{timestamp}-{uuid4().hex}.db"
            try:
                with candidate.open("xb"):
                    pass
            except FileExistsError:
                continue
            destination = candidate
        started = time.monotonic()

        def check_timeout(status, remaining, total):
            if time.monotonic() - started > 30:
                raise DatabaseBackupError("La base sigue ocupada; no se pudo crear la copia en 30 segundos")

        with closing(sqlite3.connect(source.as_uri() + "?mode=ro", uri=True)) as original:
            with closing(sqlite3.connect(destination)) as copied:
                original.backup(copied, pages=128, progress=check_timeout)
        verify_backup(destination)
        manifest = destination.with_suffix(".db.json")
        record = {
            "file": destination.name, "source": str(source), "reason": reason,
            "created_at": datetime.now(UTC).isoformat(), "sha256": _digest(destination),
        }
        with manifest.open("x", encoding="utf-8") as stream:
            manifest_created = True
            json.dump(record, stream, ensure_ascii=False, indent=2)
    except (OSError, ValueError, sqlite3.Error, DatabaseBackupError) as error:
        for path in (manifest if manifest_created else None, destination):
            if path is not None:
                try:
                    path.unlink(missing_ok=True)
                except OSError:
                    pass
        raise DatabaseBackupError(
            f"No se pudo crear una copia verificada. La operación se ha detenido: {error}"
        ) from error
    removed, warnings = _prune(root, source, prefix, keep)
    return BackupResult(destination, removed, warnings)


def backup_connection(connection: sqlite3.Connection, *, reason: str) -> BackupResult | None:
    """Las conexiones en memoria no tienen datos persistentes que respaldar."""
    path = next((row[2] for row in connection.execute("PRAGMA database_list") if row[1] == "main"), "")
    return backup_database(path, reason=reason) if path else None
