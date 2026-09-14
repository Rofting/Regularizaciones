import hashlib
import os
import re
import shutil
import sqlite3
import tempfile
from contextlib import contextmanager
from datetime import date
from pathlib import Path
from typing import Iterator

from expedient_models import (
    CaseStatus,
    RegularizationCase,
    SourceDocument,
    case_from_row,
    document_from_row,
)


ALLOWED_TRANSITIONS = {
    "draft": {"gathering_sources"},
    "gathering_sources": {"under_review"},
    "under_review": {"gathering_sources", "ready_for_calculation"},
    "ready_for_calculation": {"under_review", "calculated"},
    "calculated": {"reconciled", "under_review"},
    "reconciled": {"deliveries_generated", "under_review"},
    "deliveries_generated": {"closed", "under_review"},
    "closed": {"under_review"},
}


@contextmanager
def _transaction(connection: sqlite3.Connection) -> Iterator[None]:
    """Aísla esta operación sin confirmar una transacción que pertenezca al caller."""
    if connection.in_transaction:
        connection.execute("SAVEPOINT expedient_service")
        try:
            yield
        except Exception:
            connection.execute("ROLLBACK TO SAVEPOINT expedient_service")
            connection.execute("RELEASE SAVEPOINT expedient_service")
            raise
        else:
            connection.execute("RELEASE SAVEPOINT expedient_service")
        return

    connection.execute("BEGIN")
    try:
        yield
    except Exception:
        connection.rollback()
        raise
    else:
        connection.commit()


@contextmanager
def _registration_transaction(connection: sqlite3.Connection) -> Iterator[None]:
    """Inicia una escritura exclusiva para publicar una fuente y su fila juntas."""
    connection.execute("BEGIN IMMEDIATE")
    try:
        yield
    except Exception:
        connection.rollback()
        raise
    else:
        connection.commit()


def _case_row(connection: sqlite3.Connection, case_id: int):
    return connection.execute(
        """SELECT id_case, id_comunidad, nombre, fecha_inicio, fecha_fin, estado,
                  id_periodo
           FROM regularization_cases WHERE id_case = ?""",
        (case_id,),
    ).fetchone()


def _document_row(connection: sqlite3.Connection, document_id: int):
    return connection.execute(
        """SELECT id_document, id_case, original_name, archived_path, sha256,
                  document_kind, status
           FROM source_documents WHERE id_document = ?""",
        (document_id,),
    ).fetchone()


def _validate_transition(current: str, target: str) -> None:
    allowed = ALLOWED_TRANSITIONS.get(current, set())
    if target not in allowed:
        next_states = ", ".join(sorted(allowed)) or "ninguno"
        raise ValueError(
            f"No se puede cambiar el expediente de {current} a {target}; "
            f"siguientes estados permitidos: {next_states}"
        )


def _update_case_status(connection: sqlite3.Connection, case_id: int,
                        current: str, target: str) -> RegularizationCase:
    _validate_transition(current, target)
    connection.execute(
        """UPDATE regularization_cases
           SET estado = ?, updated_at = datetime('now')
           WHERE id_case = ?""",
        (target, case_id),
    )
    row = _case_row(connection, case_id)
    return case_from_row(row)


def create_case(connection: sqlite3.Connection, community_id: int, *, name: str,
                start_date: date, end_date: date) -> RegularizationCase:
    if end_date < start_date:
        raise ValueError(
            "La fecha de fin debe ser posterior o igual a la fecha de inicio "
            "(fin posterior requerido)"
        )

    with _transaction(connection):
        cursor = connection.execute(
            """INSERT INTO regularization_cases
               (id_comunidad, nombre, fecha_inicio, fecha_fin, estado)
               VALUES (?, ?, ?, ?, 'draft')""",
            (community_id, name, start_date.isoformat(), end_date.isoformat()),
        )
        row = _case_row(connection, cursor.lastrowid)
    return case_from_row(row)


def get_case(connection: sqlite3.Connection, case_id: int) -> RegularizationCase:
    row = _case_row(connection, case_id)
    if row is None:
        raise LookupError("Expediente no encontrado")
    return case_from_row(row)


def list_cases(connection: sqlite3.Connection, community_id: int) -> tuple[RegularizationCase, ...]:
    rows = connection.execute(
        """SELECT id_case, id_comunidad, nombre, fecha_inicio, fecha_fin, estado,
                  id_periodo
           FROM regularization_cases
           WHERE id_comunidad = ?
           ORDER BY fecha_inicio DESC, id_case DESC""",
        (community_id,),
    ).fetchall()
    return tuple(case_from_row(row) for row in rows)


def link_case_to_period(connection: sqlite3.Connection, id_case: int) -> int:
    """Enlaza un expediente a un periodo con la misma comunidad y fechas."""
    with _transaction(connection):
        case = get_case(connection, id_case)
        start = case.start_date.isoformat()
        end = case.end_date.isoformat()

        if case.period_id is not None:
            linked = connection.execute(
                """SELECT id_comunidad, fecha_inicio, fecha_fin
                   FROM periodos WHERE id_periodo = ?""",
                (case.period_id,),
            ).fetchone()
            if linked is None:
                raise ValueError("El periodo enlazado ya no existe")
            if (
                linked["id_comunidad"] != case.community_id
                or linked["fecha_inicio"] != start
                or linked["fecha_fin"] != end
            ):
                raise ValueError(
                    "El periodo enlazado pertenece a otra comunidad o tiene fechas incompatibles"
                )
            return case.period_id

        same_name = connection.execute(
            """SELECT id_periodo, fecha_inicio, fecha_fin
               FROM periodos WHERE id_comunidad = ? AND nombre = ?""",
            (case.community_id, case.name),
        ).fetchone()
        if same_name is not None and (
            same_name["fecha_inicio"] != start or same_name["fecha_fin"] != end
        ):
            raise ValueError(
                f"El periodo {case.name!r} ya existe con fechas incompatibles"
            )

        period = same_name or connection.execute(
            """SELECT id_periodo, fecha_inicio, fecha_fin
               FROM periodos
               WHERE id_comunidad = ? AND fecha_inicio = ? AND fecha_fin = ?
               ORDER BY id_periodo LIMIT 1""",
            (case.community_id, start, end),
        ).fetchone()
        if period is None:
            cursor = connection.execute(
                """INSERT INTO periodos
                   (id_comunidad, nombre, fecha_inicio, fecha_fin, estado)
                   VALUES (?, ?, ?, ?, 'abierto')""",
                (case.community_id, case.name, start, end),
            )
            period_id = int(cursor.lastrowid)
        else:
            period_id = int(period["id_periodo"])

        occupied = connection.execute(
            """SELECT id_case FROM regularization_cases
               WHERE id_comunidad = ? AND id_periodo = ? AND id_case <> ?""",
            (case.community_id, period_id, id_case),
        ).fetchone()
        if occupied is not None:
            raise ValueError("El periodo ya está vinculado a otro expediente")

        connection.execute(
            """UPDATE regularization_cases
               SET id_periodo = ?, updated_at = datetime('now')
               WHERE id_case = ?""",
            (period_id, id_case),
        )
    return period_id


def set_case_status(connection: sqlite3.Connection, case_id: int,
                    status: CaseStatus) -> RegularizationCase:
    with _transaction(connection):
        current = get_case(connection, case_id)
        updated = _update_case_status(connection, case_id, current.status, status)
    return updated


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as source:
        while chunk := source.read(64 * 1024):
            digest.update(chunk)
    return digest.hexdigest()


def _safe_filename(name: str) -> str:
    safe_name = re.sub(r"[^\w.-]", "_", name, flags=re.UNICODE).strip(".")
    return safe_name or "documento"


def register_source_document(connection: sqlite3.Connection, case_id: int, *,
                             source_path: str | Path, archive_root: str | Path,
                             document_kind: str) -> tuple[SourceDocument, bool]:
    if connection.in_transaction:
        raise RuntimeError(
            "No registre fuentes dentro de una transacción externa; "
            "registre el documento fuera de esa transacción"
        )

    source = Path(source_path)
    if not source.is_file():
        raise FileNotFoundError(f"El archivo fuente no existe o no es un archivo: {source}")

    sha256 = _sha256(source)
    destination = Path(archive_root) / str(case_id) / "fuentes" / (
        f"{sha256[:12]}_{_safe_filename(source.name)}"
    )
    staging_path: Path | None = None
    final_created = False

    try:
        with _registration_transaction(connection):
            existing = connection.execute(
                """SELECT id_document, id_case, original_name, archived_path, sha256,
                          document_kind, status
                   FROM source_documents WHERE id_case = ? AND sha256 = ?""",
                (case_id, sha256),
            ).fetchone()
            if existing is not None:
                return document_from_row(existing), False

            case = get_case(connection, case_id)
            destination.parent.mkdir(parents=True, exist_ok=True)
            if destination.exists():
                if not destination.is_file() or _sha256(destination) != sha256:
                    raise FileExistsError(f"Ya existe un archivo archivado distinto en {destination}")
            else:
                with tempfile.NamedTemporaryFile(
                    dir=destination.parent, prefix=f".{destination.name}.",
                    suffix=".tmp", delete=False,
                ) as staging_file:
                    staging_path = Path(staging_file.name)
                shutil.copy2(source, staging_path)
                destination_descriptor = os.open(
                    destination, os.O_WRONLY | os.O_CREAT | os.O_EXCL
                )
                final_created = True
                with os.fdopen(destination_descriptor, "wb") as destination_file:
                    with staging_path.open("rb") as staging_file:
                        shutil.copyfileobj(staging_file, destination_file)
                staging_path.unlink()
                staging_path = None
            cursor = connection.execute(
                """INSERT INTO source_documents
                   (id_case, original_name, archived_path, sha256, document_kind, status)
                   VALUES (?, ?, ?, ?, ?, 'registered')""",
                (case_id, source.name, str(destination), sha256, document_kind),
            )
            if case.status == "draft":
                _update_case_status(
                    connection, case_id, case.status, "gathering_sources"
                )
            elif case.status in {
                "ready_for_calculation",
                "calculated",
                "reconciled",
                "deliveries_generated",
                "closed",
            }:
                _update_case_status(
                    connection, case_id, case.status, "under_review"
                )
            document = document_from_row(_document_row(connection, cursor.lastrowid))
    except Exception:
        if staging_path is not None and staging_path.exists():
            staging_path.unlink()
        if final_created and destination.exists():
            destination.unlink()
        raise
    return document, True
