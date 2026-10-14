"""Archivo de correos y trazabilidad de sus adjuntos en el flujo de fuentes."""
from dataclasses import dataclass
import hashlib
import json
import os
from pathlib import Path
import tempfile

import case_ingestion
import community_discovery
import expedient_service
from email_sources import read_email_sources


@dataclass(frozen=True)
class EmailIngestionResult:
    email_document_id: int
    results: tuple[case_ingestion.IngestionResult, ...]
    notes: tuple[str, ...]
    errors: tuple[tuple[str, str], ...]
    foreign_count: int


def email_provenance(connection, document_id) -> tuple[str, ...]:
    rows = connection.execute("""SELECT DISTINCT d.original_name, e.subject, e.sender
        FROM email_attachments a JOIN email_imports e ON e.id_email_document=a.id_email_document
        JOIN source_documents d ON d.id_document=e.id_email_document
        WHERE a.id_document=? ORDER BY d.id_document""", (document_id,)).fetchall()
    return tuple(f"Correo de origen: {row['original_name']} · {row['subject']} · {row['sender']}" for row in rows)


def _archive_attachment(path: Path, payload: bytes, sha256: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    if path.exists():
        if not path.is_file() or hashlib.sha256(path.read_bytes()).hexdigest() != sha256:
            raise ValueError(f"El adjunto archivado ha cambiado: {path.name}")
        return
    temporary = None
    try:
        with tempfile.NamedTemporaryFile(dir=path.parent, delete=False) as stream:
            temporary = Path(stream.name)
            stream.write(payload)
        # Publicación atómica, sin dejar un adjunto parcial ni sobreescribir otro.
        try:
            os.link(temporary, path)
        except FileExistsError:
            if not path.is_file() or hashlib.sha256(path.read_bytes()).hexdigest() != sha256:
                raise ValueError(f"El adjunto archivado ha cambiado: {path.name}")
    finally:
        if temporary is not None:
            temporary.unlink(missing_ok=True)


def import_email_to_case(connection, case_id, *, source_path, archive_root, analyser=None):
    document, _ = expedient_service.register_source_document(
        connection, case_id, source_path=source_path, archive_root=archive_root,
        document_kind="email",
    )
    if document.document_kind != "email":
        raise ValueError("La fuente ya está registrada como otro tipo de documento")
    with connection:
        connection.execute("UPDATE source_documents SET status='not_applicable' WHERE id_document=?",
                           (document.id_document,))
        connection.execute("INSERT OR IGNORE INTO email_imports(id_email_document) VALUES (?)",
                           (document.id_document,))
    results, notes, errors = [], [], []
    foreign_count = 0
    try:
        if hashlib.sha256(document.archived_path.read_bytes()).hexdigest() != document.sha256:
            raise ValueError("El correo archivado ha cambiado")
        message = read_email_sources(document.archived_path)
    except Exception as error:
        errors.append((document.original_name, str(error)))
        with connection:
            connection.execute("UPDATE email_imports SET notes_json=? WHERE id_email_document=?",
                               (json.dumps([str(error)], ensure_ascii=False), document.id_document))
        return EmailIngestionResult(document.id_document, (), (), tuple(errors), 0)
    notes.extend(message.notes)
    with connection:
        connection.execute("UPDATE email_imports SET subject=?, sender=?, message_date=? WHERE id_email_document=?",
                           (message.subject, message.sender, message.message_date, document.id_document))
    case = connection.execute("""SELECT c.codigo FROM comunidades c JOIN regularization_cases r
        ON c.id_comunidad=r.id_comunidad WHERE r.id_case=?""", (case_id,)).fetchone()
    active_analyser = analyser or case_ingestion._case_analyser(connection, case_id)
    for attachment in message.attachments:
        digest = hashlib.sha256(attachment.data).hexdigest()
        path = (Path(archive_root) / str(case_id) / "correos" / document.sha256
                / str(attachment.index) / attachment.filename)
        child_id = None
        outcome = "error"
        try:
            _archive_attachment(path, attachment.data, digest)
            _, foreign = community_discovery.partition_sources_for_community(
                (path,), str(case["codigo"]), connection=connection,
            )
            connection.commit()
            if foreign:
                foreign_count += 1
                outcome = "foreign_community"
                notes.append(f"{attachment.original_name}: indica otra comunidad; se conserva en el correo y sus adjuntos")
            else:
                # Repetir el correo enlaza el documento existente sin reanalizar
                # ni invalidar las decisiones ya confirmadas por el usuario.
                existing = connection.execute("SELECT id_document FROM source_documents WHERE id_case=? AND sha256=?",
                                              (case_id, digest)).fetchone()
                if existing is not None:
                    child = case_ingestion._source_document(connection, existing["id_document"])
                    if not child.archived_path.is_file() or hashlib.sha256(child.archived_path.read_bytes()).hexdigest() != digest:
                        raise ValueError("El documento adjunto ya incorporado ha cambiado o no está disponible")
                    result = case_ingestion.IngestionResult(child, False,
                        case_ingestion._open_issue_count(connection, case_id))
                else:
                    result = case_ingestion.add_analysed_document_to_case(
                        connection, case_id, source_path=path, archive_root=archive_root,
                        analysis=active_analyser(path),
                    )
                child_id = result.document.id_document
                results.append(result)
                outcome = "imported"
        except Exception as error:
            connection.rollback()
            errors.append((attachment.original_name, str(error)))
        with connection:
            connection.execute("""INSERT INTO email_attachments
                (id_email_document, attachment_index, original_name, archived_path, sha256, id_document, outcome)
                VALUES (?,?,?,?,?,?,?) ON CONFLICT(id_email_document,attachment_index) DO UPDATE SET
                id_document=COALESCE(excluded.id_document,email_attachments.id_document), outcome=excluded.outcome""",
                (document.id_document, attachment.index, attachment.original_name, str(path), digest, child_id, outcome))
    with connection:
        connection.execute("UPDATE email_imports SET notes_json=? WHERE id_email_document=?",
                           (json.dumps(notes + [f"{name}: {error}" for name, error in errors], ensure_ascii=False),
                            document.id_document))
    return EmailIngestionResult(document.id_document, tuple(results), tuple(notes), tuple(errors), foreign_count)
