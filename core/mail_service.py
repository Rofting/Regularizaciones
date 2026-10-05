"""Borradores EML y entregas auditables de cartas PDF por propietario."""

from __future__ import annotations

import hashlib
import os
import re
import sqlite3
import string
import uuid
from dataclasses import dataclass
from email.message import EmailMessage
from email.parser import BytesParser
from email.policy import default
from email.utils import make_msgid, parseaddr
from pathlib import Path
from typing import Protocol

from office_settings import load_office_settings


DEFAULT_SUBJECT = "Regularización {community} · {period}"
DEFAULT_BODY = ("Estimado/a {owner}:\n\nAdjuntamos su carta de regularización "
                "del período {period} de {community}.\n\nUn saludo.")


@dataclass(frozen=True)
class MailRunResult:
    id_mail_run: int
    output_path: Path
    draft_count: int
    skipped_count: int
    duplicate_count: int


class MailTransport(Protocol):
    def send(self, message: EmailMessage) -> None: ...


def _transport_error(error: BaseException) -> str:
    """Motivo comprensible sin copiar la respuesta del servidor (puede tener datos)."""
    import smtplib
    import ssl
    if isinstance(error, ValueError):
        return str(error)
    if isinstance(error, smtplib.SMTPAuthenticationError):
        return "El servidor rechazó el usuario o la contraseña SMTP"
    if isinstance(error, smtplib.SMTPRecipientsRefused):
        return "El servidor rechazó la dirección del destinatario"
    if isinstance(error, smtplib.SMTPSenderRefused):
        return "El servidor no permite enviar con el correo del despacho como remitente"
    if isinstance(error, ssl.SSLError):
        return "Falló la conexión segura (TLS) con el servidor de correo"
    if isinstance(error, (smtplib.SMTPConnectError, smtplib.SMTPServerDisconnected)):
        return "No se pudo conectar con el servidor de correo (servidor, puerto o red)"
    if isinstance(error, smtplib.SMTPException):  # hereda de OSError: va antes
        return "El servidor de correo rechazó el envío"
    if isinstance(error, (TimeoutError, ConnectionError, OSError)):
        return "No se pudo conectar con el servidor de correo (servidor, puerto o red)"
    return "Error de transporte SMTP"


def _sha(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as source:
        while chunk := source.read(65536):
            digest.update(chunk)
    return digest.hexdigest()


def _email(value: object) -> str | None:
    raw = str(value or "").strip()
    name, address = parseaddr(raw)
    if name or address != raw or not re.fullmatch(r"[^\s@<>]+@[^\s@<>]+\.[^\s@<>]+", raw):
        return None
    return raw


def _render(template: str, values: dict[str, str], *, subject: bool = False) -> str:
    try:
        for _literal, field, format_spec, conversion in string.Formatter().parse(template):
            if field is not None and (field not in values or format_spec or conversion):
                raise ValueError("Sólo se admiten {community}, {period} y {owner}")
        rendered = template.format_map(values).strip()
    except (KeyError, ValueError) as error:
        raise ValueError(f"Plantilla de correo no válida: {error}") from error
    if not rendered or (subject and ("\n" in rendered or "\r" in rendered)):
        raise ValueError("El asunto o cuerpo de correo no es válido")
    return rendered


def _safe_attachment(raw: object, run_directory: Path) -> Path:
    if not raw:
        raise ValueError("Falta el PDF de una carta")
    path = Path(str(raw)).resolve()
    if path.parent != run_directory.resolve() or path.suffix.lower() != ".pdf" or not path.is_file():
        raise ValueError("El PDF de una carta no está en el lote validado")
    return path


def prepare_mail_run(
    connection: sqlite3.Connection, id_case: int, output_root: Path,
    subject_template: str = DEFAULT_SUBJECT, body_template: str = DEFAULT_BODY,
) -> MailRunResult:
    """Prepara filas por propietario sin enviar nada ni omitir problemas."""
    connection.row_factory = sqlite3.Row
    run = connection.execute(
        """SELECT l.id_letter_run,l.output_path,c.estado,co.nombre AS community,
                  p.nombre AS period
           FROM letter_generation_runs l
           JOIN regularization_cases c ON c.id_case=l.id_case
           JOIN comunidades co ON co.id_comunidad=c.id_comunidad
           JOIN periodos p ON p.id_periodo=l.id_periodo
           WHERE l.id_case=? AND l.status='completed'
           ORDER BY l.id_letter_run DESC LIMIT 1""", (id_case,),
    ).fetchone()
    if run is None or run["estado"] not in ("deliveries_generated", "closed"):
        raise ValueError("Primero genera y valida todas las cartas del expediente")
    rows = connection.execute(
        """SELECT g.id_propietario,g.pdf_path,g.pdf_pages,p.nombre_propietario,p.email
           FROM generated_letters g JOIN propietarios p USING(id_propietario)
           WHERE g.id_letter_run=? AND g.status='generated'
           ORDER BY g.id_propietario""", (run["id_letter_run"],),
    ).fetchall()
    if not rows:
        raise ValueError("El lote no contiene cartas PDF listas")
    office = load_office_settings(connection)
    sender = _email(office.email) or ""
    entries = []
    seen: set[str] = set()
    for row in rows:
        pdf = _safe_attachment(row["pdf_path"], Path(run["output_path"]))
        if not row["pdf_pages"]:
            raise ValueError("La carta PDF no tiene páginas verificadas")
        recipient = _email(row["email"])
        values = {"community": run["community"], "period": run["period"],
                  "owner": row["nombre_propietario"]}
        subject = _render(subject_template, values, subject=True)
        body = _render(body_template, values)
        attachment_sha = _sha(pdf)
        fingerprint = hashlib.sha256(
            "\0".join((recipient.casefold() if recipient else "", sender, subject,
                       body, attachment_sha)).encode("utf-8")
        ).hexdigest()
        if recipient is None:
            status, error = "skipped", "Falta un correo válido del propietario"
        elif recipient.casefold() in seen:
            status, error = "duplicate", "Correo repetido en el lote; revisar destinatarios"
        else:
            status, error = "draft", None
            seen.add(recipient.casefold())
        entries.append((row, pdf, recipient, subject, body, attachment_sha,
                        fingerprint, status, error))
    input_sha = hashlib.sha256(repr([
        (int(row["id_propietario"]), recipient, attachment_sha, fingerprint, status)
        for row, _pdf, recipient, _subject, _body, attachment_sha, fingerprint, status, _error
        in entries
    ]).encode("utf-8")).hexdigest()
    existing = connection.execute(
        "SELECT id_mail_run,output_path FROM mail_runs WHERE id_letter_run=? AND input_sha256=?",
        (run["id_letter_run"], input_sha),
    ).fetchone()
    if existing:
        counts = connection.execute(
            "SELECT status,COUNT(*) FROM mail_deliveries WHERE id_mail_run=? GROUP BY status",
            (existing["id_mail_run"],),
        ).fetchall()
        by_status = {status: count for status, count in counts}
        return MailRunResult(existing["id_mail_run"], Path(existing["output_path"]),
                             by_status.get("draft", 0), by_status.get("skipped", 0),
                             by_status.get("duplicate", 0))
    output_root = Path(output_root).resolve()
    output_root.mkdir(parents=True, exist_ok=True)
    cursor = connection.execute(
        """INSERT INTO mail_runs(id_case,id_letter_run,input_sha256,subject,body,output_path,status)
           VALUES (?,?,?,?,?,?,'draft')""",
        (id_case, run["id_letter_run"], input_sha, subject_template, body_template, str(output_root)),
    )
    mail_id = int(cursor.lastrowid)
    destination = output_root / f"lote_{mail_id}"
    connection.execute("UPDATE mail_runs SET output_path=? WHERE id_mail_run=?",
                       (str(destination), mail_id))
    for row, pdf, recipient, _subject, _body, attachment_sha, fingerprint, status, error in entries:
        connection.execute(
            """INSERT INTO mail_deliveries
               (id_mail_run,id_propietario,recipient,attachment_sha256,fingerprint,status,error_message)
               VALUES (?,?,?,?,?,?,?)""",
            (mail_id, row["id_propietario"], recipient, attachment_sha, fingerprint,
             status, error),
        )
    connection.commit()
    return MailRunResult(mail_id, destination,
                         sum(item[-2] == "draft" for item in entries),
                         sum(item[-2] == "skipped" for item in entries),
                         sum(item[-2] == "duplicate" for item in entries))


def generate_eml_drafts(connection: sqlite3.Connection, id_mail_run: int) -> MailRunResult:
    """Guarda un MIME por destinatario válido con publicación atómica."""
    connection.row_factory = sqlite3.Row
    run = connection.execute("SELECT * FROM mail_runs WHERE id_mail_run=?", (id_mail_run,)).fetchone()
    if run is None:
        raise ValueError("No existe el lote de correo")
    folder = Path(run["output_path"])
    folder.mkdir(parents=True, exist_ok=True)
    letter = connection.execute("SELECT output_path FROM letter_generation_runs WHERE id_letter_run=?",
                                (run["id_letter_run"],)).fetchone()
    sender = _email(load_office_settings(connection).email)
    rows = connection.execute(
        """SELECT d.*,p.nombre_propietario,l.pdf_path,c.nombre AS community,per.nombre AS period
           FROM mail_deliveries d JOIN propietarios p USING(id_propietario)
           JOIN generated_letters l ON l.id_letter_run=? AND l.id_propietario=d.id_propietario
           JOIN mail_runs r ON r.id_mail_run=d.id_mail_run
           JOIN comunidades c ON c.id_comunidad=p.id_comunidad
           JOIN periodos per ON per.id_periodo=(SELECT id_periodo FROM letter_generation_runs WHERE id_letter_run=r.id_letter_run)
           WHERE d.id_mail_run=? AND d.status IN ('draft','failed') ORDER BY d.id_propietario""",
        (run["id_letter_run"], id_mail_run),
    ).fetchall()
    for row in rows:
        try:
            pdf = _safe_attachment(row["pdf_path"], Path(letter["output_path"]))
            if _sha(pdf) != row["attachment_sha256"]:
                raise ValueError("El PDF cambió desde la preparación del correo")
            values = {"community": row["community"], "period": row["period"],
                      "owner": row["nombre_propietario"]}
            message = EmailMessage()
            if sender:
                message["From"] = sender
            message["To"] = row["recipient"]
            message["Subject"] = _render(run["subject"], values, subject=True)
            message["Message-ID"] = make_msgid(domain="regularizaciones.local")
            message.set_content(_render(run["body"], values))
            message.add_attachment(pdf.read_bytes(), maintype="application", subtype="pdf",
                                   filename=pdf.name)
            target = folder / f"carta_{row['id_propietario']}.eml"
            temporary = folder / f".{target.name}.{uuid.uuid4().hex}.tmp"
            try:
                temporary.write_bytes(message.as_bytes())
                os.replace(temporary, target)
            finally:
                temporary.unlink(missing_ok=True)
            connection.execute(
                "UPDATE mail_deliveries SET status='draft',eml_path=?,message_id=?,error_message=NULL WHERE id_mail_delivery=?",
                (str(target), str(message["Message-ID"]), row["id_mail_delivery"]),
            )
        except (OSError, ValueError) as error:
            connection.execute(
                "UPDATE mail_deliveries SET status='failed',eml_path=NULL,error_message=? WHERE id_mail_delivery=?",
                (str(error), row["id_mail_delivery"]),
            )
    connection.commit()
    counts = {status: count for status, count in connection.execute(
        "SELECT status,COUNT(*) FROM mail_deliveries WHERE id_mail_run=? GROUP BY status", (id_mail_run,)
    )}
    return MailRunResult(id_mail_run, folder, counts.get("draft", 0),
                         counts.get("skipped", 0), counts.get("duplicate", 0))


def exclude_mail_delivery(connection: sqlite3.Connection, id_mail_run: int,
                          id_propietario: int, *, reason: str) -> None:
    """Registra la exclusión consciente de un borrador aún no enviado."""
    if not reason.strip():
        raise ValueError("Indica el motivo de exclusión")
    row = connection.execute(
        "SELECT status,eml_path FROM mail_deliveries WHERE id_mail_run=? AND id_propietario=?",
        (id_mail_run, id_propietario),
    ).fetchone()
    if row is None or row[0] != "draft":
        raise ValueError("Sólo se puede excluir un borrador pendiente")
    folder = connection.execute("SELECT output_path FROM mail_runs WHERE id_mail_run=?",
                                (id_mail_run,)).fetchone()[0]
    if row[1]:
        eml = Path(row[1])
        if eml.parent.resolve() == Path(folder).resolve():
            eml.unlink(missing_ok=True)
    connection.execute(
        "UPDATE mail_deliveries SET status='skipped',eml_path=NULL,error_message=? WHERE id_mail_run=? AND id_propietario=?",
        ("Excluido: " + reason.strip()[:250], id_mail_run, id_propietario),
    )
    connection.commit()


def send_mail_run(connection: sqlite3.Connection, id_mail_run: int,
                  transport: MailTransport, *, confirmed_by: str) -> tuple[int, int]:
    """Envía sólo borradores confirmados; un reintento omite huellas ya enviadas."""
    if not confirmed_by.strip():
        raise ValueError("Confirma explícitamente el lote antes de enviar")
    connection.row_factory = sqlite3.Row
    run = connection.execute("SELECT * FROM mail_runs WHERE id_mail_run=?", (id_mail_run,)).fetchone()
    if run is None:
        raise ValueError("No existe el lote de correo")
    connection.execute("UPDATE mail_runs SET confirmed_by=?,confirmed_at=datetime('now') WHERE id_mail_run=?",
                       (confirmed_by.strip(), id_mail_run))
    connection.commit()
    sender = _email(load_office_settings(connection).email)
    if not sender:
        raise ValueError("Configura un correo válido del despacho antes de enviar")
    rows = connection.execute(
        "SELECT * FROM mail_deliveries WHERE id_mail_run=? AND status IN ('draft','failed') ORDER BY id_propietario",
        (id_mail_run,),
    ).fetchall()
    sent = failed = 0
    for row in rows:
        previous = connection.execute(
            "SELECT 1 FROM mail_deliveries WHERE fingerprint=? AND status='sent' LIMIT 1",
            (row["fingerprint"],),
        ).fetchone()
        if previous:
            connection.execute("UPDATE mail_deliveries SET status='sent',sent_at=datetime('now'),error_message=NULL WHERE id_mail_delivery=?",
                               (row["id_mail_delivery"],))
            connection.commit()
            continue
        try:
            eml = Path(row["eml_path"] or "")
            if eml.parent.resolve() != Path(run["output_path"]).resolve() or not eml.is_file():
                raise ValueError("Falta el borrador EML")
            message = BytesParser(policy=default).parsebytes(eml.read_bytes())
            if (message["From"] != sender or message["To"] != row["recipient"]
                    or message["Message-ID"] != row["message_id"]):
                raise ValueError("Los encabezados del borrador han cambiado")
            attachments = list(message.iter_attachments())
            if len(attachments) != 1 or hashlib.sha256(
                attachments[0].get_payload(decode=True) or b""
            ).hexdigest() != row["attachment_sha256"]:
                raise ValueError("El adjunto del borrador ha cambiado")
            letter = connection.execute(
                "SELECT pdf_path FROM generated_letters WHERE id_letter_run=? AND id_propietario=?",
                (run["id_letter_run"], row["id_propietario"]),
            ).fetchone()
            if letter is None or _sha(Path(letter["pdf_path"])) != row["attachment_sha256"]:
                raise ValueError("El PDF cambió desde la preparación del correo")
            transport.send(message)
            connection.execute("UPDATE mail_deliveries SET status='sent',sent_at=datetime('now'),error_message=NULL WHERE id_mail_delivery=?",
                               (row["id_mail_delivery"],))
            sent += 1
        except Exception as error:
            # Las respuestas de servidores ajenos pueden contener datos sensibles.
            detail = _transport_error(error)
            connection.execute("UPDATE mail_deliveries SET status='failed',error_message=? WHERE id_mail_delivery=?",
                               (detail[:300], row["id_mail_delivery"]),
            )
            failed += 1
        connection.commit()
    pending = connection.execute(
        # Sin correo, duplicados y exclusiones son decisiones registradas, no envíos pendientes.
        "SELECT COUNT(*) FROM mail_deliveries WHERE id_mail_run=? AND status IN ('draft','failed')",
        (id_mail_run,),
    ).fetchone()[0]
    connection.execute("UPDATE mail_runs SET status=? WHERE id_mail_run=?",
                       ("partial" if pending else "sent", id_mail_run))
    connection.commit()
    return sent, failed
