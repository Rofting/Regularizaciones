"""Lectura de correos guardados, sin ejecutar ni abrir sus adjuntos."""
from dataclasses import dataclass
from email import policy
from email.parser import BytesParser
from pathlib import Path
import re


EMAIL_SUFFIXES = frozenset({".eml", ".msg"})
DOCUMENT_SUFFIXES = frozenset({".pdf", ".xlsx", ".xls", ".csv"})
MAX_EMAIL_BYTES = 50 * 1024 * 1024
MAX_ATTACHMENT_BYTES = 25 * 1024 * 1024
MAX_ATTACHMENTS = 100


@dataclass(frozen=True)
class EmailAttachment:
    index: int
    original_name: str
    filename: str
    data: bytes


@dataclass(frozen=True)
class EmailSources:
    subject: str
    sender: str
    message_date: str
    attachments: tuple[EmailAttachment, ...]
    notes: tuple[str, ...]


def safe_attachment_name(name: str) -> str:
    # Los nombres de Outlook pueden usar barras Windows también en Linux.
    basename = str(name).replace("\\", "/").rsplit("/", 1)[-1]
    basename = re.sub(r'[\x00-\x1f<>:"|?*]', "_", basename).strip(" .")
    suffix = Path(basename).suffix
    stem = Path(basename).stem[:160].strip(" .") or "adjunto"
    if stem.upper().split(".")[0] in {"CON", "PRN", "AUX", "NUL", *[f"COM{i}" for i in range(1, 10)], *[f"LPT{i}" for i in range(1, 10)]}:
        stem = "_" + stem
    return stem + suffix


def read_email_sources(path: Path) -> EmailSources:
    path = Path(path)
    if path.suffix.lower() not in EMAIL_SUFFIXES:
        raise ValueError("Selecciona un correo EML o MSG")
    if path.stat().st_size > MAX_EMAIL_BYTES:
        raise ValueError("El correo supera el límite de 50 MB")
    attachments = []
    notes = []
    total_bytes = 0

    def accept(index, name, data):
        nonlocal total_bytes
        filename = safe_attachment_name(name)
        if Path(filename).suffix.lower() not in DOCUMENT_SUFFIXES:
            notes.append(f"Adjunto omitido por formato: {name}")
            return
        if not isinstance(data, bytes) or not data:
            notes.append(f"Adjunto vacío o no legible: {name}")
            return
        if len(data) > MAX_ATTACHMENT_BYTES or total_bytes + len(data) > MAX_EMAIL_BYTES:
            notes.append(f"Adjunto omitido por tamaño: {name}")
            return
        total_bytes += len(data)
        attachments.append(EmailAttachment(index, str(name), filename, data))

    if path.suffix.lower() == ".eml":
        message = BytesParser(policy=policy.default).parsebytes(path.read_bytes())
        subject = str(message.get("Subject", ""))
        sender = str(message.get("From", ""))
        message_date = str(message.get("Date", ""))
        index = 0
        for part in message.walk():
            if part.is_multipart():
                continue
            name = part.get_filename()
            if not name:
                if part.get_content_disposition() != "attachment":
                    continue
                notes.append("Adjunto sin nombre: no se puede identificar su formato")
                continue
            index += 1
            if index > MAX_ATTACHMENTS:
                notes.append("Se ha alcanzado el límite de 100 adjuntos")
                break
            try:
                payload = part.get_payload(decode=True)
                if part.defects:
                    notes.append(f"Adjunto con codificación dañada: {name}")
                    continue
                accept(index, name, payload)
            except (ValueError, TypeError) as error:
                notes.append(f"Adjunto no legible: {name}: {error}")
    else:
        import extract_msg
        with extract_msg.openMsg(str(path)) as message:
            subject = str(message.subject or "")
            sender = str(message.sender or "")
            message_date = str(message.date or "")
            for index, attachment in enumerate(message.attachments, 1):
                if index > MAX_ATTACHMENTS:
                    notes.append("Se ha alcanzado el límite de 100 adjuntos")
                    break
                try:
                    name = attachment.longFilename or attachment.shortFilename or "adjunto_sin_nombre"
                    accept(index, name, attachment.data)
                except Exception as error:
                    notes.append(f"Adjunto {index} no legible: {error}")
    if not attachments:
        notes.append("El correo no contiene adjuntos compatibles (PDF, Excel o CSV)")
    return EmailSources(subject, sender, message_date, tuple(attachments), tuple(notes))
