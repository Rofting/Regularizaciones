"""Correo de cartas: borradores .eml auditables y envío opcional sin duplicados.

Flujo:
1. ``plan_mail``: qué propietario recibiría qué carta y quién no tiene correo
   (sin escribir nada).
2. ``prepare_mail_run``: registra un lote con una entrega por propietario. Los
   que no tienen correo válido o se excluyen quedan como ``skipped`` con su
   motivo; nunca bloquean a los demás.
3. ``generate_eml_drafts``: escribe un .eml por entrega. Outlook o Thunderbird
   lo abren como borrador listo para revisar y enviar (cabecera X-Unsent).
4. ``send_mail_run``: envío SMTP opcional con confirmación explícita. Lo ya
   enviado (misma huella) no se vuelve a enviar, y un adjunto que cambió tras
   preparar el lote se rechaza.

La huella de cada entrega es SHA-256 de destinatario, carta, asunto y cuerpo.
Los cuerpos no se guardan en la base ni en registros.
"""

from __future__ import annotations

import hashlib
import mimetypes
import os
import re
import sqlite3
import string
import tempfile
from dataclasses import dataclass
from email.message import EmailMessage
from email.utils import formatdate, make_msgid
from pathlib import Path
from typing import Iterable

from mail_transport import MailTransportError

DEFAULT_SUBJECT = "Regularización {expediente} · {comunidad} · {vivienda}"
DEFAULT_BODY = (
    "Estimado/a {propietario}:\n\n"
    "Le adjuntamos la carta con la regularización del período {periodo} de la "
    "vivienda {vivienda} de {comunidad}.\n\n"
    "Para cualquier duda puede responder a este correo.\n\n"
    "Un saludo,\n{despacho}"
)
PLACEHOLDERS = ("propietario", "vivienda", "comunidad", "expediente", "periodo", "despacho")
_EMAIL = re.compile(r"^[A-Za-z0-9.!#$%&'*+/=?^_`{|}~-]+@[A-Za-z0-9-]+(?:\.[A-Za-z0-9-]+)+$")
_DOCX = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"

NO_EMAIL = "Sin correo en el listado de propietarios"
INVALID_EMAIL = "Correo no válido"
NO_LETTER = "No se encuentra el archivo de la carta"
EXCLUDED = "Excluido al preparar el lote"


@dataclass(frozen=True)
class MailCandidate:
    owner_id: int
    letter_id: int
    dwelling: str
    name: str
    email: str | None
    raw_email: str | None
    letter_path: Path | None
    problem: str | None
    duplicate: bool = False

    @property
    def sendable(self) -> bool:
        return self.problem is None


@dataclass(frozen=True)
class MailPlan:
    id_case: int
    letter_run_id: int | None
    candidates: tuple[MailCandidate, ...]

    @property
    def sendable(self) -> tuple[MailCandidate, ...]:
        return tuple(item for item in self.candidates if item.sendable)

    @property
    def without_email(self) -> tuple[MailCandidate, ...]:
        return tuple(item for item in self.candidates if not item.sendable)

    @property
    def duplicates(self) -> tuple[MailCandidate, ...]:
        return tuple(item for item in self.candidates if item.duplicate)


@dataclass(frozen=True)
class DeliveryRow:
    id_delivery: int
    owner_id: int
    dwelling: str
    name: str
    recipient: str | None
    status: str
    reason: str | None
    eml_path: str | None
    sent_at: str | None


@dataclass(frozen=True)
class MailRunSummary:
    id_mail_run: int
    status: str
    output_path: str | None
    deliveries: tuple[DeliveryRow, ...]

    def count(self, status: str) -> int:
        return sum(row.status == status for row in self.deliveries)


def normalise_email(raw: object) -> str | None:
    text = str(raw or "").strip().strip("<>").strip()
    if text.lower().startswith("mailto:"):
        text = text[7:]
    if not text or not _EMAIL.fullmatch(text):
        return None
    local, domain = text.rsplit("@", 1)
    return f"{local}@{domain.lower()}"


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def latest_letter_run(connection: sqlite3.Connection, id_case: int):
    return connection.execute(
        """SELECT id_letter_run, output_path FROM letter_generation_runs
           WHERE id_case=? AND status IN ('completed','incomplete')
           ORDER BY id_letter_run DESC LIMIT 1""",
        (id_case,),
    ).fetchone()


def plan_mail(connection: sqlite3.Connection, id_case: int) -> MailPlan:
    """Destinatario y carta de cada propietario del último lote de cartas."""
    run = latest_letter_run(connection, id_case)
    if run is None:
        return MailPlan(id_case, None, ())
    rows = connection.execute(
        """SELECT g.id_generated_letter, g.id_propietario, g.output_path, g.status,
                  p.codigo_vivienda, p.nombre_propietario, p.email
           FROM generated_letters g JOIN propietarios p USING(id_propietario)
           WHERE g.id_letter_run=? ORDER BY p.codigo_vivienda""",
        (run[0],),
    ).fetchall()
    emails = [normalise_email(row[6]) for row in rows]
    # «Ana@x.es» y «ana@x.es» son el mismo buzón en la práctica.
    folded = [email.casefold() if email else None for email in emails]
    repeated = {email for email in folded if email and folded.count(email) > 1}
    candidates = []
    for row, email in zip(rows, emails):
        path = Path(row[2]) if row[2] else None
        if row[3] != "generated" or path is None or not path.is_file():
            problem = NO_LETTER
        elif not str(row[6] or "").strip():
            problem = NO_EMAIL
        elif email is None:
            problem = INVALID_EMAIL
        else:
            problem = None
        candidates.append(MailCandidate(
            owner_id=int(row[1]), letter_id=int(row[0]), dwelling=str(row[4]), name=str(row[5]),
            email=email, raw_email=row[6], letter_path=path, problem=problem,
            duplicate=bool(email and email.casefold() in repeated),
        ))
    return MailPlan(id_case, int(run[0]), tuple(candidates))


def validate_template(template: str) -> None:
    unknown = sorted({name for _text, name, _spec, _conversion in string.Formatter().parse(template)
                      if name is not None and name not in PLACEHOLDERS})
    if unknown:
        raise ValueError(
            "Campos no reconocidos: " + ", ".join("{" + name + "}" for name in unknown)
            + ". Se pueden usar: " + ", ".join("{" + name + "}" for name in PLACEHOLDERS) + "."
        )


def _context(connection: sqlite3.Connection, id_case: int, owner_id: int) -> dict[str, str]:
    from office_settings import load_office_settings
    row = connection.execute(
        """SELECT r.nombre, r.fecha_inicio, r.fecha_fin, c.codigo, c.nombre,
                  p.nombre_propietario, p.codigo_vivienda
           FROM regularization_cases r JOIN comunidades c USING(id_comunidad)
           JOIN propietarios p ON p.id_propietario=?
           WHERE r.id_case=?""",
        (owner_id, id_case),
    ).fetchone()
    start, end = (f"{value[8:10]}/{value[5:7]}/{value[:4]}" for value in (row[1], row[2]))
    return {
        "expediente": row[0], "periodo": f"{start} – {end}", "comunidad": f"{row[4] or row[3]}",
        "propietario": row[5], "vivienda": row[6],
        "despacho": load_office_settings(connection).letter_signature,
    }


def _render(template: str, context: dict[str, str]) -> str:
    validate_template(template)
    return template.format_map(context)


def _fingerprint(recipient: str, attachment_sha256: str, subject: str, body: str) -> str:
    material = "\x1f".join((recipient.lower(), attachment_sha256, subject, body))
    return hashlib.sha256(material.encode("utf-8")).hexdigest()


def prepare_mail_run(
    connection: sqlite3.Connection,
    id_case: int,
    *,
    subject_template: str = DEFAULT_SUBJECT,
    body_template: str = DEFAULT_BODY,
    excluded_owner_ids: Iterable[int] = (),
    created_by: str,
) -> int:
    """Registra un lote con una entrega por propietario de la última carta."""
    if not created_by.strip():
        raise ValueError("Indica quién prepara el lote.")
    validate_template(subject_template)
    validate_template(body_template)
    plan = plan_mail(connection, id_case)
    if plan.letter_run_id is None:
        raise ValueError("Genera primero las cartas del expediente.")
    excluded = {int(owner) for owner in excluded_owner_ids}
    with connection:
        run_id = connection.execute(
            """INSERT INTO mail_runs(id_case,id_letter_run,subject_template,body_template,created_by)
               VALUES (?,?,?,?,?)""",
            (id_case, plan.letter_run_id, subject_template, body_template, created_by),
        ).lastrowid
        for item in plan.candidates:
            values = dict(recipient=item.email, attachment=str(item.letter_path) if item.letter_path else None,
                          sha=None, subject=None, fingerprint=None, status="skipped", reason=item.problem)
            if item.sendable and item.owner_id in excluded:
                values["reason"] = EXCLUDED
            elif item.sendable:
                context = _context(connection, id_case, item.owner_id)
                subject = _render(subject_template, context)
                body = _render(body_template, context)
                sha = _sha256(item.letter_path)
                values.update(sha=sha, subject=subject, status="pending", reason=None,
                              fingerprint=_fingerprint(item.email, sha, subject, body))
            connection.execute(
                """INSERT INTO mail_deliveries(id_mail_run,id_propietario,id_generated_letter,recipient,
                       attachment_path,attachment_sha256,subject,fingerprint,status,reason)
                   VALUES (?,?,?,?,?,?,?,?,?,?)""",
                (run_id, item.owner_id, item.letter_id, values["recipient"], values["attachment"],
                 values["sha"], values["subject"], values["fingerprint"], values["status"], values["reason"]),
            )
    return int(run_id)


def _run(connection: sqlite3.Connection, id_mail_run: int):
    run = connection.execute(
        "SELECT id_case, subject_template, body_template, status, output_path FROM mail_runs WHERE id_mail_run=?",
        (id_mail_run,),
    ).fetchone()
    if run is None:
        raise LookupError("No existe ese lote de correo.")
    return run


def _message(connection, run, delivery, sender: str | None) -> EmailMessage:
    context = _context(connection, int(run[0]), int(delivery["id_propietario"]))
    body = _render(run[2], context)
    attachment = Path(delivery["attachment_path"])
    if _fingerprint(delivery["recipient"], delivery["attachment_sha256"], delivery["subject"], body) != delivery["fingerprint"]:
        raise MailTransportError("Los datos del correo cambiaron después de preparar el lote; prepara uno nuevo.")
    if not attachment.is_file():
        raise MailTransportError("La carta ya no está en su carpeta; genera las cartas o prepara un lote nuevo.")
    if _sha256(attachment) != delivery["attachment_sha256"]:
        raise MailTransportError("La carta cambió después de preparar el lote; prepara un lote nuevo.")
    message = EmailMessage()
    if sender:
        message["From"] = sender
    message["To"] = delivery["recipient"]
    message["Subject"] = delivery["subject"]
    message["Date"] = formatdate(localtime=True)
    message["Message-ID"] = make_msgid(domain=(sender or "regularizaciones.local").split("@")[-1])
    message.set_content(body)
    maintype, subtype = (mimetypes.guess_type(attachment.name)[0] or _DOCX).split("/", 1)
    message.add_attachment(attachment.read_bytes(), maintype=maintype, subtype=subtype, filename=attachment.name)
    return message


def _safe_name(value: str) -> str:
    return re.sub(r"[^\w.-]+", "_", value, flags=re.UNICODE).strip("_")[:60] or "entrega"


def _write_atomic(path: Path, data: bytes) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    handle, temporary = tempfile.mkstemp(dir=path.parent, prefix=".", suffix=".tmp")
    try:
        with os.fdopen(handle, "wb") as stream:
            stream.write(data)
        os.replace(temporary, path)
    except Exception:
        Path(temporary).unlink(missing_ok=True)
        raise


def _deliveries(connection, id_mail_run: int, statuses: tuple[str, ...]):
    cursor = connection.cursor()
    cursor.row_factory = sqlite3.Row
    return cursor.execute(
        f"""SELECT d.*, p.codigo_vivienda, p.nombre_propietario FROM mail_deliveries d
            JOIN propietarios p USING(id_propietario)
            WHERE d.id_mail_run=? AND d.status IN ({",".join("?" * len(statuses))})
            ORDER BY p.codigo_vivienda""",
        (id_mail_run, *statuses),
    ).fetchall()


def _update(connection, id_delivery: int, **fields) -> None:
    assignments = ", ".join(f"{key}=?" for key in fields)
    connection.execute(
        f"UPDATE mail_deliveries SET {assignments}, updated_at=datetime('now') WHERE id_delivery=?",
        (*fields.values(), id_delivery),
    )
    connection.commit()


def generate_eml_drafts(
    connection: sqlite3.Connection, id_mail_run: int, output_root: Path, *, sender: str | None = None,
) -> MailRunSummary:
    """Un .eml por entrega pendiente; se reutiliza el de una huella idéntica."""
    run = _run(connection, id_mail_run)
    case = connection.execute(
        "SELECT c.codigo, r.nombre FROM regularization_cases r JOIN comunidades c USING(id_comunidad) WHERE r.id_case=?",
        (run[0],),
    ).fetchone()
    folder = Path(output_root) / "correo" / _safe_name(str(case[0])) / _safe_name(str(case[1])) / f"lote_{id_mail_run}"
    folder.mkdir(parents=True, exist_ok=True)
    for delivery in _deliveries(connection, id_mail_run, ("pending",)):
        previous = connection.execute(
            """SELECT eml_path FROM mail_deliveries
               WHERE fingerprint=? AND eml_path IS NOT NULL AND id_delivery<>? ORDER BY id_delivery DESC""",
            (delivery["fingerprint"], delivery["id_delivery"]),
        ).fetchall()
        reused = next((row[0] for row in previous if Path(row[0]).is_file()), None)
        try:
            message = _message(connection, run, delivery, sender)
        except MailTransportError as error:
            _update(connection, delivery["id_delivery"], status="failed", reason=str(error))
            continue
        if reused:
            _update(connection, delivery["id_delivery"], status="draft", eml_path=reused,
                    reason="Borrador idéntico ya generado")
            continue
        message["X-Unsent"] = "1"  # Outlook lo abre como borrador para enviar
        path = folder / f"{_safe_name(delivery['codigo_vivienda'])}_{_safe_name(delivery['nombre_propietario'])}.eml"
        _write_atomic(path, bytes(message))
        _update(connection, delivery["id_delivery"], status="draft", eml_path=str(path), reason=None)
    connection.execute(
        "UPDATE mail_runs SET status='drafted', output_path=?, updated_at=datetime('now') WHERE id_mail_run=?",
        (str(folder), id_mail_run),
    )
    connection.commit()
    return mail_run_summary(connection, id_mail_run)


def send_mail_run(
    connection: sqlite3.Connection, id_mail_run: int, transport, *, confirmed_by: str, sender: str,
) -> MailRunSummary:
    """Envía las entregas pendientes o fallidas tras una confirmación explícita.

    Nunca reenvía una huella ya enviada (en este lote o en otro) y deja el lote
    en ``completed`` o ``incomplete`` aunque falle la conexión.
    """
    if not str(confirmed_by or "").strip():
        raise PermissionError("El envío necesita la confirmación de quien lo realiza.")
    if not str(sender or "").strip():
        raise ValueError("Falta el remitente del correo.")
    run = _run(connection, id_mail_run)
    connection.execute(
        "UPDATE mail_runs SET status='sending', confirmed_by=?, updated_at=datetime('now') WHERE id_mail_run=?",
        (confirmed_by.strip(), id_mail_run),
    )
    connection.commit()
    try:
        pending = _deliveries(connection, id_mail_run, ("pending", "draft", "failed"))
        work = []
        for delivery in pending:
            already = connection.execute(
                "SELECT sent_at FROM mail_deliveries WHERE fingerprint=? AND status='sent' AND id_delivery<>? LIMIT 1",
                (delivery["fingerprint"], delivery["id_delivery"]),
            ).fetchone()
            if already is not None:
                _update(connection, delivery["id_delivery"], status="skipped",
                        reason=f"Ya enviado el {already[0]}; no se reenvía")
                continue
            try:
                work.append((delivery, _message(connection, run, delivery, sender)))
            except MailTransportError as error:
                _update(connection, delivery["id_delivery"], status="failed", reason=str(error))
        if work:
            try:
                with transport as opened:
                    for delivery, message in work:
                        try:
                            message_id = opened.send(message)
                        except MailTransportError as error:
                            _update(connection, delivery["id_delivery"], status="failed", reason=str(error)[:300])
                            continue
                        connection.execute(
                            """UPDATE mail_deliveries SET status='sent', reason=NULL, message_id=?,
                                   sent_at=datetime('now'), updated_at=datetime('now') WHERE id_delivery=?""",
                            (message_id, delivery["id_delivery"]),
                        )
                        connection.commit()
            except MailTransportError as error:
                for delivery, _message_obj in work:
                    current = connection.execute(
                        "SELECT status FROM mail_deliveries WHERE id_delivery=?", (delivery["id_delivery"],),
                    ).fetchone()[0]
                    if current != "sent":
                        _update(connection, delivery["id_delivery"], status="failed", reason=str(error)[:300])
    finally:
        open_items = connection.execute(
            "SELECT COUNT(*) FROM mail_deliveries WHERE id_mail_run=? AND status IN ('pending','draft','failed')",
            (id_mail_run,),
        ).fetchone()[0]
        connection.execute(
            """UPDATE mail_runs SET status=?, updated_at=datetime('now'),
                   completed_at=CASE WHEN ?='completed' THEN datetime('now') ELSE completed_at END
               WHERE id_mail_run=?""",
            ("incomplete" if open_items else "completed", "incomplete" if open_items else "completed", id_mail_run),
        )
        connection.commit()
    return mail_run_summary(connection, id_mail_run)


def mail_run_summary(connection: sqlite3.Connection, id_mail_run: int) -> MailRunSummary:
    run = _run(connection, id_mail_run)
    rows = connection.execute(
        """SELECT d.id_delivery, d.id_propietario, p.codigo_vivienda, p.nombre_propietario,
                  d.recipient, d.status, d.reason, d.eml_path, d.sent_at
           FROM mail_deliveries d JOIN propietarios p USING(id_propietario)
           WHERE d.id_mail_run=? ORDER BY p.codigo_vivienda""",
        (id_mail_run,),
    ).fetchall()
    return MailRunSummary(id_mail_run, str(run[3]), run[4], tuple(DeliveryRow(*row) for row in rows))


def latest_mail_run(connection: sqlite3.Connection, id_case: int) -> int | None:
    row = connection.execute(
        "SELECT id_mail_run FROM mail_runs WHERE id_case=? ORDER BY id_mail_run DESC LIMIT 1", (id_case,),
    ).fetchone()
    return int(row[0]) if row else None


__all__ = [
    "DEFAULT_BODY", "DEFAULT_SUBJECT", "DeliveryRow", "MailCandidate", "MailPlan", "MailRunSummary",
    "PLACEHOLDERS", "generate_eml_drafts", "latest_letter_run", "latest_mail_run", "mail_run_summary",
    "normalise_email", "plan_mail", "prepare_mail_run", "send_mail_run", "validate_template",
]
