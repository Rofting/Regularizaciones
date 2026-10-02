"""Transportes de correo: SMTP real (opcional) y uno falso para las pruebas.

La contraseña SMTP no se guarda en la base, en JSON ni en registros: se pide
al almacén de credenciales del sistema (Administrador de credenciales de
Windows) mediante ``keyring``, con el usuario como identificador.
"""

from __future__ import annotations

import json
import re
import smtplib
import sqlite3
import ssl
from dataclasses import asdict, dataclass, field
from email.message import EmailMessage
from email.utils import make_msgid
from typing import Callable, Protocol

KEYRING_SERVICE = "Regularizaciones-SMTP"
_SECURITY = ("ssl", "starttls")


class MailTransportError(RuntimeError):
    """Error de envío con un mensaje apto para mostrar y guardar."""


class MailTransport(Protocol):
    def send(self, message: EmailMessage) -> str:
        """Envía el mensaje y devuelve su Message-ID."""


@dataclass(frozen=True)
class SmtpSettings:
    host: str = ""
    port: int = 587
    username: str = ""
    security: str = "starttls"   # "starttls" (587) o "ssl" (465)
    sender: str = ""             # remitente visible; por defecto el usuario
    timeout_seconds: int = 30

    @property
    def configured(self) -> bool:
        return bool(self.host.strip() and self.username.strip())

    @property
    def from_address(self) -> str:
        return (self.sender or self.username).strip()

    def validate(self) -> None:
        if not self.host.strip():
            raise ValueError("Indica el servidor SMTP (p. ej. smtp.office365.com).")
        if not self.username.strip():
            raise ValueError("Indica el usuario de la cuenta de correo.")
        if self.security not in _SECURITY:
            raise ValueError("La seguridad debe ser STARTTLS o SSL.")
        if not 0 < int(self.port) < 65536:
            raise ValueError("El puerto no es válido.")


def load_smtp_settings(connection: sqlite3.Connection) -> SmtpSettings:
    row = connection.execute("SELECT settings_json FROM mail_settings WHERE id=1").fetchone()
    if row is None:
        return SmtpSettings()
    data = json.loads(row[0])
    known = {key: data[key] for key in SmtpSettings.__dataclass_fields__ if key in data}
    return SmtpSettings(**known)


def save_smtp_settings(connection: sqlite3.Connection, settings: SmtpSettings) -> None:
    settings.validate()
    connection.execute(
        """INSERT INTO mail_settings(id, settings_json) VALUES (1, ?)
           ON CONFLICT(id) DO UPDATE SET settings_json=excluded.settings_json,
                                         updated_at=datetime('now')""",
        (json.dumps(asdict(settings), ensure_ascii=False),),
    )
    connection.commit()


def _keyring():
    try:
        import keyring
    except ImportError as error:  # instalación antigua sin la librería
        raise MailTransportError(
            "Falta el componente de credenciales (keyring). Ejecuta INSTALAR.bat de nuevo."
        ) from error
    return keyring


def store_password(username: str, password: str) -> None:
    """Guarda la contraseña en el almacén del sistema, nunca en la aplicación."""
    if not username.strip() or not password:
        raise ValueError("Indica usuario y contraseña.")
    _keyring().set_password(KEYRING_SERVICE, username.strip(), password)


def has_password(username: str) -> bool:
    try:
        return bool(_keyring().get_password(KEYRING_SERVICE, username.strip()))
    except Exception:
        return False


def keyring_password(username: str) -> str:
    try:
        password = _keyring().get_password(KEYRING_SERVICE, username.strip())
    except MailTransportError:
        raise
    except Exception as error:
        raise MailTransportError(f"No se pudo leer la contraseña guardada: {type(error).__name__}") from error
    if not password:
        raise MailTransportError("No hay contraseña guardada para esta cuenta. Configura el envío de nuevo.")
    return password


def sanitise_error(error: BaseException, *secrets: str) -> str:
    """Mensaje corto sin contraseñas ni respuestas largas del servidor."""
    text = f"{type(error).__name__}: {error}"
    for secret in secrets:
        if secret:
            text = text.replace(secret, "***")
    text = re.sub(r"\s+", " ", text).strip()
    return text[:300]


class SmtpTransport:
    """Envía con TLS obligatorio (STARTTLS o SSL) y tiempo límite."""

    def __init__(self, settings: SmtpSettings, password_provider: Callable[[str], str] = keyring_password):
        settings.validate()
        self.settings = settings
        self._password_provider = password_provider
        self._connection: smtplib.SMTP | None = None
        self._password = ""

    def __enter__(self) -> "SmtpTransport":
        self._password = self._password_provider(self.settings.username)
        context = ssl.create_default_context()
        try:
            if self.settings.security == "ssl":
                server = smtplib.SMTP_SSL(self.settings.host, self.settings.port,
                                          timeout=self.settings.timeout_seconds, context=context)
            else:
                server = smtplib.SMTP(self.settings.host, self.settings.port,
                                      timeout=self.settings.timeout_seconds)
                server.ehlo()
                server.starttls(context=context)
                server.ehlo()
            server.login(self.settings.username, self._password)
        except (smtplib.SMTPException, OSError) as error:
            raise MailTransportError(
                "No se pudo conectar con el servidor de correo: " + sanitise_error(error, self._password)
            ) from error
        self._connection = server
        return self

    def __exit__(self, *_exc) -> None:
        if self._connection is not None:
            try:
                self._connection.quit()
            except Exception:
                pass
        self._connection = None
        self._password = ""

    def send(self, message: EmailMessage) -> str:
        if self._connection is None:
            raise MailTransportError("La conexión SMTP no está abierta.")
        if "Message-ID" not in message:
            message["Message-ID"] = make_msgid(domain=self.settings.from_address.split("@")[-1] or None)
        try:
            refused = self._connection.send_message(message)
        except (smtplib.SMTPException, OSError) as error:
            raise MailTransportError(sanitise_error(error, self._password)) from error
        if refused:
            raise MailTransportError("El servidor rechazó el destinatario: " + ", ".join(refused))
        return str(message["Message-ID"])


@dataclass
class FakeMailTransport:
    """Transporte de pruebas: no envía nada; guarda los mensajes."""

    fail_for: set[str] = field(default_factory=set)
    sent: list[EmailMessage] = field(default_factory=list)

    def __enter__(self) -> "FakeMailTransport":
        return self

    def __exit__(self, *_exc) -> None:
        return None

    def send(self, message: EmailMessage) -> str:
        if str(message["To"]) in self.fail_for:
            raise MailTransportError("550 Buzón inexistente")
        self.sent.append(message)
        return message["Message-ID"] or make_msgid(domain="prueba.local")


__all__ = [
    "FakeMailTransport", "KEYRING_SERVICE", "MailTransport", "MailTransportError", "SmtpSettings",
    "SmtpTransport", "has_password", "keyring_password", "load_smtp_settings", "sanitise_error",
    "save_smtp_settings", "store_password",
]
