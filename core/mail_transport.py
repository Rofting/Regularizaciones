"""Transporte SMTP opcional; el secreto permanece en el gestor del sistema."""

from __future__ import annotations

import smtplib
import ssl
from dataclasses import dataclass, field
from email.message import EmailMessage


@dataclass(frozen=True)
class SmtpSettings:
    host: str
    port: int
    username: str
    credential_key: str
    timeout: int = 20


class SmtpTransport:
    def __init__(self, settings: SmtpSettings, *, session_password: str | None = None):
        if not settings.host.strip() or not settings.username.strip() or not settings.credential_key.strip():
            raise ValueError("Falta la configuración SMTP")
        if not 1 <= settings.port <= 65535 or settings.timeout <= 0:
            raise ValueError("Puerto o tiempo de espera SMTP no válido")
        self.settings = settings
        self._session_password = session_password

    def send(self, message: EmailMessage) -> None:
        password = self._session_password
        if password is None:
            import keyring
            password = keyring.get_password("Regularizaciones SMTP", self.settings.credential_key)
        if not password:
            raise ValueError("No hay contraseña SMTP en el gestor de credenciales")
        with smtplib.SMTP(self.settings.host, self.settings.port,
                          timeout=self.settings.timeout) as client:
            client.ehlo()
            client.starttls(context=ssl.create_default_context())
            client.ehlo()
            client.login(self.settings.username, password)
            client.send_message(message)


@dataclass
class FakeMailTransport:
    messages: list[EmailMessage] = field(default_factory=list)

    def send(self, message: EmailMessage) -> None:
        self.messages.append(message)
