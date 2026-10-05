import sys
import types
import unittest
from email.message import EmailMessage
from pathlib import Path
from unittest.mock import patch


CORE = Path(__file__).resolve().parents[1] / "core"
if str(CORE) not in sys.path:
    sys.path.insert(0, str(CORE))

from mail_transport import SmtpSettings, SmtpTransport


class SmtpTransportTest(unittest.TestCase):
    def test_uses_system_secret_and_starttls_before_authentication(self):
        calls = []

        class Client:
            def __init__(self, host, port, timeout):
                calls.append(("connect", host, port, timeout))

            def __enter__(self):
                return self

            def __exit__(self, *_args):
                calls.append(("close",))

            def ehlo(self):
                calls.append(("ehlo",))

            def starttls(self, *, context):
                calls.append(("starttls", context is not None))

            def login(self, user, password):
                calls.append(("login", user, password))

            def send_message(self, message):
                calls.append(("send", message["To"]))

        message = EmailMessage()
        message["To"] = "test@example.org"
        with patch.dict(sys.modules, {"keyring": types.SimpleNamespace(get_password=lambda *_: "secret")}), \
             patch("mail_transport.smtplib.SMTP", Client):
            SmtpTransport(SmtpSettings("mail.example.org", 587, "usuario", "cuenta")).send(message)
        self.assertEqual(["connect", "ehlo", "starttls", "ehlo", "login", "send", "close"],
                         [call[0] for call in calls])

    def test_session_password_works_without_keyring(self):
        with patch("mail_transport.smtplib.SMTP") as client:
            SmtpTransport(SmtpSettings("mail.example.org", 587, "usuario", "cuenta"),
                          session_password="secret").send(EmailMessage())
        client.return_value.__enter__.return_value.login.assert_called_once_with("usuario", "secret")


if __name__ == "__main__":
    unittest.main()
