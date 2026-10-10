"""Prueba gráfica real del alta en una base temporal (Windows o Xvfb)."""
import os
import sqlite3
import sys
import tempfile
import time
import traceback
from contextlib import closing
from pathlib import Path
from unittest.mock import patch


def main():
    with tempfile.TemporaryDirectory(prefix="regularizaciones-ui-") as home:
        os.environ["REGULARIZACIONES_HOME"] = home
        sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "core"))
        import app

        errors = []
        with patch.object(app.AppGestionFincas, "_comprobar_datos_despacho"), \
                patch.object(app.messagebox, "showinfo") as info, \
                patch.object(app.messagebox, "showerror") as error:
            root = app.AppGestionFincas()
            root.report_callback_exception = lambda *exc: errors.append(exc)

            def pump(predicate, timeout=5):
                deadline = time.monotonic() + timeout
                while time.monotonic() < deadline:
                    root.update()
                    if predicate():
                        return
                    time.sleep(.01)
                windows = [(w.title(), w.state(), w.winfo_viewable())
                           for w in root.winfo_children() if isinstance(w, app.ctk.CTkToplevel)]
                raise AssertionError(f"La interfaz no respondió dentro del plazo: {windows}; callbacks: {errors}")

            def descendants(widget):
                for child in widget.winfo_children():
                    yield child
                    yield from descendants(child)

            try:
                pump(lambda: root.winfo_viewable())
                next(w for w in descendants(root) if isinstance(w, app.ctk.CTkButton)
                     and w.cget("text") == "Nueva comunidad").invoke()
                pump(lambda: root.grab_current() is not None)
                dialog = root.grab_current()
                assert dialog.title() == "Añadir comunidad"
                assert dialog.winfo_viewable() and float(dialog.attributes("-alpha")) == 1
                # Elegir la opción visible de registro rápido, sin saltar su callback.
                labels = [w for w in descendants(dialog) if isinstance(w, app.ctk.CTkLabel)
                          and w.cget("text") == "Registro rápido"]
                row = labels[0].master.master
                next(w for w in descendants(row) if isinstance(w, app.ctk.CTkButton)).invoke()
                pump(lambda: root.grab_current() is not None)
                form = root.grab_current()
                assert form.title() == "Nueva Comunidad" and form.winfo_viewable()
                entries = [w for w in descendants(form) if isinstance(w, app.ctk.CTkEntry)]
                for entry, value in zip(entries, ("TEST", "Comunidad de prueba", "", "4")):
                    entry.insert(0, value)
                next(w for w in descendants(form) if isinstance(w, app.ctk.CTkButton)
                     and w.cget("text") == "Guardar comunidad").invoke()
                pump(lambda: root.grab_current() is None)
                assert info.call_count == 1, "No se confirmó el alta"
                assert not error.called, error.call_args
                assert not errors, errors
                with closing(sqlite3.connect(app.RUTA_BD)) as connection:
                    assert connection.execute(
                        "SELECT nombre, num_viviendas FROM comunidades WHERE codigo='TEST'"
                    ).fetchone() == ("Comunidad de prueba", 4)
                root._nueva_comunidad()
                pump(lambda: root.grab_current() is not None)
                root.grab_current().event_generate("<Escape>")
                pump(lambda: root.grab_current() is None)
                assert not errors, errors
                print("Interfaz verificada: abrir alta, guardar comunidad y cerrar con Escape.")
            finally:
                root.destroy()


if __name__ == "__main__":
    try:
        main()
    except Exception:
        detail = traceback.format_exc().replace("%", "%25").replace("\r", "%0D").replace("\n", "%0A")
        print("::error title=Prueba gráfica de alta::" + detail)
        raise
