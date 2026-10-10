import sys
import unittest
from pathlib import Path
from unittest.mock import Mock

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "core"))
from app import AppGestionFincas


class DialogVisibilityTest(unittest.TestCase):
    def test_hidden_window_does_not_capture_clicks_and_retries(self):
        dialog = Mock()
        dialog.winfo_viewable.return_value = False
        AppGestionFincas._grab_seguro(dialog)
        dialog.grab_set.assert_not_called()
        dialog.winfo_viewable.return_value = True
        dialog.after.call_args.args[1]()
        dialog.lift.assert_called()
        dialog.focus_set.assert_called_once()
        dialog.grab_set.assert_called_once()

    def test_permanently_hidden_window_is_destroyed_without_grab(self):
        dialog = Mock()
        dialog.winfo_viewable.return_value = False
        AppGestionFincas._grab_seguro(dialog, intentos=0)
        dialog.grab_set.assert_not_called()
        dialog.destroy.assert_called_once()

    def test_closed_window_is_not_reopened_by_delayed_callback(self):
        dialog = Mock()
        dialog.winfo_exists.return_value = False
        AppGestionFincas._grab_seguro(dialog)
        dialog.deiconify.assert_not_called()
        dialog.grab_set.assert_not_called()
