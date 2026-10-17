"""Проверки действий, которые вызывает Qt-меню."""

import sys
import unittest
from unittest.mock import Mock, patch

if sys.platform != "win32":
    raise unittest.SkipTest("Действия трея используют Windows API")

import tray


class TrayActionTests(unittest.TestCase):
    def test_conflicting_hotkey_is_not_saved(self):
        actions = {
            "translate": ("смены раскладки", lambda: {"modifiers": [], "key": "F4"},
                          Mock(), Mock()),
            "exit": ("выхода", lambda: {"modifiers": [], "key": "F10"},
                     Mock(), Mock()),
        }
        with patch.object(tray, "HOTKEY_ACTIONS", actions), patch.object(tray, "_show_error") as error:
            self.assertFalse(tray._set_hotkey_and_apply([], "F10", "translate"))
        actions["translate"][2].assert_not_called()
        error.assert_called_once()

    def test_valid_hotkey_is_saved_and_applied(self):
        write = Mock(return_value=True)
        update = Mock(return_value=True)
        actions = {
            "translate": ("смены раскладки", lambda: {"modifiers": [], "key": "F4"}, write, update),
            "exit": ("выхода", lambda: {"modifiers": [], "key": "F10"},
                     Mock(), Mock()),
        }
        with patch.object(tray, "HOTKEY_ACTIONS", actions):
            self.assertTrue(tray._set_hotkey_and_apply(["CTRL"], "T", "translate"))
        write.assert_called_once_with(["CTRL"], "T")
        update.assert_called_once()


if __name__ == "__main__":
    unittest.main()
