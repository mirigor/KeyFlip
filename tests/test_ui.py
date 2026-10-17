"""Проверки панели и её связи с настройками без регистрации системных хоткеев."""

import os
import sys
import tempfile
import unittest
from unittest.mock import patch

if sys.platform != "win32":
    raise unittest.SkipTest("KeyFlip UI imports Windows-only actions")

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtCore import QPoint, QRect, QSize
from PySide6.QtWidgets import QApplication

import config
import ui


class PopupPositionTests(unittest.TestCase):
    def test_popup_stays_inside_each_monitor_and_taskbar_edge(self):
        for available in (QRect(0, 0, 1920, 1040), QRect(-1920, 0, 1920, 1040)):
            anchors = [QPoint(available.left() + 100, 1060),
                       QPoint(available.left() + 100, -20),
                       QPoint(available.left() - 20, 300),
                       QPoint(available.right() + 20, 300)]
            for anchor in anchors:
                with self.subTest(available=available, anchor=anchor):
                    size = QSize(340, 308)
                    position = ui.popup_position(QRect(anchor.x() - 12, anchor.y() - 12, 24, 24),
                                                 available, size, anchor)
                    self.assertTrue(available.contains(QRect(position, size)))


class UITests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])
        cls.app.setQuitOnLastWindowClosed(False)

    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.config_patch = patch.object(config, "CONFIG_FILE", os.path.join(self.directory.name, "config.json"))
        self.config_patch.start()
        config.load_config()
        self.widgets = []

    def tearDown(self):
        for widget in self.widgets:
            widget.close()
        self.app.processEvents()
        self.config_patch.stop()
        self.directory.cleanup()

    def test_popup_refresh_shows_current_enabled_state_and_hotkeys(self):
        popup = ui.KeyFlipPopup()
        self.widgets.append(popup)
        config.set_enabled(False)
        config.write_translate_hotkey(["CTRL", "ALT"], "T")
        popup.refresh()
        self.assertFalse(popup.toggle.isChecked())
        self.assertEqual(popup.status.text(), "На паузе")
        self.assertIn("Смена раскладки", popup.shortcuts["translate"].text())
        self.assertIn("Ctrl + Alt + T", popup.shortcuts["translate"].text())
        self.assertFalse(hasattr(popup, "settings"))

    def test_popup_shortcut_opens_editor_for_selected_action(self):
        controller = ui.TrayController()
        self.widgets.append(controller)
        with patch.object(ui, "capture_hotkey_and_apply_via_thread") as capture:
            controller.popup.shortcuts["case"].click()
            capture.assert_called_once_with("case")

    def test_popup_settings_apply_and_refresh_saved_state(self):
        controller = ui.TrayController()
        self.widgets.append(controller)
        with patch.object(ui.actions, "toggle_autorun") as autorun, \
                patch.object(ui.actions, "toggle_file_logging") as logging:
            controller.popup.autorun_toggle.click()
            controller.popup.logging_toggle.click()
            autorun.assert_called_once_with()
            logging.assert_called_once_with()
        self.assertFalse(controller.popup.autorun_toggle.isChecked())
        self.assertFalse(controller.popup.logging_toggle.isChecked())

        config.write_autorun_flag(True)
        config.write_file_logging_flag(True)
        controller.refresh()
        self.assertTrue(controller.popup.autorun_toggle.isChecked())
        self.assertTrue(controller.popup.logging_toggle.isChecked())

    def test_custom_menu_actions_open_the_correct_editor(self):
        controller = ui.TrayController()
        self.widgets.append(controller)
        for target, menu in controller._hotkey_menus.items():
            with self.subTest(target=target), patch.object(controller, "edit_hotkey") as edit:
                menu.actions()[-1].trigger()
                edit.assert_called_once_with(target)

    def test_exit_event_requests_qt_shutdown(self):
        controller = ui.TrayController()
        self.widgets.append(controller)
        with patch.object(ui, "exit_event") as event, patch.object(self.app, "quit") as quit_app:
            event.is_set.return_value = True
            controller._check_exit()
            quit_app.assert_called_once_with()


if __name__ == "__main__":
    unittest.main()
