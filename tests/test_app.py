import os
import runpy
import sys
import tempfile
import unittest
from unittest.mock import Mock, patch

if sys.platform != "win32":
    raise unittest.SkipTest("Windows-only application entry point")

import config
import logging_setup
import ui
import winapi

ENTRY = runpy.run_path(os.path.join(config.BASE_DIR, "app.pyw"))


class StartupTests(unittest.TestCase):
    def test_installer_operation_rejects_unknown_argument(self):
        self.assertEqual(ENTRY["installer_operation"](["--unexpected"]), 2)

    def test_installer_operation_detects_running_instance(self):
        operation = ENTRY["installer_operation"]
        with patch.object(ENTRY["win32event"], "CreateMutex", return_value=123), \
                patch.object(ENTRY["win32api"], "GetLastError", return_value=183), \
                patch.object(ENTRY["win32api"], "CloseHandle") as close:
            self.assertEqual(operation(["--installer-check"]), 3)
            close.assert_called_once_with(123)

    def test_uninstall_cleanup_removes_only_own_autorun(self):
        operation = ENTRY["installer_operation"]
        with tempfile.TemporaryDirectory() as directory, \
                patch.object(config, "_startup_shortcut_path", return_value=os.path.join(directory, "KeyFlip.lnk")), \
                patch.object(ENTRY["win32event"], "CreateMutex", return_value=123), \
                patch.object(ENTRY["win32api"], "GetLastError", return_value=0), \
                patch.object(ENTRY["win32api"], "CloseHandle"), \
                patch.object(ENTRY["win32com"].client, "Dispatch") as dispatch:
            shortcut_path = os.path.join(directory, "KeyFlip.lnk")
            with open(shortcut_path, "wb") as shortcut_file:
                shortcut_file.write(b"test")
            dispatch.return_value.CreateShortcut.return_value.TargetPath = "C:\\Other\\KeyFlip.exe"
            self.assertEqual(operation(["--uninstall-cleanup"]), 0)
            self.assertTrue(os.path.isfile(shortcut_path))
            dispatch.return_value.CreateShortcut.return_value.TargetPath = sys.executable
            self.assertEqual(operation(["--uninstall-cleanup"]), 0)
            self.assertFalse(os.path.exists(shortcut_path))

    def test_console_handler_does_not_write_to_missing_pythonw_stream(self):
        import logging
        with patch.object(logging_setup.sys, "stdout", None):
            self.assertIsInstance(logging_setup._create_console_handler(), logging.NullHandler)

    def test_startup_errors_are_saved_even_with_file_logging_disabled(self):
        report = ENTRY["_report_startup_error"]
        with tempfile.TemporaryDirectory() as directory, \
                patch.dict(report.__globals__, {"STARTUP_ERROR_FILE": os.path.join(directory, "error.log")}), \
                patch.object(ENTRY["win32api"], "MessageBox") as message:
            try:
                raise RuntimeError("test startup failure")
            except RuntimeError as error:
                report(error)
            with open(os.path.join(directory, "error.log"), encoding="utf-8") as saved:
                self.assertIn("RuntimeError: test startup failure", saved.read())
            message.assert_called_once()

    def test_startup_failure_closes_controller_and_waits_for_worker(self):
        main = ENTRY["main"]
        with patch.dict(main.__globals__, {"load_config": lambda: None, "apply_autorun_setting": lambda: None}), \
                patch.object(winapi, "exit_event") as exit_event, \
                patch.object(winapi, "register_exit_handler"), \
                patch.object(winapi.hotkey_ready, "wait", return_value=True), \
                patch.object(winapi, "HOTKEY_THREAD_ID", 0), \
                patch.object(winapi, "wait_for_pending_operations") as pending, \
                patch.object(ui, "TrayController") as controller, \
                patch("PySide6.QtGui.QIcon", return_value=Mock()), \
                patch.object(main.__globals__["threading"], "Thread") as worker, \
                patch.dict(main.__globals__, {"QApplication": Mock()}), \
                patch.dict(main.__globals__, {"_create_single_instance_mutex": lambda: 123,
                                             "_release_mutex": Mock()}) as mutex_scope:
            release = mutex_scope["_release_mutex"]
            controller.return_value.start.side_effect = RuntimeError("test failure")
            with self.assertRaisesRegex(RuntimeError, "test failure"):
                main()
            exit_event.set.assert_called_once()
            controller.return_value.close.assert_called_once()
            worker.return_value.join.assert_called_once()
            pending.assert_called_once()
            release.assert_called_once_with(123)


if __name__ == "__main__":
    unittest.main()
