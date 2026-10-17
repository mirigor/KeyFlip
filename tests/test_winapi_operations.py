import ctypes
import sys
import threading
import unittest
from contextlib import ExitStack
from unittest.mock import Mock, patch

if sys.platform != "win32":
    raise unittest.SkipTest("Windows-only input operations")

import winapi


class InputTests(unittest.TestCase):
    def test_input_structure_matches_platform_abi(self):
        self.assertEqual(ctypes.sizeof(winapi.INPUT), 40 if ctypes.sizeof(ctypes.c_void_p) == 8 else 28)

    def test_emoji_and_newlines_are_encoded_without_plain_enter(self):
        events = []

        def send(count, inputs, size):
            events.extend((inputs[i].ki.wVk, inputs[i].ki.wScan, inputs[i].ki.dwFlags) for i in range(count))
            return count

        with patch.object(winapi, "SendInput", side_effect=send):
            self.assertTrue(winapi.send_unicode_via_sendinput("🙂\r\n", 0))
        self.assertEqual([event[1] for event in events[:4]], [0xD83D, 0xD83D, 0xDE42, 0xDE42])
        self.assertEqual([event[0] for event in events[4:]],
                         [winapi.VK_SHIFT, winapi.VK_RETURN, winapi.VK_RETURN, winapi.VK_SHIFT])

    def test_partial_input_is_distinct_from_zero_input(self):
        with patch.object(winapi, "SendInput", side_effect=[1, 1]) as send:
            with self.assertRaises(winapi.PartialInputError):
                winapi.send_unicode_via_sendinput("A", 0)
            self.assertEqual(send.call_count, 2)  # Отпускается незавершённое нажатие
        with patch.object(winapi, "SendInput", return_value=0):
            self.assertFalse(winapi.send_unicode_via_sendinput("A", 0))


class OperationTests(unittest.TestCase):
    def prepare(self, stack):
        stack.enter_context(patch.object(winapi, "is_enabled", return_value=True))
        event = stack.enter_context(patch.object(winapi, "exit_event"))
        event.is_set.return_value = False
        native = stack.enter_context(patch.object(winapi, "user32"))
        native.GetForegroundWindow.return_value = 101
        native.GetKeyboardLayout.return_value = 0x0409
        stack.enter_context(patch.object(winapi, "get_active_window_info", return_value=("editor", 1, "editor.exe")))
        stack.enter_context(patch.object(winapi, "_save_clipboard_for_restore", return_value=(True, [])))
        restore = stack.enter_context(patch.object(winapi, "_restore_clipboard_snapshot"))
        stack.enter_context(patch.object(winapi, "_release_modifiers_after_hotkey"))
        copy = stack.enter_context(patch.object(winapi, "safe_copy_from_selection", return_value=("abc", set())))
        send = stack.enter_context(patch.object(winapi, "send_unicode_via_sendinput", return_value=True))
        paste = stack.enter_context(patch.object(winapi, "send_ctrl_v"))
        clipboard = stack.enter_context(patch.object(winapi.pyperclip, "copy"))
        return native, copy, send, paste, clipboard, restore

    def test_focus_change_cancels_both_operations(self):
        for handler in (winapi.handle_hotkey_transform, winapi.handle_hotkey_case):
            with self.subTest(handler=handler.__name__), ExitStack() as stack:
                native, copy, send, paste, clipboard, restore = self.prepare(stack)
                copy.side_effect = lambda **kwargs: (setattr(native.GetForegroundWindow, "return_value", 202)
                                                     or ("abc", set()))
                handler()
                send.assert_not_called()
                paste.assert_not_called()
                clipboard.assert_not_called()
                restore.assert_called_once()

    def test_partial_input_never_retries_full_text(self):
        for handler in (winapi.handle_hotkey_transform, winapi.handle_hotkey_case):
            with self.subTest(handler=handler.__name__), ExitStack() as stack:
                native, copy, send, paste, clipboard, restore = self.prepare(stack)
                send.side_effect = winapi.PartialInputError("partial")
                handler()
                send.assert_called_once()
                clipboard.assert_not_called()
                paste.assert_not_called()
                native.PostMessageW.assert_not_called()
                restore.assert_called_once()

    def test_original_window_is_kept_if_snapshot_takes_time(self):
        for handler in (winapi.handle_hotkey_transform, winapi.handle_hotkey_case):
            with self.subTest(handler=handler.__name__), ExitStack() as stack:
                native, copy, send, _paste, _clipboard, _restore = self.prepare(stack)

                def snapshot(_context):
                    native.GetForegroundWindow.return_value = 202
                    return True, []

                stack.enter_context(patch.object(winapi, "_save_clipboard_for_restore", side_effect=snapshot))
                handler()
                self.assertEqual(copy.call_args.kwargs["expected_hwnd"], 101)
                send.assert_not_called()

    def test_disabled_handler_does_not_touch_modifier_keys(self):
        with patch.object(winapi, "is_enabled", return_value=False), \
                patch.object(winapi, "_release_modifiers_after_hotkey") as release:
            winapi.handle_hotkey_transform()
            winapi.handle_hotkey_case()
            release.assert_not_called()

    def test_failed_clipboard_snapshot_aborts_copy(self):
        for handler in (winapi.handle_hotkey_transform, winapi.handle_hotkey_case):
            with self.subTest(handler=handler.__name__), ExitStack() as stack:
                _native, copy, send, _paste, _clipboard, _restore = self.prepare(stack)
                stack.enter_context(patch.object(winapi, "_save_clipboard_for_restore", return_value=(False, [])))
                handler()
                copy.assert_not_called()
                send.assert_not_called()

    def test_worker_start_failure_releases_lock(self):
        with patch.object(winapi, "is_enabled", return_value=True), \
                patch.object(winapi, "exit_event") as event, \
                patch.object(winapi, "_handler_lock", threading.Lock()) as lock, \
                patch.object(winapi, "_last_translate_ts", 0), \
                patch.object(winapi.threading, "Thread") as worker:
            event.is_set.return_value = False
            worker.return_value.start.side_effect = RuntimeError("no threads")
            with self.assertRaises(RuntimeError):
                winapi._invoke_text_handler(Mock())
            self.assertFalse(lock.locked())
            self.assertEqual(winapi._last_translate_ts, 0)

    def test_busy_worker_does_not_extend_debounce(self):
        lock = threading.Lock()
        lock.acquire()
        with patch.object(winapi, "is_enabled", return_value=True), \
                patch.object(winapi, "exit_event") as event, \
                patch.object(winapi, "_handler_lock", lock), \
                patch.object(winapi, "_last_translate_ts", 10):
            event.is_set.return_value = False
            winapi._invoke_text_handler(Mock())
            self.assertEqual(winapi._last_translate_ts, 10)

    def test_registration_disables_key_repeat(self):
        with patch.object(winapi, "is_enabled", return_value=True), \
                patch.object(winapi, "user32") as native:
            for update in (winapi.update_exit_hotkey_in_thread, winapi.update_translate_hotkey_in_thread,
                           winapi.update_case_hotkey_in_thread):
                update()
                self.assertTrue(native.RegisterHotKey.call_args.args[2] & winapi.MOD_NOREPEAT)


if __name__ == "__main__":
    unittest.main()
