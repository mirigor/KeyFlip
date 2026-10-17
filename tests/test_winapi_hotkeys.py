# -*- coding: utf-8 -*-

import os
import sys
import tempfile
import unittest
import logging
from unittest.mock import patch

if sys.platform != "win32":
    raise unittest.SkipTest("winapi hotkey helpers are Windows-only")

import config as _config

logging.getLogger("KeyFlip").disabled = True
_ORIGINAL_CONFIG_FILE = _config.CONFIG_FILE
_config.CONFIG_FILE = os.path.join(tempfile.gettempdir(), "keyflip-test-no-file.json")
try:
    import winapi
finally:
    _config.CONFIG_FILE = _ORIGINAL_CONFIG_FILE


class WinapiHotkeyHelperTests(unittest.TestCase):
    def test_key_name_to_vk_for_common_keys(self) -> None:
        self.assertEqual(winapi.key_name_to_vk("F4"), 0x73)
        self.assertEqual(winapi.key_name_to_vk("esc"), 0x1B)
        self.assertEqual(winapi.key_name_to_vk("space"), 0x20)
        self.assertEqual(winapi.key_name_to_vk("a"), ord("A"))
        self.assertEqual(winapi.key_name_to_vk("Z"), ord("Z"))
        self.assertEqual(winapi.key_name_to_vk("7"), ord("7"))
        self.assertEqual(winapi.key_name_to_vk("VK_7A"), 0x7A)

    def test_modifiers_list_to_mask_accepts_aliases_and_ignores_unknown_values(self) -> None:
        mask = winapi.modifiers_list_to_mask(["ctrl", "CONTROL", "ALT", "shift", "WIN", "bad"])

        self.assertEqual(
            mask,
            winapi.MOD_CONTROL | winapi.MOD_ALT | winapi.MOD_SHIFT | winapi.MOD_WIN,
        )

    def test_hotkey_tuple_from_config(self) -> None:
        self.assertEqual(
            winapi.hotkey_tuple_from_config(["CTRL", "SHIFT"], "U"),
            (winapi.MOD_CONTROL | winapi.MOD_SHIFT, ord("U")),
        )

    def test_hotkeys_conflict(self) -> None:
        self.assertTrue(winapi.hotkeys_conflict(winapi.MOD_CONTROL, ord("Q"), winapi.MOD_CONTROL, ord("Q")))
        self.assertFalse(winapi.hotkeys_conflict(winapi.MOD_CONTROL, ord("Q"), winapi.MOD_ALT, ord("Q")))
        self.assertFalse(winapi.hotkeys_conflict(winapi.MOD_CONTROL, ord("Q"), winapi.MOD_CONTROL, ord("W")))

    def test_hotkey_lists_equal_ignores_modifier_order_and_case(self) -> None:
        self.assertTrue(winapi.hotkey_lists_equal(["ctrl", "SHIFT"], "u", ["SHIFT", "CTRL"], "U"))
        self.assertFalse(winapi.hotkey_lists_equal(["CTRL"], "U", ["CTRL", "SHIFT"], "U"))
        self.assertFalse(winapi.hotkey_lists_equal(["CTRL"], "U", ["CTRL"], "Y"))

    def test_post_helpers_return_false_when_hotkey_thread_is_not_ready(self) -> None:
        old_thread_id = winapi.HOTKEY_THREAD_ID
        try:
            winapi.HOTKEY_THREAD_ID = 0
            with patch("winapi.logger.warning"):
                self.assertFalse(winapi.post_register_translate(True))
                self.assertFalse(winapi.post_register_case(True))
                self.assertFalse(winapi.post_update_exit_hotkey())
                self.assertFalse(winapi.post_update_translate_hotkey())
                self.assertFalse(winapi.post_update_case_hotkey())
        finally:
            winapi.HOTKEY_THREAD_ID = old_thread_id


if __name__ == "__main__":
    unittest.main()
