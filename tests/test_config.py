# -*- coding: utf-8 -*-

import json
import os
import tempfile
import unittest
from unittest.mock import patch

import config


class ConfigTests(unittest.TestCase):
    def setUp(self) -> None:
        self.tmpdir = tempfile.TemporaryDirectory()
        self.old_config_file = config.CONFIG_FILE
        self.old_enabled = config.is_enabled()
        config.CONFIG_FILE = os.path.join(self.tmpdir.name, "keyflip.json")

    def tearDown(self) -> None:
        config.CONFIG_FILE = self.old_config_file
        with config._enabled_lock:  # type: ignore[attr-defined]
            config._enabled = self.old_enabled  # type: ignore[attr-defined]
        self.tmpdir.cleanup()

    def _read_disk_json(self) -> dict:
        with open(config.CONFIG_FILE, "r", encoding="utf-8") as f:
            return json.load(f)

    def test_read_json_missing_or_invalid_returns_empty_dict(self) -> None:
        self.assertEqual(config.read_json_config(), {})

        with open(config.CONFIG_FILE, "w", encoding="utf-8") as f:
            f.write("{not valid json")
        self.assertEqual(config.read_json_config(), {})

        with open(config.CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(["not", "a", "dict"], f)
        self.assertEqual(config.read_json_config(), {})

    def test_write_json_roundtrip_preserves_unicode(self) -> None:
        payload = {"enabled": True, "text": "Привет мир", "emoji": "🙂"}
        self.assertTrue(config.write_json_config(payload))
        self.assertEqual(config.read_json_config(), payload)
        self.assertFalse(os.path.exists(config.CONFIG_FILE + ".tmp"))

    def test_replace_failure_preserves_previous_config(self):
        previous = {"enabled": True, "custom": "original"}
        self.assertTrue(config.write_json_config(previous))
        with patch.object(config.os, "replace", side_effect=PermissionError("locked")):
            self.assertFalse(config.write_json_config({"enabled": False}))
        self.assertEqual(self._read_disk_json(), previous)

    def test_failed_save_rolls_back_enabled_state(self):
        config.load_config()
        with patch.object(config, "write_json_config", return_value=False):
            self.assertFalse(config.set_enabled(False))
        self.assertTrue(config.is_enabled())
        self.assertTrue(self._read_disk_json()["enabled"])

    def test_hotkeys_recover_from_invalid_nested_values(self):
        for value in ({"modifiers": "CTRL", "key": "Q"},
                      {"modifiers": [], "key": 123}, {"modifiers": [], "key": "INVALID"}):
            with self.subTest(value=value):
                config.write_json_config({"case_hotkey": value})
                self.assertEqual(config.read_case_hotkey(), config.default_case_hotkey())

    def test_writer_rejects_invalid_key_without_overwriting_config(self):
        config.load_config()
        before = config.read_json_config()
        self.assertFalse(config.write_translate_hotkey([], "INVALID"))
        self.assertEqual(config.read_json_config(), before)

    def test_modifiers_deduplicate_and_keys_normalize(self):
        self.assertTrue(config.write_case_hotkey(["ctrl", "CONTROL", None, "shift"], " vk_55 "))
        self.assertEqual(config.read_case_hotkey(), {"modifiers": ["CTRL", "SHIFT"], "key": "U"})

    def test_autorun_failure_does_not_save_enabled_flag(self):
        with patch.object(config, "create_startup_shortcut", return_value=False), \
                patch.object(config, "write_autorun_flag") as save:
            self.assertFalse(config.set_autorun_enabled(True))
            save.assert_not_called()

    def test_autorun_rolls_back_shortcut_if_config_write_fails(self):
        with patch.object(config, "create_startup_shortcut", return_value=True), \
                patch.object(config, "remove_startup_shortcut", return_value=True) as remove, \
                patch.object(config, "write_autorun_flag", return_value=False):
            self.assertFalse(config.set_autorun_enabled(True))
            remove.assert_called_once_with()

    def test_load_config_creates_defaults_and_sets_enabled(self) -> None:
        config.load_config()

        data = self._read_disk_json()
        self.assertTrue(data["enabled"])
        self.assertFalse(data["file_logging"])
        self.assertFalse(data["autorun"])
        self.assertEqual(data["exit_hotkey"], {"modifiers": [], "key": "F10"})
        self.assertEqual(data["translate_hotkey"], {"modifiers": [], "key": "F4"})
        self.assertEqual(data["case_hotkey"], {"modifiers": ["CTRL", "SHIFT"], "key": "U"})
        self.assertTrue(config.is_enabled())

    def test_load_config_preserves_existing_enabled_false(self) -> None:
        self.assertTrue(config.write_json_config({"enabled": False}))

        config.load_config()

        data = self._read_disk_json()
        self.assertFalse(data["enabled"])
        self.assertFalse(config.is_enabled())
        self.assertEqual(data["translate_hotkey"], {"modifiers": [], "key": "F4"})

    def test_flag_read_write_helpers(self) -> None:
        self.assertFalse(config.read_file_logging_flag())
        self.assertFalse(config.read_autorun_flag())

        self.assertTrue(config.write_file_logging_flag(True))
        self.assertTrue(config.write_autorun_flag(True))

        self.assertTrue(config.read_file_logging_flag())
        self.assertTrue(config.read_autorun_flag())

    def test_hotkey_defaults_when_config_value_is_malformed(self) -> None:
        self.assertTrue(
            config.write_json_config(
                {
                    "exit_hotkey": "bad",
                    "translate_hotkey": None,
                    "case_hotkey": [],
                }
            )
        )

        self.assertEqual(config.read_exit_hotkey(), {"modifiers": [], "key": "F10"})
        self.assertEqual(config.read_translate_hotkey(), {"modifiers": [], "key": "F4"})
        self.assertEqual(config.read_case_hotkey(), {"modifiers": ["CTRL", "SHIFT"], "key": "U"})

    def test_hotkey_writers_normalize_modifiers(self) -> None:
        self.assertTrue(config.write_exit_hotkey(["control", "menu", "shift", "windows", "bad"], "F8"))
        self.assertEqual(
            config.read_exit_hotkey(),
            {"modifiers": ["CTRL", "ALT", "SHIFT", "WIN"], "key": "F8"},
        )

        self.assertTrue(config.write_translate_hotkey(["ctrl", "alt"], "T"))
        self.assertEqual(config.read_translate_hotkey(), {"modifiers": ["CTRL", "ALT"], "key": "T"})

        self.assertTrue(config.write_case_hotkey(["ctrl", "shift"], "U"))
        self.assertEqual(config.read_case_hotkey(), {"modifiers": ["CTRL", "SHIFT"], "key": "U"})

    def test_set_enabled_persists_and_preserves_existing_keys(self) -> None:
        self.assertTrue(config.write_json_config({"custom": 123, "enabled": True}))

        config.set_enabled(False)

        self.assertFalse(config.is_enabled())
        self.assertEqual(self._read_disk_json(), {"custom": 123, "enabled": False})


if __name__ == "__main__":
    unittest.main()
