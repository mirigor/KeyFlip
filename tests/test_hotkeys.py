import unittest

from hotkeys import hotkeys_equal, key_name_to_vk, normalize_key, normalize_modifiers


class HotkeyTests(unittest.TestCase):
    def test_unknown_keys_are_rejected(self):
        for key in ("", "bad", "VK_00", "VK_FF", "VK_ZZ", None, 42):
            with self.subTest(key=key), self.assertRaises(ValueError):
                key_name_to_vk(key)

    def test_extended_function_keys_and_aliases(self):
        self.assertEqual(key_name_to_vk("F24"), 0x87)
        self.assertEqual(normalize_key(" return "), "ENTER")
        self.assertEqual(normalize_key("VK_79"), "F10")
        self.assertEqual(normalize_key("VK_41"), "A")
        self.assertEqual(normalize_key("PAGEUP"), "PGUP")

    def test_modifier_aliases_and_malformed_values(self):
        self.assertEqual(normalize_modifiers([" control ", "CTRL", "menu", 3, None, "windows"]),
                         ["CTRL", "ALT", "WIN"])
        self.assertEqual(normalize_modifiers("CTRL"), [])

    def test_physical_keys_are_compared_including_aliases(self):
        self.assertTrue(hotkeys_equal(["CONTROL", "ALT"], "VK_0D", ["MENU", "CTRL"], "RETURN"))
        self.assertFalse(hotkeys_equal([], "UNKNOWN", [], "UNKNOWN"))


if __name__ == "__main__":
    unittest.main()
