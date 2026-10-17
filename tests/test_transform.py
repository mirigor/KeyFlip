# -*- coding: utf-8 -*-

import unittest
from unittest.mock import patch

from transform import (
    EN_TO_RU,
    RU_TO_EN,
    change_case_by_logic,
    transform_text_by_keyboard_layout_based_on_hkl,
)


class TransformTests(unittest.TestCase):
    def test_en_to_ru_by_english_layout(self) -> None:
        self.assertEqual(transform_text_by_keyboard_layout_based_on_hkl("ghbdtn", 0x0409), "привет")

    def test_ru_to_en_by_russian_layout(self) -> None:
        self.assertEqual(transform_text_by_keyboard_layout_based_on_hkl("привет", 0x0419), "ghbdtn")

    def test_words_with_uppercase_are_converted(self) -> None:
        self.assertEqual(transform_text_by_keyboard_layout_based_on_hkl("Ghbdtn vbh", 0x0409), "Привет мир")
        self.assertEqual(transform_text_by_keyboard_layout_based_on_hkl("GHBDTN VBH", 0x0409), "ПРИВЕТ МИР")
        self.assertEqual(transform_text_by_keyboard_layout_based_on_hkl("Привет мир", 0x0419), "Ghbdtn vbh")
        self.assertEqual(transform_text_by_keyboard_layout_based_on_hkl("ПРИВЕТ МИР", 0x0419), "GHBDTN VBH")

    def test_mixed_text_translates_chars_in_both_layouts(self) -> None:
        self.assertEqual(transform_text_by_keyboard_layout_based_on_hkl("hello мир", 0x0409), "руддщ vbh")
        self.assertEqual(transform_text_by_keyboard_layout_based_on_hkl("руддщ vbh", 0x0419), "hello мир")

    def test_punctuation_mapping(self) -> None:
        self.assertEqual(transform_text_by_keyboard_layout_based_on_hkl("[]{};:'\"`~", 0x0409), "хъХЪжЖэЭёЁ")
        self.assertEqual(transform_text_by_keyboard_layout_based_on_hkl("хъХЪжЖэЭёЁ", 0x0419), "[]{};:'\"`~")
        self.assertEqual(transform_text_by_keyboard_layout_based_on_hkl(".,/?", 0x0409), "юб.,")
        self.assertEqual(transform_text_by_keyboard_layout_based_on_hkl("юб.,", 0x0419), ".,/?")

    def test_digits_whitespace_and_unknown_symbols_are_stable(self) -> None:
        self.assertEqual(transform_text_by_keyboard_layout_based_on_hkl("123  \t🙂", 0x0409), "123  \t🙂")
        self.assertEqual(transform_text_by_keyboard_layout_based_on_hkl("ქართული 🙂", 0x0419), "ქართული 🙂")

    def test_unknown_or_invalid_hkl_defaults_to_en_to_ru(self) -> None:
        self.assertEqual(transform_text_by_keyboard_layout_based_on_hkl("ghbdtn", 0), "привет")
        with patch("transform.logger.exception"):
            self.assertEqual(transform_text_by_keyboard_layout_based_on_hkl("ghbdtn", None), "привет")  # type: ignore[arg-type]

    def test_mapping_tables_have_expected_reverse_entries(self) -> None:
        self.assertEqual(EN_TO_RU["q"], "й")
        self.assertEqual(EN_TO_RU["Q"], "Й")
        self.assertEqual(RU_TO_EN["й"], "q")
        self.assertEqual(RU_TO_EN["Й"], "Q")
        self.assertEqual(RU_TO_EN["ё"], "`")
        self.assertEqual(RU_TO_EN["Ё"], "~")

    def test_change_case_logic_basic(self) -> None:
        self.assertEqual(change_case_by_logic("привет"), "ПРИВЕТ")
        self.assertEqual(change_case_by_logic("ПРИВЕТ"), "привет")
        self.assertEqual(change_case_by_logic("ПрИвЕт"), "привет")
        self.assertEqual(change_case_by_logic("hello, мир!"), "HELLO, МИР!")
        self.assertEqual(change_case_by_logic("HELLO, МИР!"), "hello, мир!")

    def test_change_case_keeps_non_letters_and_empty_text(self) -> None:
        self.assertEqual(change_case_by_logic(""), "")
        self.assertEqual(change_case_by_logic("123!? "), "123!? ")
        self.assertEqual(change_case_by_logic("ё"), "Ё")
        self.assertEqual(change_case_by_logic("Ё"), "ё")


if __name__ == "__main__":
    unittest.main()
