"""Действия панели и контекстного меню KeyFlip."""

import win32api

from config import (
    is_enabled,
    read_autorun_flag,
    read_case_hotkey,
    read_exit_hotkey,
    read_file_logging_flag,
    read_translate_hotkey,
    set_autorun_enabled,
    set_enabled,
    write_case_hotkey,
    write_exit_hotkey,
    write_file_logging_flag,
    write_translate_hotkey,
)
from hotkeys import hotkeys_equal
from logging_setup import logger, set_file_logging_enabled
from winapi import (
    exit_event,
    post_register_case,
    post_register_translate,
    post_update_case_hotkey,
    post_update_exit_hotkey,
    post_update_translate_hotkey,
)


hotkey_lists_equal = hotkeys_equal

HOTKEY_ACTIONS = {
    "translate": ("смены раскладки", read_translate_hotkey, write_translate_hotkey, post_update_translate_hotkey),
    "case": ("смены регистра", read_case_hotkey, write_case_hotkey, post_update_case_hotkey),
    "exit": ("выхода", read_exit_hotkey, write_exit_hotkey, post_update_exit_hotkey),
}


def _show_error(message: str) -> None:
    try:
        win32api.MessageBox(0, message, "KeyFlip", 0)
    except Exception:
        logger.exception("Не удалось показать сообщение об ошибке: %s", message)


def toggle_enabled() -> None:
    enabled = not is_enabled()
    if not set_enabled(enabled):
        _show_error("Не удалось сохранить состояние KeyFlip. Проверьте доступ к настройкам.")
        return
    for post_register in (post_register_translate, post_register_case):
        if not post_register(enabled):
            logger.warning("Не удалось обновить регистрацию горячих клавиш: %s", post_register.__name__)


def toggle_file_logging() -> None:
    enabled = not read_file_logging_flag()
    if write_file_logging_flag(enabled):
        set_file_logging_enabled(enabled)
    else:
        _show_error("Не удалось сохранить настройку логирования.")


def toggle_autorun() -> None:
    if not set_autorun_enabled(not read_autorun_flag()):
        _show_error("Не удалось изменить автозапуск. Проверьте доступ к папке автозагрузки и настройкам.")


def _set_hotkey_and_apply(modifiers: list[str], key: str, target: str) -> bool:
    """Сохранить сочетание после проверки конфликтов и обновить WinAPI-регистрацию."""
    if target not in HOTKEY_ACTIONS:
        logger.error("Неизвестное действие горячей клавиши: %r", target)
        return False

    for other_target, (label, read, _write, _update) in HOTKEY_ACTIONS.items():
        if other_target == target:
            continue
        other = read()
        if hotkeys_equal(modifiers, key, other.get("modifiers", []), other.get("key", "")):
            _show_error(f"Эта комбинация уже используется для {label}.")
            return False

    _label, _read, write, update = HOTKEY_ACTIONS[target]
    if not write(modifiers, key):
        _show_error("Не удалось сохранить комбинацию. Проверьте доступ к настройкам.")
        return False
    if not update():
        logger.warning("Комбинация сохранена, но не удалось обновить её регистрацию: %s", target)
    return True


def on_exit() -> None:
    logger.info("Выход запрошен через трей или горячую клавишу")
    exit_event.set()
