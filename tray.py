"""
Трэй, иконка и меню для KeyFlip.

Содержит:
- подготовку изображения для иконки
- callbacks меню (вкл/выкл, логирование, автозапуск, выборе хоткеев)
- запуск pystray Icon.run (tray_worker)
"""

import os

import win32api
from PIL import Image, ImageDraw
from pystray import Menu, MenuItem, Icon

from config import (
    ICON_ON,
    ICON_OFF,
    read_file_logging_flag,
    write_file_logging_flag,
    read_autorun_flag,
    write_autorun_flag,
    apply_autorun_setting,
    read_exit_hotkey,
    read_translate_hotkey,
    read_case_hotkey,
    write_exit_hotkey,
    write_translate_hotkey,
    write_case_hotkey,
    is_enabled,
    set_enabled,
)
from logging_setup import logger
from winapi import (
    post_register_translate,
    post_update_exit_hotkey,
    post_update_translate_hotkey,
    post_register_case,
    post_update_case_hotkey,
    capture_hotkey_and_apply_via_thread,
    exit_event,
)


# ---------------- Подготовка иконки ----------------
def prepare_tray_icon_image(enabled: bool | None = None) -> Image.Image:
    """
    Подготовить PIL.Image для иконки в трее (использовать ICON_ON/ICON_OFF если есть,
    иначе нарисовать простую заглушку).
    """
    if enabled is None:
        enabled = is_enabled()
    path = ICON_ON if enabled else ICON_OFF
    if os.path.isfile(path):
        try:
            img = Image.open(path)
            logger.debug("Tray icon: using %s", path)
            return img
        except Exception:  # noqa
            logger.exception("Tray icon: cannot open %s", path)
    # простая сгенерированная иконка
    size = (64, 64)
    bg = (76, 175, 80, 255) if enabled else (220, 53, 69, 255)
    img = Image.new("RGBA", size, bg)
    d = ImageDraw.Draw(img)
    try:
        d.rectangle((8, 16, 56, 48), outline=(255, 255, 255), width=2)
        d.text((16, 12), "KF", fill=(255, 255, 255))
    except Exception:  # noqa
        pass
    return img


# ---------------- Меню: утилиты ----------------
def enabled_menu_text(_item: MenuItem | None = None) -> str:
    """Текст пункта меню 'Вкл/Выкл' в зависимости от состояния."""
    return "❌ Выключить" if is_enabled() else "✅ Включить"


def toggle_enabled(_icon: Icon | None, _item: MenuItem | None) -> None:
    """Переключить глобальное включение/выключение приложения и обновить иконку."""
    try:
        new_state = not is_enabled()
        set_enabled(new_state)
        logger.info("Tray: toggled enabled -> %s", new_state)

        try:
            ok = post_register_translate(new_state)
            if not ok:
                logger.warning("toggle_enabled: post_register_translate returned False")
        except Exception:  # noqa
            logger.exception("toggle_enabled: failed to post register/unregister request")

        try:
            ok2 = post_register_case(new_state)
            if not ok2:
                logger.warning("toggle_enabled: post_register_case returned False")
        except Exception:  # noqa
            logger.exception("toggle_enabled: failed to post register/unregister request for case")


        # обновление иконки (если вызов пришёл из pystray)
        try:
            if _icon is not None:
                _icon.icon = prepare_tray_icon_image(new_state)
        except Exception:  # noqa
            logger.exception("Tray: failed to update icon image after toggle")
    except Exception:  # noqa
        logger.exception("toggle_enabled: exception")


def toggle_file_logging(_icon: Icon | None, _item: MenuItem | None) -> None:
    """Переключить логирование в файл (чтение/запись в конфиг)."""
    try:
        current = read_file_logging_flag()
        new = not current
        ok = write_file_logging_flag(new)
        if ok:
            logger.info("Tray: file_logging toggled -> %s", new)
        else:
            logger.warning("Tray: file_logging toggle attempted but write failed")
    except Exception:  # noqa
        logger.exception("toggle_file_logging: exception")


def toggle_autorun(_icon: Icon | None, _item: MenuItem | None) -> None:
    """Переключить автозапуск: создать/удалить ярлык и записать в конфиг."""
    try:
        current = read_autorun_flag()
        new = not current
        ok = write_autorun_flag(new)
        if ok:
            logger.info("Tray: autorun toggled -> %s", new)
            try:
                apply_autorun_setting()
            except Exception:  # noqa
                logger.exception("toggle_autorun: apply_autorun_setting failed")
        else:
            logger.warning("Tray: autorun toggle attempted but write failed")
    except Exception:  # noqa
        logger.exception("toggle_autorun: exception")


# ---------------- Хоткеи: сравнение / форматирование ----------------
def _normalize_mods(mods: list | None) -> set[str]:
    """Нормализовать список модификаторов в set строк в верхнем регистре."""
    return set((m or "").upper() for m in (mods or []))


def hotkey_lists_equal(a_mods: list | None, a_key: str, b_mods: list | None, b_key: str) -> bool:
    """Сравнить две комбинации модификаторов+клавиша (нормализуя)."""
    return _normalize_mods(a_mods) == _normalize_mods(b_mods) and ((a_key or "").upper() == (b_key or "").upper())


def is_exit_hotkey_equal(expected_mods: list[str], expected_key: str) -> bool:
    """Проверить, совпадает ли текущая exit-комбинация с ожидаемой."""
    eh = read_exit_hotkey()
    return hotkey_lists_equal(eh.get("modifiers", []) or [], eh.get("key", ""), expected_mods, expected_key)


def is_translate_hotkey_equal(expected_mods: list[str], expected_key: str) -> bool:
    """Проверить, совпадает ли текущая translate-комбинация с ожидаемой."""
    th = read_translate_hotkey()
    return hotkey_lists_equal(th.get("modifiers", []) or [], th.get("key", ""), expected_mods, expected_key)


def is_case_hotkey_equal(expected_mods: list[str], expected_key: str) -> bool:
    """Проверить, совпадает ли текущая комбинация 'регист' с ожидаемой."""
    ch = read_case_hotkey()
    return hotkey_lists_equal(ch.get("modifiers", []) or [], ch.get("key", ""), expected_mods, expected_key)

def format_exit_hotkey_display(_item: MenuItem | None = None) -> str:
    """Формат строки для отображения комбинации выхода в трее."""
    eh = read_exit_hotkey()
    mods = eh.get("modifiers") or []
    key = (eh.get("key") or "").upper()
    if not key:
        return "Выход: (нет)"
    if not mods:
        if key == "F10":
            return "Выход: F10 (по умолчанию)"
        return f"Выход: {key}"
    return f"Выход: {'+'.join(m.upper() for m in mods)}+{key}"


def format_translate_hotkey_display(_item: MenuItem | None = None) -> str:
    """Формат строки для отображения комбинации перевода в трее."""
    th = read_translate_hotkey()
    mods = th.get("modifiers") or []
    key = (th.get("key") or "").upper()
    if not key:
        return "Перевод: (нет)"
    if not mods:
        if key == "F4":
            return "Перевод: F4 (по умолчанию)"
        return f"Перевод: {key}"
    return f"Перевод: {'+'.join(m.upper() for m in mods)}+{key}"


def format_case_hotkey_display(_item: MenuItem | None = None) -> str:
    """Формат строки для отображения комбинации изменения регистра в трее."""
    ch = read_case_hotkey()
    mods = ch.get("modifiers") or []
    key = (ch.get("key") or "").upper()
    if not key:
        return "Регистр: (нет)"
    if not mods:
        if key == "U":
            return "Регистр: Ctrl+Shift+U (по умолчанию)"
        return f"Регистр: {key}"
    return f"Регистр: {'+'.join(m.upper() for m in mods)}+{key}"


# ---------------- Запись хоткеев с проверкой конфликтов ----------------
def _show_conflict_messagebox(msg: str) -> None:
    """Локальный MessageBox (попытка, без краха при ошибке)."""
    try:
        win32api.MessageBox(0, msg, "KeyFlip", 0)
    except Exception:  # noqa
        # если не удалось показать MessageBox — просто логируем
        logger.debug("MessageBox not available for: %s", msg)


def _set_hotkey_and_apply(mods: list[str], key: str, target: str) -> bool:
    """
    Установить хоткей и применить:
    - target == 'exit' -> write_exit_hotkey + post_update_exit_hotkey
    - target == 'translate' -> write_translate_hotkey + post_update_translate_hotkey
    - target == 'case' -> write_case_hotkey + post_update_case_hotkey
    """
    try:
        if target == "exit":
            other = read_translate_hotkey()
            if hotkey_lists_equal(mods, key, other.get("modifiers", []) or [], other.get("key", "")):
                _show_conflict_messagebox(
                    f"Нельзя установить комбинацию выхода {'+'.join(mods) + '+' if mods else ''}{key} — она уже используется для перевода."
                )
                return False
            other2 = read_case_hotkey()
            if hotkey_lists_equal(mods, key, other2.get("modifiers", []) or [], other2.get("key", "")):
                _show_conflict_messagebox(
                    f"Нельзя установить комбинацию регистра {'+'.join(mods) + '+' if mods else ''}{key} — она уже используется для изменения регистра."
                )
                return False
            ok = write_exit_hotkey(mods, key)
            if ok:
                post_update_exit_hotkey()
            return ok

        elif target == "translate":
            other = read_exit_hotkey()
            if hotkey_lists_equal(mods, key, other.get("modifiers", []) or [], other.get("key", "")):
                _show_conflict_messagebox(
                    f"Нельзя установить комбинацию перевода {'+'.join(mods) + '+' if mods else ''}{key} — она уже используется для выхода."
                )
                return False
            other2 = read_case_hotkey()
            if hotkey_lists_equal(mods, key, other2.get("modifiers", []) or [], other2.get("key", "")):
                _show_conflict_messagebox(
                    f"Нельзя установить комбинацию регистра {'+'.join(mods) + '+' if mods else ''}{key} — она уже используется для изменения регистра."
                )
                return False
            ok = write_translate_hotkey(mods, key)
            if ok:
                post_update_translate_hotkey()
            return ok

        elif target == "case":
            other = read_exit_hotkey()
            if hotkey_lists_equal(mods, key, other.get("modifiers", []) or [], other.get("key", "")):
                _show_conflict_messagebox(
                    f"Нельзя установить комбинацию регистра {'+'.join(mods) + '+' if mods else ''}{key} — она уже используется для выхода."
                )
                return False
            other2 = read_translate_hotkey()
            if hotkey_lists_equal(mods, key, other2.get("modifiers", []) or [], other2.get("key", "")):
                _show_conflict_messagebox(
                    f"Нельзя установить комбинацию регистра {'+'.join(mods) + '+' if mods else ''}{key} — она уже используется для перевода."
                )
                return False
            ok = write_case_hotkey(mods, key)
            if ok:
                post_update_case_hotkey()
            return ok
        else:
            logger.error("_set_hotkey_and_apply: unknown target %r", target)
            return False
    except Exception:  # noqa
        logger.exception("_set_hotkey_and_apply: exception")
        return False


def _make_preset_setter(mods: list[str], key: str, target: str) -> object:
    """
    Фабрика для создания пресетных колбэков меню (уменьшает повторение).
    Возвращает функцию-подходящую для pystray (callback).
    """

    def _fn(_icon: Icon | None, _item: MenuItem | None) -> None:
        try:
            _set_hotkey_and_apply(mods, key, target)
        except Exception:  # noqa
            logger.exception("preset setter exception for %s %s", mods, key)

    return _fn


# ---------------- Пресеты и capture ----------------
menu_set_exit_f10 = _make_preset_setter([], "F10", "exit")
menu_set_exit_ctrl_q = _make_preset_setter(["CTRL"], "Q", "exit")
menu_set_exit_ctrl_alt_x = _make_preset_setter(["CTRL", "ALT"], "X", "exit")

menu_set_translate_f4 = _make_preset_setter([], "F4", "translate")
menu_set_translate_ctrl_alt_t = _make_preset_setter(["CTRL", "ALT"], "T", "translate")
menu_set_translate_ctrl_shift_y = _make_preset_setter(["CTRL", "SHIFT"], "Y", "translate")

menu_set_case_ctrl_alt_u = _make_preset_setter(["CTRL", "SHIFT"], "U", "case")

def menu_set_exit_custom_capture(_icon: Icon | None, _item: MenuItem | None) -> None:
    """Поймать комбинацию для выхода в отдельном потоке."""
    try:
        capture_hotkey_and_apply_via_thread("exit")
    except Exception:  # noqa
        logger.exception("menu_set_exit_custom_capture: исключение при запуске capture thread")


def menu_set_translate_custom_capture(_icon: Icon | None, _item: MenuItem | None) -> None:
    """Поймать комбинацию для перевода в отдельном потоке."""
    try:
        capture_hotkey_and_apply_via_thread("translate")
    except Exception:  # noqa
        logger.exception("menu_set_translate_custom_capture: исключение при запуске capture thread")


def menu_set_case_custom_capture(_icon, _item) -> None:
    """Поймать комбинацию для 'регистра' в отдельном потоке."""
    try:
        capture_hotkey_and_apply_via_thread("case")
    except Exception:  # noqa
        logger.exception("menu_set_case_custom_capture: exception starting capture thread")


# ---------------- Выход через трей ----------------
def on_exit(_icon: Icon | None = None, _item: MenuItem | None = None) -> None:
    """
    Завершить работу приложения (через трей).
    Устанавливаем exit_event и останавливаем иконку.
    """
    logger.info("Выход запрошен (через трей/горячую клавишу)")
    try:
        exit_event.set()
    except Exception:  # noqa
        logger.exception("on_exit: не удалось установить exit_event")
    try:
        if _icon:
            _icon.stop()
    except Exception:  # noqa
        logger.exception("on_exit: ошибка при остановке иконки")


# ---------------- Tray worker ----------------
def tray_worker() -> None:
    """
    Запустить pystray icon + меню (в отдельном потоке).
    Этот метод предназначен для запуска как daemon-поток.
    """
    img = prepare_tray_icon_image()
    settings_menu = Menu(
        MenuItem(format_translate_hotkey_display, None, enabled=False),
        MenuItem(format_exit_hotkey_display, None, enabled=False),
        MenuItem(format_case_hotkey_display, None, enabled=False),
        Menu.SEPARATOR,
        MenuItem("Логирование в файл", toggle_file_logging, checked=lambda item: read_file_logging_flag()),
        MenuItem("Автозапуск при старте Windows", toggle_autorun, checked=lambda item: read_autorun_flag()),
        MenuItem(
            "Комбинация выхода",
            Menu(
                MenuItem("F10 (по умолчанию)", menu_set_exit_f10,
                         checked=lambda item: is_exit_hotkey_equal([], "F10")),
                MenuItem("Ctrl+Q", menu_set_exit_ctrl_q,
                         checked=lambda item: is_exit_hotkey_equal(["CTRL"], "Q")),
                MenuItem("Ctrl+Alt+X", menu_set_exit_ctrl_alt_x,
                         checked=lambda item: is_exit_hotkey_equal(["CTRL", "ALT"], "X")),
                MenuItem("Ввести свою комбинацию...", menu_set_exit_custom_capture,
                         checked=lambda item: not (
                                 is_exit_hotkey_equal([], "F10")
                                 or is_exit_hotkey_equal(["CTRL"], "Q")
                                 or is_exit_hotkey_equal(["CTRL", "ALT"], "X")
                         )),
            )
        ),
        Menu.SEPARATOR,
        MenuItem(
            "Комбинация перевода",
            Menu(
                MenuItem("F4 (по умолчанию)", menu_set_translate_f4,
                         checked=lambda item: is_translate_hotkey_equal([], "F4")),
                MenuItem("Ctrl+Alt+T", menu_set_translate_ctrl_alt_t,
                         checked=lambda item: is_translate_hotkey_equal(["CTRL", "ALT"], "T")),
                MenuItem("Ctrl+Shift+Y", menu_set_translate_ctrl_shift_y,
                         checked=lambda item: is_translate_hotkey_equal(["CTRL", "SHIFT"], "Y")),
                MenuItem("Ввести свою комбинацию...", menu_set_translate_custom_capture,
                         checked=lambda item: not (
                                 is_translate_hotkey_equal([], "F4")
                                 or is_translate_hotkey_equal(["CTRL", "ALT"], "T")
                                 or is_translate_hotkey_equal(["CTRL", "SHIFT"], "Y")
                         )),
            )
        ),
        Menu.SEPARATOR,
        MenuItem(
            "Комбинация регистра",
            Menu(
                MenuItem("Ctrl+Shift+U (по умолчанию)", menu_set_case_ctrl_alt_u,
                         checked=lambda item: is_case_hotkey_equal(["CTRL", "SHIFT"], "U")),
                MenuItem("Ввести свою комбинацию...", menu_set_case_custom_capture,
                         checked=lambda item: not is_case_hotkey_equal(["CTRL", "SHIFT"], "U")),
            )
        ),
    )
    menu = Menu(
        MenuItem(enabled_menu_text, toggle_enabled),
        MenuItem("Настройки", settings_menu),
        MenuItem("Выход", on_exit),
    )
    icon = Icon("KeyFlip", img, "KeyFlip", menu=menu)
    try:
        icon.icon = prepare_tray_icon_image(is_enabled())
    except Exception:  # noqa
        logger.exception("tray_worker: failed to set initial icon")
    try:
        logger.debug("tray_worker: запуск иконки")
        icon.run()
    except Exception:  # noqa
        logger.exception("tray_worker: exception while running icon")
