"""
WinAPI-слой для KeyFlip.

Содержит:
- ctypes / SendInput / клавиатурные утилиты
- чтение активного окна
- безопасное чтение выделения (через Ctrl+C / Ctrl+Insert + sentinel fallback)
- low-level keyboard hook для захвата комбинации
- регистрация глобальных hotkey'ов (RegisterHotKey) и Windows message loop для них
- обработчик translate (вызов transform и вставка через SendInput)
- обработчик case (изменение регистра) — добавлено
"""

import ctypes
import threading
import time
import tkinter as tk
import uuid
from collections.abc import Callable
from ctypes import wintypes

import psutil
import pyperclip
import win32api
import win32con

from config import (
    read_translate_hotkey,
    read_exit_hotkey,
    read_case_hotkey,
    write_translate_hotkey,
    write_exit_hotkey,
    write_case_hotkey,
    is_enabled,
)
from logging_setup import logger
from transform import transform_text_by_keyboard_layout_based_on_hkl, change_case_by_logic

# ---------------- ctypes helpers ----------------
user32 = ctypes.windll.user32
kernel32 = ctypes.windll.kernel32

ULONG_PTR = wintypes.WPARAM  # alias для совместимости

# ---------------- Constants ----------------
INPUT_KEYBOARD = 1
KEYEVENTF_UNICODE = 0x0004
KEYEVENTF_KEYUP = 0x0002

# Общие VK
VK_CONTROL = 0x11
VK_MENU = 0x12
VK_SHIFT = 0x10
VK_DELETE = 0x2E
VK_INSERT = 0x2D
VK_ESCAPE = 0x1B
VK_RETURN = 0x0D
VK_TAB = 0x09

# Для регистратора сочетаний
WH_KEYBOARD_LL = 13
WM_KEYDOWN = 0x0100
WM_SYSKEYDOWN = 0x0104
WM_KEYUP = 0x0101
WM_SYSKEYUP = 0x0105

# Hotkey ids и сообщения
HOTKEY_ID_TRANSLATE = 1
HOTKEY_ID_EXIT = 2
HOTKEY_ID_CASE = 3
MOD_NONE = 0

WM_USER = 0x0400
MSG_REGISTER_TRANSLATE = WM_USER + 1
MSG_UNREGISTER_TRANSLATE = WM_USER + 2
MSG_UPDATE_EXIT_HOTKEY = WM_USER + 3
MSG_UPDATE_TRANSLATE_HOTKEY = WM_USER + 4
MSG_REGISTER_CASE = WM_USER + 5
MSG_UNREGISTER_CASE = WM_USER + 6
MSG_UPDATE_CASE_HOTKEY = WM_USER + 7

HOTKEY_THREAD_ID = 0

# Side-specific VK constants (для safe_copy и hook)
VK_LSHIFT = 0xA0
VK_RSHIFT = 0xA1
VK_LCONTROL = 0xA2
VK_RCONTROL = 0xA3
VK_LMENU = 0xA4
VK_RMENU = 0xA5
VK_LWIN = 0x5B
VK_RWIN = 0x5C


# MOD masks для RegisterHotKey
MOD_ALT = 0x0001
MOD_CONTROL = 0x0002
MOD_SHIFT = 0x0004
MOD_WIN = 0x0008


# ---------------- SendInput structures ----------------
class KEYBDINPUT(ctypes.Structure):
    _fields_ = [
        ("wVk", wintypes.WORD),
        ("wScan", wintypes.WORD),
        ("dwFlags", wintypes.DWORD),
        ("time", wintypes.DWORD),
        ("dwExtraInfo", ULONG_PTR),
    ]


class InputUnion(ctypes.Union):
    _fields_ = [("ki", KEYBDINPUT), ("padding", wintypes.ULONG * 8)]


class INPUT(ctypes.Structure):
    _anonymous_ = ("union",)
    _fields_ = [("type", wintypes.DWORD), ("union", InputUnion)]


SendInput = user32.SendInput
SendInput.argtypes = (wintypes.UINT, ctypes.POINTER(INPUT), ctypes.c_int)
SendInput.restype = wintypes.UINT

# ---------------- VK map for readable names ----------------
VK_MAP = {
    "F1": 0x70, "F2": 0x71, "F3": 0x72, "F4": 0x73, "F5": 0x74, "F6": 0x75,
    "F7": 0x76, "F8": 0x77, "F9": 0x78, "F10": 0x79, "F11": 0x7A, "F12": 0x7B,
    "ESC": 0x1B, "TAB": 0x09, "ENTER": 0x0D, "RETURN": 0x0D, "SPACE": 0x20,
    "INSERT": 0x2D, "DELETE": 0x2E, "HOME": 0x24, "END": 0x23, "PGUP": 0x21, "PGDN": 0x22,
    "LEFT": 0x25, "UP": 0x26, "RIGHT": 0x27, "DOWN": 0x28,
}

# ---------------- Public primitives ----------------
exit_event = threading.Event()

# exit handlers (optional callbacks to be called when exit hotkey pressed)
_exit_handlers: list[Callable[[], None]] = []


def register_exit_handler(cb: Callable[[], None]) -> None:
    """Зарегистрировать callback для обработки выхода (например tray.on_exit)."""
    try:
        if cb not in _exit_handlers:
            _exit_handlers.append(cb)
    except Exception:  # noqa
        logger.exception("register_exit_handler failed")


def _invoke_exit_handlers() -> None:
    """Вызвать все зарегистрированные обработчики выхода."""
    for cb in list(_exit_handlers):
        try:
            cb()
        except Exception:  # noqa
            logger.exception("exit handler raised exception")


# ---------------- Helpers: low-level key emulation ----------------
def _key_down(vk: int) -> None:
    try:
        user32.keybd_event(vk, 0, 0, 0)
    except Exception:  # noqa
        logger.exception("_key_down failed for vk=%s", vk)


def _key_up(vk: int) -> None:
    try:
        user32.keybd_event(vk, 0, KEYEVENTF_KEYUP, 0)
    except Exception:  # noqa
        logger.exception("_key_up failed for vk=%s", vk)


def send_ctrl_c() -> None:
    """Послать Ctrl+C клавишами (быстрая комбинация)."""
    _key_down(VK_CONTROL)
    _key_down(ord('C'))
    time.sleep(0.015)
    _key_up(ord('C'))
    _key_up(VK_CONTROL)


def send_ctrl_v() -> None:
    """Послать Ctrl+V клавишами."""
    _key_down(VK_CONTROL)
    _key_down(ord('V'))
    time.sleep(0.015)
    _key_up(ord('V'))
    _key_up(VK_CONTROL)


def send_delete() -> None:
    """Послать Delete клавишей."""
    _key_down(VK_DELETE)
    time.sleep(0.01)
    _key_up(VK_DELETE)



def _append_vk_press(inputs: list, vk: int) -> None:
    """Добавить в inputs нажатие виртуальной клавиши vk (down + up)."""
    try:
        ki_down = KEYBDINPUT(vk, 0, 0, 0, 0)
        inputs.append(INPUT(INPUT_KEYBOARD, InputUnion(ki=ki_down)))
        ki_up = KEYBDINPUT(vk, 0, KEYEVENTF_KEYUP, 0, 0)
        inputs.append(INPUT(INPUT_KEYBOARD, InputUnion(ki=ki_up)))
    except Exception:
        pass


def _append_shifted_enter(inputs: list) -> None:
    """Добавить в inputs комбинацию Shift+Enter (SHIFT down, ENTER down/up, SHIFT up)."""
    try:
        # SHIFT down
        ki_shift_down = KEYBDINPUT(VK_SHIFT, 0, 0, 0, 0)
        inputs.append(INPUT(INPUT_KEYBOARD, InputUnion(ki=ki_shift_down)))
        # ENTER down/up
        ki_enter_down = KEYBDINPUT(VK_RETURN, 0, 0, 0, 0)
        inputs.append(INPUT(INPUT_KEYBOARD, InputUnion(ki=ki_enter_down)))
        ki_enter_up = KEYBDINPUT(VK_RETURN, 0, KEYEVENTF_KEYUP, 0, 0)
        inputs.append(INPUT(INPUT_KEYBOARD, InputUnion(ki=ki_enter_up)))
        # SHIFT up
        ki_shift_up = KEYBDINPUT(VK_SHIFT, 0, KEYEVENTF_KEYUP, 0, 0)
        inputs.append(INPUT(INPUT_KEYBOARD, InputUnion(ki=ki_shift_up)))
    except Exception:
        pass


def send_unicode_via_sendinput(text: str, delay_between_keys: float = 0.001) -> None:
    """
    Вставить текст через SendInput (UNICODE), но во всех случаях заменять переносы строк
    на Shift+Enter, чтобы избежать отправки сообщений (например в Telegram).
    Табуляция вставляется через VK_TAB, остальные символы — через KEYEVENTF_UNICODE.
    """
    if not text:
        return
    try:
        inputs: list[INPUT] = []
        i = 0
        L = len(text)
        used_shift_enter = False
        while i < L:
            ch = text[i]
            # Обработка CRLF / CR / LF -> всегда SHIFT+ENTER
            if ch == '\r' or ch == '\n':
                # если последовательность CRLF — пропускаем второй символ
                if ch == '\r' and i + 1 < L and text[i + 1] == '\n':
                    i += 1  # пропустить LF, обработаем Enter один раз
                _append_shifted_enter(inputs)
                used_shift_enter = True
                i += 1
                continue

            # Табуляция -> VK_TAB
            if ch == '\t':
                _append_vk_press(inputs, VK_TAB)
                i += 1
                continue

            # По-умолчанию — Unicode-символ через KEYEVENTF_UNICODE
            code = ord(ch)
            try:
                ki_down = KEYBDINPUT(0, code, KEYEVENTF_UNICODE, 0, 0)
                inputs.append(INPUT(INPUT_KEYBOARD, InputUnion(ki=ki_down)))
                ki_up = KEYBDINPUT(0, code, KEYEVENTF_UNICODE | KEYEVENTF_KEYUP, 0, 0)
                inputs.append(INPUT(INPUT_KEYBOARD, InputUnion(ki=ki_up)))
            except Exception:
                # fallback: постим WM_CHAR в окно (в редких случаях)
                try:
                    hwnd_fore = user32.GetForegroundWindow()
                    if hwnd_fore:
                        user32.PostMessageW(hwnd_fore, win32con.WM_CHAR, code, 0)
                except Exception:
                    pass
            i += 1

        # отправляем пакет
        n = len(inputs)
        if n == 0:
            return
        arr_type = INPUT * n
        arr = arr_type(*inputs)
        p = ctypes.pointer(arr[0])
        sent = SendInput(n, p, ctypes.sizeof(INPUT))
        if sent != n:
            logger.warning("send_unicode_via_sendinput: SendInput sent %d of %d events", sent, n)
        if used_shift_enter:
            logger.debug("send_unicode_via_sendinput: использован Shift+Enter для переносов строк")
        if delay_between_keys > 0:
            time.sleep(delay_between_keys)
    except Exception:  # noqa
        logger.exception("send_unicode_via_sendinput: exception")


# ---------------- Active window info ----------------
def get_active_window_info() -> tuple[str, int, str | None]:
    """
    Вернуть (title, pid, proc_name) активного окна; при ошибке вернуть заглушки.
    """
    try:
        hwnd = user32.GetForegroundWindow()
    except Exception:  # noqa
        return "<unknown>", 0, None

    try:
        length = user32.GetWindowTextLengthW(hwnd)
        buff = ctypes.create_unicode_buffer(length + 1)
        user32.GetWindowTextW(hwnd, buff, length + 1)
        title = buff.value
    except Exception:  # noqa
        title = "<unknown>"

    pid = 0
    proc_name: str | None = None
    try:
        pid_c = ctypes.c_ulong()
        user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid_c))
        pid = int(pid_c.value)
        if psutil:
            try:
                proc = psutil.Process(pid)
                proc_name = proc.name()
            except Exception:  # noqa
                proc_name = None
    except Exception:  # noqa
        pass
    return title, pid, proc_name


# ---------------- Clipboard helper (safe copy) ----------------
def safe_copy_from_selection(timeout_per_attempt: float = 0.6, max_attempts: int = 2) -> tuple[str, set[int]]:
    """
    Осторожно скопировать выделение (попытки Ctrl+C, Ctrl+Insert).
    Возвращает (строка, set_of_pressed_side_specific_vk).
    """
    title, pid, proc_name = get_active_window_info()
    logger.debug("safe_copy: start. active window: %r pid=%s proc=%r", title, pid, proc_name)
    try:
        initial_hwnd = user32.GetForegroundWindow()
    except Exception as e:  # noqa
        logger.debug("safe_copy: GetForegroundWindow failed: %s", e)
        initial_hwnd = None

    def foreground_changed() -> bool:
        try:
            cur_hwnd = user32.GetForegroundWindow()
            return initial_hwnd is not None and cur_hwnd is not None and cur_hwnd != initial_hwnd
        except Exception:  # noqa
            return False

    try:
        old_seq = user32.GetClipboardSequenceNumber()
    except Exception as e:  # noqa
        logger.debug("safe_copy: GetClipboardSequenceNumber failed: %s", e)
        old_seq = None

    last_exception = None

    # helper: detect currently pressed side-specific modifier VKs
    def _detect_pressed_modifier_vks() -> set[int]:
        pressed: set[int] = set()
        try:
            for vk in (VK_LCONTROL, VK_RCONTROL, VK_LMENU, VK_RMENU, VK_LSHIFT, VK_RSHIFT, VK_LWIN, VK_RWIN):
                try:
                    if user32.GetAsyncKeyState(vk) & 0x8000:
                        pressed.add(vk)
                except Exception:  # noqa
                    pass
        except Exception:  # noqa
            pass
        return pressed

    def _release_vks(vks: set[int]) -> None:
        try:
            for vk in vks:
                try:
                    _key_up(vk)
                except Exception:  # noqa
                    pass
        except Exception:  # noqa
            pass

    pressed_vks = _detect_pressed_modifier_vks()
    if pressed_vks:
        logger.debug("safe_copy: temporarily releasing modifiers vks=%r", pressed_vks)
        _release_vks(pressed_vks)

    try:
        if old_seq is not None:
            for attempt in range(1, max_attempts + 1):
                if foreground_changed():
                    logger.debug("safe_copy: foreground changed before Ctrl+C -> abort")
                    return "", pressed_vks
                try:
                    send_ctrl_c()
                    logger.debug("safe_copy: sent ctrl+c (attempt %d)", attempt)
                except Exception as e:  # noqa
                    logger.exception("safe_copy: ctrl+c exception: %s", e)
                    last_exception = e

                t0 = time.time()
                changed = False
                while time.time() - t0 < timeout_per_attempt:
                    if foreground_changed():
                        logger.debug("safe_copy: foreground changed during wait after Ctrl+C -> abort")
                        return "", pressed_vks
                    try:
                        seq = user32.GetClipboardSequenceNumber()
                    except Exception:  # noqa
                        seq = None
                    if seq is not None and seq != old_seq:
                        changed = True
                        break
                    time.sleep(0.02)
                if changed:
                    try:
                        pasted = pyperclip.paste()
                        if pasted == "":
                            logger.debug("safe_copy: clipboard changed but empty -> treat as no selection")
                            return "", pressed_vks
                        logger.debug("safe_copy: buffer changed after ctrl+c (len=%d)", len(pasted))
                        return pasted, pressed_vks
                    except Exception as e:  # noqa
                        logger.exception("safe_copy: paste() after ctrl+c failed: %s", e)
                        last_exception = e

                # fallback Ctrl+Insert
                if foreground_changed():
                    logger.debug("safe_copy: foreground changed before Ctrl+Insert -> abort")
                    return "", pressed_vks
                try:
                    _key_down(VK_CONTROL)
                    _key_down(VK_INSERT)
                    time.sleep(0.01)
                    _key_up(VK_INSERT)
                    _key_up(VK_CONTROL)
                    logger.debug("safe_copy: sent ctrl+insert (attempt %d)", attempt)
                except Exception as e:  # noqa
                    logger.exception("safe_copy: ctrl+insert exception: %s", e)
                    last_exception = e

                t0 = time.time()
                changed = False
                while time.time() - t0 < timeout_per_attempt:
                    if foreground_changed():
                        logger.debug("safe_copy: foreground changed during wait after Ctrl+Insert -> abort")
                        return "", pressed_vks
                    try:
                        seq = user32.GetClipboardSequenceNumber()
                    except Exception:  # noqa
                        seq = None
                    if seq is not None and seq != old_seq:
                        changed = True
                        break
                    time.sleep(0.02)
                if changed:
                    try:
                        pasted = pyperclip.paste()
                        if pasted == "":
                            logger.debug("safe_copy: clipboard changed after insert but empty -> no selection")
                            return "", pressed_vks
                        logger.debug("safe_copy: buffer changed after ctrl+insert (len=%d)", len(pasted))
                        return pasted, pressed_vks
                    except Exception as e:  # noqa
                        logger.exception("safe_copy: paste() after ctrl+insert failed: %s", e)
                        last_exception = e

                logger.debug("safe_copy: no clipboard change after attempt %d", attempt)
                break

            logger.debug("safe_copy: sequence not changed -> no selection")
            return "", pressed_vks

        # ---- fallback sentinel approach ----
        if foreground_changed():
            logger.debug("safe_copy: foreground changed before fallback -> abort")
            return "", pressed_vks
        sentinel = f"__KEYFLIP_SENTINEL__{uuid.uuid4()}__"
        try:
            pyperclip.copy(sentinel)
            logger.debug("safe_copy: sentinel written for fallback")
        except Exception as e:  # noqa
            logger.exception("safe_copy: cannot write sentinel: %s", e)
            return "", pressed_vks
        try:
            if foreground_changed():
                logger.debug("safe_copy: foreground changed before fallback Ctrl+C -> abort")
                return "", pressed_vks
            send_ctrl_c()
            logger.debug("safe_copy: sent ctrl+c (fallback)")
        except Exception as e:  # noqa
            logger.exception("safe_copy: ctrl+c exception (fallback): %s", e)
            last_exception = e
        t0 = time.time()
        while time.time() - t0 < timeout_per_attempt:
            if foreground_changed():
                logger.debug("safe_copy: foreground changed during fallback wait -> abort")
                return "", pressed_vks
            try:
                pasted = pyperclip.paste()
            except Exception as e:  # noqa
                logger.debug("safe_copy: paste exception while waiting (fallback): %s", e)
                pasted = None
            if pasted is None:
                time.sleep(0.02)
                continue
            if pasted != sentinel:
                if pasted == "":
                    return "", pressed_vks
                logger.debug("safe_copy: buffer changed (fallback) len=%d", len(pasted))
                return pasted, pressed_vks
            time.sleep(0.02)
        # last attempt: Ctrl+Insert
        if foreground_changed():
            logger.debug("safe_copy: foreground changed before fallback Ctrl+Insert -> abort")
            return "", pressed_vks
        try:
            _key_down(VK_CONTROL)
            _key_down(VK_INSERT)
            time.sleep(0.01)
            _key_up(VK_INSERT)
            _key_up(VK_CONTROL)
            logger.debug("safe_copy: sent ctrl+insert (fallback last)")
        except Exception as e:  # noqa
            logger.exception("safe_copy: ctrl+insert exception (fallback last): %s", e)
            last_exception = e
        t0 = time.time()
        while time.time() - t0 < timeout_per_attempt:
            if foreground_changed():
                logger.debug("safe_copy: foreground changed during fallback insert wait -> abort")
                return "", pressed_vks
            try:
                pasted = pyperclip.paste()
            except Exception as e:  # noqa
                logger.debug("safe_copy: paste exception after insert (fallback): %s", e)
                pasted = None
            if pasted is None:
                time.sleep(0.02)
                continue
            if pasted != sentinel:
                if pasted == "":
                    return "", pressed_vks
                logger.debug("safe_copy: buffer changed after ctrl+insert (fallback) len=%d", len(pasted))
                return pasted, pressed_vks
            time.sleep(0.02)
        logger.warning("safe_copy: НЕ удалось получить выделение. last_exc=%r", repr(last_exception))
        return "", pressed_vks
    finally:
        logger.debug("safe_copy: returning with pressed_vks=%r", pressed_vks)


# ---------------- Restore modifiers ----------------
def _restore_modifiers_from_vks(vks: set[int]) -> None:
    """
    Восстановить логические модификаторы, нажав соответствующие left-side VK.
    """
    try:
        to_press: set[int] = set()
        for vk in vks:
            if vk in (VK_LCONTROL, VK_RCONTROL):
                to_press.add(VK_LCONTROL)
            elif vk in (VK_LMENU, VK_RMENU):
                to_press.add(VK_LMENU)
            elif vk in (VK_LSHIFT, VK_RSHIFT):
                to_press.add(VK_LSHIFT)
            elif vk in (VK_LWIN, VK_RWIN):
                to_press.add(VK_LWIN)
        for vk in to_press:
            try:
                _key_down(vk)
            except Exception:  # noqa
                pass
    except Exception:  # noqa
        pass


# ---------------- Translate handler (debounce + worker) ----------------
_handler_lock = threading.Lock()
_last_translate_ts = 0.0
_TRANSLATE_DEBOUNCE_SEC = 0.6


def _translate_invoker() -> None:
    """Дебаунс и запуск worker-потока для обработки translate hotkey."""
    global _last_translate_ts
    if not is_enabled():
        logger.debug("Translate: ignored because enabled==False")
        return
    now = time.time()
    if now - _last_translate_ts < _TRANSLATE_DEBOUNCE_SEC:
        logger.debug("Translate: debounce ignored (delta=%.3f)", now - _last_translate_ts)
        return
    _last_translate_ts = now
    if not _handler_lock.acquire(blocking=False):
        logger.debug("Translate: handler busy, ignored")
        return

    def _worker() -> None:
        try:
            handle_hotkey_transform()
        finally:
            try:
                _handler_lock.release()
            except Exception:  # noqa
                pass

    threading.Thread(target=_worker, daemon=True).start()


def handle_hotkey_transform() -> None:
    """Прочитать выделение, преобразовать раскладку и вставить обратно."""
    logger.info("Translate обработка — начинаю преобразование выделения")
    pressed_vks: set[int] = set()
    try:
        if not is_enabled():
            logger.debug("handle: disabled, returning")
            return
        title, pid, proc_name = get_active_window_info()
        logger.info("handle: active window: %r pid=%s proc=%r", title, pid, proc_name)
        try:
            try:
                saved = pyperclip.paste()
            except Exception:  # noqa
                saved = None
            logger.debug("handle: прочитал (но не буду восстанавливать) буфер text-len=%s",
                         None if saved is None else len(saved))
        except Exception:  # noqa
            pass

        hwnd = None
        try:
            hwnd = user32.GetForegroundWindow()
            pid_c = ctypes.c_ulong()
            thread_id = user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid_c))
            hkl = user32.GetKeyboardLayout(thread_id) & 0xFFFF
            logger.debug("handle: foreground thread hkl=0x%04X", hkl)
        except Exception:  # noqa
            hkl = 0
            logger.exception("handle: не удалось получить HKL, предполагаем EN")

        res = safe_copy_from_selection(timeout_per_attempt=0.6, max_attempts=2)
        if isinstance(res, tuple):
            selected, pressed_vks = res
        else:
            selected = res
            pressed_vks = set()

        if selected is None:
            selected = ""
        if not isinstance(pressed_vks, set):
            try:
                pressed_vks = set(pressed_vks or [])
            except Exception:  # noqa
                pressed_vks = set()

        logger.info("handle: прочитано выделение (len=%d)", len(selected))
        if not selected:
            logger.info("handle: выделение пустое — ничего не делаю.")
            return

        converted = transform_text_by_keyboard_layout_based_on_hkl(selected, hkl)
        logger.info("handle: преобразование выполнено (len=%d).", len(converted))
        logger.debug("handle: исходное=%r converted=%r", selected, converted)
        if converted == selected:
            logger.info("handle: преобразованный текст совпадает с исходным — ничего не меняю.")
            return

        try:
            send_delete()
            time.sleep(0.02)
        except Exception as e:  # noqa
            logger.exception("handle: delete exception: %s", e)

        try:
            send_unicode_via_sendinput(converted, delay_between_keys=0.001)
            logger.debug("handle: вставил через SendInput (unicode) — без использования clipboard")
        except Exception as e:  # noqa
            logger.exception("handle: send_unicode_via_sendinput failed: %s", e)
            try:
                pyperclip.copy(converted)
                time.sleep(0.02)
                send_ctrl_v()
                logger.debug("handle: fallback: вставил через буфер (ctrl+v)")
            except Exception as e2:  # noqa
                logger.exception("handle: fallback paste failed: %s", e2)

        try:
            if hwnd:
                if hkl == 0x0419:
                    new_klid = "00000409"
                else:
                    new_klid = "00000419"
                hkl_new = user32.LoadKeyboardLayoutW(new_klid, 1)
                WM_INPUTLANGCHANGEREQUEST = 0x0050
                res_post = user32.PostMessageW(hwnd, WM_INPUTLANGCHANGEREQUEST, 0, ctypes.c_void_p(hkl_new))
                if res_post == 0:
                    logger.debug("handle: PostMessageW failed (res=0), falling back to ActivateKeyboardLayout")
                    try:
                        user32.ActivateKeyboardLayout(hkl_new, 0)
                        logger.debug("handle: ActivateKeyboardLayout used as fallback")
                    except Exception:  # noqa
                        logger.exception("handle: fallback ActivateKeyboardLayout also failed")
                else:
                    logger.debug("handle: posted WM_INPUTLANGCHANGEREQUEST -> success")
        except Exception:  # noqa
            logger.exception("handle: switch layout failed (foreground-thread)")

        logger.info("handle: завершено успешно.")
    except Exception as e:  # noqa
        logger.exception("handle_hotkey_transform: исключение при обработке: %s", e)
    finally:
        try:
            if pressed_vks:
                logger.debug("handle: restoring previously released modifiers vks=%r", pressed_vks)
                _restore_modifiers_from_vks(pressed_vks)
        except Exception:  # noqa
            logger.exception("handle: failed to restore modifiers")


# ---------------- Case handler (debounce + worker) ----------------
def _case_invoker() -> None:
    """Дебаунс и запуск worker-потока для обработки case hotkey."""
    global _last_translate_ts
    if not is_enabled():
        logger.debug("Case: ignored because enabled==False")
        return
    now = time.time()
    if now - _last_translate_ts < _TRANSLATE_DEBOUNCE_SEC:
        logger.debug("Case: debounce ignored (delta=%.3f)", now - _last_translate_ts)
        return
    _last_translate_ts = now
    if not _handler_lock.acquire(blocking=False):
        logger.debug("Case: handler busy, ignored")
        return

    def _worker() -> None:
        try:
            handle_hotkey_case()
        finally:
            try:
                _handler_lock.release()
            except Exception:  # noqa
                pass

    threading.Thread(target=_worker, daemon=True).start()


def handle_hotkey_case() -> None:
    """Прочитать выделение, изменить регистр и вставить обратно."""
    logger.info("Case обработка — начинаю изменение регистра выделения")
    pressed_vks: set[int] = set()
    try:
        if not is_enabled():
            logger.debug("handle_case: disabled, returning")
            return
        title, pid, proc_name = get_active_window_info()
        logger.info("handle_case: active window: %r pid=%s proc=%r", title, pid, proc_name)
        try:
            try:
                saved = pyperclip.paste()
            except Exception:  # noqa
                saved = None
            logger.debug("handle_case: прочитал (но не буду восстанавливать) буфер text-len=%s",
                         None if saved is None else len(saved))
        except Exception:  # noqa
            pass

        res = safe_copy_from_selection(timeout_per_attempt=0.6, max_attempts=2)
        if isinstance(res, tuple):
            selected, pressed_vks = res
        else:
            selected = res
            pressed_vks = set()

        if selected is None:
            selected = ""
        if not isinstance(pressed_vks, set):
            try:
                pressed_vks = set(pressed_vks or [])
            except Exception:  # noqa
                pressed_vks = set()

        logger.info("handle_case: прочитано выделение (len=%d)", len(selected))
        if not selected:
            logger.info("handle_case: выделение пустое — ничего не делаю.")
            return

        converted = change_case_by_logic(selected)
        if converted == selected:
            logger.info("handle_case: регистр не изменился — ничего не делаю.")
            return

        try:
            send_delete()
            time.sleep(0.02)
        except Exception as e:  # noqa
            logger.exception("handle_case: delete exception: %s", e)

        try:
            send_unicode_via_sendinput(converted, delay_between_keys=0.001)
            logger.debug("handle_case: вставил через SendInput (unicode)")
        except Exception as e:  # noqa
            logger.exception("handle_case: send_unicode_via_sendinput failed: %s", e)
            try:
                pyperclip.copy(converted)
                time.sleep(0.02)
                send_ctrl_v()
                logger.debug("handle_case: fallback: вставил через буфер (ctrl+v)")
            except Exception as e2:  # noqa
                logger.exception("handle_case: fallback paste failed: %s", e2)

        logger.info("handle_case: завершено успешно.")
    except Exception as e:  # noqa
        logger.exception("handle_hotkey_case: исключение при обработке: %s", e)
    finally:
        try:
            if pressed_vks:
                logger.debug("handle_case: restoring previously released modifiers vks=%r", pressed_vks)
                _restore_modifiers_from_vks(pressed_vks)
        except Exception:  # noqa
            logger.exception("handle_case: failed to restore modifiers")


# ---------------- Hotkey messaging helpers ----------------
def _post_thread_message_check_thread() -> bool:
    """Проверка, готов ли hotkey-поток для приема сообщений."""
    if HOTKEY_THREAD_ID == 0:
        logger.warning("hotkey thread id not ready yet -> cannot post")
        return False
    return True


def post_register_translate(should_register: bool) -> bool:
    """Послать потоковое сообщение для регистрации/отмены комбинации перевода."""
    global HOTKEY_THREAD_ID
    if not _post_thread_message_check_thread():
        return False
    msg = MSG_REGISTER_TRANSLATE if should_register else MSG_UNREGISTER_TRANSLATE
    wparam = 1 if should_register else 0
    try:
        res = user32.PostThreadMessageW(HOTKEY_THREAD_ID, msg, wparam, 0)
        if res == 0:
            err = kernel32.GetLastError()
            logger.error("post_register_translate: PostThreadMessageW failed (err=%d)", err)
            return False
        logger.debug("post_register_translate: posted msg=%d wParam=%d to thread %d", msg, wparam, HOTKEY_THREAD_ID)
        return True
    except Exception:  # noqa
        logger.exception("post_register_translate: exception while posting")
        return False


def post_update_exit_hotkey() -> bool:
    """Послать сообщение для обновления exit-hotkey в hotkey-потоке."""
    global HOTKEY_THREAD_ID
    if not _post_thread_message_check_thread():
        return False
    try:
        res = user32.PostThreadMessageW(HOTKEY_THREAD_ID, MSG_UPDATE_EXIT_HOTKEY, 0, 0)
        if res == 0:
            err = kernel32.GetLastError()
            logger.error("post_update_exit_hotkey: PostThreadMessageW failed (err=%d)", err)
            return False
        logger.debug("post_update_exit_hotkey: posted MSG_UPDATE_EXIT_HOTKEY to thread %d", HOTKEY_THREAD_ID)
        return True
    except Exception:  # noqa
        logger.exception("post_update_exit_hotkey: exception while posting")
        return False


def post_update_translate_hotkey() -> bool:
    """Послать сообщение для обновления translate-hotkey в hotkey-потоке."""
    global HOTKEY_THREAD_ID
    if not _post_thread_message_check_thread():
        return False
    try:
        res = user32.PostThreadMessageW(HOTKEY_THREAD_ID, MSG_UPDATE_TRANSLATE_HOTKEY, 0, 0)
        if res == 0:
            err = kernel32.GetLastError()
            logger.error("post_update_translate_hotkey: PostThreadMessageW failed (err=%d)", err)
            return False
        logger.debug("post_update_translate_hotkey: posted MSG_UPDATE_TRANSLATE_HOTKEY to thread %d", HOTKEY_THREAD_ID)
        return True
    except Exception:  # noqa
        logger.exception("post_update_translate_hotkey: exception while posting")
        return False


def post_register_case(should_register: bool) -> bool:
    """Послать потоковое сообщение для регистрации/отмены комбинации перевода."""
    global HOTKEY_THREAD_ID
    if not _post_thread_message_check_thread():
        return False
    msg = MSG_REGISTER_CASE if should_register else MSG_UNREGISTER_CASE
    wparam = 1 if should_register else 0
    try:
        res = user32.PostThreadMessageW(HOTKEY_THREAD_ID, msg, wparam, 0)
        if res == 0:
            err = kernel32.GetLastError()
            logger.error("post_register_case: PostThreadMessageW failed (err=%d)", err)
            return False
        logger.debug("post_register_case: posted msg=%d wParam=%d to thread %d", msg, wparam, HOTKEY_THREAD_ID)
        return True
    except Exception:  # noqa
        logger.exception("post_register_case: exception while posting")
        return False


def post_update_case_hotkey() -> bool:
    """Послать сообщение для обновления case-hotkey в hotkey-потоке."""
    global HOTKEY_THREAD_ID
    if not _post_thread_message_check_thread():
        return False
    try:
        res = user32.PostThreadMessageW(HOTKEY_THREAD_ID, MSG_UPDATE_CASE_HOTKEY, 0, 0)
        if res == 0:
            err = kernel32.GetLastError()
            logger.error("post_update_case_hotkey: PostThreadMessageW failed (err=%d)", err)
            return False
        logger.debug("post_update_case_hotkey: posted MSG_UPDATE_CASE_HOTKEY to thread %d", HOTKEY_THREAD_ID)
        return True
    except Exception:  # noqa
        logger.exception("post_update_case_hotkey: exception while posting")
        return False


# ---------------- Helpers: key names / masks ----------------
def key_name_to_vk(name: str) -> int:
    """Преобразовать читабельное имя клавиши в VK."""
    n = (name or "").upper()
    if not n:
        return VK_MAP.get("F10", 0x79)
    if n in VK_MAP:
        return VK_MAP[n]
    if len(n) == 1:
        ch = n
        if 'A' <= ch <= 'Z':
            return ord(ch)
        if '0' <= ch <= '9':
            return ord(ch)
    try:
        if n.startswith("VK_"):
            return int(getattr(win32con, n, 0))
    except Exception:  # noqa
        pass
    return VK_MAP.get("F10", 0x79)


def modifiers_list_to_mask(mods: list[str]) -> int:
    """Преобразовать список модификаторов в маску для RegisterHotKey."""
    mask = 0
    for m in mods or []:
        mm = (m or "").upper()
        if mm in ("ALT",):
            mask |= MOD_ALT
        elif mm in ("CTRL", "CONTROL"):
            mask |= MOD_CONTROL
        elif mm in ("SHIFT",):
            mask |= MOD_SHIFT
        elif mm in ("WIN", "WINDOWS"):
            mask |= MOD_WIN
    return mask


def hotkey_tuple_from_config(mods: list[str], key: str) -> tuple[int, int]:
    """Вернуть (mask, vk) для комбинации из конфига."""
    return modifiers_list_to_mask(mods), key_name_to_vk(key)


def hotkeys_conflict(mask1: int, vk1: int, mask2: int, vk2: int) -> bool:
    """Проверить, конфликтуют ли две комбинации (одинаковые mask+vk)."""
    return mask1 == mask2 and vk1 == vk2


# ---------------- Update registration helpers (run in hotkey thread) ----------------
def update_exit_hotkey_in_thread() -> None:
    """В hotkey-потоке: снять старую комбинацию выхода и зарегистрировать новую."""
    try:
        try:
            user32.UnregisterHotKey(None, HOTKEY_ID_EXIT)
        except Exception:  # noqa
            pass
        eh = read_exit_hotkey()
        mods = eh.get("modifiers", []) or []
        key = eh.get("key", "F10") or "F10"
        mask, vk = hotkey_tuple_from_config(mods, key)

        th = read_translate_hotkey()
        th_mask, th_vk = hotkey_tuple_from_config(th.get("modifiers", []) or [], th.get("key", "F4"))
        if hotkeys_conflict(mask, vk, th_mask, th_vk):
            logger.error("update_exit_hotkey_in_thread: конфликт с translate hotkey; пропускаю регистрацию exit")
            try:
                win32api.MessageBox(
                    0,
                    f"Не удалось зарегистрировать комбинацию выхода {'+'.join(mods) + '+' if mods else ''}{key} — конфликт с комбинацией перевода.",
                    "KeyFlip",
                    0,
                )
            except Exception:  # noqa
                pass
            return

        case = read_case_hotkey()
        case_mask, case_vk = hotkey_tuple_from_config(case.get("modifiers", []) or [], case.get("key", "U"))
        if hotkeys_conflict(mask, vk, case_mask, case_vk):
            logger.error("update_exit_hotkey_in_thread: конфликт с case hotkey; пропускаю регистрацию exit")
            try:
                win32api.MessageBox(
                    0,
                    f"Не удалось зарегистрировать комбинацию выхода {'+'.join(mods) + '+' if mods else ''}{key} — конфликт с комбинацией изменения регистра.",
                    "KeyFlip",
                    0,
                )
            except Exception:  # noqa
                pass
            return

        ok = user32.RegisterHotKey(None, HOTKEY_ID_EXIT, mask, vk)
        if ok:
            logger.debug(
                "update_exit_hotkey_in_thread: registered exit hotkey %s + %s (mask=0x%X vk=0x%X)",
                "+".join(mods) if mods else "(no modifiers)",
                key,
                mask,
                vk,
            )
        else:
            err = kernel32.GetLastError()
            logger.error(
                "update_exit_hotkey_in_thread: failed to register exit hotkey (err=%d) mask=0x%X vk=0x%X",
                err,
                mask,
                vk,
            )
            try:
                win32api.MessageBox(
                    0,
                    f"Не удалось зарегистрировать комбинацию выхода {'+'.join(mods) + '+' if mods else ''}{key}.\nКод ошибки: {err}",
                    "KeyFlip",
                    0,
                )
            except Exception:  # noqa
                pass
    except Exception:  # noqa
        logger.exception("update_exit_hotkey_in_thread: exception")


def update_translate_hotkey_in_thread() -> None:
    """В hotkey-потоке: снять старую комбинацию перевода и зарегистрировать новую (с учётом enabled)."""
    try:
        try:
            user32.UnregisterHotKey(None, HOTKEY_ID_TRANSLATE)
        except Exception:  # noqa
            pass
        th = read_translate_hotkey()
        mods = th.get("modifiers", []) or []
        key = th.get("key", "F4") or "F4"
        mask, vk = hotkey_tuple_from_config(mods, key)

        if not is_enabled():
            logger.debug("update_translate_hotkey_in_thread: приложение выключено, пропускаю регистрацию translate")
            return

        eh = read_exit_hotkey()
        eh_mask, eh_vk = hotkey_tuple_from_config(eh.get("modifiers", []) or [], eh.get("key", "F10"))
        if hotkeys_conflict(mask, vk, eh_mask, eh_vk):
            logger.error("update_translate_hotkey_in_thread: конфликт с exit hotkey; пропускаю регистрацию translate")
            try:
                win32api.MessageBox(
                    0,
                    f"Не удалось зарегистрировать комбинацию перевода {'+'.join(mods) + '+' if mods else ''}{key} — конфликт с комбинацией выхода.",
                    "KeyFlip",
                    0,
                )
            except Exception:  # noqa
                pass
            return

        case = read_case_hotkey()
        case_mask, case_vk = hotkey_tuple_from_config(case.get("modifiers", []) or [], case.get("key", "U"))
        if hotkeys_conflict(mask, vk, case_mask, case_vk):
            logger.error("update_translate_hotkey_in_thread: конфликт с case hotkey; пропускаю регистрацию translate")
            try:
                win32api.MessageBox(
                    0,
                    f"Не удалось зарегистрировать комбинацию перевода {'+'.join(mods) + '+' if mods else ''}{key} — конфликт с комбинацией изменения регистра.",
                    "KeyFlip",
                    0,
                )
            except Exception:  # noqa
                pass
            return

        ok = user32.RegisterHotKey(None, HOTKEY_ID_TRANSLATE, mask, vk)
        if ok:
            logger.debug(
                "update_translate_hotkey_in_thread: registered translate hotkey %s + %s (mask=0x%X vk=0x%X)",
                "+".join(mods) if mods else "(no modifiers)",
                key,
                mask,
                vk,
            )
        else:
            err = kernel32.GetLastError()
            logger.error(
                "update_translate_hotkey_in_thread: failed to register translate hotkey (err=%d) mask=0x%X vk=0x%X",
                err,
                mask,
                vk,
            )
            try:
                win32api.MessageBox(
                    0,
                    f"Не удалось зарегистрировать комбинацию перевода {'+'.join(mods) + '+' if mods else ''}{key}.\nКод ошибки: {err}",
                    "KeyFlip",
                    0,
                )
            except Exception:  # noqa
                pass
    except Exception:  # noqa
        logger.exception("update_translate_hotkey_in_thread: exception")



def update_case_hotkey_in_thread() -> None:
    """В hotkey-потоке: снять старую комбинацию case и зарегистрировать новую."""
    try:
        try:
            user32.UnregisterHotKey(None, HOTKEY_ID_CASE)
        except Exception:  # noqa
            pass
        ch = read_case_hotkey()
        mods = ch.get("modifiers", []) or []
        key = ch.get("key", "U") or "U"
        mask, vk = hotkey_tuple_from_config(mods, key)

        th = read_translate_hotkey()
        th_mask, th_vk = hotkey_tuple_from_config(th.get("modifiers", []) or [], th.get("key", "F4"))
        if hotkeys_conflict(mask, vk, th_mask, th_vk):
            logger.error("update_exit_hotkey_in_thread: конфликт с translate hotkey; пропускаю регистрацию case")
            try:
                win32api.MessageBox(
                    0,
                    f"Не удалось зарегистрировать комбинацию выхода {'+'.join(mods) + '+' if mods else ''}{key} — конфликт с комбинацией перевода.",
                    "KeyFlip",
                    0,
                )
            except Exception:  # noqa
                pass
            return

        eh = read_exit_hotkey()
        eh_mask, eh_vk = hotkey_tuple_from_config(eh.get("modifiers", []) or [], eh.get("key", "F10"))
        if hotkeys_conflict(mask, vk, eh_mask, eh_vk):
            logger.error("update_translate_hotkey_in_thread: конфликт с exit hotkey; пропускаю регистрацию case")
            try:
                win32api.MessageBox(
                    0,
                    f"Не удалось зарегистрировать комбинацию перевода {'+'.join(mods) + '+' if mods else ''}{key} — конфликт с комбинацией выхода.",
                    "KeyFlip",
                    0,
                )
            except Exception:  # noqa
                pass
            return

        ok = user32.RegisterHotKey(None, HOTKEY_ID_CASE, mask, vk)
        if ok:
            logger.debug("update_case_hotkey_in_thread: registered case hotkey %s + %s (mask=0x%X vk=0x%X)",
                         "+".join(mods) if mods else "(no modifiers)", key, mask, vk)
        else:
            err = kernel32.GetLastError()
            logger.error("update_case_hotkey_in_thread: failed to register case hotkey (err=%d) mask=0x%X vk=0x%X",
                         err, mask, vk)
            try:
                win32api.MessageBox(0,
                                    f"Не удалось зарегистрировать комбинацию регистра {'+'.join(mods) + '+' if mods else ''}{key}.\nКод ошибки: {err}",
                                    "KeyFlip", 0)
            except Exception:  # noqa
                pass
    except Exception:  # noqa
        logger.exception("update_case_hotkey_in_thread: exception")


# ---------------- Low-level hook structures and hook proc ----------------
LRESULT = ctypes.c_longlong if ctypes.sizeof(ctypes.c_void_p) == 8 else ctypes.c_long


class KBDLLHOOKSTRUCT(ctypes.Structure):
    _fields_ = [
        ("vkCode", wintypes.DWORD),
        ("scanCode", wintypes.DWORD),
        ("flags", wintypes.DWORD),
        ("time", wintypes.DWORD),
        ("dwExtraInfo", ULONG_PTR),
    ]


user32.SetWindowsHookExW.argtypes = (wintypes.INT, ctypes.c_void_p, wintypes.HINSTANCE, wintypes.DWORD)
user32.SetWindowsHookExW.restype = wintypes.HHOOK
user32.UnhookWindowsHookEx.argtypes = (wintypes.HHOOK,)
user32.UnhookWindowsHookEx.restype = wintypes.BOOL
user32.CallNextHookEx.argtypes = (wintypes.HHOOK, ctypes.c_int, wintypes.WPARAM, wintypes.LPARAM)
user32.CallNextHookEx.restype = LRESULT


def _vk_is_modifier(vk: int) -> bool:
    """Является ли vk модификатором (ctrl/alt/shift/win)."""
    return vk in (
        VK_CONTROL, VK_MENU, VK_SHIFT, VK_LSHIFT, VK_RSHIFT, VK_LCONTROL, VK_RCONTROL, VK_LMENU, VK_RMENU, VK_LWIN,
        VK_RWIN
    )


def _current_modifiers_list() -> list[str]:
    """Собрать список логических модификаторов, которые сейчас нажаты."""
    mods: list[str] = []
    try:
        if user32.GetAsyncKeyState(VK_LCONTROL) & 0x8000 or user32.GetAsyncKeyState(VK_RCONTROL) & 0x8000:
            mods.append("CTRL")
        if user32.GetAsyncKeyState(VK_LMENU) & 0x8000 or user32.GetAsyncKeyState(VK_RMENU) & 0x8000:
            mods.append("ALT")
        if user32.GetAsyncKeyState(VK_LSHIFT) & 0x8000 or user32.GetAsyncKeyState(VK_RSHIFT) & 0x8000:
            mods.append("SHIFT")
        if user32.GetAsyncKeyState(VK_LWIN) & 0x8000 or user32.GetAsyncKeyState(VK_RWIN) & 0x8000:
            mods.append("WIN")
    except Exception:  # noqa
        pass
    return mods


def vk_to_key_name(vk: int) -> str:
    """Преобразовать vk в читабельное имя клавиши."""
    for name, code in VK_MAP.items():
        if code == vk:
            return name
    if 0x41 <= vk <= 0x5A:
        return chr(vk)
    if 0x30 <= vk <= 0x39:
        return chr(vk)
    if vk == 0x20:
        return "SPACE"
    if vk == 0x0D:
        return "ENTER"
    return f"VK_{vk:02X}"


def capture_hotkey_via_hook_blocking(timeout: float | None = 10.0, show_dialog: bool = True) -> dict | None:
    """
    Заблокированно захватить комбинацию клавиш через WH_KEYBOARD_LL.
    Возвращает {'modifiers': [...], 'key': 'X'} или None.
    """
    result: dict | None = {"modifiers": [], "key": ""}
    done_event = threading.Event()
    hook_handle = None

    HOOKPROC = ctypes.WINFUNCTYPE(LRESULT, ctypes.c_int, wintypes.WPARAM, wintypes.LPARAM)

    def _low_level_proc(nCode: int, wParam, lParam):
        nonlocal hook_handle, result, done_event
        try:
            if nCode == 0 and (wParam == WM_KEYDOWN or wParam == WM_SYSKEYDOWN):
                k = ctypes.cast(lParam, ctypes.POINTER(KBDLLHOOKSTRUCT)).contents
                vk = int(k.vkCode)
                if vk == VK_ESCAPE:
                    result = None
                    done_event.set()
                    return user32.CallNextHookEx(hook_handle, nCode, wParam, lParam)
                if _vk_is_modifier(vk):
                    return user32.CallNextHookEx(hook_handle, nCode, wParam, lParam)
                mods = _current_modifiers_list()
                keyname = vk_to_key_name(vk)
                result = {"modifiers": mods, "key": keyname}
                done_event.set()
                return user32.CallNextHookEx(hook_handle, nCode, wParam, lParam)
            return user32.CallNextHookEx(hook_handle, nCode, wParam, lParam)
        except Exception:  # noqa
            logger.exception("capture_hotkey: hook proc exception")
            try:
                return user32.CallNextHookEx(hook_handle, nCode, wParam, lParam)
            except Exception:  # noqa
                return LRESULT(0)

    hook_proc_ptr = HOOKPROC(_low_level_proc)
    tk_root = None

    try:
        try:
            hmod = kernel32.GetModuleHandleW(None)
        except Exception:  # noqa
            hmod = None

        hook_handle = user32.SetWindowsHookExW(WH_KEYBOARD_LL, hook_proc_ptr, hmod or None, 0)
        if not hook_handle:
            err = kernel32.GetLastError()
            logger.debug("capture_hotkey: SetWindowsHookExW failed (err=%d) hmod=%r", err, hmod)
            hook_handle = user32.SetWindowsHookExW(WH_KEYBOARD_LL, hook_proc_ptr, None, 0)
            if not hook_handle:
                err2 = kernel32.GetLastError()
                logger.error("capture_hotkey: SetWindowsHookExW final failed (err=%d)", err2)
                return None

        logger.debug("capture_hotkey: SetWindowsHookExW succeeded (hook=%r)", hook_handle)

        try:
            if show_dialog:
                try:
                    tk_root = tk.Tk()
                    tk_root.title("KeyFlip — введите комбинацию")
                    tk_root.attributes("-topmost", True)
                    tk_root.geometry("360x80")
                    label = tk.Label(tk_root, text="Нажми требуемую комбинацию (ESC — отмена)", padx=8, pady=12)
                    label.pack(fill="both", expand=True)

                    def on_close():
                        nonlocal result
                        result = None
                        done_event.set()

                    tk_root.protocol("WM_DELETE_WINDOW", on_close)
                    tk_root.update()
                except Exception:  # noqa
                    tk_root = None
        except Exception:  # noqa
            tk_root = None

        t0 = time.time()
        while True:
            if done_event.is_set():
                break
            if timeout is not None and (time.time() - t0) > timeout:
                logger.debug("capture_hotkey: timeout expired")
                result = None
                break
            # process messages so hook works
            msg = wintypes.MSG()
            has = user32.PeekMessageW(ctypes.byref(msg), None, 0, 0, 1)  # PM_REMOVE = 1
            if has:
                user32.TranslateMessage(ctypes.byref(msg))
                user32.DispatchMessageW(ctypes.byref(msg))
            if tk_root is not None:
                try:
                    tk_root.update()
                except Exception:  # noqa
                    try:
                        tk_root.destroy()
                    except Exception:  # noqa
                        pass
                    tk_root = None
            else:
                time.sleep(0.01)

    except Exception:  # noqa
        logger.exception("capture_hotkey: exception while hooking/looping")
        result = None
    finally:
        try:
            if hook_handle:
                user32.UnhookWindowsHookEx(hook_handle)
                logger.debug("capture_hotkey: UnhookWindowsHookEx called")
        except Exception:  # noqa
            logger.exception("capture_hotkey: failed to unhook")
        try:
            if "tk_root" in locals() and tk_root is not None:
                try:
                    tk_root.destroy()
                except Exception:  # noqa
                    pass
        except Exception:  # noqa
            pass
        try:
            hook_proc_ptr = None
        except Exception:  # noqa
            pass
    return result


def hotkey_lists_equal(a_mods: list | None, a_key: str, b_mods: list | None, b_key: str) -> bool:
    """Сравнить две комбинации модификаторов+клавиша."""
    a_norm = set((m or "").upper() for m in (a_mods or []))
    b_norm = set((m or "").upper() for m in (b_mods or []))
    return a_norm == b_norm and ((a_key or "").upper() == (b_key or "").upper())


def capture_hotkey_and_apply_via_thread(target: str) -> None:
    """
    Запустить поток, который захватит комбинацию, затем запишет её в конфиг.
    target: 'exit' или 'translate' или 'case'
    """

    def _runner() -> None:
        logger.info("Capture thread: ожидаю комбинацию. Нажми требуемую комбинацию (ESC для отмены).")
        res = capture_hotkey_via_hook_blocking(timeout=None, show_dialog=True)
        if not res:
            logger.info("Capture thread: захват комбинации отменён или не получен.")
            return
        mods = res.get("modifiers", [])  # type: ignore[assignment]
        key = res.get("key", "F10")  # type: ignore[assignment]

        if target == "exit":
            other = read_translate_hotkey()
            if hotkey_lists_equal(mods, key, other.get("modifiers", []) or [], other.get("key", "")):
                try:
                    win32api.MessageBox(
                        0,
                        f"Нельзя установить комбинацию выхода {'+'.join(mods) + '+' if mods else ''}{key} — она уже используется для перевода.",
                        "KeyFlip",
                        0,
                    )
                except Exception:  # noqa
                    pass
                logger.info("Capture thread: попытка установки exit hotkey отклонена — конфликт с translate")
                return
            other2 = read_case_hotkey()
            if hotkey_lists_equal(mods, key, other2.get("modifiers", []) or [], other2.get("key", "")):
                try:
                    win32api.MessageBox(
                        0,
                        f"Нельзя установить комбинацию регистра {'+'.join(mods) + '+' if mods else ''}{key} — она уже используется для изменения регистра.",
                        "KeyFlip",
                        0
                    )
                except Exception:  # noqa
                    pass
                logger.info("Capture thread: попытка установки translate hotkey отклонена — конфликт с case")
                return
            ok = write_exit_hotkey(mods, key)
            if ok:
                post_update_exit_hotkey()
                logger.info(
                    "Capture thread: установлена новая комбинация выхода: %s + %s",
                    "+".join(mods) if mods else "(none)",
                    key,
                )
            else:
                logger.warning("Capture thread: не удалось записать новую комбинацию выхода")
        elif target == "translate":
            other = read_exit_hotkey()
            if hotkey_lists_equal(mods, key, other.get("modifiers", []) or [], other.get("key", "")):
                try:
                    win32api.MessageBox(
                        0,
                        f"Нельзя установить комбинацию перевода {'+'.join(mods) + '+' if mods else ''}{key} — она уже используется для выхода.",
                        "KeyFlip",
                        0,
                    )
                except Exception:  # noqa
                    pass
                logger.info("Capture thread: попытка установки translate hotkey отклонена — конфликт с exit")
                return
            other2 = read_case_hotkey()
            if hotkey_lists_equal(mods, key, other2.get("modifiers", []) or [], other2.get("key", "")):
                try:
                    win32api.MessageBox(
                        0,
                        f"Нельзя установить комбинацию регистра {'+'.join(mods) + '+' if mods else ''}{key} — она уже используется для изменения регистра.",
                        "KeyFlip",
                        0
                    )
                except Exception:  # noqa
                    pass
                logger.info("Capture thread: попытка установки translate hotkey отклонена — конфликт с case")
                return
            ok = write_translate_hotkey(mods, key)
            if ok:
                post_update_translate_hotkey()
                logger.info(
                    "Capture thread: установлена новая комбинация перевода: %s + %s",
                    "+".join(mods) if mods else "(none)",
                    key,
                )
            else:
                logger.warning("Capture thread: не удалось записать новую комбинацию перевода")
        elif target == "case":
            other = read_exit_hotkey()
            if hotkey_lists_equal(mods, key, other.get("modifiers", []) or [], other.get("key", "")):
                try:
                    win32api.MessageBox(
                        0,
                        f"Нельзя установить комбинацию регистра {'+'.join(mods) + '+' if mods else ''}{key} — она уже используется для выхода.",
                        "KeyFlip",
                        0
                    )
                except Exception:  # noqa
                    pass
                logger.info("Capture thread: попытка установки case hotkey отклонена — конфликт с exit")
                return
            other2 = read_translate_hotkey()
            if hotkey_lists_equal(mods, key, other2.get("modifiers", []) or [], other2.get("key", "")):
                try:
                    win32api.MessageBox(
                        0,
                        f"Нельзя установить комбинацию регистра {'+'.join(mods) + '+' if mods else ''}{key} — она уже используется для перевода.",
                        "KeyFlip",
                        0
                    )
                except Exception:  # noqa
                    pass
                logger.info("Capture thread: попытка установки case hotkey отклонена — конфликт с translate")
                return
            ok = write_case_hotkey(mods, key)
            if ok:
                post_update_case_hotkey()
                logger.info(
                    "Capture thread: установлена новая комбинация регистра: %s + %s",
                    "+".join(mods) if mods else "(none)",
                    key
                )
            else:
                logger.warning("Capture thread: не удалось записать новую комбинацию регистра")
        else:
            logger.warning("Capture thread: неизвестная цель capture: %r", target)

    name = "ExitCaptureThread" if target == "exit" else ("TranslateCaptureThread" if target == "translate" else "CaseCaptureThread")
    t = threading.Thread(target=_runner, daemon=True, name=name)
    t.start()


# ---------------- Win hotkey message loop ----------------
def win_hotkey_loop() -> None:
    """
    Поток, который слушает Windows сообщения и обрабатывает WM_HOTKEY и внутренние MSG_*.
    """
    global HOTKEY_THREAD_ID
    try:
        HOTKEY_THREAD_ID = kernel32.GetCurrentThreadId()
        logger.debug("win_hotkey_loop: HOTKEY_THREAD_ID = %d", HOTKEY_THREAD_ID)

        # Убедимся, что exit зарегистрирован
        try:
            update_exit_hotkey_in_thread()
        except Exception:  # noqa
            logger.exception("win_hotkey_loop: update_exit_hotkey_in_thread initial failed")

        # Попробуем зарегистрировать translate (если enabled)
        try:
            update_translate_hotkey_in_thread()
        except Exception:  # noqa
            logger.exception("win_hotkey_loop: update_translate_hotkey_in_thread initial failed")

        # Попробуем зарегистрировать case
        try:
            update_case_hotkey_in_thread()
        except Exception:  # noqa
            logger.exception("win_hotkey_loop: update_case_hotkey_in_thread initial failed")

        msg = wintypes.MSG()
        while True:
            # Блокирующее ожидание сообщения
            has = user32.GetMessageW(ctypes.byref(msg), None, 0, 0)
            if has == 0:
                # WM_QUIT
                logger.debug("win_hotkey_loop: GetMessage returned 0 -> quitting loop")
                break
            if has == -1:
                logger.error("win_hotkey_loop: GetMessage error")
                break

            if msg.message == win32con.WM_HOTKEY:
                wParam = msg.wParam
                hot_id = int(wParam)
                logger.debug("win_hotkey_loop: WM_HOTKEY id=%d", hot_id)
                if hot_id == HOTKEY_ID_EXIT:
                    logger.info("win_hotkey_loop: exit hotkey pressed -> initiating exit")
                    # Установим exit_event и вызовем зарегистрированные обработчики
                    exit_event.set()
                    try:
                        _invoke_exit_handlers()
                    except Exception:  # noqa
                        logger.exception("win_hotkey_loop: exit handlers raised")
                    # Прекращаем регистрировать хоткеи
                    try:
                        user32.UnregisterHotKey(None, HOTKEY_ID_EXIT)
                    except Exception:  # noqa
                        pass
                    try:
                        user32.UnregisterHotKey(None, HOTKEY_ID_TRANSLATE)
                    except Exception:  # noqa
                        pass
                    try:
                        user32.UnregisterHotKey(None, HOTKEY_ID_CASE)
                    except Exception:  # noqa
                        pass
                    break
                elif hot_id == HOTKEY_ID_TRANSLATE:
                    logger.debug("win_hotkey_loop: translate hotkey pressed")
                    _translate_invoker()
                elif hot_id == HOTKEY_ID_CASE:
                    logger.debug("win_hotkey_loop: case hotkey pressed")
                    _case_invoker()
                else:
                    logger.debug("win_hotkey_loop: unknown hotkey id=%s", hot_id)

            elif msg.message == MSG_REGISTER_TRANSLATE:
                logger.debug("win_hotkey_loop: MSG_REGISTER_TRANSLATE received")
                try:
                    update_translate_hotkey_in_thread()
                except Exception:  # noqa
                    logger.exception("win_hotkey_loop: update_translate_hotkey_in_thread failed in handler")
            elif msg.message == MSG_UNREGISTER_TRANSLATE:
                logger.debug("win_hotkey_loop: MSG_UNREGISTER_TRANSLATE received")
                try:
                    user32.UnregisterHotKey(None, HOTKEY_ID_TRANSLATE)
                except Exception:  # noqa
                    logger.exception("win_hotkey_loop: UnregisterHotKey translate failed")
            elif msg.message == MSG_REGISTER_CASE:
                logger.debug("win_hotkey_loop: MSG_REGISTER_CASE received")
                try:
                    update_case_hotkey_in_thread()
                except Exception:  # noqa
                    logger.exception("win_hotkey_loop: update_case_hotkey_in_thread failed in handler")
            elif msg.message == MSG_UNREGISTER_CASE:
                logger.debug("win_hotkey_loop: MSG_UNREGISTER_CASE received")
                try:
                    user32.UnregisterHotKey(None, HOTKEY_ID_CASE)
                except Exception:  # noqa
                    logger.exception("win_hotkey_loop: UnregisterHotKey case failed")
            elif msg.message == MSG_UPDATE_EXIT_HOTKEY:
                logger.debug("win_hotkey_loop: MSG_UPDATE_EXIT_HOTKEY received")
                try:
                    update_exit_hotkey_in_thread()
                except Exception:  # noqa
                    logger.exception("win_hotkey_loop: update_exit_hotkey_in_thread failed in handler")
            elif msg.message == MSG_UPDATE_TRANSLATE_HOTKEY:
                logger.debug("win_hotkey_loop: MSG_UPDATE_TRANSLATE_HOTKEY received")
                try:
                    update_translate_hotkey_in_thread()
                except Exception:  # noqa
                    logger.exception("win_hotkey_loop: update_translate_hotkey_in_thread failed in handler")
            elif msg.message == MSG_UPDATE_CASE_HOTKEY:
                logger.debug("win_hotkey_loop: MSG_UPDATE_CASE_HOTKEY received")
                try:
                    update_case_hotkey_in_thread()
                except Exception:  # noqa
                    logger.exception("win_hotkey_loop: update_case_hotkey_in_thread failed in handler")
            else:
                # стандартная обработка сообщений
                user32.TranslateMessage(ctypes.byref(msg))
                user32.DispatchMessageW(ctypes.byref(msg))

    except Exception:  # noqa
        logger.exception("win_hotkey_loop: fatal exception")
    finally:
        logger.info("win_hotkey_loop: exiting, cleaning up")
        try:
            user32.UnregisterHotKey(None, HOTKEY_ID_EXIT)
        except Exception:  # noqa
            pass
        try:
            user32.UnregisterHotKey(None, HOTKEY_ID_TRANSLATE)
        except Exception:  # noqa
            pass
        try:
            user32.UnregisterHotKey(None, HOTKEY_ID_CASE)
        except Exception:  # noqa
            pass
        try:
            exit_event.set()
            _invoke_exit_handlers()
        except Exception:  # noqa
            pass
