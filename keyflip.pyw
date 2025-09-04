# keyflip_winapi.py
from __future__ import annotations
import ctypes
import json
import logging
import os
import sys
import threading
import time
import uuid
from logging.handlers import RotatingFileHandler
from typing import Dict, Optional, Tuple

# внешние библиотеки
try:
    import pyperclip
    from PIL import Image, ImageDraw
    import pystray
    import psutil
    import win32event, win32api, win32con
except Exception:
    print(
        "❌ Требуются библиотеки: pyperclip, Pillow, pystray, psutil, pywin32\n"
        "Установи командой:\n"
        "   pip install pyperclip pillow pystray psutil pywin32"
    )
    raise

# ctypes helpers
user32 = ctypes.windll.user32
from ctypes import wintypes

# --- Константы / пути ---
APP_NAME = "KeyFlip"
MUTEX_NAME = r"Global\KeyFlipMutex"
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
LOG_FILE = os.path.join(BASE_DIR, "keyflip.log")
ICON_ON = os.path.join(BASE_DIR, "icon_on.ico")
ICON_OFF = os.path.join(BASE_DIR, "icon_off.ico")
CONFIG_FILE = os.path.join(BASE_DIR, "keyflip.json")

# --- Один экземпляр (mutex) ---
mutex_handle = win32event.CreateMutex(None, False, MUTEX_NAME)
ERROR_ALREADY_EXISTS = 183
if win32api.GetLastError() == ERROR_ALREADY_EXISTS:
    try:
        win32api.MessageBox(0, f"Программа {APP_NAME} уже запущена!", APP_NAME, 0)
    except Exception:
        print(f"{APP_NAME} уже запущена!")
    sys.exit(0)

# --- Логирование ---
logger = logging.getLogger(APP_NAME)
logger.setLevel(logging.DEBUG)
handler = RotatingFileHandler(LOG_FILE, maxBytes=5 * 1024 * 1024, backupCount=3, encoding="utf-8")
fmt = logging.Formatter("%(asctime)s %(levelname)s [%(threadName)s] %(message)s", "%Y-%m-%d %H:%M:%S")
handler.setFormatter(fmt)
logger.addHandler(handler)
console = logging.StreamHandler(sys.stdout)
console.setFormatter(fmt)
console.setLevel(logging.INFO)
logger.addHandler(console)
logger.info("%s запуск — лог: %s", APP_NAME, LOG_FILE)

# ------------ Mapping: explicit, детерминированные таблицы ------------
EN_TO_RU: Dict[str, str] = {
    # letters
    'q': 'й', 'w': 'ц', 'e': 'у', 'r': 'к', 't': 'е', 'y': 'н', 'u': 'г', 'i': 'ш', 'o': 'щ', 'p': 'з',
    'a': 'ф', 's': 'ы', 'd': 'в', 'f': 'а', 'g': 'п', 'h': 'р', 'j': 'о', 'k': 'л', 'l': 'д',
    'z': 'я', 'x': 'ч', 'c': 'с', 'v': 'м', 'b': 'и', 'n': 'т', 'm': 'ь',

    # punctuation & digits typical mapping (Windows Russian)
    '1': '1', '!': '!',
    '2': '2', '@': '"',
    '3': '3', '#': '№',
    '4': '4', '$': ';',
    '5': '5', '%': '%',
    '6': '6', '^': ':',
    '7': '7', '&': '?',
    '8': '8', '*': '*',
    '9': '9', '(': '(',
    '0': '0', ')': ')',

    ',': 'б', '<': 'Б',
    '.': 'ю', '>': 'Ю',
    '/': '.', '?': ',',
    ';': 'ж', ':': 'Ж',
    "'": 'э', '"': 'Э',
    '[': 'х', '{': 'Х',
    ']': 'ъ', '}': 'Ъ',
    '\\': '\\', '|': '|',
    '`': 'ё', '~': 'Ё',
    '-': '-', '_': '_',
    '=': '=', '+': '+',
}

# Добавим заглавные соответствия для букв/символов
for k, v in list(EN_TO_RU.items()):
    if k.isalpha():
        EN_TO_RU[k.upper()] = v.upper()

# RU -> EN (обратная таблица)
RU_TO_EN: Dict[str, str] = {v: k for k, v in EN_TO_RU.items()}
RU_TO_EN['ё'] = '`'
RU_TO_EN['Ё'] = '~'

def transform_text_by_keyboard_layout_based_on_hkl(s: str, hkl: int) -> str:
    lang = hkl & 0xFFFF
    if lang == 0x0419:
        mapping = RU_TO_EN
        reverse = EN_TO_RU
        prefer = 'RU->EN'
    else:
        mapping = EN_TO_RU
        reverse = RU_TO_EN
        prefer = 'EN->RU'

    out = []
    for ch in s:
        if ch in mapping:
            out.append(mapping[ch])
            continue
        if ch in reverse:
            out.append(reverse[ch])
            continue
        out.append(ch)
    logger.debug("transform: used %s mapping", prefer)
    return ''.join(out)

# ------------ Active window info ------------
def get_active_window_info() -> Tuple[str, int, Optional[str]]:
    try:
        hwnd = user32.GetForegroundWindow()
    except Exception:
        return "<unknown>", 0, None
    title = "<unknown>"
    try:
        length = user32.GetWindowTextLengthW(hwnd)
        buff = ctypes.create_unicode_buffer(length + 1)
        user32.GetWindowTextW(hwnd, buff, length + 1)
        title = buff.value
    except Exception:
        title = "<unknown>"
    pid = 0
    proc_name = None
    try:
        pid_c = ctypes.c_ulong()
        user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid_c))
        pid = int(pid_c.value)
        if psutil:
            try:
                proc = psutil.Process(pid)
                proc_name = proc.name()
            except Exception:
                proc_name = None
    except Exception:
        pass
    return title, pid, proc_name

# ------------ Helper: эмуляция клавиш через keybd_event ------------
def _key_down(vk: int):
    user32.keybd_event(vk, 0, 0, 0)

def _key_up(vk: int):
    KEYEVENTF_KEYUP = 0x0002
    user32.keybd_event(vk, 0, KEYEVENTF_KEYUP, 0)

def send_ctrl_c():
    VK_CONTROL = 0x11
    _key_down(VK_CONTROL)
    _key_down(ord('C'))
    time.sleep(0.015)
    _key_up(ord('C'))
    _key_up(VK_CONTROL)

def send_ctrl_v():
    VK_CONTROL = 0x11
    _key_down(VK_CONTROL)
    _key_down(ord('V'))
    time.sleep(0.015)
    _key_up(ord('V'))
    _key_up(VK_CONTROL)

def send_delete():
    VK_DELETE = 0x2E
    _key_down(VK_DELETE)
    time.sleep(0.01)
    _key_up(VK_DELETE)

# ------------ Clipboard helper (бережно) ------------
def safe_restore_clipboard(old_clip: Optional[str]) -> None:
    try:
        if old_clip is None:
            pyperclip.copy('')
            logger.debug("safe_restore_clipboard: восстановлен пустой буфер")
        else:
            pyperclip.copy(old_clip)
            logger.debug("safe_restore_clipboard: буфер восстановлен (len=%s)", len(old_clip))
    except Exception as e:
        logger.exception("safe_restore_clipboard: не удалось восстановить буфер: %s", e)

def safe_copy_from_selection(timeout_per_attempt: float = 0.6, max_attempts: int = 2) -> str:
    title, pid, proc_name = get_active_window_info()
    logger.debug("safe_copy: start. active window: %r pid=%s proc=%r", title, pid, proc_name)

    try:
        old_clip = pyperclip.paste()
    except Exception as e:
        logger.debug("safe_copy: can't read old clipboard: %s", e)
        old_clip = None

    try:
        initial_hwnd = user32.GetForegroundWindow()
    except Exception as e:
        logger.debug("safe_copy: GetForegroundWindow failed: %s", e)
        initial_hwnd = None

    def foreground_changed() -> bool:
        try:
            cur = user32.GetForegroundWindow()
            return initial_hwnd is not None and cur is not None and cur != initial_hwnd
        except Exception:
            return False

    try:
        old_seq = user32.GetClipboardSequenceNumber()
    except Exception as e:
        logger.debug("safe_copy: GetClipboardSequenceNumber failed: %s", e)
        old_seq = None

    last_exception = None

    if old_seq is not None:
        for attempt in range(1, max_attempts + 1):
            if foreground_changed():
                logger.debug("safe_copy: foreground changed before Ctrl+C -> abort")
                return ""
            try:
                send_ctrl_c()
                logger.debug("safe_copy: sent Ctrl+C (attempt %d)", attempt)
            except Exception as e:
                logger.exception("safe_copy: ctrl+c exception: %s", e)
                last_exception = e

            t0 = time.time()
            changed = False
            while time.time() - t0 < timeout_per_attempt:
                if foreground_changed():
                    logger.debug("safe_copy: foreground changed during wait -> abort")
                    return ""
                try:
                    seq = user32.GetClipboardSequenceNumber()
                except Exception:
                    seq = None
                if seq is not None and seq != old_seq:
                    changed = True
                    break
                time.sleep(0.02)

            if changed:
                try:
                    cur = pyperclip.paste()
                    if cur == "":
                        logger.debug("safe_copy: clipboard changed but empty -> no selection")
                        return ""
                    logger.debug("safe_copy: buffer changed after Ctrl+C (len=%d)", len(cur))
                    return cur
                except Exception as e:
                    logger.exception("safe_copy: paste() after ctrl+c failed: %s", e)
                    last_exception = e

            # fallback Ctrl+Insert
            if foreground_changed():
                logger.debug("safe_copy: foreground changed before Ctrl+Insert -> abort")
                return ""
            try:
                VK_CONTROL = 0x11
                VK_INSERT = 0x2D
                _key_down(VK_CONTROL)
                _key_down(VK_INSERT)
                time.sleep(0.01)
                _key_up(VK_INSERT)
                _key_up(VK_CONTROL)
                logger.debug("safe_copy: sent Ctrl+Insert (attempt %d)", attempt)
            except Exception as e:
                logger.exception("safe_copy: ctrl+insert exception: %s", e)
                last_exception = e

            t0 = time.time()
            changed = False
            while time.time() - t0 < timeout_per_attempt:
                if foreground_changed():
                    logger.debug("safe_copy: foreground changed during wait after Insert -> abort")
                    return ""
                try:
                    seq = user32.GetClipboardSequenceNumber()
                except Exception:
                    seq = None
                if seq is not None and seq != old_seq:
                    changed = True
                    break
                time.sleep(0.02)

            if changed:
                try:
                    cur = pyperclip.paste()
                    if cur == "":
                        logger.debug("safe_copy: clipboard changed after insert but empty -> no selection")
                        return ""
                    logger.debug("safe_copy: buffer changed after Ctrl+Insert (len=%d)", len(cur))
                    return cur
                except Exception as e:
                    logger.exception("safe_copy: paste() after insert failed: %s", e)
                    last_exception = e

            logger.debug("safe_copy: no clipboard change after attempt %d", attempt)
            break

        logger.debug("safe_copy: sequence not changed -> no selection")
        return ""

    # ---- fallback: sentinel approach but careful ----
    if foreground_changed():
        logger.debug("safe_copy: foreground changed before fallback -> abort")
        return ""

    sentinel = f"__{APP_NAME.upper()}_SENTINEL__{uuid.uuid4()}__"
    try:
        pyperclip.copy(sentinel)
        logger.debug("safe_copy: sentinel written for fallback")
    except Exception as e:
        logger.exception("safe_copy: cannot write sentinel: %s", e)
        return ""

    try:
        if foreground_changed():
            logger.debug("safe_copy: foreground changed before fallback Ctrl+C -> restore and abort")
            safe_restore_clipboard(old_clip)
            return ""
        send_ctrl_c()
        logger.debug("safe_copy: sent Ctrl+C (fallback)")
    except Exception as e:
        logger.exception("safe_copy: ctrl+c exception (fallback): %s", e)
        last_exception = e

    t0 = time.time()
    while time.time() - t0 < timeout_per_attempt:
        if foreground_changed():
            logger.debug("safe_copy: foreground changed during fallback wait -> restore and abort")
            try:
                safe_restore_clipboard(old_clip)
            except Exception:
                logger.exception("safe_copy: failed restore old_clip on abort")
            return ""
        try:
            cur = pyperclip.paste()
        except Exception as e:
            logger.debug("safe_copy: paste exception while waiting (fallback): %s", e)
            cur = None
        if cur is None:
            time.sleep(0.02)
            continue
        if cur != sentinel:
            if cur == "":
                try:
                    safe_restore_clipboard(old_clip)
                except Exception:
                    logger.exception("safe_copy: failed restore old_clip after empty fallback")
                return ""
            try:
                safe_restore_clipboard(old_clip)
            except Exception:
                logger.exception("safe_copy: failed restore old_clip after success fallback")
            logger.debug("safe_copy: buffer changed (fallback) len=%d", len(cur))
            return cur
        time.sleep(0.02)

    # last attempt: Ctrl+Insert
    if foreground_changed():
        logger.debug("safe_copy: foreground changed before fallback Insert -> restore and abort")
        try:
            safe_restore_clipboard(old_clip)
        except Exception:
            logger.exception("safe_copy: failed restore old_clip on abort2")
        return ""

    try:
        VK_CONTROL = 0x11
        VK_INSERT = 0x2D
        _key_down(VK_CONTROL)
        _key_down(VK_INSERT)
        time.sleep(0.01)
        _key_up(VK_INSERT)
        _key_up(VK_CONTROL)
        logger.debug("safe_copy: sent Ctrl+Insert (fallback last)")
    except Exception as e:
        logger.exception("safe_copy: ctrl+insert exception (fallback last): %s", e)
        last_exception = e

    t0 = time.time()
    while time.time() - t0 < timeout_per_attempt:
        if foreground_changed():
            logger.debug("safe_copy: foreground changed during fallback insert wait -> restore and abort")
            try:
                safe_restore_clipboard(old_clip)
            except Exception:
                logger.exception("safe_copy: failed restore old_clip on abort3")
            return ""
        try:
            cur = pyperclip.paste()
        except Exception as e:
            logger.debug("safe_copy: paste exception after insert (fallback): %s", e)
            cur = None
        if cur is None:
            time.sleep(0.02)
            continue
        if cur != sentinel:
            if cur == "":
                try:
                    safe_restore_clipboard(old_clip)
                except Exception:
                    logger.exception("safe_copy: failed restore old_clip after empty fallback2")
                return ""
            try:
                safe_restore_clipboard(old_clip)
            except Exception:
                logger.exception("safe_copy: failed restore old_clip after success fallback2")
            logger.debug("safe_copy: buffer changed after ctrl+insert (fallback) len=%d", len(cur))
            return cur
        time.sleep(0.02)

    # nothing worked
    try:
        safe_restore_clipboard(old_clip)
    except Exception:
        logger.exception("safe_copy: failed final restore old_clip")
    logger.warning("safe_copy: НЕ удалось получить выделение. last_exc=%r", repr(last_exception))
    return ""

# ------------ Config (enabled) ------------
_enabled_lock = threading.Lock()
_enabled: bool = True  # default; will be overwritten by load_config()

def load_config():
    """
    При старте всегда включаем функционал (enabled=True) и сохраняем это в файл.
    Это гарантирует, что даже если в конфиге было false — после запуска он станет true.
    """
    global _enabled
    try:
        # Принудительно включаем при старте
        with _enabled_lock:
            _enabled = True

        # Сохраняем текущее состояние (перезапишем конфиг)
        save_config()
        logger.debug("config: startup forced enabled=True (config saved)")
    except Exception:
        logger.exception("config: forced enable on startup failed; using default enabled=True")
        with _enabled_lock:
            _enabled = True


def save_config():
    try:
        with _enabled_lock:
            j = {"enabled": bool(_enabled)}
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(j, f)
        logger.debug("config: saved enabled=%s to %s", j["enabled"], CONFIG_FILE)
    except Exception:
        logger.exception("config: save failed")

def is_enabled() -> bool:
    with _enabled_lock:
        return bool(_enabled)

def set_enabled(val: bool):
    global _enabled
    with _enabled_lock:
        _enabled = bool(val)
    save_config()
    logger.info("state: enabled set to %s", _enabled)

# Load config on startup
load_config()

# ------------ Hotkey registration helpers ------------
HOTKEY_ID_F4 = 1
HOTKEY_ID_F10 = 2
MOD_NONE = 0
VK_F4 = 0x73
VK_F10 = 0x79

def register_f4():
    try:
        ok = user32.RegisterHotKey(None, HOTKEY_ID_F4, MOD_NONE, VK_F4)
        if ok:
            logger.debug("register_f4: F4 registered")
        else:
            logger.error("register_f4: failed to register F4 (maybe already registered by another app)")
    except Exception:
        logger.exception("register_f4: exception while registering F4")

def unregister_f4():
    try:
        ok = user32.UnregisterHotKey(None, HOTKEY_ID_F4)
        if ok:
            logger.debug("unregister_f4: F4 unregistered")
        else:
            logger.debug("unregister_f4: UnregisterHotKey returned False (probably not registered)")
    except Exception:
        logger.exception("unregister_f4: exception while unregistering F4")

# ------------ Handler & Hotkey logic ------------
exit_event = threading.Event()
_handler_lock = threading.Lock()
_last_f4_ts = 0.0
_F4_DEBOUNCE_SEC = 0.6

def _f4_invoker():
    global _last_f4_ts
    if not is_enabled():
        logger.debug("F4: ignored because enabled==False")
        return

    now = time.time()
    if now - _last_f4_ts < _F4_DEBOUNCE_SEC:
        logger.debug("F4: debounce ignored (delta=%.3f)", now - _last_f4_ts)
        return
    _last_f4_ts = now

    if not _handler_lock.acquire(blocking=False):
        logger.debug("F4: handler busy, ignored")
        return

    def _worker():
        try:
            handle_hotkey_transform()
        finally:
            try:
                _handler_lock.release()
            except Exception:
                pass

    threading.Thread(target=_worker, daemon=True).start()

def handle_hotkey_transform():
    logger.info("F4 обработка — начинаю преобразование выделения")
    try:
        if not is_enabled():
            logger.debug("handle: disabled, returning")
            return

        title, pid, proc_name = get_active_window_info()
        logger.info("handle: active window: %r pid=%s proc=%r", title, pid, proc_name)

        # save current clipboard to restore later
        try:
            saved = pyperclip.paste()
            logger.debug("handle: сохранён буфер (len=%s)", None if saved is None else len(saved))
        except Exception as e:
            logger.exception("handle: не удалось сохранить буфер: %s", e)
            saved = None

        # get foreground thread HKL so we know direction
        try:
            hwnd = user32.GetForegroundWindow()
            pid_c = ctypes.c_ulong()
            thread_id = user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid_c))
            hkl = user32.GetKeyboardLayout(thread_id) & 0xFFFF
            logger.debug("handle: foreground thread hkl=0x%04X", hkl)
        except Exception:
            hkl = 0
            logger.exception("handle: не удалось получить HKL, предполагаем EN")

        # read selection carefully
        selected = safe_copy_from_selection(timeout_per_attempt=0.6, max_attempts=2)
        if selected is None:
            selected = ""
        logger.info("handle: прочитано выделение (len=%d)", len(selected))

        # if nothing selected -> do nothing (don't loop sentinel into clipboard!)
        if not selected:
            logger.info("handle: выделение пустое — ничего не делаю. Восстанавливаю буфер.")
            safe_restore_clipboard(saved)
            return

        # choose mapping based on hkl
        converted = transform_text_by_keyboard_layout_based_on_hkl(selected, hkl)
        logger.info("handle: преобразование выполнено (len=%d).", len(converted))
        logger.debug("handle: исходное=%r converted=%r", selected, converted)

        if converted == selected:
            logger.info("handle: преобразованный текст совпадает с исходным — ничего не меняю.")
            safe_restore_clipboard(saved)
            return

        # delete selection
        try:
            send_delete()
            time.sleep(0.02)
        except Exception as e:
            logger.exception("handle: delete exception: %s", e)

        # paste via clipboard + Ctrl+V (reliable across apps)
        try:
            pyperclip.copy(converted)
            time.sleep(0.02)
            send_ctrl_v()
            logger.debug("handle: вставил через буфер (ctrl+v)")
        except Exception as e:
            logger.exception("handle: paste exception: %s", e)

        # переключаем раскладку в окне (меняем на противоположную)
        try:
            if hwnd:
                if hkl == 0x0419:
                    new_klid = "00000409"
                else:
                    new_klid = "00000419"
                hkl_new = user32.LoadKeyboardLayoutW(new_klid, 1)
                WM_INPUTLANGCHANGEREQUEST = 0x0050
                res = user32.PostMessageW(hwnd, WM_INPUTLANGCHANGEREQUEST, 0, ctypes.c_void_p(hkl_new))
                if res == 0:
                    logger.debug("handle: PostMessageW failed (res=0), falling back to ActivateKeyboardLayout")
                    try:
                        user32.ActivateKeyboardLayout(hkl_new, 0)
                        logger.debug("handle: ActivateKeyboardLayout used as fallback")
                    except Exception:
                        logger.exception("handle: fallback ActivateKeyboardLayout also failed")
                else:
                    logger.debug("handle: posted WM_INPUTLANGCHANGEREQUEST -> success")
        except Exception:
            logger.exception("handle: switch layout failed (foreground-thread)")

        # restore saved clipboard
        safe_restore_clipboard(saved)
        logger.info("handle: завершено успешно, буфер восстановлен")
    except Exception as e:
        logger.exception("handle_hotkey_transform: исключение при обработке: %s", e)
        try:
            safe_restore_clipboard(None)
        except Exception:
            pass

# ------------ Tray / Icon (с пунктом Вкл/Выкл и иконками) ----------------
def prepare_tray_icon_image(enabled: Optional[bool] = None) -> Image.Image:
    if enabled is None:
        enabled = is_enabled()

    path = ICON_ON if enabled else ICON_OFF
    if os.path.isfile(path):
        try:
            img = Image.open(path)
            logger.debug("Tray icon: using %s", path)
            return img
        except Exception:
            logger.exception("Tray icon: cannot open %s", path)

    size = (64, 64)
    bg = (76, 175, 80, 255) if enabled else (220, 53, 69, 255)
    img = Image.new('RGBA', size, bg)
    d = ImageDraw.Draw(img)
    try:
        d.rectangle((8, 16, 56, 48), outline=(255, 255, 255), width=2)
        d.text((16, 12), "KF", fill=(255, 255, 255))
    except Exception:
        pass
    return img

def enabled_menu_text(item):
    return "❌ Выключить" if is_enabled() else "✅ Включить"

def toggle_enabled(icon, item):
    new_state = not is_enabled()
    set_enabled(new_state)
    logger.info("Tray: toggled enabled -> %s", new_state)
    # immediately update hotkey registration
    try:
        if new_state:
            register_f4()
        else:
            unregister_f4()
    except Exception:
        logger.exception("toggle_enabled: failed to update hotkey registration")

    # Обновляем иконку прямо сейчас (pystray Icon передаётся первым аргументом)
    try:
        if icon is not None:
            icon.icon = prepare_tray_icon_image(new_state)
    except Exception:
        logger.exception("Tray: failed to update icon image after toggle")

def on_exit(icon=None, item=None):
    logger.info("Выход запрошен (через трей/F10)")
    exit_event.set()
    try:
        if icon:
            icon.stop()
    except Exception:
        logger.exception("on_exit: ошибка при остановке иконки")
    try:
        win32event.ReleaseMutex(mutex_handle)
    except Exception:
        pass
    logger.info("Завершение работы.")
    return

def tray_worker():
    img = prepare_tray_icon_image()
    menu = pystray.Menu(
        pystray.MenuItem(enabled_menu_text, toggle_enabled),
        pystray.MenuItem("Выход", on_exit)
    )
    icon = pystray.Icon(APP_NAME, img, APP_NAME, menu=menu)
    try:
        icon.icon = prepare_tray_icon_image(is_enabled())
    except Exception:
        logger.exception("tray_worker: failed to set initial icon")

    try:
        logger.debug("tray_worker: запуск иконки")
        icon.run()
    except Exception as e:
        logger.exception("tray_worker: исключение: %s", e)

# ------------ WinAPI hotkey loop ------------
def win_hotkey_loop():
    # Register F10 always (for exit)
    if not user32.RegisterHotKey(None, HOTKEY_ID_F10, MOD_NONE, VK_F10):
        logger.error("Не удалось зарегистрировать F10 через RegisterHotKey")
    else:
        logger.debug("Зарегистрирован F10 via RegisterHotKey")

    # Register F4 only if enabled
    if is_enabled():
        register_f4()
    else:
        # ensure it's not registered
        try:
            user32.UnregisterHotKey(None, HOTKEY_ID_F4)
        except Exception:
            pass

    try:
        msg = wintypes.MSG()
        while not exit_event.is_set():
            has = user32.GetMessageW(ctypes.byref(msg), None, 0, 0)
            if has == 0:
                break
            if msg.message == 0x0312:  # WM_HOTKEY
                hotkey_id = msg.wParam
                if hotkey_id == HOTKEY_ID_F4:
                    logger.debug("WM_HOTKEY: F4 received")
                    _f4_invoker()
                elif hotkey_id == HOTKEY_ID_F10:
                    logger.debug("WM_HOTKEY: F10 received -> exit")
                    on_exit()
            user32.TranslateMessage(ctypes.byref(msg))
            user32.DispatchMessageW(ctypes.byref(msg))
    finally:
        try:
            user32.UnregisterHotKey(None, HOTKEY_ID_F4)
            user32.UnregisterHotKey(None, HOTKEY_ID_F10)
        except Exception:
            pass

# ------------ Main ------------
def main():
    t = threading.Thread(target=tray_worker, daemon=True, name="TrayThread")
    t.start()
    logger.debug("main: трей поток запущен")

    t2 = threading.Thread(target=win_hotkey_loop, daemon=True, name="WinHotkeyThread")
    t2.start()
    logger.debug("main: WinAPI hotkey thread started")

    logger.info("%s запущен. Нажми F4 чтобы преобразовать выделение, F10 чтобы выйти.", APP_NAME)
    try:
        while not exit_event.is_set():
            time.sleep(0.3)
    except KeyboardInterrupt:
        logger.info("main: KeyboardInterrupt")
    finally:
        try:
            win32event.ReleaseMutex(mutex_handle)
        except Exception:
            pass
        logger.info("main: завершение")

if __name__ == "__main__":
    main()
