import ctypes
import json
import logging
import os
import sys
import threading
import time
import uuid
from logging.handlers import RotatingFileHandler
from typing import Optional

# внешние библиотеки
try:
    import pyperclip
    from PIL import Image, ImageDraw
    from pystray import Menu, MenuItem, Icon
    import psutil
    import win32event, win32api, win32con
except Exception:  # noqa
    print(
        "❌ Требуются библиотеки: pyperclip, Pillow, pystray, psutil, pywin32\n"
        "Установи командой:\n"
        "   pip install pyperclip pillow pystray psutil pywin32"
    )
    raise

# ctypes helpers
user32 = ctypes.windll.user32
kernel32 = ctypes.windll.kernel32
from ctypes import wintypes

# алиас для указателя (устраняет замечания анализаторов)
ULONG_PTR = wintypes.WPARAM

# ---------------- Константы / пути ----------------
APP_NAME = "KeyFlip"
MUTEX_NAME = r"Global\KeyFlipMutex"
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
LOG_FILE = os.path.join(BASE_DIR, "keyflip.log")
ICON_ON = os.path.join(BASE_DIR, "icon_on.ico")
ICON_OFF = os.path.join(BASE_DIR, "icon_off.ico")
CONFIG_FILE = os.path.join(BASE_DIR, "keyflip.json")

# ---------------- Константы клавиш / SendInput ----------------
INPUT_KEYBOARD = 1
KEYEVENTF_UNICODE = 0x0004
KEYEVENTF_KEYUP = 0x0002

# Общие VK (используются в разных местах — определяем глобально)
VK_CONTROL = 0x11
VK_MENU = 0x12
VK_SHIFT = 0x10
VK_DELETE = 0x2E
VK_INSERT = 0x2D
VK_ESCAPE = 0x1B

# Для регистратора сочетаний
WH_KEYBOARD_LL = 13
WM_KEYDOWN = 0x0100
WM_SYSKEYDOWN = 0x0104
WM_KEYUP = 0x0101
WM_SYSKEYUP = 0x0105

# ---------------- Помощники работы с конфигом (JSON) ----------------
def read_json_config() -> dict:
    """Прочитать JSON-конфиг, вернуть словарь (или пустой)"""
    try:
        if not os.path.isfile(CONFIG_FILE):
            return {}
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            return json.load(f) or {}
    except Exception:  # noqa
        return {}

def write_json_config(j: dict) -> bool:
    """Атомарная запись JSON: в tmp-файл, потом replace"""
    try:
        tmp = CONFIG_FILE + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(j, f, ensure_ascii=False, indent=2)
        try:
            os.replace(tmp, CONFIG_FILE)
        except Exception:  # noqa
            # fallback: удалить старый и переименовать
            if os.path.exists(CONFIG_FILE):
                try:
                    os.remove(CONFIG_FILE)
                except Exception:  # noqa
                    pass
            os.rename(tmp, CONFIG_FILE)
        return True
    except Exception:  # noqa
        logger.exception("config: write_json_config failed")
        return False

def read_file_logging_flag() -> bool:
    try:
        j = read_json_config()
        return bool(j.get("file_logging", False))
    except Exception:  # noqa
        return False

def write_file_logging_flag(val: bool) -> bool:
    try:
        j = read_json_config()
        j.setdefault("enabled", True)
        j.setdefault("autorun", False)
        j["file_logging"] = bool(val)
        ok = write_json_config(j)
        if ok:
            logger.info("config: записан file_logging=%s", j["file_logging"])
        return ok
    except Exception:  # noqa
        logger.exception("config: write_file_logging_flag failed")
        return False

def read_autorun_flag() -> bool:
    try:
        j = read_json_config()
        return bool(j.get("autorun", False))
    except Exception:  # noqa
        return False

def write_autorun_flag(val: bool) -> bool:
    try:
        j = read_json_config()
        j.setdefault("enabled", True)
        j.setdefault("file_logging", False)
        j["autorun"] = bool(val)
        ok = write_json_config(j)
        if ok:
            logger.info("config: записан autorun=%s", j["autorun"])
        return ok
    except Exception:  # noqa
        logger.exception("config: write_autorun_flag failed")
        return False

def default_exit_hotkey() -> dict:
    """Значение по умолчанию для комбинации выхода"""
    return {"modifiers": [], "key": "F10"}

def read_exit_hotkey() -> dict:
    """Прочитать exit_hotkey из конфига — вернуть {'modifiers': [...], 'key': 'F10'}"""
    try:
        j = read_json_config()
        eh = j.get("exit_hotkey")
        if not isinstance(eh, dict):
            return default_exit_hotkey()
        mods = eh.get("modifiers", []) or []
        key = eh.get("key", "F10") or "F10"
        return {"modifiers": list(mods), "key": str(key)}
    except Exception:  # noqa
        return default_exit_hotkey()

def write_exit_hotkey(modifiers: list[str], key: str) -> bool:
    """Нормализовать и записать exit_hotkey в конфиг"""
    try:
        j = read_json_config()
        j.setdefault("enabled", True)
        j.setdefault("file_logging", False)
        j.setdefault("autorun", False)
        mods: list[str] = []
        for m in modifiers or []:
            mm = (m or "").upper()
            if mm in ("CTRL", "CONTROL"):
                mods.append("CTRL")
            elif mm in ("ALT", "MENU"):
                mods.append("ALT")
            elif mm in ("SHIFT",):
                mods.append("SHIFT")
            elif mm in ("WIN", "WINDOWS"):
                mods.append("WIN")
        j["exit_hotkey"] = {"modifiers": mods, "key": str(key)}
        ok = write_json_config(j)
        if ok:
            logger.info("config: записан exit_hotkey=%s", j["exit_hotkey"])
        return ok
    except Exception:  # noqa
        logger.exception("config: write_exit_hotkey failed")
        return False

# ---------------- translate hotkey ----------------
def default_translate_hotkey() -> dict:
    """Значение по умолчанию для комбинации перевода (F4)"""
    return {"modifiers": [], "key": "F4"}

def read_translate_hotkey() -> dict:
    """Прочитать translate_hotkey из конфига — вернуть {'modifiers': [...], 'key': 'F4'}"""
    try:
        j = read_json_config()
        th = j.get("translate_hotkey")
        if not isinstance(th, dict):
            return default_translate_hotkey()
        mods = th.get("modifiers", []) or []
        key = th.get("key", "F4") or "F4"
        return {"modifiers": list(mods), "key": str(key)}
    except Exception:  # noqa
        return default_translate_hotkey()

def write_translate_hotkey(modifiers: list[str], key: str) -> bool:
    """Нормализовать и записать translate_hotkey в конфиг"""
    try:
        j = read_json_config()
        j.setdefault("enabled", True)
        j.setdefault("file_logging", False)
        j.setdefault("autorun", False)
        mods: list[str] = []
        for m in modifiers or []:
            mm = (m or "").upper()
            if mm in ("CTRL", "CONTROL"):
                mods.append("CTRL")
            elif mm in ("ALT", "MENU"):
                mods.append("ALT")
            elif mm in ("SHIFT",):
                mods.append("SHIFT")
            elif mm in ("WIN", "WINDOWS"):
                mods.append("WIN")
        j["translate_hotkey"] = {"modifiers": mods, "key": str(key)}
        ok = write_json_config(j)
        if ok:
            logger.info("config: записан translate_hotkey=%s", j["translate_hotkey"])
        return ok
    except Exception:  # noqa
        logger.exception("config: write_translate_hotkey failed")
        return False

# ---------------- Autorun: создание/удаление ярлыка в автозагрузке ----------------
def _startup_shortcut_path() -> str:
    appdata = os.environ.get("APPDATA") or os.path.expanduser("~\\AppData\\Roaming")
    startup_dir = os.path.join(appdata, "Microsoft", "Windows", "Start Menu", "Programs", "Startup")
    return os.path.join(startup_dir, f"{APP_NAME}.lnk")

def create_startup_shortcut() -> bool:
    """Создать .lnk в папке автозагрузки текущего пользователя"""
    shortcut_path = _startup_shortcut_path()
    try:
        import win32com.client  # type: ignore
        shell = win32com.client.Dispatch("WScript.Shell")
        shortcut = shell.CreateShortcut(shortcut_path)
        target = sys.executable
        script = os.path.abspath(sys.argv[0])
        shortcut.TargetPath = target
        if script.endswith(".py") or script.endswith(".pyw"):
            shortcut.Arguments = f'"{script}"'
        else:
            shortcut.Arguments = ""
        shortcut.WorkingDirectory = BASE_DIR
        if os.path.isfile(ICON_ON):
            shortcut.IconLocation = ICON_ON
        else:
            shortcut.IconLocation = target
        shortcut.Save()
        logger.info("startup: ярлык создан: %s", shortcut_path)
        return True
    except Exception:  # noqa
        logger.exception("startup: не удалось создать ярлык автозагрузки")
        return False

def remove_startup_shortcut() -> bool:
    """Удалить .lnк из автозагрузки (если есть)"""
    sp = _startup_shortcut_path()
    try:
        if os.path.isfile(sp):
            os.remove(sp)
            logger.info("startup: ярлык удалён: %s", sp)
        else:
            logger.debug("startup: ярлык не найден: %s", sp)
        return True
    except Exception:  # noqa
        logger.exception("startup: не удалось удалить ярлык: %s", sp)
        return False

def apply_autorun_setting() -> None:
    """Применить текущее значение autorun (создать/удалить ярлык)"""
    try:
        logger.warning("AUTORUN: apply_autorun_setting CALLED")
        if read_autorun_flag():
            create_startup_shortcut()
        else:
            remove_startup_shortcut()
    except Exception:  # noqa
        logger.exception("apply_autorun_setting: ошибка при применении autorun")

# ---------------- Один экземпляр (mutex) ----------------
mutex_handle = win32event.CreateMutex(None, False, MUTEX_NAME)
ERROR_ALREADY_EXISTS = 183
if win32api.GetLastError() == ERROR_ALREADY_EXISTS:
    try:
        win32api.MessageBox(0, f"Программа {APP_NAME} уже запущена!", APP_NAME, 0)
    except Exception:  # noqa
        print(f"{APP_NAME} уже запущена!")
    sys.exit(0)

# ---------------- Логирование ----------------
logger = logging.getLogger(APP_NAME)
logger.setLevel(logging.DEBUG)
handler = RotatingFileHandler(LOG_FILE, maxBytes=5 * 1024 * 1024, backupCount=3, encoding="utf-8")
fmt = logging.Formatter("%(asctime)s %(levelname)s [%(threadName)s] %(message)s", "%Y-%m-%d %H:%M:%S")
handler.setFormatter(fmt)

class FileLoggingFilter(logging.Filter):
    """Фильтр, который при каждой записи проверяет флаг file_logging в конфиге"""
    def filter(self, record: logging.LogRecord) -> bool:
        try:
            return read_file_logging_flag()
        except Exception:  # noqa
            return False

handler.addFilter(FileLoggingFilter())
logger.addHandler(handler)
console = logging.StreamHandler(sys.stdout)
console.setFormatter(fmt)
console.setLevel(logging.INFO)
logger.addHandler(console)
logger.info("%s запуск — лог: %s", APP_NAME, LOG_FILE)

# ---------------- Таблицы соответствий раскладок ----------------
EN_TO_RU: dict[str, str] = {
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
RU_TO_EN: dict[str, str] = {v: k for k, v in EN_TO_RU.items()}
RU_TO_EN['ё'] = '`'
RU_TO_EN['Ё'] = '~'

def transform_text_by_keyboard_layout_based_on_hkl(s: str, hkl: int) -> str:
    """
    Преобразовать строку в соответствии с раскладкой:
    если LANGID == 0x0419 (русский) — RU->EN, иначе EN->RU.
    """
    lang = hkl & 0xFFFF
    if lang == 0x0419:
        mapping = RU_TO_EN
        reverse = EN_TO_RU
        prefer = 'RU->EN'
    else:
        mapping = EN_TO_RU
        reverse = RU_TO_EN
        prefer = 'EN->RU'
    out: list[str] = []
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

# ---------------- Информация об активном окне ----------------
def get_active_window_info() -> tuple[str, int, Optional[str]]:
    """Вернуть (title, pid, proc_name) активного окна; при ошибке вернуть заглушки"""
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
    proc_name: Optional[str] = None
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

# ---------------- Эмуляция нажатий клавиш ----------------
def _key_down(vk: int) -> None:
    user32.keybd_event(vk, 0, 0, 0)

def _key_up(vk: int) -> None:
    user32.keybd_event(vk, 0, KEYEVENTF_KEYUP, 0)

def send_ctrl_c() -> None:
    _key_down(VK_CONTROL)
    _key_down(ord('C'))
    time.sleep(0.015)
    _key_up(ord('C'))
    _key_up(VK_CONTROL)

def send_ctrl_v() -> None:
    _key_down(VK_CONTROL)
    _key_down(ord('V'))
    time.sleep(0.015)
    _key_up(ord('V'))
    _key_up(VK_CONTROL)

def send_delete() -> None:
    _key_down(VK_DELETE)
    time.sleep(0.01)
    _key_up(VK_DELETE)

# ---------------- SendInput (UNICODE) ----------------
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

def send_unicode_via_sendinput(text: str, delay_between_keys: float = 0.001) -> None:
    """Вставить текст через SendInput (UNICODE) — не трогаем буфер обмена"""
    if not text:
        return
    inputs: list[INPUT] = []
    for ch in text:
        code = ord(ch)
        ki_down = KEYBDINPUT(0, code, KEYEVENTF_UNICODE, 0, 0)
        inp_down = INPUT(INPUT_KEYBOARD, InputUnion(ki=ki_down))
        inputs.append(inp_down)
        ki_up = KEYBDINPUT(0, code, KEYEVENTF_UNICODE | KEYEVENTF_KEYUP, 0, 0)
        inp_up = INPUT(INPUT_KEYBOARD, InputUnion(ki=ki_up))
        inputs.append(inp_up)
    n = len(inputs)
    arr_type = INPUT * n
    arr = arr_type(*inputs)
    p = ctypes.pointer(arr[0])
    sent = SendInput(n, p, ctypes.sizeof(INPUT))
    if sent != n:
        logger.warning("send_unicode_via_sendinput: SendInput sent %d of %d events", sent, n)
    if delay_between_keys > 0:
        time.sleep(delay_between_keys)

# ---------------- Clipboard helper (бережно) ----------------
def safe_copy_from_selection(timeout_per_attempt: float = 0.6, max_attempts: int = 2) -> tuple[str, set]:
    """
    Осторожно скопировать выделение (попытки Ctrl+C, Ctrl+Insert).
    Возвращает (строка, set_of_pressed_vks).
    Мы ВРЕМЕННО отпускаем физически удерживаемые side-specific модификаторы
    и НЕ восстанавливаем их внутри функции — вызывающий должен восстановить их
    после вставки.
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
    def _detect_pressed_modifier_vks() -> set:
        pressed = set()
        try:
            for vk in (VK_LCONTROL, VK_RCONTROL, VK_LMENU, VK_RMENU, VK_LSHIFT, VK_RSHIFT, VK_LWIN, VK_RWIN):
                try:
                    if user32.GetAsyncKeyState(vk) & 0x8000:
                        pressed.add(vk)
                except Exception:
                    pass
        except Exception:
            pass
        return pressed

    # helper: release side-specific modifier VKs
    def _release_vks(vks: set) -> None:
        try:
            for vk in vks:
                try:
                    _key_up(vk)
                except Exception:
                    pass
        except Exception:
            pass

    pressed_vks = _detect_pressed_modifier_vks()
    if pressed_vks:
        logger.debug("safe_copy: temporarily releasing modifiers vks=%r", pressed_vks)
        _release_vks(pressed_vks)

    # --- Основной цикл попыток чтения выделения ---
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
        sentinel = f"__{APP_NAME.upper()}_SENTINEL__{uuid.uuid4()}__"
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
        # Важно: НЕ восстанавливаем модификаторы здесь — вызывающий код должен восстановить их
        logger.debug("safe_copy: returning with pressed_vks=%r", pressed_vks)

# ---------------- Config (enabled + defaults) ----------------
_enabled_lock = threading.Lock()
_enabled: bool = True

def load_config() -> None:
    """Загрузить конфиг, установить дефолты при их отсутствии"""
    global _enabled
    changed = False
    j = read_json_config()
    if not j:
        j = {}
        changed = True
    if "enabled" not in j:
        j["enabled"] = True
        changed = True
    _enabled = bool(j["enabled"])
    if "file_logging" not in j:
        j["file_logging"] = False
        changed = True
    if "autorun" not in j:
        j["autorun"] = False
        changed = True
    if "exit_hotkey" not in j:
        j["exit_hotkey"] = default_exit_hotkey()
        changed = True
    if "translate_hotkey" not in j:
        j["translate_hotkey"] = default_translate_hotkey()
        changed = True
    if changed:
        write_json_config(j)

def save_config() -> None:
    """Сохранить минимальную информацию (enabled)"""
    try:
        with _enabled_lock:
            j = {"enabled": bool(_enabled)}
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(j, f)
        logger.debug("config: saved enabled=%s to %s", j["enabled"], CONFIG_FILE)
    except Exception:  # noqa
        logger.exception("config: save failed")

def is_enabled() -> bool:
    with _enabled_lock:
        return bool(_enabled)

def set_enabled(val: bool) -> None:
    global _enabled
    with _enabled_lock:
        _enabled = bool(val)
    save_config()
    logger.info("state: enabled set to %s", _enabled)

# загрузка конфига и применение autorun при старте
load_config()
try:
    apply_autorun_setting()
except Exception:  # noqa
    logger.exception("startup: apply_autorun_setting on startup failed")

# ---------------- Hotkey: регистрация и обмен сообщениями ----------------
# ID хоткеев: translate (динамический), exit (динамический, должен работать всегда)
HOTKEY_ID_TRANSLATE = 1
HOTKEY_ID_EXIT = 2
MOD_NONE = 0

WM_USER = 0x0400
MSG_REGISTER_TRANSLATE = WM_USER + 1
MSG_UNREGISTER_TRANSLATE = WM_USER + 2
MSG_UPDATE_EXIT_HOTKEY = WM_USER + 3
MSG_UPDATE_TRANSLATE_HOTKEY = WM_USER + 4

HOTKEY_THREAD_ID = 0

def post_register_translate(should_register: bool) -> bool:
    """Послать потоковое сообщение для регистрации/отмены комбинации перевода"""
    global HOTKEY_THREAD_ID
    if HOTKEY_THREAD_ID == 0:
        logger.warning("post_register_translate: HOTKEY_THREAD_ID not ready yet -> cannot post")
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
    """Попросить hotkey-поток перерегистрировать комбинацию выхода (читает из файла)"""
    global HOTKEY_THREAD_ID
    if HOTKEY_THREAD_ID == 0:
        logger.warning("post_update_exit_hotkey: HOTKEY_THREAD_ID not ready yet -> cannot post")
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
    """Попросить hotkey-поток перерегистрировать комбинацию перевода (читает из файла)"""
    global HOTKEY_THREAD_ID
    if HOTKEY_THREAD_ID == 0:
        logger.warning("post_update_translate_hotkey: HOTKEY_THREAD_ID not ready yet -> cannot post")
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

# ---------------- Helpers: key name <-> vk, modifiers mask ----------------
MOD_ALT = 0x0001
MOD_CONTROL = 0x0002
MOD_SHIFT = 0x0004
MOD_WIN = 0x0008

VK_MAP = {
    "F1": 0x70, "F2": 0x71, "F3": 0x72, "F4": 0x73, "F5": 0x74, "F6": 0x75,
    "F7": 0x76, "F8": 0x77, "F9": 0x78, "F10": 0x79, "F11": 0x7A, "F12": 0x7B,
    "ESC": 0x1B, "TAB": 0x09, "ENTER": 0x0D, "RETURN": 0x0D, "SPACE": 0x20,
    "INSERT": 0x2D, "DELETE": 0x2E, "HOME": 0x24, "END": 0x23, "PGUP": 0x21, "PGDN": 0x22,
    "LEFT": 0x25, "UP": 0x26, "RIGHT": 0x27, "DOWN": 0x28,
}

def key_name_to_vk(name: str) -> int:
    """Преобразовать читаемое имя клавиши в VK-код"""
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
    """Преобразовать список модификаторов в маску для RegisterHotKey"""
    mask = 0
    for m in mods:
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
    """Вернуть (mask, vk) из списка модификаторов и ключа"""
    return modifiers_list_to_mask(mods), key_name_to_vk(key)

def hotkeys_conflict(mask1: int, vk1: int, mask2: int, vk2: int) -> bool:
    """Проверка, конфликты ли две комбинации (совпадают ли mask и vk)"""
    return mask1 == mask2 and vk1 == vk2

# ---------------- Registration logic (exit & translate) ----------------
def update_exit_hotkey_in_thread() -> None:
    """В hotkey-потоке: снять старую комбинацию выхода и зарегистрировать новую"""
    try:
        try:
            user32.UnregisterHotKey(None, HOTKEY_ID_EXIT)
        except Exception:  # noqa
            pass
        eh = read_exit_hotkey()
        mods = eh.get("modifiers", []) or []
        key = eh.get("key", "F10") or "F10"
        mask, vk = hotkey_tuple_from_config(mods, key)

        # проверить конфликт с translate (если translate уже зарегистрирован)
        th = read_translate_hotkey()
        th_mask, th_vk = hotkey_tuple_from_config(th.get("modifiers", []) or [], th.get("key", "F4"))
        if hotkeys_conflict(mask, vk, th_mask, th_vk):
            # конфликт — сообщаем и не регистрируем
            logger.error("update_exit_hotkey_in_thread: конфлик с translate hotkey; пропускаю регистрацию exit")
            try:
                win32api.MessageBox(0,
                    f"Не удалось зарегистрировать комбинацию выхода { '+'.join(mods) + '+' if mods else ''}{key} — конфликт с комбинацией перевода.",
                    APP_NAME, 0)
            except Exception:
                pass
            return

        ok = user32.RegisterHotKey(None, HOTKEY_ID_EXIT, mask, vk)
        if ok:
            logger.debug("update_exit_hotkey_in_thread: registered exit hotkey %s + %s (mask=0x%X vk=0x%X)",
                         "+".join(mods) if mods else "(no modifiers)", key, mask, vk)
        else:
            err = kernel32.GetLastError()
            logger.error("update_exit_hotkey_in_thread: failed to register exit hotkey (err=%d) mask=0x%X vk=0x%X",
                         err, mask, vk)
            try:
                win32api.MessageBox(0,
                    f"Не удалось зарегистрировать комбинацию выхода { '+'.join(mods) + '+' if mods else ''}{key}.\nКод ошибки: {err}",
                    APP_NAME, 0)
            except Exception:
                pass
    except Exception:  # noqa
        logger.exception("update_exit_hotkey_in_thread: exception")

def update_translate_hotkey_in_thread() -> None:
    """В hotkey-потоке: снять старую комбинацию перевода и зарегистрировать новую (с учётом enabled)"""
    try:
        try:
            user32.UnregisterHotKey(None, HOTKEY_ID_TRANSLATE)
        except Exception:  # noqa
            pass
        th = read_translate_hotkey()
        mods = th.get("modifiers", []) or []
        key = th.get("key", "F4") or "F4"
        mask, vk = hotkey_tuple_from_config(mods, key)

        # если app выключен — не регистрируем translate
        if not is_enabled():
            logger.debug("update_translate_hotkey_in_thread: приложение выключено, пропускаю регистрацию translate")
            return

        # проверить конфликт с exit
        eh = read_exit_hotkey()
        eh_mask, eh_vk = hotkey_tuple_from_config(eh.get("modifiers", []) or [], eh.get("key", "F10"))
        if hotkeys_conflict(mask, vk, eh_mask, eh_vk):
            logger.error("update_translate_hotkey_in_thread: конфликт с exit hotkey; пропускаю регистрацию translate")
            try:
                win32api.MessageBox(0,
                    f"Не удалось зарегистрировать комбинацию перевода { '+'.join(mods) + '+' if mods else ''}{key} — конфликт с комбинацией выхода.",
                    APP_NAME, 0)
            except Exception:
                pass
            return

        ok = user32.RegisterHotKey(None, HOTKEY_ID_TRANSLATE, mask, vk)
        if ok:
            logger.debug("update_translate_hotkey_in_thread: registered translate hotkey %s + %s (mask=0x%X vk=0x%X)",
                         "+".join(mods) if mods else "(no modifiers)", key, mask, vk)
        else:
            err = kernel32.GetLastError()
            logger.error("update_translate_hotkey_in_thread: failed to register translate hotkey (err=%d) mask=0x%X vk=0x%X",
                         err, mask, vk)
            try:
                win32api.MessageBox(0,
                    f"Не удалось зарегистрировать комбинацию перевода { '+'.join(mods) + '+' if mods else ''}{key}.\nКод ошибки: {err}",
                    APP_NAME, 0)
            except Exception:
                pass
    except Exception:  # noqa
        logger.exception("update_translate_hotkey_in_thread: exception")

# ---------------- Обработчик F4 (перевод) и основная логика преобразования ----------------
exit_event = threading.Event()
_handler_lock = threading.Lock()
_last_translate_ts = 0.0
_TRANSLATE_DEBOUNCE_SEC = 0.6

def _translate_invoker() -> None:
    """Дебаунс и запуск worker-потока для обработки translate hotkey"""
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

def _restore_modifiers_from_vks(vks: set) -> None:
    """
    Восстановить логические модификаторы, нажав соответствующие left-side VK.
    Принимает множество side-specific VK (VK_LCONTROL, VK_RCONTROL, ...).
    Нажимает левую версию модификатора (LCTRL, LALT, LSHIFT, LWIN).
    """
    try:
        to_press = set()
        for vk in vks:
            if vk in (VK_LCONTROL, VK_RCONTROL,):
                to_press.add(VK_LCONTROL)
            elif vk in (VK_LMENU, VK_RMENU,):
                to_press.add(VK_LMENU)
            elif vk in (VK_LSHIFT, VK_RSHIFT,):
                to_press.add(VK_LSHIFT)
            elif vk in (VK_LWIN, VK_RWIN,):
                to_press.add(VK_LWIN)
        for vk in to_press:
            try:
                _key_down(vk)
            except Exception:
                pass
    except Exception:
        pass



# --- заменить существующую функцию handle_hotkey_transform на этот вариант ---
def handle_hotkey_transform() -> None:
    """Главная логика: прочитать выделение, преобразовать раскладку и вставить обратно"""
    logger.info("Translate обработка — начинаю преобразование выделения")
    pressed_vks: set = set()
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

        # ВАЖНО: safe_copy_from_selection возвращает (text, pressed_vks)
        res = safe_copy_from_selection(timeout_per_attempt=0.6, max_attempts=2)
        if isinstance(res, tuple):
            selected, pressed_vks = res
        else:
            # старый стиль — на случай, если функция не была обновлена
            selected = res
            pressed_vks = set()

        if selected is None:
            selected = ""
        if not isinstance(pressed_vks, set):
            # обеспечить корректный тип
            try:
                pressed_vks = set(pressed_vks or [])
            except Exception:
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

        logger.info("handle: завершено успешно. (буфер оставлен как есть — в нём будет выделение)")
    except Exception as e:  # noqa
        logger.exception("handle_hotkey_transform: исключение при обработке: %s", e)
    finally:
        # Восстановим ранее отпущенные модификаторы (если были)
        try:
            if pressed_vks:
                logger.debug("handle: restoring previously released modifiers vks=%r", pressed_vks)
                _restore_modifiers_from_vks(pressed_vks)
        except Exception:
            logger.exception("handle: failed to restore modifiers")

# ---------------- Tray / Icon / Меню ----------------
def prepare_tray_icon_image(enabled: Optional[bool] = None) -> Image.Image:
    """Подготовить PIL.Image для иконки в трее (файлы icon_on/icon_off или fallback)"""
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
    size = (64, 64)
    bg = (76, 175, 80, 255) if enabled else (220, 53, 69, 255)
    img = Image.new('RGBA', size, bg)
    d = ImageDraw.Draw(img)
    try:
        d.rectangle((8, 16, 56, 48), outline=(255,255,255), width=2)
        d.text((16, 12), "KF", fill=(255,255,255))
    except Exception:  # noqa
        pass
    return img

def enabled_menu_text(_item) -> str:
    """Текст пункта меню Вкл/Выкл в зависимости от состояния"""
    return "❌ Выключить" if is_enabled() else "✅ Включить"

def toggle_enabled(_icon, _item) -> None:
    """Переключить глобальное включение/выключение приложения (и icon)"""
    new_state = not is_enabled()
    set_enabled(new_state)
    logger.info("Tray: toggled enabled -> %s", new_state)
    try:
        ok = post_register_translate(new_state)
        if not ok:
            logger.warning("toggle_enabled: post_register_translate returned False")
    except Exception:  # noqa
        logger.exception("toggle_enabled: failed to post register/unregister request")
    # обновление иконки (если вызов пришёл из pystray)
    try:
        if _icon is not None:
            _icon.icon = prepare_tray_icon_image(new_state)
    except Exception:  # noqa
        logger.exception("Tray: failed to update icon image after toggle")

def toggle_file_logging(_icon, _item) -> None:
    """Переключить логирование в файл (чтение/запись в конфиг)"""
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

def toggle_autorun(_icon, _item) -> None:
    """Переключить автозапуск (создать/удалить ярлык и записать в конфиг)"""
    try:
        current = read_autorun_flag()
        new = not current
        ok = write_autorun_flag(new)
        if ok:
            logger.info("Tray: autorun toggled -> %s", new)
            apply_autorun_setting()
        else:
            logger.warning("Tray: autorun toggle attempted but write failed")
    except Exception:  # noqa
        logger.exception("toggle_autorun: exception")

# ---------------- Hotkey захват: диалог + low-level hook ----------------
# Вспомогательные VK для low-level hook
VK_LSHIFT = 0xA0
VK_RSHIFT = 0xA1
VK_LCONTROL = 0xA2
VK_RCONTROL = 0xA3
VK_LMENU = 0xA4
VK_RMENU = 0xA5
VK_LWIN = 0x5B
VK_RWIN = 0x5C

# LRESULT тип в зависимости от разрядности
LRESULT = ctypes.c_longlong if ctypes.sizeof(ctypes.c_void_p) == 8 else ctypes.c_long

class KBDLLHOOKSTRUCT(ctypes.Structure):
    """Структура параметров, передаваемых low-level hook'ом клавиатуры"""
    _fields_ = [
        ("vkCode", wintypes.DWORD),
        ("scanCode", wintypes.DWORD),
        ("flags", wintypes.DWORD),
        ("time", wintypes.DWORD),
        ("dwExtraInfo", ULONG_PTR),
    ]

# корректные argtypes/restype для функций хука
user32.SetWindowsHookExW.argtypes = (wintypes.INT, ctypes.c_void_p, wintypes.HINSTANCE, wintypes.DWORD)
user32.SetWindowsHookExW.restype = wintypes.HHOOK
user32.UnhookWindowsHookEx.argtypes = (wintypes.HHOOK,)
user32.UnhookWindowsHookEx.restype = wintypes.BOOL
user32.CallNextHookEx.argtypes = (wintypes.HHOOK, ctypes.c_int, wintypes.WPARAM, wintypes.LPARAM)
user32.CallNextHookEx.restype = LRESULT

def _vk_is_modifier(vk: int) -> bool:
    """Является ли VK модификатором (Ctrl/Alt/Shift/Win)"""
    return vk in (
        VK_CONTROL, VK_MENU, VK_SHIFT,
        VK_LSHIFT, VK_RSHIFT, VK_LCONTROL, VK_RCONTROL, VK_LMENU, VK_RMENU,
        VK_LWIN, VK_RWIN
    )

def _current_modifiers_list() -> list[str]:
    """Вернуть список удерживаемых модификаторов в момент вызова"""
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
    """Преобразовать VK в удобочитаемое имя клавиши (F1..F12, A..Z, 0..9, SPACE, ENTER и т.д.)"""
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

def capture_hotkey_via_hook_blocking(timeout: Optional[float] = 10.0, show_dialog: bool = True) -> Optional[dict]:
    """
    Заблокированно захватить комбинацию клавиш:
    — показать всплывающее окно (tkinter),
    — слушать WH_KEYBOARD_LL, записать первую ненулевую (не-модификаторную) клавишу с текущими модификаторами.
    Возвращает {'modifiers': [...], 'key': 'X'} или None при отмене/таймауте/ошибке.
    """
    result: Optional[dict] = {"modifiers": [], "key": ""}
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
                    # игнорируем одиночный модификатор
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

        tk_root = None
        try:
            if show_dialog:
                try:
                    import tkinter as tk
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
            # процесс сообщений, чтобы хук работал
            msg = wintypes.MSG()
            has = user32.PeekMessageW(ctypes.byref(msg), None, 0, 0, 1)  # PM_REMOVE = 1
            if has:
                user32.TranslateMessage(ctypes.byref(msg))
                user32.DispatchMessageW(ctypes.byref(msg))
            # обновление tkinter окна
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
            if 'tk_root' in locals() and tk_root is not None:
                try:
                    tk_root.destroy()
                except Exception:  # noqa
                    pass
        except Exception:  # noqa
            pass
        # избавиться от ссылки на callback
        try:
            hook_proc_ptr = None
        except Exception:  # noqa
            pass
    return result

def hotkey_lists_equal(a_mods: list[str], a_key: str, b_mods: list[str], b_key: str) -> bool:
    """Сравнить две комбинации (нормализуя модификаторы и ключи)"""
    a_norm = set((m or "").upper() for m in (a_mods or []))
    b_norm = set((m or "").upper() for m in (b_mods or []))
    return a_norm == b_norm and ( (a_key or "").upper() == (b_key or "").upper() )

def capture_hotkey_and_apply_via_thread(target: str, _icon=None) -> None:
    """
    Запустить поток, который покажет окно и захватит комбинацию, затем запишет её в конфиг.
    target: 'exit' или 'translate'
    """
    def _runner() -> None:
        thread_name = threading.current_thread().name
        logger.info("Tray: ожидаю комбинацию. Нажми требуемую комбинацию (ESC для отмены).")
        res = capture_hotkey_via_hook_blocking(timeout=None, show_dialog=True)
        if not res:
            logger.info("Tray: захват комбинации отменён или не получен.")
            return
        mods = res.get("modifiers", [])
        key = res.get("key", "F10")

        # проверка конфликтов: не позволяем захватить комбинацию, уже используемую другой функцией
        if target == "exit":
            other = read_translate_hotkey()
            if hotkey_lists_equal(mods, key, other.get("modifiers", []) or [], other.get("key", "")):
                try:
                    win32api.MessageBox(0,
                        f"Нельзя установить комбинацию выхода { '+'.join(mods) + '+' if mods else ''}{key} — она уже используется для перевода.",
                        APP_NAME, 0)
                except Exception:
                    pass
                logger.info("Tray: попытка установки exit hotkey отклонена — конфликт с translate")
                return
            ok = write_exit_hotkey(mods, key)
            if ok:
                post_update_exit_hotkey()
                logger.info("Tray: установлена новая комбинация выхода: %s + %s", "+".join(mods) if mods else "(none)", key)
            else:
                logger.warning("Tray: не удалось записать новую комбинацию выхода")
        else:  # translate
            other = read_exit_hotkey()
            if hotkey_lists_equal(mods, key, other.get("modifiers", []) or [], other.get("key", "")):
                try:
                    win32api.MessageBox(0,
                        f"Нельзя установить комбинацию перевода { '+'.join(mods) + '+' if mods else ''}{key} — она уже используется для выхода.",
                        APP_NAME, 0)
                except Exception:
                    pass
                logger.info("Tray: попытка установки translate hotkey отклонена — конфликт с exit")
                return
            ok = write_translate_hotkey(mods, key)
            if ok:
                post_update_translate_hotkey()
                logger.info("Tray: установлена новая комбинация перевода: %s + %s", "+".join(mods) if mods else "(none)", key)
            else:
                logger.warning("Tray: не удалось записать новую комбинацию перевода")
    name = "ExitCaptureThread" if target == "exit" else "TranslateCaptureThread"
    t = threading.Thread(target=_runner, daemon=True, name=name)
    t.start()

# ---------------- Меню: отображение текущих комбинаций и опции ----------------
def is_exit_hotkey_equal(expected_mods: list[str], expected_key: str) -> bool:
    eh = read_exit_hotkey()
    return hotkey_lists_equal(eh.get("modifiers", []) or [], eh.get("key", ""), expected_mods, expected_key)

def is_translate_hotkey_equal(expected_mods: list[str], expected_key: str) -> bool:
    th = read_translate_hotkey()
    return hotkey_lists_equal(th.get("modifiers", []) or [], th.get("key", ""), expected_mods, expected_key)

def format_exit_hotkey_display(_item=None) -> str:
    """Формат для строки отображения комбинации выхода в трее"""
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

def format_translate_hotkey_display(_item=None) -> str:
    """Формат для строки отображения комбинации перевода в трее"""
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

def _set_exit_hotkey_and_apply(mods: list[str], key: str) -> bool:
    # проверка конфликта с translate перед записью
    other = read_translate_hotkey()
    if hotkey_lists_equal(mods, key, other.get("modifiers", []) or [], other.get("key", "")):
        try:
            win32api.MessageBox(0,
                f"Нельзя установить комбинацию выхода { '+'.join(mods) + '+' if mods else ''}{key} — она уже используется для перевода.",
                APP_NAME, 0)
        except Exception:
            pass
        return False
    ok = write_exit_hotkey(mods, key)
    if ok:
        post_update_exit_hotkey()
    return ok

def _set_translate_hotkey_and_apply(mods: list[str], key: str) -> bool:
    # проверка конфликта с exit перед записью
    other = read_exit_hotkey()
    if hotkey_lists_equal(mods, key, other.get("modifiers", []) or [], other.get("key", "")):
        try:
            win32api.MessageBox(0,
                f"Нельзя установить комбинацию перевода { '+'.join(mods) + '+' if mods else ''}{key} — она уже используется для выхода.",
                APP_NAME, 0)
        except Exception:
            pass
        return False
    ok = write_translate_hotkey(mods, key)
    if ok:
        post_update_translate_hotkey()
    return ok

def menu_set_exit_f10(_icon, _item) -> None:
    _set_exit_hotkey_and_apply([], "F10")

def menu_set_exit_ctrl_q(_icon, _item) -> None:
    _set_exit_hotkey_and_apply(["CTRL"], "Q")

def menu_set_exit_ctrl_alt_x(_icon, _item) -> None:
    _set_exit_hotkey_and_apply(["CTRL", "ALT"], "X")

def menu_set_exit_custom_capture(_icon, _item) -> None:
    try:
        capture_hotkey_and_apply_via_thread("exit", _icon)
    except Exception:  # noqa
        logger.exception("menu_set_exit_custom_capture: исключение при запуске capture thread")

def menu_set_translate_f4(_icon, _item) -> None:
    _set_translate_hotkey_and_apply([], "F4")

def menu_set_translate_ctrl_alt_t(_icon, _item) -> None:
    _set_translate_hotkey_and_apply(["CTRL", "ALT"], "T")

def menu_set_translate_ctrl_shift_y(_icon, _item) -> None:
    _set_translate_hotkey_and_apply(["CTRL", "SHIFT"], "Y")

def menu_set_translate_custom_capture(_icon, _item) -> None:
    try:
        capture_hotkey_and_apply_via_thread("translate", _icon)
    except Exception:  # noqa
        logger.exception("menu_set_translate_custom_capture: исключение при запуске capture thread")

def on_exit(_icon=None, _item=None) -> None:
    """Завершить работу приложения (через трей или хоткей выхода)"""
    logger.info("Выход запрошен (через трей/горячую клавишу)")
    exit_event.set()
    try:
        if _icon:
            _icon.stop()
    except Exception:  # noqa
        logger.exception("on_exit: ошибка при остановке иконки")
    try:
        win32event.ReleaseMutex(mutex_handle)
    except Exception:  # noqa
        pass
    logger.info("Завершение работы.")
    return

def tray_worker() -> None:
    """Запустить pystray icon + меню (в отдельном потоке)"""
    img = prepare_tray_icon_image()
    settings_menu = Menu(
        MenuItem(format_translate_hotkey_display, None, enabled=False),
        MenuItem(format_exit_hotkey_display, None, enabled=False),
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
    )
    menu = Menu(
        MenuItem(enabled_menu_text, toggle_enabled),
        MenuItem("Настройки", settings_menu),
        MenuItem("Выход", on_exit)
    )
    icon = Icon(APP_NAME, img, APP_NAME, menu=menu)
    try:
        icon.icon = prepare_tray_icon_image(is_enabled())
    except Exception:  # noqa
        logger.exception("tray_worker: failed to set initial icon")
    try:
        logger.debug("tray_worker: запуск иконки")
        icon.run()
    except Exception as e:  # noqa
        logger.exception("tray_worker: исключение: %s", e)

# ---------------- WinAPI hotkey loop (главный hotkey-поток) ----------------
def win_hotkey_loop() -> None:
    """Поток, который слушает сообщения Windows и обрабатывает WM_HOTKEY и наши внутренние сообщения"""
    global HOTKEY_THREAD_ID
    try:
        HOTKEY_THREAD_ID = kernel32.GetCurrentThreadId()
        logger.debug("win_hotkey_loop: HOTKEY_THREAD_ID = %d", HOTKEY_THREAD_ID)
    except Exception:  # noqa
        HOTKEY_THREAD_ID = 0
        logger.exception("win_hotkey_loop: failed to get thread id")

    # регистрируем exit/translate согласно файлу (внутри функций они учитывают enabled и конфликты)
    update_exit_hotkey_in_thread()
    update_translate_hotkey_in_thread()

    try:
        msg = wintypes.MSG()
        while not exit_event.is_set():
            has = user32.GetMessageW(ctypes.byref(msg), None, 0, 0)
            if has == 0:
                break
            if msg.message == MSG_REGISTER_TRANSLATE:
                logger.debug("win_hotkey_loop: MSG_REGISTER_TRANSLATE received wParam=%s", int(msg.wParam))
                try:
                    update_translate_hotkey_in_thread()
                except Exception:  # noqa
                    logger.exception("win_hotkey_loop: update_translate_hotkey_in_thread failed")
                continue
            elif msg.message == MSG_UNREGISTER_TRANSLATE:
                logger.debug("win_hotkey_loop: MSG_UNREGISTER_TRANSLATE received")
                try:
                    user32.UnregisterHotKey(None, HOTKEY_ID_TRANSLATE)
                except Exception:  # noqa
                    logger.exception("win_hotkey_loop: unregister translate failed")
                continue
            elif msg.message == MSG_UPDATE_EXIT_HOTKEY:
                logger.debug("win_hotkey_loop: MSG_UPDATE_EXIT_HOTKEY received -> update exit hotkey")
                try:
                    update_exit_hotkey_in_thread()
                except Exception:  # noqa
                    logger.exception("win_hotkey_loop: update_exit_hotkey_in_thread failed")
                continue
            elif msg.message == MSG_UPDATE_TRANSLATE_HOTKEY:
                logger.debug("win_hotkey_loop: MSG_UPDATE_TRANSLATE_HOTKEY received -> update translate hotkey")
                try:
                    update_translate_hotkey_in_thread()
                except Exception:  # noqa
                    logger.exception("win_hotkey_loop: update_translate_hotkey_in_thread failed")
                continue
            if msg.message == 0x0312:  # WM_HOTKEY
                hotkey_id = msg.wParam
                if hotkey_id == HOTKEY_ID_TRANSLATE:
                    if not is_enabled():
                        logger.debug("WM_HOTKEY: translate received but currently disabled -> ignored")
                    else:
                        logger.debug("WM_HOTKEY: translate received")
                        _translate_invoker()
                elif hotkey_id == HOTKEY_ID_EXIT:
                    logger.debug("WM_HOTKEY: EXIT hotkey received -> exiting")
                    on_exit()
            user32.TranslateMessage(ctypes.byref(msg))
            user32.DispatchMessageW(ctypes.byref(msg))
    finally:
        try:
            user32.UnregisterHotKey(None, HOTKEY_ID_TRANSLATE)
            user32.UnregisterHotKey(None, HOTKEY_ID_EXIT)
        except Exception:  # noqa
            pass

# ---------------- Main ----------------
def main() -> None:
    t = threading.Thread(target=tray_worker, daemon=True, name="TrayThread")
    t.start()
    logger.debug("main: трей поток запущен")

    t2 = threading.Thread(target=win_hotkey_loop, daemon=True, name="WinHotkeyThread")
    t2.start()
    logger.debug("main: WinAPI hotkey thread started")

    try:
        post_update_exit_hotkey()
        post_update_translate_hotkey()
    except Exception:  # noqa
        logger.exception("main: failed to post initial update messages")

    try:
        while not exit_event.is_set():
            time.sleep(0.3)
    except KeyboardInterrupt:
        logger.info("main: KeyboardInterrupt")
    finally:
        try:
            win32event.ReleaseMutex(mutex_handle)
        except Exception:  # noqa
            pass
        logger.info("main: завершение")

if __name__ == "__main__":
    main()
