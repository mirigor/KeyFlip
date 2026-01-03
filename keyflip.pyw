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

# --- Константы / пути ---
APP_NAME = "KeyFlip"
MUTEX_NAME = r"Global\KeyFlipMutex"
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
LOG_FILE = os.path.join(BASE_DIR, "keyflip.log")
ICON_ON = os.path.join(BASE_DIR, "icon_on.ico")
ICON_OFF = os.path.join(BASE_DIR, "icon_off.ico")
CONFIG_FILE = os.path.join(BASE_DIR, "keyflip.json")

def read_file_logging_flag() -> bool:
    """
    Читает CONFIG_FILE и возвращает булево поле file_logging.
    Если файла нет или парсинг не удался — возвращает False.
    Вызывается при каждой попытке записи в файловый handler.
    """
    try:
        if not os.path.isfile(CONFIG_FILE):
            return False
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            j = json.load(f)
        return bool(j.get("file_logging", False))
    except Exception:
        # в случае ошибок считаем file logging выключенным
        return False

def write_file_logging_flag(val: bool) -> bool:
    """
    Записывает флаг file_logging в CONFIG_FILE, оставляя при этом поля enabled/autorun, если они есть.
    Возвращает True при успехе, False при ошибке.
    """
    try:
        j = {}
        if os.path.isfile(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    j = json.load(f)
            except Exception:
                j = {}
        # сохраняем/устанавливаем поля
        if "enabled" not in j:
            j["enabled"] = True
        if "autorun" not in j:
            j["autorun"] = False
        j["file_logging"] = bool(val)
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(j, f, ensure_ascii=False, indent=2)
        logger.info("config: записан file_logging=%s в %s", j["file_logging"], CONFIG_FILE)
        return True
    except Exception:
        logger.exception("config: не удалось записать file_logging в %s", CONFIG_FILE)
        return False

def read_autorun_flag() -> bool:
    """
    Читает CONFIG_FILE и возвращает булево поле autorun.
    Если файла нет или парсинг не удался — возвращает False.
    """
    try:
        if not os.path.isfile(CONFIG_FILE):
            return False
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            j = json.load(f)
        return bool(j.get("autorun", False))
    except Exception:
        return False

def write_autorun_flag(val: bool) -> bool:
    """
    Записывает флаг autorun в CONFIG_FILE, оставляя при этом поля enabled/file_logging, если они есть.
    Возвращает True при успехе, False при ошибке.
    """
    try:
        j = {}
        if os.path.isfile(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    j = json.load(f)
            except Exception:
                j = {}
        if "enabled" not in j:
            j["enabled"] = True
        if "file_logging" not in j:
            j["file_logging"] = False
        j["autorun"] = bool(val)
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(j, f, ensure_ascii=False, indent=2)
        logger.info("config: записан autorun=%s в %s", j["autorun"], CONFIG_FILE)
        return True
    except Exception:
        logger.exception("config: не удалось записать autorun в %s", CONFIG_FILE)
        return False

# ---------------- startup shortcut helpers ----------------
def _startup_shortcut_path() -> str:
    """
    Возвращает путь к .lnk в папке автозагрузки текущего пользователя.
    """
    appdata = os.environ.get("APPDATA")
    if not appdata:
        # на всякий случай fallback
        appdata = os.path.expanduser("~\\AppData\\Roaming")
    startup_dir = os.path.join(appdata, "Microsoft", "Windows", "Start Menu", "Programs", "Startup")
    return os.path.join(startup_dir, f"{APP_NAME}.lnk")

def create_startup_shortcut() -> bool:
    """
    Создаёт ярлык в папке автозагрузки. Возвращает True при успехе.
    Ярлык указывает на python-исполняемый файл и передаёт скрипт (или exe) как аргумент.
    """
    shortcut_path = _startup_shortcut_path()
    try:
        # Отдельно импортируем win32com, чтобы не падать при наличии pywin32 но без COM (редко)
        try:
            import win32com.client  # type: ignore
        except Exception:
            logger.exception("create_startup_shortcut: win32com.client недоступен, не могу создать ярлык")
            return False

        shell = win32com.client.Dispatch("WScript.Shell")
        shortcut = shell.CreateShortcut(shortcut_path)
        # Если приложение собрано в exe (pyinstaller), sys.executable указывает на exe.
        # Иначе — указываем на интерпретатор и добавляем путь к скрипту в Arguments.
        target = sys.executable
        script = os.path.abspath(sys.argv[0])
        # Если исполняемый файл - python.exe/wpython, передаём скрипт в аргументах.
        shortcut.TargetPath = target
        # аргументы — скрипт в кавычках, если это .py
        if script.endswith(".py") or script.endswith(".pyw"):
            shortcut.Arguments = f'"{script}"'
        else:
            # возможно уже exe — не передаём аргументы
            shortcut.Arguments = ""
        shortcut.WorkingDirectory = BASE_DIR
        # Иконка: используем ICON_ON если есть, иначе сам exe
        if os.path.isfile(ICON_ON):
            shortcut.IconLocation = ICON_ON
        else:
            shortcut.IconLocation = target
        shortcut.Save()
        logger.info("startup: ярлык создан: %s", shortcut_path)
        return True
    except Exception:
        logger.exception("startup: не удалось создать ярлык автозагрузки")
        return False

def remove_startup_shortcut() -> bool:
    """
    Удаляет ярлык из автозагрузки, если он существует.
    """
    sp = _startup_shortcut_path()
    try:
        if os.path.isfile(sp):
            os.remove(sp)
            logger.info("startup: ярлык удалён: %s", sp)
        else:
            logger.debug("startup: ярлык не найден: %s", sp)
        return True
    except Exception:
        logger.exception("startup: не удалось удалить ярлык: %s", sp)
        return False

def apply_autorun_setting() -> None:
    """
    Применяет текущее значение autorun (читая из файла) — создаёт или удаляет ярлык.
    Вызывать на старте и при переключении через меню.
    """
    logger.warning("AUTORUN: apply_autorun_setting CALLED")

    try:
        if read_autorun_flag():
            create_startup_shortcut()
        else:
            remove_startup_shortcut()
    except Exception:
        logger.exception("apply_autorun_setting: ошибка при применении autorun")

# --- Один экземпляр (mutex) ---
mutex_handle = win32event.CreateMutex(None, False, MUTEX_NAME)
ERROR_ALREADY_EXISTS = 183
if win32api.GetLastError() == ERROR_ALREADY_EXISTS:
    try:
        win32api.MessageBox(0, f"Программа {APP_NAME} уже запущена!", APP_NAME, 0)
    except Exception:  # noqa
        print(f"{APP_NAME} уже запущена!")
    sys.exit(0)

# --- Логирование ---
logger = logging.getLogger(APP_NAME)
logger.setLevel(logging.DEBUG)
handler = RotatingFileHandler(LOG_FILE, maxBytes=5 * 1024 * 1024, backupCount=3, encoding="utf-8")
fmt = logging.Formatter("%(asctime)s %(levelname)s [%(threadName)s] %(message)s", "%Y-%m-%d %H:%M:%S")
handler.setFormatter(fmt)

# прикрепляем фильтр, который читает CONFIG_FILE при каждом лог-записи
class FileLoggingFilter(logging.Filter):
    """Фильтр для file_handler — пускает запись в файл только если flag в config True"""
    def filter(self, record: logging.LogRecord) -> bool:
        try:
            return read_file_logging_flag()
        except Exception:
            return False

handler.addFilter(FileLoggingFilter())

logger.addHandler(handler)
console = logging.StreamHandler(sys.stdout)
console.setFormatter(fmt)
console.setLevel(logging.INFO)
logger.addHandler(console)
logger.info("%s запуск — лог: %s", APP_NAME, LOG_FILE)

# ------------ Mapping: explicit, детерминированные таблицы ------------
EN_TO_RU: dict[str, str] = {
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
RU_TO_EN: dict[str, str] = {v: k for k, v in EN_TO_RU.items()}
RU_TO_EN['ё'] = '`'
RU_TO_EN['Ё'] = '~'

def transform_text_by_keyboard_layout_based_on_hkl(s: str, hkl: int) -> str:
    """
    hkl: низшие 16 бит — LANGID (0x0419 = Russian)
    Если текущая раскладка — Russian (0x0419) -> переводим RU -> EN.
    Иначе -> EN -> RU.
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
def get_active_window_info() -> tuple[str, int, Optional[str]]:
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
    proc_name = None
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

# ------------ Send Unicode via SendInput (не трогает clipboard) ------------
INPUT_KEYBOARD = 1
KEYEVENTF_UNICODE = 0x0004
KEYEVENTF_KEYUP = 0x0002
wintypes.ULONG_PTR = wintypes.WPARAM

class KEYBDINPUT(ctypes.Structure):
    _fields_ = [
        ("wVk", wintypes.WORD),
        ("wScan", wintypes.WORD),
        ("dwFlags", wintypes.DWORD),
        ("time", wintypes.DWORD),
        ("dwExtraInfo", wintypes.ULONG_PTR)
    ]

class _INPUT_union(ctypes.Union):
    _fields_ = [("ki", KEYBDINPUT), ("padding", wintypes.ULONG * 8)]

class INPUT(ctypes.Structure):
    _anonymous_ = ("union",)
    _fields_ = [("type", wintypes.DWORD), ("union", _INPUT_union)]

SendInput = user32.SendInput
SendInput.argtypes = (wintypes.UINT, ctypes.POINTER(INPUT), ctypes.c_int)
SendInput.restype = wintypes.UINT

def send_unicode_via_sendinput(text: str, delay_between_keys: float = 0.001):
    """Посылает символы через SendInput (UNICODE). Без использования буфера обмена"""
    if not text:
        return

    inputs = []
    for ch in text:
        code = ord(ch)
        # key down
        ki_down = KEYBDINPUT(0, code, KEYEVENTF_UNICODE, 0, 0)
        inp_down = INPUT(INPUT_KEYBOARD, _INPUT_union(ki=ki_down))
        inputs.append(inp_down)
        # key up
        ki_up = KEYBDINPUT(0, code, KEYEVENTF_UNICODE | KEYEVENTF_KEYUP, 0, 0)
        inp_up = INPUT(INPUT_KEYBOARD, _INPUT_union(ki=ki_up))
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


# ------------ Clipboard helper (бережно) ------------
def safe_restore_clipboard(old_clip: Optional[str]) -> None:
    """Оставлена для совместимости; но по умолчанию мы НЕ восстанавливаем старый буфер"""
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
    """
    Бережное копирование выделения:
    - Опираемся на GetClipboardSequenceNumber (предпочтительно).
    - НО: не восстанавливаем старый буфер. Возвращаем скопированный (текст).
    """
    title, pid, proc_name = get_active_window_info()
    logger.debug("safe_copy: start. active window: %r pid=%s proc=%r", title, pid, proc_name)

    try:
        initial_hwnd = user32.GetForegroundWindow()
    except Exception as e:
        logger.debug("safe_copy: GetForegroundWindow failed: %s", e)
        initial_hwnd = None

    def foreground_changed() -> bool:
        try:
            cur = user32.GetForegroundWindow()
            return initial_hwnd is not None and cur is not None and cur != initial_hwnd
        except Exception:  # noqa
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
                logger.debug("safe_copy: sent ctrl+c (attempt %d)", attempt)
            except Exception as e:
                logger.exception("safe_copy: ctrl+c exception: %s", e)

            t0 = time.time()
            changed = False
            while time.time() - t0 < timeout_per_attempt:
                if foreground_changed():
                    logger.debug("safe_copy: foreground changed during wait after Ctrl+C -> abort")
                    return ""
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
                    cur = pyperclip.paste()
                    if cur == "":
                        logger.debug("safe_copy: clipboard changed but empty -> treat as no selection")
                        return ""
                    logger.debug("safe_copy: buffer changed after ctrl+c (len=%d)", len(cur))
                    return cur
                except Exception as e:
                    logger.exception("safe_copy: paste() after ctrl+c failed: %s", e)

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
                logger.debug("safe_copy: sent ctrl+insert (attempt %d)", attempt)
            except Exception as e:
                logger.exception("safe_copy: ctrl+insert exception: %s", e)

            t0 = time.time()
            changed = False
            while time.time() - t0 < timeout_per_attempt:
                if foreground_changed():
                    logger.debug("safe_copy: foreground changed during wait after Ctrl+Insert -> abort")
                    return ""
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
                    cur = pyperclip.paste()
                    if cur == "":
                        logger.debug("safe_copy: clipboard changed after insert but empty -> no selection")
                        return ""
                    logger.debug("safe_copy: buffer changed after ctrl+insert (len=%d)", len(cur))
                    return cur
                except Exception as e:
                    logger.exception("safe_copy: paste() after ctrl+insert failed: %s", e)

            logger.debug("safe_copy: no clipboard change after attempt %d", attempt)
            break

        logger.debug("safe_copy: sequence not changed -> no selection")
        return ""

    # ---- fallback: sentinel approach ----
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
            logger.debug("safe_copy: foreground changed before fallback Ctrl+C -> abort")
            return ""
        send_ctrl_c()
        logger.debug("safe_copy: sent ctrl+c (fallback)")
    except Exception as e:
        logger.exception("safe_copy: ctrl+c exception (fallback): %s", e)
        last_exception = e

    t0 = time.time()
    while time.time() - t0 < timeout_per_attempt:
        if foreground_changed():
            logger.debug("safe_copy: foreground changed during fallback wait -> abort")
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
                return ""
            logger.debug("safe_copy: buffer changed (fallback) len=%d", len(cur))
            return cur
        time.sleep(0.02)

    # last attempt: Ctrl+Insert
    if foreground_changed():
        logger.debug("safe_copy: foreground changed before fallback Ctrl+Insert -> abort")
        return ""

    try:
        VK_CONTROL = 0x11
        VK_INSERT = 0x2D
        _key_down(VK_CONTROL)
        _key_down(VK_INSERT)
        time.sleep(0.01)
        _key_up(VK_INSERT)
        _key_up(VK_CONTROL)
        logger.debug("safe_copy: sent ctrl+insert (fallback last)")
    except Exception as e:
        logger.exception("safe_copy: ctrl+insert exception (fallback last): %s", e)
        last_exception = e

    t0 = time.time()
    while time.time() - t0 < timeout_per_attempt:
        if foreground_changed():
            logger.debug("safe_copy: foreground changed during fallback insert wait -> abort")
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
                return ""
            logger.debug("safe_copy: buffer changed after ctrl+insert (fallback) len=%d", len(cur))
            return cur
        time.sleep(0.02)

    # nothing worked
    logger.warning("safe_copy: НЕ удалось получить выделение. last_exc=%r", repr(last_exception))
    return ""

# ------------ Config (enabled) ------------
_enabled_lock = threading.Lock()
_enabled: bool = True  # default; will be overwritten by load_config()

def load_config():
    """
    Загружает keyflip.json.
    Если файла нет — создаёт с дефолтами.
    Если нет file_logging/autorun — добавляет их.
    """
    global _enabled

    changed = False

    if os.path.isfile(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                j = json.load(f)
        except Exception:
            j = {}
            changed = True
    else:
        j = {}
        changed = True

    # enabled
    if "enabled" not in j:
        j["enabled"] = True
        changed = True

    _enabled = bool(j["enabled"])

    # file_logging
    if "file_logging" not in j:
        j["file_logging"] = False
        changed = True

    # autorun
    if "autorun" not in j:
        j["autorun"] = False
        changed = True

    if changed:
        try:
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(j, f, ensure_ascii=False, indent=2)
        except Exception:
            logger.exception("config: failed to write %s", CONFIG_FILE)


def save_config():
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

def set_enabled(val: bool):
    global _enabled
    with _enabled_lock:
        _enabled = bool(val)
    save_config()
    logger.info("state: enabled set to %s", _enabled)

# Load config on startup
load_config()
# Применяем autorun согласно настройке в файле (если пользователь вручную изменил файл)
try:
    apply_autorun_setting()
except Exception:
    logger.exception("startup: apply_autorun_setting on startup failed")

# ------------ Hotkey constants & thread messaging ------------
HOTKEY_ID_F4 = 1
HOTKEY_ID_F10 = 2
MOD_NONE = 0
VK_F4 = 0x73
VK_F10 = 0x79

# We'll use PostThreadMessage to request registration/unregistration in the hotkey thread
WM_USER = 0x0400
MSG_REGISTER_F4 = WM_USER + 1
MSG_UNREGISTER_F4 = WM_USER + 2

# id потока, где вызывается win_hotkey_loop (будет установлен в начале цикла)
HOTKEY_THREAD_ID = 0  # заполнится в win_hotkey_loop

def post_register_f4(should_register: bool):
    """Посылает сообщение потоку hotkey чтобы тот зарегистрировал/удалил F4"""
    global HOTKEY_THREAD_ID
    if HOTKEY_THREAD_ID == 0:
        logger.warning("post_register_f4: HOTKEY_THREAD_ID not ready yet -> cannot post")
        return False
    msg = MSG_REGISTER_F4 if should_register else MSG_UNREGISTER_F4
    # wParam = 1 for register, 0 for unregister (необязательно но удобно)
    wparam = 1 if should_register else 0
    try:
        res = user32.PostThreadMessageW(HOTKEY_THREAD_ID, msg, wparam, 0)
        if res == 0:
            err = kernel32.GetLastError()
            logger.error("post_register_f4: PostThreadMessageW failed (err=%d)", err)
            return False
        logger.debug("post_register_f4: posted msg=%d wParam=%d to thread %d", msg, wparam, HOTKEY_THREAD_ID)
        return True
    except Exception:  # noqa
        logger.exception("post_register_f4: exception while posting")
        return False

# ------------ Hotkey register helpers (used only from hotkey thread) ------------
def register_f4_in_thread():
    try:
        ok = user32.RegisterHotKey(None, HOTKEY_ID_F4, MOD_NONE, VK_F4)
        if ok:
            logger.debug("register_f4: F4 registered in hotkey thread")
        else:
            err = kernel32.GetLastError()
            logger.error("register_f4: failed to register F4 (err=%d)", err)
    except Exception:  # noqa
        logger.exception("register_f4: exception while registering F4")

def unregister_f4_in_thread():
    try:
        ok = user32.UnregisterHotKey(None, HOTKEY_ID_F4)
        if ok:
            logger.debug("unregister_f4: F4 unregistered in hotkey thread")
        else:
            logger.debug("unregister_f4: UnregisterHotKey returned False (probably not registered)")
    except Exception:  # noqa
        logger.exception("unregister_f4: exception while unregistering F4")

# ------------ Handler & Hotkey logic ------------
exit_event = threading.Event()
_handler_lock = threading.Lock()
_last_f4_ts = 0.0
_F4_DEBOUNCE_SEC = 0.6

def _f4_invoker():
    global _last_f4_ts
    # if disabled -> ignore immediately
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
            except Exception:  # noqa
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

        # save current clipboard text only for debug/diagnostics (we НЕ будем восстанавливать его)
        try:
            saved = None
            try:
                saved = pyperclip.paste()
            except Exception:  # noqa
                saved = None
            logger.debug("handle: прочитал (но не буду восстанавливать) буфер text-len=%s", None if saved is None else len(saved))
        except Exception:  # noqa
            saved = None

        # get foreground thread HKL so we know direction
        try:
            hwnd = user32.GetForegroundWindow()
            pid_c = ctypes.c_ulong()
            thread_id = user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid_c))
            hkl = user32.GetKeyboardLayout(thread_id) & 0xFFFF
            logger.debug("handle: foreground thread hkl=0x%04X", hkl)
        except Exception:  # noqa
            hkl = 0
            logger.exception("handle: не удалось получить HKL, предполагаем EN")

        # read selection carefully (this WILL change clipboard to the selection)
        selected = safe_copy_from_selection(timeout_per_attempt=0.6, max_attempts=2)
        if selected is None:
            selected = ""
        logger.info("handle: прочитано выделение (len=%d)", len(selected))

        # if nothing selected -> do nothing
        if not selected:
            logger.info("handle: выделение пустое — ничего не делаю.")
            return

        # choose mapping based on hkl
        converted = transform_text_by_keyboard_layout_based_on_hkl(selected, hkl)
        logger.info("handle: преобразование выполнено (len=%d).", len(converted))
        logger.debug("handle: исходное=%r converted=%r", selected, converted)

        if converted == selected:
            logger.info("handle: преобразованный текст совпадает с исходным — ничего не меняю.")
            return

        # delete selection
        try:
            send_delete()
            time.sleep(0.02)
        except Exception as e:
            logger.exception("handle: delete exception: %s", e)

        # вставляем через SendInput (не трогаем clipboard)
        try:
            send_unicode_via_sendinput(converted, delay_between_keys=0.001)
            logger.debug("handle: вставил через SendInput (unicode) — без использования clipboard")
        except Exception as e:
            logger.exception("handle: send_unicode_via_sendinput failed: %s", e)
            # fallback на clipboard paste (опасно для non-text formats, но это fallback)
            try:
                pyperclip.copy(converted)
                time.sleep(0.02)
                send_ctrl_v()
                logger.debug("handle: fallback: вставил через буфер (ctrl+v)")
            except Exception as e2:
                logger.exception("handle: fallback paste failed: %s", e2)

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
                    except Exception:  # noqa
                        logger.exception("handle: fallback ActivateKeyboardLayout also failed")
                else:
                    logger.debug("handle: posted WM_INPUTLANGCHANGEREQUEST -> success")
        except Exception:  # noqa
            logger.exception("handle: switch layout failed (foreground-thread)")

        # ВАЖНО: не восстанавливаем старый буфер — оставляем то, что сейчас в clipboard (обычно выделение)
        logger.info("handle: завершено успешно. (буфер оставлен как есть — в нём будет выделение)")
    except Exception as e:
        logger.exception("handle_hotkey_transform: исключение при обработке: %s", e)

# ------------ Tray / Icon (с пунктом Вкл/Выкл и иконками) ----------------
def prepare_tray_icon_image(enabled: Optional[bool] = None) -> Image.Image:
    """
    Возвращает PIL.Image для трея. Если файлы icon_on/icon_off присутствуют, используем их,
    иначе генерируем простую квадратную иконку с K и цветом (зелёный/красный).
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

    # fallback: draw simple image
    size = (64, 64)
    bg = (76, 175, 80, 255) if enabled else (220, 53, 69, 255)  # green / red
    img = Image.new('RGBA', size, bg)
    d = ImageDraw.Draw(img)
    try:
        d.rectangle((8, 16, 56, 48), outline=(255,255,255), width=2)
        d.text((16, 12), "KF", fill=(255,255,255))
    except Exception:  # noqa
        pass
    return img

def enabled_menu_text(item):
    return "❌ Выключить" if is_enabled() else "✅ Включить"

def toggle_enabled(icon, item):
    """Action called from tray menu: flips state, updates config and icon image immediately"""
    new_state = not is_enabled()
    set_enabled(new_state)
    logger.info("Tray: toggled enabled -> %s", new_state)
    # Вместо прямого вызова Register/Unregister ми постим сообщение в hotkey-thread
    try:
        ok = post_register_f4(new_state)
        if not ok:
            logger.warning("toggle_enabled: post_register_f4 returned False")
    except Exception:  # noqa
        logger.exception("toggle_enabled: failed to post register/unregister request")

    # Обновляем иконку прямо сейчас (pystray Icon передаётся первым аргументом)
    try:
        if icon is not None:
            icon.icon = prepare_tray_icon_image(new_state)
    except Exception:  # noqa
        logger.exception("Tray: failed to update icon image after toggle")

def toggle_file_logging(icon, item):
    """Переключает запись логов в файл (меняет file_logging в keyflip.json)"""
    try:
        current = read_file_logging_flag()
        new = not current
        ok = write_file_logging_flag(new)
        if ok:
            logger.info("Tray: file_logging toggled -> %s", new)
        else:
            logger.warning("Tray: file_logging toggle attempted but write failed")
    except Exception:
        logger.exception("toggle_file_logging: исключение")

def toggle_autorun(icon, item):
    """Переключает автозапуск при старте Windows (меняет autorun в keyflip.json и применяет)"""
    try:
        current = read_autorun_flag()
        new = not current
        ok = write_autorun_flag(new)
        if ok:
            logger.info("Tray: autorun toggled -> %s", new)
            apply_autorun_setting()
        else:
            logger.warning("Tray: autorun toggle attempted but write failed")
    except Exception:
        logger.exception("toggle_autorun: исключение")

def on_exit(icon=None, item=None):
    logger.info("Выход запрошен (через трей/F10)")
    exit_event.set()
    try:
        if icon:
            icon.stop()
    except Exception:  # noqa
        logger.exception("on_exit: ошибка при остановке иконки")
    try:
        win32event.ReleaseMutex(mutex_handle)
    except Exception:  # noqa
        pass
    logger.info("Завершение работы.")
    return

def tray_worker():
    img = prepare_tray_icon_image()

    # settings menu: добавляем чекбоксы, которые читают флаги из файла при отображении
    settings_menu = Menu(
        MenuItem("Логирование в файл", toggle_file_logging, checked=lambda item: read_file_logging_flag()),
        MenuItem("Автозапуск при старте Windows", toggle_autorun, checked=lambda item: read_autorun_flag()),
    )
    menu = Menu(
        MenuItem(enabled_menu_text, toggle_enabled),
        MenuItem("Настройки", settings_menu),
        MenuItem("Выход", on_exit)
    )
    icon = Icon(APP_NAME, img, APP_NAME, menu=menu)
    # ensure icon image matches saved state (in case load_config changed it earlier)
    try:
        icon.icon = prepare_tray_icon_image(is_enabled())
    except Exception:  # noqa
        logger.exception("tray_worker: failed to set initial icon")

    try:
        logger.debug("tray_worker: запуск иконки")
        icon.run()
    except Exception as e:
        logger.exception("tray_worker: исключение: %s", e)

# ------------ WinAPI hotkey loop ------------
def win_hotkey_loop():
    global HOTKEY_THREAD_ID
    # устанавливаем id текущего потока — другим потокам можно будет посылать PostThreadMessageW
    try:
        HOTKEY_THREAD_ID = kernel32.GetCurrentThreadId()
        logger.debug("win_hotkey_loop: HOTKEY_THREAD_ID = %d", HOTKEY_THREAD_ID)
    except Exception:  # noqa
        HOTKEY_THREAD_ID = 0
        logger.exception("win_hotkey_loop: failed to get thread id")

    # Register F10 всегда (для выхода) — делаем это в этом потоке
    if not user32.RegisterHotKey(None, HOTKEY_ID_F10, MOD_NONE, VK_F10):
        logger.error("Не удалось зарегистрировать F10 через RegisterHotKey")
    else:
        logger.debug("Зарегистрирован F10 via RegisterHotKey")

    # Register F4 only if enabled (this registration happens in this thread's message queue)
    if is_enabled():
        register_f4_in_thread()
    else:
        # ensure it's not registered in this thread
        try:
            user32.UnregisterHotKey(None, HOTKEY_ID_F4)
        except Exception:  # noqa
            pass

    try:
        msg = wintypes.MSG()
        while not exit_event.is_set():
            has = user32.GetMessageW(ctypes.byref(msg), None, 0, 0)
            if has == 0:
                break

            # обработка наших внутренних сообщений (регистрация/отмена регистрации)
            if msg.message == MSG_REGISTER_F4:
                logger.debug("win_hotkey_loop: MSG_REGISTER_F4 received wParam=%s", int(msg.wParam))
                try:
                    register_f4_in_thread()
                except Exception:  # noqa
                    logger.exception("win_hotkey_loop: register_f4_in_thread failed")
                # не DispatchMessage для этого пользовательского сообщения
                continue
            elif msg.message == MSG_UNREGISTER_F4:
                logger.debug("win_hotkey_loop: MSG_UNREGISTER_F4 received")
                try:
                    unregister_f4_in_thread()
                except Exception:  # noqa
                    logger.exception("win_hotkey_loop: unregister_f4_in_thread failed")
                continue

            if msg.message == 0x0312:  # WM_HOTKEY
                hotkey_id = msg.wParam
                if hotkey_id == HOTKEY_ID_F4:
                    # если пришёл F4 — сработает только если он зарегистрирован в этом потоке
                    # дополнительная защита: если состояние выключено — игнорируем
                    if not is_enabled():
                        logger.debug("WM_HOTKEY: F4 received but currently disabled -> ignored")
                    else:
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
        except Exception:  # noqa
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
        except Exception:  # noqa
            pass
        logger.info("main: завершение")

if __name__ == "__main__":
    main()
