import ctypes
import json
import logging
import os
import sys
import threading
import time
import uuid
from logging.handlers import RotatingFileHandler
from typing import Optional, Tuple, List, Dict

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

# --- Пару глобальных помощников для безопасности при записи/чтении конфига ---
_config_write_lock = threading.Lock()  # защищает запись в файл
_config_cache_lock = threading.Lock()
_config_cache: Optional[dict] = None
_config_cache_mtime: float = 0.0

def _read_config_raw() -> dict:
    """Прочитать JSON файл (не использую кэш) — возвращает {} при ошибках."""
    try:
        if not os.path.isfile(CONFIG_FILE):
            return {}
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            return json.load(f) or {}
    except Exception:
        return {}

def _reload_config_if_needed() -> dict:
    """
    Возвращает кэшированный config dict. Перезагружает, если файл изменился по mtime
    или если кэша ещё нет.
    """
    global _config_cache, _config_cache_mtime
    try:
        with _config_cache_lock:
            if not os.path.isfile(CONFIG_FILE):
                # файл отсутствует — сброс кэша
                if _config_cache is None:
                    _config_cache = {}
                    _config_cache_mtime = 0.0
                return dict(_config_cache)
            m = os.path.getmtime(CONFIG_FILE)
            if _config_cache is None or m != _config_cache_mtime:
                # reload
                try:
                    with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                        j = json.load(f) or {}
                except Exception:
                    j = {}
                _config_cache = j
                _config_cache_mtime = m
            return dict(_config_cache)
    except Exception:
        return {}

def _atomic_write_json(path: str, obj: dict) -> bool:
    """
    Атомарно записывает JSON: пишем в временный файл, flush+fsync, затем os.replace.
    Возвращает True при успехе.
    """
    tmp = path + ".tmp"
    try:
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(obj, f, ensure_ascii=False, indent=2)
            f.flush()
            try:
                os.fsync(f.fileno())
            except Exception:
                # на некоторых файловых системах может не сработать, но это ок
                pass
        os.replace(tmp, path)
        # обновим кэш
        try:
            with _config_cache_lock:
                _config_cache = obj.copy()
                _config_cache_mtime = os.path.getmtime(path)
        except Exception:
            pass
        return True
    except Exception:
        try:
            if os.path.exists(tmp):
                os.remove(tmp)
        except Exception:
            pass
        return False

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

# Фильтр: файл логов включается/выключается флагом в CONFIG_FILE
class FileLoggingFilter(logging.Filter):
    def filter(self, record: logging.LogRecord) -> bool:
        cfg = _reload_config_if_needed()
        return bool(cfg.get("file_logging", False))

handler.addFilter(FileLoggingFilter())

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
    if not text:
        return
    inputs = []
    for ch in text:
        code = ord(ch)
        ki_down = KEYBDINPUT(0, code, KEYEVENTF_UNICODE, 0, 0)
        inp_down = INPUT(INPUT_KEYBOARD, _INPUT_union(ki=ki_down))
        inputs.append(inp_down)
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
    logger.warning("safe_copy: НЕ удалось получить выделение. last_exc=%r", repr(last_exception))
    return ""

# ------------ Config (enabled + defaults) ------------
_enabled_lock = threading.Lock()
_enabled: bool = True  # default; будет перезаписан load_config()

def default_exit_hotkey() -> Dict:
    return {"modifiers": [], "key": "F10"}

def load_config():
    """
    Загружает keyflip.json (создаёт с дефолтами, если нет).
    Устанавливает _enabled и гарантирует, что в файле есть file_logging/autorun/exit_hotkey.
    """
    try:
        j = _read_config_raw()
        changed = False
        if "enabled" not in j:
            j["enabled"] = True
            changed = True
        if "file_logging" not in j:
            j["file_logging"] = False
            changed = True
        if "autorun" not in j:
            j["autorun"] = False
            changed = True
        if "exit_hotkey" not in j:
            j["exit_hotkey"] = default_exit_hotkey()
            changed = True
        if changed:
            # атомарно записываем
            with _config_write_lock:
                _atomic_write_json(CONFIG_FILE, j)
        # применим enabled в память
        global _enabled
        with _enabled_lock:
            _enabled = bool(j.get("enabled", True))
        # обновим кэш сразу
        with _config_cache_lock:
            global _config_cache, _config_cache_mtime
            _config_cache = j.copy()
            try:
                _config_cache_mtime = os.path.getmtime(CONFIG_FILE)
            except Exception:
                _config_cache_mtime = 0.0
    except Exception:
        logger.exception("config: load_config unexpected exception; using defaults")
        with _enabled_lock:
            _enabled = True

def save_config():
    try:
        with _enabled_lock:
            j = {"enabled": bool(_enabled)}
        with _config_write_lock:
            _atomic_write_json(CONFIG_FILE, j)
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

# Функции чтения/записи отдельных флагов используют кэш и атомарную запись
def read_file_logging_flag() -> bool:
    cfg = _reload_config_if_needed()
    return bool(cfg.get("file_logging", False))

def write_file_logging_flag(val: bool) -> bool:
    with _config_write_lock:
        j = _reload_config_if_needed()
        j["file_logging"] = bool(val)
        # оставляем enabled/autorun/exit_hotkey как есть
        ok = _atomic_write_json(CONFIG_FILE, j)
        if ok:
            logger.info("config: записан file_logging=%s в %s", j["file_logging"], CONFIG_FILE)
        else:
            logger.exception("config: не удалось записать file_logging в %s", CONFIG_FILE)
        return ok

def read_autorun_flag() -> bool:
    cfg = _reload_config_if_needed()
    return bool(cfg.get("autorun", False))

def write_autorun_flag(val: bool) -> bool:
    with _config_write_lock:
        j = _reload_config_if_needed()
        j["autorun"] = bool(val)
        ok = _atomic_write_json(CONFIG_FILE, j)
        if ok:
            logger.info("config: записан autorun=%s в %s", j["autorun"], CONFIG_FILE)
        else:
            logger.exception("config: не удалось записать autorun в %s", CONFIG_FILE)
        return ok

def read_exit_hotkey() -> Dict:
    """
    Возвращает {"modifiers":[...], "key":"F10"} — без поддержки list-формата.
    """
    cfg = _reload_config_if_needed()
    eh = cfg.get("exit_hotkey")
    if not isinstance(eh, dict):
        return default_exit_hotkey()
    mods = eh.get("modifiers", []) or []
    key = eh.get("key", "F10") or "F10"
    return {"modifiers": list(mods), "key": str(key)}

def normalize_modifiers(modifiers: List[str]) -> List[str]:
    """
    Нормализует список модификаторов: uppercase, CONTROL->CTRL, убирает дубли, сохраняет порядок.
    Допустимые: CTRL, ALT, SHIFT, WIN
    """
    out = []
    seen = set()
    for m in modifiers:
        if not m:
            continue
        mm = str(m).upper()
        if mm == "CONTROL":
            mm = "CTRL"
        if mm == "WINDOWS":
            mm = "WIN"
        if mm not in ("CTRL", "ALT", "SHIFT", "WIN"):
            continue
        if mm in seen:
            continue
        seen.add(mm)
        out.append(mm)
    return out

def write_exit_hotkey(modifiers: List[str], key: str) -> bool:
    """
    Записывает exit_hotkey в CONFIG_FILE (атомарно). Нормализует модификаторы.
    """
    mods = normalize_modifiers(modifiers or [])
    key = (key or "F10").upper()
    with _config_write_lock:
        j = _reload_config_if_needed()
        j["exit_hotkey"] = {"modifiers": mods, "key": key}
        ok = _atomic_write_json(CONFIG_FILE, j)
        if ok:
            logger.info("config: записан exit_hotkey=%s", j["exit_hotkey"])
        else:
            logger.exception("config: не удалось записать exit_hotkey в %s", CONFIG_FILE)
        return ok

# Load config on startup
load_config()

# ---------------- startup shortcut helpers ----------------
def _startup_shortcut_path() -> str:
    appdata = os.environ.get("APPDATA")
    if not appdata:
        appdata = os.path.expanduser("~\\AppData\\Roaming")
    startup_dir = os.path.join(appdata, "Microsoft", "Windows", "Start Menu", "Programs", "Startup")
    return os.path.join(startup_dir, f"{APP_NAME}.lnk")

def create_startup_shortcut() -> bool:
    shortcut_path = _startup_shortcut_path()
    try:
        try:
            import win32com.client  # type: ignore
        except Exception:
            logger.exception("create_startup_shortcut: win32com.client недоступен, не могу создать ярлык")
            return False
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
    except Exception:
        logger.exception("startup: не удалось создать ярлык автозагрузки")
        return False

def remove_startup_shortcut() -> bool:
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
    logger.warning("AUTORUN: apply_autorun_setting CALLED")
    try:
        if read_autorun_flag():
            create_startup_shortcut()
        else:
            remove_startup_shortcut()
    except Exception:
        logger.exception("apply_autorun_setting: ошибка при применении autorun")

# Применим autorun при старте
try:
    apply_autorun_setting()
except Exception:
    logger.exception("startup: apply_autorun_setting on startup failed")

# ------------ Hotkey constants & thread messaging ------------
HOTKEY_ID_F4 = 1
HOTKEY_ID_EXIT = 2  # id для динамической комбинации выхода
MOD_NONE = 0

# We'll use PostThreadMessage to request registration/unregistration in the hotkey thread
WM_USER = 0x0400
MSG_REGISTER_F4 = WM_USER + 1
MSG_UNREGISTER_F4 = WM_USER + 2
MSG_UPDATE_EXIT_HOTKEY = WM_USER + 3

HOTKEY_THREAD_ID = 0  # заполнится в win_hotkey_loop

def post_register_f4(should_register: bool):
    global HOTKEY_THREAD_ID
    if HOTKEY_THREAD_ID == 0:
        logger.warning("post_register_f4: HOTKEY_THREAD_ID not ready yet -> cannot post")
        return False
    msg = MSG_REGISTER_F4 if should_register else MSG_UNREGISTER_F4
    wparam = 1 if should_register else 0
    try:
        res = user32.PostThreadMessageW(HOTKEY_THREAD_ID, msg, wparam, 0)
        if res == 0:
            err = kernel32.GetLastError()
            logger.error("post_register_f4: PostThreadMessageW failed (err=%d)", err)
            return False
        logger.debug("post_register_f4: posted msg=%d wParam=%d to thread %d", msg, wparam, HOTKEY_THREAD_ID)
        return True
    except Exception:
        logger.exception("post_register_f4: exception while posting")
        return False

def post_update_exit_hotkey():
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
    except Exception:
        logger.exception("post_update_exit_hotkey: exception while posting")
        return False

# ------------ Hotkey register helpers (used only from hotkey thread) ------------
def register_f4_in_thread():
    try:
        ok = user32.RegisterHotKey(None, HOTKEY_ID_F4, MOD_NONE, 0x73)  # F4 VK = 0x73
        if ok:
            logger.debug("register_f4: F4 registered in hotkey thread")
        else:
            err = kernel32.GetLastError()
            logger.error("register_f4: failed to register F4 (err=%d)", err)
    except Exception:
        logger.exception("register_f4: exception while registering F4")

def unregister_f4_in_thread():
    try:
        ok = user32.UnregisterHotKey(None, HOTKEY_ID_F4)
        if ok:
            logger.debug("unregister_f4: F4 unregistered in hotkey thread")
        else:
            logger.debug("unregister_f4: UnregisterHotKey returned False (probably not registered)")
    except Exception:
        logger.exception("unregister_f4: exception while unregistering F4")

# ------------- Exit hotkey helpers (dynamic) ----------------
# Modifiers masks for RegisterHotKey
MOD_ALT = 0x0001
MOD_CONTROL = 0x0002
MOD_SHIFT = 0x0004
MOD_WIN = 0x0008

# map textual names to VK codes for common keys
VK_MAP = {
    "F1": 0x70, "F2": 0x71, "F3": 0x72, "F4": 0x73, "F5": 0x74, "F6": 0x75,
    "F7": 0x76, "F8": 0x77, "F9": 0x78, "F10": 0x79, "F11": 0x7A, "F12": 0x7B,
    "ESC": 0x1B, "TAB": 0x09, "RETURN": 0x0D, "ENTER": 0x0D, "SPACE": 0x20,
    "INSERT": 0x2D, "DELETE": 0x2E, "HOME": 0x24, "END": 0x23, "PGUP": 0x21, "PGDN": 0x22,
    "LEFT": 0x25, "UP": 0x26, "RIGHT": 0x27, "DOWN": 0x28,
}

def key_name_to_vk(name: str) -> int:
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
    except Exception:
        pass
    return VK_MAP.get("F10", 0x79)

def modifiers_list_to_mask(mods: List[str]) -> int:
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

def update_exit_hotkey_in_thread():
    try:
        try:
            user32.UnregisterHotKey(None, HOTKEY_ID_EXIT)
        except Exception:
            pass
        eh = read_exit_hotkey()
        mods = eh.get("modifiers", []) or []
        key = eh.get("key", "F10") or "F10"
        mask = modifiers_list_to_mask(mods)
        vk = key_name_to_vk(key)
        ok = user32.RegisterHotKey(None, HOTKEY_ID_EXIT, mask, vk)
        if ok:
            logger.debug("update_exit_hotkey_in_thread: registered exit hotkey %s + %s (mask=0x%X vk=0x%X)",
                         "+".join(mods) if mods else "(no modifiers)", key, mask, vk)
        else:
            err = kernel32.GetLastError()
            logger.error("update_exit_hotkey_in_thread: failed to register exit hotkey (err=%d) mask=0x%X vk=0x%X",
                         err, mask, vk)
    except Exception:
        logger.exception("update_exit_hotkey_in_thread: exception")

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
        try:
            saved = None
            try:
                saved = pyperclip.paste()
            except Exception:
                saved = None
            logger.debug("handle: прочитал (но не буду восстанавливать) буфер text-len=%s", None if saved is None else len(saved))
        except Exception:
            saved = None
        try:
            hwnd = user32.GetForegroundWindow()
            pid_c = ctypes.c_ulong()
            thread_id = user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid_c))
            hkl = user32.GetKeyboardLayout(thread_id) & 0xFFFF
            logger.debug("handle: foreground thread hkl=0x%04X", hkl)
        except Exception:
            hkl = 0
            logger.exception("handle: не удалось получить HKL, предполагаем EN")
        selected = safe_copy_from_selection(timeout_per_attempt=0.6, max_attempts=2)
        if selected is None:
            selected = ""
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
        except Exception as e:
            logger.exception("handle: delete exception: %s", e)
        try:
            send_unicode_via_sendinput(converted, delay_between_keys=0.001)
            logger.debug("handle: вставил через SendInput (unicode) — без использования clipboard")
        except Exception as e:
            logger.exception("handle: send_unicode_via_sendinput failed: %s", e)
            try:
                pyperclip.copy(converted)
                time.sleep(0.02)
                send_ctrl_v()
                logger.debug("handle: fallback: вставил через буфер (ctrl+v)")
            except Exception as e2:
                logger.exception("handle: fallback paste failed: %s", e2)
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
        logger.info("handle: завершено успешно. (буфер оставлен как есть — в нём будет выделение)")
    except Exception as e:
        logger.exception("handle_hotkey_transform: исключение при обработке: %s", e)

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
        d.rectangle((8, 16, 56, 48), outline=(255,255,255), width=2)
        d.text((16, 12), "KF", fill=(255,255,255))
    except Exception:
        pass
    return img

def enabled_menu_text(item):
    return "❌ Выключить" if is_enabled() else "✅ Включить"

def toggle_enabled(icon, item):
    new_state = not is_enabled()
    set_enabled(new_state)
    logger.info("Tray: toggled enabled -> %s", new_state)
    try:
        ok = post_register_f4(new_state)
        if not ok:
            logger.warning("toggle_enabled: post_register_f4 returned False")
    except Exception:
        logger.exception("toggle_enabled: failed to post register/unregister request")
    try:
        if icon is not None:
            icon.icon = prepare_tray_icon_image(new_state)
    except Exception:
        logger.exception("Tray: failed to update icon image after toggle")

def toggle_file_logging(icon, item):
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

# ---- Exit hotkey menu actions
def _set_exit_hotkey_and_apply(mods: List[str], key: str):
    ok = write_exit_hotkey(mods, key)
    if ok:
        post_update_exit_hotkey()
    return ok

def menu_set_exit_f10(icon, item):
    _set_exit_hotkey_and_apply([], "F10")
def menu_set_exit_ctrl_q(icon, item):
    _set_exit_hotkey_and_apply(["CTRL"], "Q")
def menu_set_exit_ctrl_alt_x(icon, item):
    _set_exit_hotkey_and_apply(["CTRL","ALT"], "X")

def on_exit(icon=None, item=None):
    logger.info("Выход запрошен (через трей/горячую клавишу)")
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
    settings_menu = Menu(
        MenuItem("Логирование в файл", toggle_file_logging, checked=lambda item: read_file_logging_flag()),
        MenuItem("Автозапуск при старте Windows", toggle_autorun, checked=lambda item: read_autorun_flag()),
        MenuItem(
            "Комбинация выхода",
            Menu(
                MenuItem(
                    "F10 (по умолчанию)",
                    menu_set_exit_f10,
                    checked=lambda item: (
                            read_exit_hotkey().get("key", "") == "F10"
                            and not read_exit_hotkey().get("modifiers")
                    )
                ),
                MenuItem(
                    "Ctrl+Q",
                    menu_set_exit_ctrl_q,
                    checked=lambda item: (
                            read_exit_hotkey().get("key", "") == "Q"
                            and "CTRL" in read_exit_hotkey().get("modifiers", [])
                    )
                ),
                MenuItem(
                    "Ctrl+Alt+X",
                    menu_set_exit_ctrl_alt_x,
                    checked=lambda item: (
                            read_exit_hotkey().get("key", "") == "X"
                            and {"CTRL", "ALT"} == set(read_exit_hotkey().get("modifiers", []))
                    ),
                ),
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
    except Exception:
        logger.exception("tray_worker: failed to set initial icon")
    try:
        logger.debug("tray_worker: запуск иконки")
        icon.run()
    except Exception as e:
        logger.exception("tray_worker: исключение: %s", e)

# ------------ WinAPI hotkey loop ------------
def win_hotkey_loop():
    global HOTKEY_THREAD_ID
    try:
        HOTKEY_THREAD_ID = kernel32.GetCurrentThreadId()
        logger.debug("win_hotkey_loop: HOTKEY_THREAD_ID = %d", HOTKEY_THREAD_ID)
    except Exception:
        HOTKEY_THREAD_ID = 0
        logger.exception("win_hotkey_loop: failed to get thread id")

    # Register dynamic exit hotkey here by calling updater (so it's done in this thread)
    update_exit_hotkey_in_thread()

    if is_enabled():
        register_f4_in_thread()
    else:
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
            if msg.message == MSG_REGISTER_F4:
                logger.debug("win_hotkey_loop: MSG_REGISTER_F4 received wParam=%s", int(msg.wParam))
                try:
                    register_f4_in_thread()
                except Exception:
                    logger.exception("win_hotkey_loop: register_f4_in_thread failed")
                continue
            elif msg.message == MSG_UNREGISTER_F4:
                logger.debug("win_hotkey_loop: MSG_UNREGISTER_F4 received")
                try:
                    unregister_f4_in_thread()
                except Exception:
                    logger.exception("win_hotkey_loop: unregister_f4_in_thread failed")
                continue
            elif msg.message == MSG_UPDATE_EXIT_HOTKEY:
                logger.debug("win_hotkey_loop: MSG_UPDATE_EXIT_HOTKEY received -> update exit hotkey")
                try:
                    update_exit_hotkey_in_thread()
                except Exception:
                    logger.exception("win_hotkey_loop: update_exit_hotkey_in_thread failed")
                continue
            if msg.message == 0x0312:  # WM_HOTKEY
                hotkey_id = msg.wParam
                if hotkey_id == HOTKEY_ID_F4:
                    if not is_enabled():
                        logger.debug("WM_HOTKEY: F4 received but currently disabled -> ignored")
                    else:
                        logger.debug("WM_HOTKEY: F4 received")
                        _f4_invoker()
                elif hotkey_id == HOTKEY_ID_EXIT:
                    logger.debug("WM_HOTKEY: EXIT hotkey received -> exiting")
                    on_exit()
            user32.TranslateMessage(ctypes.byref(msg))
            user32.DispatchMessageW(ctypes.byref(msg))
    finally:
        try:
            user32.UnregisterHotKey(None, HOTKEY_ID_F4)
            user32.UnregisterHotKey(None, HOTKEY_ID_EXIT)
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
    try:
        post_update_exit_hotkey()
    except Exception:
        logger.exception("main: failed to post initial update_exit_hotkey")
    logger.info("%s запущен. Нажми F4 чтобы преобразовать выделение", APP_NAME)
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
