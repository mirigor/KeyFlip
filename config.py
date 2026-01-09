"""
Конфигурация и состояние приложения KeyFlip.

Содержит:
- пути и константы (APP_NAME, BASE_DIR, CONFIG_FILE, ICON_* и т.д.)
- чтение/запись JSON-конфига
- флаги file_logging, autorun
- чтение/запись хоткеев (exit/translate)
- управление enabled (в памяти + сохранение)
- создание/удаление ярлыка в автозагрузке
"""

import json
import logging
import os
import threading

APP_NAME = "KeyFlip"
MUTEX_NAME = r"Global\KeyFlipMutex"
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
LOG_FILE = os.path.join(BASE_DIR, "keyflip.log")
ICON_ON = os.path.join(BASE_DIR, "icon_on.ico")
ICON_OFF = os.path.join(BASE_DIR, "icon_off.ico")
CONFIG_FILE = os.path.join(BASE_DIR, "keyflip.json")

# локальный логгер (handlers будут настроены в logging_setup.py)
logger = logging.getLogger(APP_NAME)


# ---------------- Работа с JSON конфигом ----------------
def read_json_config() -> dict:
    """Прочитать JSON-конфиг и вернуть словарь (или пустой словарь при ошибке/отсутствии)."""
    try:
        if not os.path.isfile(CONFIG_FILE):
            return {}
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            return json.load(f) or {}
    except Exception:
        # Не используем logger.exception тут слишком много — достаточно debug/info
        logger.debug("read_json_config: не удалось прочитать конфиг, возвращаю {}")
        return {}


def write_json_config(j: dict) -> bool:
    """Атомарная запись JSON: записать в tmp-файл, затем заменить основной файл."""
    try:
        tmp = CONFIG_FILE + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(j, f, ensure_ascii=False, indent=2)
        try:
            os.replace(tmp, CONFIG_FILE)
        except Exception:
            # fallback: удалить старый и переименовать
            if os.path.exists(CONFIG_FILE):
                try:
                    os.remove(CONFIG_FILE)
                except Exception:
                    pass
            os.rename(tmp, CONFIG_FILE)
        return True
    except Exception:
        logger.exception("config: write_json_config failed")
        return False


# ---------------- file_logging ----------------
def read_file_logging_flag() -> bool:
    try:
        j = read_json_config()
        return bool(j.get("file_logging", False))
    except Exception:
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
    except Exception:
        logger.exception("config: write_file_logging_flag failed")
        return False


# ---------------- autorun ----------------
def _startup_shortcut_path() -> str:
    """Вернуть путь к .lnk в папке автозагрузки текущего пользователя."""
    appdata = os.environ.get("APPDATA") or os.path.expanduser("~\\AppData\\Roaming")
    startup_dir = os.path.join(appdata, "Microsoft", "Windows", "Start Menu", "Programs", "Startup")
    return os.path.join(startup_dir, f"{APP_NAME}.lnk")


def create_startup_shortcut() -> bool:
    """Создать .lnk в папке автозагрузки текущего пользователя. Возвращает True/False."""
    shortcut_path = _startup_shortcut_path()
    try:
        # импортируем локально, чтобы не требовать pywin32 при частых операциях
        import win32com.client  # type: ignore
        import sys
        shell = win32com.client.Dispatch("WScript.Shell")
        shortcut = shell.CreateShortcut(shortcut_path)
        target = sys.executable
        script = os.path.abspath(sys.argv[0]) if hasattr(sys, "argv") and sys.argv else ""
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
    """Удалить .lnk из автозагрузки (если есть)."""
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


def read_autorun_flag() -> bool:
    try:
        j = read_json_config()
        return bool(j.get("autorun", False))
    except Exception:
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
    except Exception:
        logger.exception("config: write_autorun_flag failed")
        return False


def apply_autorun_setting() -> None:
    """Применить текущее значение autorun: создать или удалить ярлык."""
    try:
        logger.debug("AUTORUN: apply_autorun_setting CALLED")
        if read_autorun_flag():
            create_startup_shortcut()
        else:
            remove_startup_shortcut()
    except Exception:
        logger.exception("apply_autorun_setting: ошибка при применении autorun")


# ---------------- Exit hotkey (конфиг) ----------------
def default_exit_hotkey() -> dict:
    """Значение по умолчанию для комбинации выхода."""
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
    except Exception:
        return default_exit_hotkey()


def write_exit_hotkey(modifiers: list[str], key: str) -> bool:
    """Нормализовать и записать exit_hotkey в конфиг."""
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
    except Exception:
        logger.exception("config: write_exit_hotkey failed")
        return False


# ---------------- Translate hotkey (конфиг) ----------------
def default_translate_hotkey() -> dict:
    """Значение по умолчанию для комбинации перевода."""
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
    except Exception:
        return default_translate_hotkey()


def write_translate_hotkey(modifiers: list[str], key: str) -> bool:
    """Нормализовать и записать translate_hotkey в конфиг."""
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
    except Exception:
        logger.exception("config: write_translate_hotkey failed")
        return False


# ---------------- enabled state ----------------
_enabled_lock = threading.Lock()
_enabled: bool = True


def load_config() -> None:
    """Загрузить конфиг и установить дефолты при их отсутствии."""
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
    """Сохранить минимальное состояние (enabled) в конфиг."""
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


def set_enabled(val: bool) -> None:
    global _enabled
    with _enabled_lock:
        _enabled = bool(val)
    save_config()
    logger.info("state: enabled set to %s", _enabled)
