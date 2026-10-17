"""
Точка входа приложения KeyFlip.

Оркестрация:
- проверка mutex (один экземпляр)
- загрузка конфигурации и применение autorun
- запуск потока обработки Win hotkey сообщений
- интерфейс Qt и трей в главном потоке
- завершение через exit_event
"""

import importlib.util
import os
import sys
import threading
import traceback
from tkinter import messagebox, Tk
from typing import Optional

from config import APP_NAME, DATA_DIR, MUTEX_NAME, load_config, apply_autorun_setting
from logging_setup import logger

_DEPENDENCY_IMPORT_ERROR: Exception | None = None
STARTUP_ERROR_FILE = os.path.join(DATA_DIR, "keyflip-startup-error.log")

try:
    import pyperclip  # noqa: F401
    import psutil  # noqa: F401
    import win32api
    import win32con  # noqa: F401
    import win32event
    import win32com.client  # noqa: F401
    from PySide6.QtWidgets import QApplication  # noqa: F401
except Exception as e:  # noqa
    _DEPENDENCY_IMPORT_ERROR = e


def _module_available(module_name: str) -> bool:
    """Проверить наличие модуля без аварийного падения при странном состоянии пакета."""
    try:
        return importlib.util.find_spec(module_name) is not None
    except Exception:  # noqa
        return False


def check_dependencies() -> None:
    """Проверяем все необходимые пакеты. Если чего-то нет — сразу говорим и выходим."""

    if getattr(sys, "frozen", False) and _DEPENDENCY_IMPORT_ERROR is None:
        return

    missing = []

    required = {
        "pyperclip": "pyperclip",
        "PySide6": "PySide6",
        "psutil": "psutil",
        "win32api": "pywin32",
        "win32com.client": "pywin32",
        "win32con": "pywin32",
        "win32event": "pywin32",
    }

    for module_name, pip_name in required.items():
        if not _module_available(module_name) and pip_name not in missing:
            missing.append(pip_name)

    if not missing and _DEPENDENCY_IMPORT_ERROR is None:
        return

    title = "KeyFlip — ошибка зависимостей"
    details = ""
    if _DEPENDENCY_IMPORT_ERROR is not None:
        details = (
            "\n\nДетали ошибки импорта:\n"
            f"{type(_DEPENDENCY_IMPORT_ERROR).__name__}: {_DEPENDENCY_IMPORT_ERROR}"
        )
    install_hint = (
        "Переустанови KeyFlip из установщика."
        if getattr(sys, "frozen", False)
        else "Установи одной командой:\npip install -r requirements.txt"
    )
    msg = (
        "❌ KeyFlip не может запуститься — отсутствуют или некорректно работают библиотеки:\n\n"
        f"{' '.join(missing) if missing else 'см. детали ниже'}\n\n"
        f"{install_hint}\n\n"
        f"После установки перезапусти приложение.{details}"
    )
    try:
        root = Tk()
        root.withdraw()  # скрываем главное окно
        messagebox.showerror(title, msg)
        root.destroy()
    except Exception:  # noqa
        print(msg)
    sys.exit(1)


# Код ошибки, когда mutex уже существует (Windows)
ERROR_ALREADY_EXISTS = 183


def _create_single_instance_mutex() -> Optional[int]:
    """
    Попытаться создать глобальный mutex и вернуть handle.
    Если другой экземпляр уже запущен — показать MessageBox и вернуть None.
    """
    try:
        mutex = win32event.CreateMutex(None, False, MUTEX_NAME)
    except Exception as e:  # noqa
        logger.exception("Создание mutex не удалось: %s", e)
        return None

    try:
        last = win32api.GetLastError()
        if last == ERROR_ALREADY_EXISTS:
            # Другой процесс уже создал mutex
            try:
                win32api.MessageBox(0, f"Программа {APP_NAME} уже запущена!", APP_NAME, 0)
            except Exception:  # noqa
                print(f"{APP_NAME} уже запущена!")
            # освободим handle, если он всё же вернулся — безопасно закрыть
            try:
                if mutex:
                    win32api.CloseHandle(mutex)
            except Exception:  # noqa
                pass
            return None
    except Exception:  # noqa
        # не удалось прочитать код ошибки — считаем, что mutex создан успешно
        pass

    return mutex


def _release_mutex(mutex_handle: Optional[int]) -> None:
    """Безопасно освободить mutex (если он есть)."""
    if not mutex_handle:
        return
    try:
        # CreateMutex(..., False, ...) не захватывает владение: достаточно закрыть handle.
        win32api.CloseHandle(mutex_handle)
    except Exception:  # noqa
        logger.exception("_release_mutex: unexpected exception")


def installer_operation(arguments: list[str]) -> int:
    """Проверка для установщика без запуска трея и вмешательства в горячие клавиши."""
    if arguments not in (["--installer-check"], ["--uninstall-cleanup"]):
        return 2
    try:
        mutex = win32event.CreateMutex(None, False, MUTEX_NAME)
        try:
            if win32api.GetLastError() == ERROR_ALREADY_EXISTS:
                return 3
        finally:
            if mutex:
                win32api.CloseHandle(mutex)
        if arguments == ["--uninstall-cleanup"]:
            from config import _startup_shortcut_path

            shortcut_path = _startup_shortcut_path()
            if os.path.isfile(shortcut_path):
                shell = win32com.client.Dispatch("WScript.Shell")
                shortcut = shell.CreateShortcut(shortcut_path)
                if os.path.normcase(os.path.abspath(shortcut.TargetPath)) == os.path.normcase(
                    os.path.abspath(sys.executable)
                ):
                    os.remove(shortcut_path)
        return 0
    except Exception:
        logger.exception("Операция установщика не удалась")
        return 1


def main() -> None:
    """Главная функция запуска приложения."""
    logger.info("%s: старт приложения", APP_NAME)

    # Один экземпляр
    mutex_handle = _create_single_instance_mutex()
    if mutex_handle is None:
        logger.info("%s: обнаружен другой экземпляр -> выход", APP_NAME)
        sys.exit(0)

    winapi = None
    controller = None
    hk_thread = None
    try:
        import winapi
        from tray import on_exit
        from ui import TrayController

        app = QApplication(sys.argv)
        app.setApplicationName(APP_NAME)
        from PySide6.QtGui import QIcon
        from config import ICON_ON
        app.setWindowIcon(QIcon(ICON_ON))
        app.setQuitOnLastWindowClosed(False)
        winapi.register_exit_handler(on_exit)
        load_config()
        try:
            apply_autorun_setting()
        except Exception:  # noqa
            logger.exception("apply_autorun_setting failed on startup")
        hk_thread = threading.Thread(target=winapi.win_hotkey_loop, name="WinHotkeyThread", daemon=True)
        hk_thread.start()
        if not winapi.hotkey_ready.wait(timeout=3):
            raise RuntimeError("Поток горячих клавиш не запустился")
        controller = TrayController()
        app.aboutToQuit.connect(controller.close)
        controller.start()
        logger.info("Интерфейс Qt запущен")
        app.exec()
    except KeyboardInterrupt:
        if winapi is not None:
            winapi.exit_event.set()
    finally:
        if winapi is not None:
            winapi.exit_event.set()
        if controller is not None:
            controller.close()
        if hk_thread is not None:
            try:
                if winapi.HOTKEY_THREAD_ID:
                    win32api.PostThreadMessage(winapi.HOTKEY_THREAD_ID, win32con.WM_QUIT, 0, 0)
            except Exception:  # noqa
                logger.debug("Поток горячих клавиш уже завершён", exc_info=True)
            hk_thread.join(timeout=1)
        if winapi is not None:
            winapi.wait_for_pending_operations()
        _release_mutex(mutex_handle)
        logger.info("%s: завершение main()", APP_NAME)


def _report_startup_error(error: Exception) -> None:
    """При запуске через pythonw ошибка видна пользователю и доступна в отдельном журнале."""
    logger.exception("Ошибка запуска KeyFlip")
    details = f"{type(error).__name__}: {error}"
    try:
        with open(STARTUP_ERROR_FILE, "w", encoding="utf-8") as report:
            report.write(traceback.format_exc())
        details += f"\n\nПодробности: {STARTUP_ERROR_FILE}"
    except OSError:
        logger.exception("Не удалось записать ошибку запуска")
    try:
        win32api.MessageBox(0, details, APP_NAME + " — ошибка запуска", 0x10)
    except Exception:
        print(details, file=sys.stderr)


if __name__ == "__main__":
    try:
        check_dependencies()
        if len(sys.argv) > 1:
            sys.exit(installer_operation(sys.argv[1:]))
        main()
    except Exception as error:
        _report_startup_error(error)
        sys.exit(1)
