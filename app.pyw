"""
Точка входа приложения KeyFlip.

Оркестрация:
- проверка mutex (один экземпляр)
- загрузка конфигурации и применение autorun
- запуск потока обработки Win hotkey сообщений
- запуск трей-иконки (pystray) в отдельном потоке
- ожидание завершения через exit_event
"""

import sys
import threading
import time
from typing import Optional, Tuple

from config import APP_NAME, MUTEX_NAME, load_config, apply_autorun_setting
from logging_setup import logger
from tray import tray_worker, on_exit
from winapi import win_hotkey_loop, exit_event, register_exit_handler

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
        try:
            win32event.ReleaseMutex(mutex_handle)
        except Exception:  # noqa
            # иногда ReleaseMutex может выдать ошибку — логируем и пытаемся закрыть handle
            logger.debug("Не удалось ReleaseMutex, пробуем CloseHandle")
        try:
            win32api.CloseHandle(mutex_handle)
        except Exception:  # noqa
            logger.exception("Не удалось CloseHandle для mutex_handle")
    except Exception:  # noqa
        logger.exception("_release_mutex: unexpected exception")


def _start_worker_threads() -> Tuple[threading.Thread, threading.Thread]:
    """
    Запустить два daemon-потока:
    - WinHotkeyThread: loop для обработки WM_HOTKEY и внутренних сообщений
    - TrayThread: pystray icon.run()
    Возвращает (hk_thread, tray_thread).
    """
    hk_thread = threading.Thread(target=win_hotkey_loop, name="WinHotkeyThread", daemon=True)
    tray_thread = threading.Thread(target=tray_worker, name="TrayThread", daemon=True)
    hk_thread.start()
    logger.debug("Запущен поток WinHotkeyThread")
    tray_thread.start()
    logger.debug("Запущен поток TrayThread")
    return hk_thread, tray_thread


def main() -> None:
    """Главная функция запуска приложения."""
    logger.info("%s: старт приложения", APP_NAME)

    # Один экземпляр
    mutex_handle = _create_single_instance_mutex()
    if mutex_handle is None:
        logger.info("%s: обнаружен другой экземпляр -> выход", APP_NAME)
        sys.exit(0)

    # Подключаем обработчик выхода (чтобы on_exit вызывался из winapi при нажатии exit-hotkey)
    try:
        register_exit_handler(on_exit)
    except Exception:  # noqa
        logger.exception("Не удалось зарегистрировать exit handler")

    # Загрузка конфига и применение autorun (если нужно)
    try:
        load_config()
    except Exception:  # noqa
        logger.exception("Ошибка при load_config() — продолжаем попытку запуска")

    try:
        apply_autorun_setting()
    except Exception:  # noqa
        logger.exception("apply_autorun_setting failed on startup")

    # Запуск потоков
    try:
        hk_thread, tray_thread = _start_worker_threads()
    except Exception as e:  # noqa
        logger.exception("Не удалось запустить потоки: %s", e)
        _release_mutex(mutex_handle)
        raise

    # Основное ожидание: пока не будет установлен exit_event
    try:
        logger.info("Ожидание события завершения...")
        exit_event.wait()  # блокирующее ожидание
        logger.info("Сигнал выхода получен.")
    except KeyboardInterrupt:
        logger.info("KeyboardInterrupt получен, инициирую завершение.")
        try:
            on_exit()
        except Exception:  # noqa
            logger.exception("Ошибка при on_exit() на KeyboardInterrupt")
    except Exception as e:  # noqa
        logger.exception("main: unexpected exception while waiting: %s", e)
    finally:
        # Гарантированное освобождение mutex
        try:
            _release_mutex(mutex_handle)
        except Exception:  # noqa
            logger.exception("Не удалось освободить mutex при завершении")
        # Небольшая пауза, чтобы потоки успели корректно завершиться
        try:
            time.sleep(0.15)
        except Exception:  # noqa
            pass
        logger.info("%s: завершение main()", APP_NAME)


if __name__ == "__main__":
    main()
