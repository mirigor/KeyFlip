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
from typing import Optional

from config import APP_NAME, MUTEX_NAME, load_config, apply_autorun_setting
from logging_setup import logger
from tray import tray_worker, on_exit
from winapi import win_hotkey_loop, exit_event

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
    Если другой экземпляр уже запущен — показать MessageBox и завершить процесс.
    """
    try:
        mutex = win32event.CreateMutex(None, False, MUTEX_NAME)
    except Exception as e:
        logger.exception("Создание mutex не удалось: %s", e)
        return None

    try:
        last = win32api.GetLastError()
        if last == ERROR_ALREADY_EXISTS:
            # Попытка уведомить пользователя о том, что приложение уже запущено
            try:
                win32api.MessageBox(0, f"Программа {APP_NAME} уже запущена!", APP_NAME, 0)
            except Exception:
                # fallback в консоль
                print(f"{APP_NAME} уже запущена!")
            return None
    except Exception:
        # если не можем прочитать ошибку — считаем что mutex создан
        pass

    return mutex


def main() -> None:
    logger.info("%s: старт приложения", APP_NAME)

    # Один экземпляр
    mutex_handle = _create_single_instance_mutex()
    if mutex_handle is None:
        logger.info("%s: обнаружен другой экземпляр -> выход", APP_NAME)
        sys.exit(0)

    # Загрузка конфига и применение autorun (если нужно)
    try:
        load_config()
    except Exception:
        logger.exception("Ошибка при load_config() — продолжаем попытку запуска")

    try:
        # apply_autorun_setting может использовать win32com, обработаем исключения внутри
        apply_autorun_setting()
    except Exception:
        logger.exception("apply_autorun_setting failed on startup")

    # Запуск потоков: hotkey loop и трей
    try:
        hk_thread = threading.Thread(target=win_hotkey_loop, name="WinHotkeyThread", daemon=True)
        hk_thread.start()
        logger.debug("Запущен поток WinHotkeyThread")

        tray_thread = threading.Thread(target=tray_worker, name="TrayThread", daemon=True)
        tray_thread.start()
        logger.debug("Запущен поток TrayThread")
    except Exception as e:
        logger.exception("Не удалось запустить потоки: %s", e)
        # в случае критической ошибки завершаем
        try:
            win32event.ReleaseMutex(mutex_handle)
        except Exception:
            pass
        raise

    # Основное ожидание: пока не будет установлен exit_event
    try:
        logger.info("Ожидание события завершения...")
        # exit_event — threading.Event из winapi
        exit_event.wait()
        logger.info("Сигнал выхода получен.")
    except KeyboardInterrupt:
        logger.info("KeyboardInterrupt получен, инициирую завершение.")
        try:
            # корректно попросим завершиться
            on_exit()
        except Exception:
            logger.exception("Ошибка при on_exit() на KeyboardInterrupt")
    finally:
        # Гарантированное освобождение mutex
        try:
            if mutex_handle:
                win32event.ReleaseMutex(mutex_handle)
                logger.debug("Mutex освобождён")
        except Exception:
            logger.exception("Не удалось освободить mutex при завершении")

        # Небольшая пауза, чтобы потоки успели завершиться нормально (не блокирующая)
        time.sleep(0.15)
        logger.info("%s: завершение main()", APP_NAME)


if __name__ == "__main__":
    main()
