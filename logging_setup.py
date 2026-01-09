"""
Инициализация логгера для KeyFlip.

Этот модуль настраивает:
- ротационный файловый хендлер (с фильтром, который читает флаг file_logging из конфига)
- консольный хендлер (stdout)
- общий логгер доступен как `logger`
"""

import logging
import sys
from logging.handlers import RotatingFileHandler

from config import LOG_FILE, APP_NAME, read_file_logging_flag

# Получаем логгер (config.py тоже использует logging.getLogger(APP_NAME) — ссылка на один объект)
logger = logging.getLogger(APP_NAME)
logger.setLevel(logging.DEBUG)

# Форматтер
fmt = logging.Formatter("%(asctime)s %(levelname)s [%(threadName)s] %(message)s", "%Y-%m-%d %H:%M:%S")

# ---------------- Файловый хендлер с фильтром ----------------
def _create_file_handler() -> RotatingFileHandler:
    """
    Создать RotatingFileHandler с фильтром, который проверяет флаг file_logging.
    Возвращает созданный handler (еще не привязан к логгеру).
    """
    fh = RotatingFileHandler(LOG_FILE, maxBytes=5 * 1024 * 1024, backupCount=3, encoding="utf-8")
    fh.setFormatter(fmt)
    fh.setLevel(logging.DEBUG)

    class FileLoggingFilter(logging.Filter):
        """
        Фильтр, который при каждой попытке логгирования проверяет флаг file_logging в конфиге.
        Если флаг выключен — записи в файл не происходит.
        """

        def filter(self, record: logging.LogRecord) -> bool:  # type: ignore[override]
            try:
                return bool(read_file_logging_flag())
            except Exception:  # noqa
                # При ошибке безопасности — не логировать в файл.
                return False

    fh.addFilter(FileLoggingFilter())
    return fh


# ---------------- Консольный (stdout) хендлер ----------------
def _create_console_handler() -> logging.StreamHandler:
    """Создать и вернуть консольный хендлер (stdout)."""
    ch = logging.StreamHandler(sys.stdout)
    ch.setFormatter(fmt)
    ch.setLevel(logging.INFO)
    return ch


# ---------------- Установка хендлеров на логгер ----------------
def _install_handlers() -> None:
    """
    Установить (инициализировать) хендлеры на глобальный logger.
    Вызывается при импорте модуля.
    """
    # чтобы не дублировать хендлеры при повторном импорте/повторной инициализации —
    # очистим существующие обработчики, если они есть (но не удаляем root handlers).
    try:
        # оставляем только хендлеры, которые НЕ принадлежат нашему логгеру, чтобы избежать дублирования
        existing = list(logger.handlers)
        if existing:
            # Если уже есть наши хендлеры — удалим их (чтобы переустановить свежие)
            for h in existing:
                try:
                    logger.removeHandler(h)
                except Exception:  # noqa
                    pass
    except Exception:  # noqa
        # не критично, продолжим
        pass

    try:
        fh = _create_file_handler()
        ch = _create_console_handler()
        logger.addHandler(fh)
        logger.addHandler(ch)
    except Exception:  # noqa
        # Если установка хендлеров провалилась — попробуем хотя бы консольный хендлер.
        try:
            ch = _create_console_handler()
            logger.addHandler(ch)
        except Exception:  # noqa
            # в крайнем случае — молча пропустить
            pass


# Выполняем установку при импорте модуля
_install_handlers()

# Информационное сообщение о старте логирования
try:
    logger.info("%s запуск — лог: %s", APP_NAME, LOG_FILE)
except Exception:  # noqa
    # аккуратно молчим, если даже логирование старта невозможно
    pass
