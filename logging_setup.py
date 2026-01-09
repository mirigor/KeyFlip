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

# Получаем логгер (config.py создаёт такой же getLogger, поэтому ссылка будет единой)
logger = logging.getLogger(APP_NAME)
logger.setLevel(logging.DEBUG)

# Форматтер
fmt = logging.Formatter("%(asctime)s %(levelname)s [%(threadName)s] %(message)s", "%Y-%m-%d %H:%M:%S")

# ---------------- Файловый хендлер с фильтром ----------------
file_handler = RotatingFileHandler(LOG_FILE, maxBytes=5 * 1024 * 1024, backupCount=3, encoding="utf-8")
file_handler.setFormatter(fmt)


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


file_handler.addFilter(FileLoggingFilter())
logger.addHandler(file_handler)

# ---------------- Консольный (stdout) хендлер ----------------
console = logging.StreamHandler(sys.stdout)
console.setFormatter(fmt)
console.setLevel(logging.INFO)
logger.addHandler(console)

# Информационное сообщение о старте логирования
try:
    logger.info("%s запуск — лог: %s", APP_NAME, LOG_FILE)
except Exception:
    # В случаях, когда логгер ещё не полностью инициализирован, аккуратно молчим
    pass
