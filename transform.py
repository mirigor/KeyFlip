"""
Модуль, содержащий таблицы соответствия раскладок и функцию преобразования текста.

Здесь — только чистая бизнес-логика (без побочных эффектов).
"""

from logging_setup import logger

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

# Обратная мапа RU -> EN
RU_TO_EN: dict[str, str] = {v: k for k, v in EN_TO_RU.items()}
# Убедимся, что 'ё' и 'Ё' также присутствуют (особенности маппинга)
RU_TO_EN['ё'] = '`'
RU_TO_EN['Ё'] = '~'


# ---------------- Вспомогательные функции ----------------
def _select_mappings_by_hkl(hkl: int) -> tuple[dict[str, str], dict[str, str], str]:
    """
    Выбрать маппинги (mapping, reverse, prefer) по HKL.
    Возвращает (mapping, reverse, prefer_name).
    """
    try:
        lang = hkl & 0xFFFF
    except Exception:  # noqa
        logger.exception("_select_mappings_by_hkl: не удалось вычислить lang из hkl")
        lang = 0

    if lang == 0x0419:
        # Если текущая раскладка русская — преобразуем RU->EN
        return RU_TO_EN, EN_TO_RU, "RU->EN"
    # Иначе — EN->RU
    return EN_TO_RU, RU_TO_EN, "EN->RU"


# ---------------- Логика преобразования ----------------
def transform_text_by_keyboard_layout_based_on_hkl(s: str, hkl: int) -> str:
    """
    Преобразовать строку `s` в соответствии с текущей раскладкой, определяемой по HKL.
    Если LANGID == 0x0419 (русский) — выполняем RU->EN, иначе EN->RU.

    Аргументы:
        s: исходная строка
        hkl: значение HKL (обычно получаемое через GetKeyboardLayout)

    Возвращает:
        преобразованную строку
    """
    mapping, reverse, prefer = _select_mappings_by_hkl(hkl)

    out_chars: list[str] = []
    for ch in s:
        # сначала пробуем основную мапу
        if ch in mapping:
            out_chars.append(mapping[ch])
            continue
        # затем пробуем обратную (поддержка случаев двойной путаницы раскладки)
        if ch in reverse:
            out_chars.append(reverse[ch])
            continue
        # иначе — оставляем символ как есть
        out_chars.append(ch)

    logger.debug("transform: used %s mapping", prefer)
    return "".join(out_chars)

def change_case_by_logic(s: str) -> str:
    """
    Изменить регистр строки по правилам:
      - если в строке нет букв — вернуть исходную строку
      - если все буквы маленькие -> сделать ВСЕ БОЛЬШИМИ
      - если все буквы большие -> сделать все маленькими
      - если часть букв большая/часть маленькая -> сделать все маленькими

    Возвращает новую строку (или исходную, если изменений не требуется).
    """
    if not s:
        return s

    letters = [ch for ch in s if ch.isalpha()]
    if not letters:
        return s

    if all(ch.islower() for ch in letters):
        return s.upper()
    return s.lower()
