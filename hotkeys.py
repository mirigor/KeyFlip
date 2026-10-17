"""Общие правила разбора сочетаний, независимые от Windows и интерфейса."""

VK_MAP = {
    **{f"F{number}": 0x6F + number for number in range(1, 25)},
    "ESC": 0x1B, "TAB": 0x09, "ENTER": 0x0D, "RETURN": 0x0D, "SPACE": 0x20,
    "BACKSPACE": 0x08, "INSERT": 0x2D, "DELETE": 0x2E,
    "HOME": 0x24, "END": 0x23, "PGUP": 0x21, "PGDN": 0x22,
    "PAGEUP": 0x21, "PAGEDOWN": 0x22,
    "LEFT": 0x25, "UP": 0x26, "RIGHT": 0x27, "DOWN": 0x28,
}
MODIFIER_ALIASES = {
    "CTRL": "CTRL", "CONTROL": "CTRL", "ALT": "ALT", "MENU": "ALT",
    "SHIFT": "SHIFT", "WIN": "WIN", "WINDOWS": "WIN",
}
KEY_ALIASES = {"RETURN": "ENTER", "PAGEUP": "PGUP", "PAGEDOWN": "PGDN"}


def normalize_modifiers(values):
    """Нормализует модификаторы, сохраняя порядок и убирая повторы."""
    if not isinstance(values, (list, tuple, set)):
        return []
    result = []
    for value in values:
        if isinstance(value, str):
            modifier = MODIFIER_ALIASES.get(value.strip().upper())
            if modifier is not None and modifier not in result:
                result.append(modifier)
    return result


def key_name_to_vk(name):
    """Возвращает VK или отклоняет неизвестную клавишу без скрытой подмены."""
    if not isinstance(name, str):
        raise ValueError("Название клавиши должно быть строкой")
    key = name.strip().upper()
    if key in VK_MAP:
        return VK_MAP[key]
    if len(key) == 1 and key in "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789":
        return ord(key)
    if key.startswith("VK_"):
        try:
            value = int(key[3:], 16)
        except ValueError:
            pass
        else:
            if 1 <= value <= 254:
                return value
    raise ValueError(f"Неизвестная клавиша: {name!r}")


def normalize_key(name):
    key = name.strip().upper() if isinstance(name, str) else name
    vk = key_name_to_vk(key)
    key = KEY_ALIASES.get(key, key)
    if key.startswith("VK_"):
        for label, code in VK_MAP.items():
            if code == vk:
                return KEY_ALIASES.get(label, label)
        if chr(vk) in "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789":
            return chr(vk)
        return f"VK_{vk:02X}"
    return key


def hotkeys_equal(a_mods, a_key, b_mods, b_key):
    try:
        return (set(normalize_modifiers(a_mods)) == set(normalize_modifiers(b_mods))
                and key_name_to_vk(a_key) == key_name_to_vk(b_key))
    except ValueError:
        return False
