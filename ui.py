"""Компактная панель и меню Qt в стиле NightLightTray."""

from PySide6.QtCore import QEvent, QPoint, QRectF, Qt, QTimer, Signal
from PySide6.QtGui import QActionGroup, QColor, QCursor, QIcon, QPainter, QPalette, QPen
from PySide6.QtWidgets import (
    QApplication, QFrame, QHBoxLayout, QLabel,
    QMenu, QPushButton, QSystemTrayIcon, QVBoxLayout, QWidget,
)
from PySide6.QtSvgWidgets import QSvgWidget
import config
import tray as actions
from winapi import capture_hotkey_and_apply_via_thread, exit_event


HOTKEYS = {
    "translate": ("Смена раскладки", config.read_translate_hotkey,
                  [((), "F4"), (("CTRL", "ALT"), "T"), (("CTRL", "SHIFT"), "Y")]),
    "case": ("Смена регистра", config.read_case_hotkey, [(("CTRL", "SHIFT"), "U")]),
    "exit": ("Выход", config.read_exit_hotkey,
             [((), "F10"), (("CTRL",), "Q"), (("CTRL", "ALT"), "X")]),
}
MODIFIER_LABELS = {"CTRL": "Ctrl", "ALT": "Alt", "SHIFT": "Shift", "WIN": "Win"}


def hotkey_text(value):
    modifiers = [MODIFIER_LABELS.get(mod.upper(), mod) for mod in value.get("modifiers", [])]
    key = (value.get("key") or "").upper()
    return " + ".join(modifiers + [key]) if key else "Не задано"


def popup_position(tray_rect, available, size, anchor):
    """Располагает панель рядом с треем на любом мониторе и стороне панели задач."""
    width, height = size.width(), size.height()
    x, y = anchor.x() - width // 2, anchor.y() - height - 8
    if tray_rect.isValid() and anchor.x() < available.left():
        x, y = available.left() + 8, anchor.y() - height // 2
    elif tray_rect.isValid() and anchor.x() > available.right():
        x, y = available.right() - width - 7, anchor.y() - height // 2
    elif anchor.y() < available.center().y():
        y = (tray_rect.bottom() if tray_rect.isValid() else anchor.y()) + 8
    inset_x = min(8, max(0, (available.width() - width) // 2))
    inset_y = min(8, max(0, (available.height() - height) // 2))
    return QPoint(
        max(available.left() + inset_x, min(x, available.right() + 1 - width - inset_x)),
        max(available.top() + inset_y, min(y, available.bottom() + 1 - height - inset_y)),
    )


def theme_colors(dark=None):
    app = QApplication.instance()
    if dark is None:
        scheme = app.styleHints().colorScheme()
        dark = scheme == Qt.ColorScheme.Dark or (
            scheme == Qt.ColorScheme.Unknown
            and app.palette().color(QPalette.ColorRole.Window).lightness() < 128
        )
    if dark:
        return dict(bg="#202020", text="#f4f4f4", muted="#a8a8a8", border="#404040",
                    card="#292929", hover="#353535", accent="#8bc5ff",
                    accent_text="#142333")
    return dict(bg="#f7f7f7", text="#202020", muted="#656565", border="#d6d6d6",
                card="#ffffff", hover="#eaeaea", accent="#0067c0",
                accent_text="#ffffff")


def theme_stylesheet(colors):
    return """
        QWidget { color: %(text)s; font-family: 'Segoe UI'; font-size: 12px; }
        QFrame#panel { background: %(bg)s; border: 1px solid %(border)s; border-radius: 12px; }
        QLabel { background: transparent; border: none; }
        QLabel#title { font-size: 16px; font-weight: 600; }
        QLabel#subtitle, QLabel#hint { color: %(muted)s; font-size: 11px; }
        QLabel#section { color: %(muted)s; font-size: 11px; font-weight: 600; }
        QLabel#status { color: %(accent)s; font-size: 11px; }
        QFrame#divider { background: %(border)s; border: none; min-height: 1px; max-height: 1px; }
        QPushButton { background: %(card)s; border: 1px solid %(border)s; border-radius: 7px; padding: 6px 10px; }
        QPushButton:hover { background: %(hover)s; }
        QPushButton:focus { border-color: %(accent)s; }
        QPushButton#hotkeyRow { background: %(card)s; text-align: left; padding: 8px 10px; }
        QPushButton#hotkeyRow:hover { background: %(hover)s; border-color: %(accent)s; }
        QMenu { background: %(bg)s; color: %(text)s; font-size: 12px; border: 1px solid %(border)s; border-radius: 8px; padding: 4px; }
        QMenu::item { padding: 5px 20px 5px 10px; border-radius: 4px; }
        QMenu::item:selected { background: %(hover)s; }
        QMenu::item:disabled { color: %(muted)s; }
        QMenu::separator { height: 1px; background: %(border)s; margin: 3px 4px; }
        QToolTip { background: %(card)s; color: %(text)s; border: 1px solid %(border)s; padding: 4px; }
    """ % colors


class ToggleSwitch(QPushButton):
    """Переключатель с отрисовкой, не зависящей от системного стиля кнопок."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setCheckable(True)
        self.setFocusPolicy(Qt.FocusPolicy.TabFocus)
        self.setFixedSize(42, 24)
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.setAccessibleName("Включить обработку текста")

    def paintEvent(self, event):
        colors = theme_colors()
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        painter.setPen(QPen(QColor(colors["accent"] if self.isChecked() else colors["muted"]), 1))
        painter.setBrush(QColor(colors["accent"] if self.isChecked() else colors["card"]))
        painter.drawRoundedRect(QRectF(1, 2, 40, 20), 10, 10)
        painter.setPen(Qt.PenStyle.NoPen)
        painter.setBrush(QColor(colors["accent_text"] if self.isChecked() else colors["muted"]))
        painter.drawEllipse(QRectF(24 if self.isChecked() else 5, 6, 12, 12))


class KeyFlipPopup(QWidget):
    hotkey_requested = Signal(str)
    toggle_requested = Signal()
    autorun_requested = Signal()
    logging_requested = Signal()

    def __init__(self):
        super().__init__(None, Qt.WindowType.Tool | Qt.WindowType.FramelessWindowHint
                         | Qt.WindowType.WindowStaysOnTopHint)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self.setFixedWidth(300)
        self.setWindowTitle("KeyFlip")
        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        self.panel = QFrame(self)
        self.panel.setObjectName("panel")
        outer.addWidget(self.panel)
        layout = QVBoxLayout(self.panel)
        layout.setContentsMargins(14, 14, 14, 12)
        layout.setSpacing(10)

        header = QHBoxLayout()
        self.brand = QSvgWidget(config.ICON_ON_SVG)
        self.brand.setFixedSize(36, 36)
        self.brand.setAccessibleName("Значок KeyFlip")
        header.addWidget(self.brand)
        labels = QVBoxLayout()
        labels.setSpacing(2)
        title = QLabel("KeyFlip")
        title.setObjectName("title")
        self.status = QLabel()
        self.status.setObjectName("status")
        labels.addWidget(title)
        labels.addWidget(self.status)
        header.addLayout(labels)
        header.addStretch()
        self.toggle = ToggleSwitch()
        self.toggle.clicked.connect(lambda: self.toggle_requested.emit())
        header.addWidget(self.toggle)
        layout.addLayout(header)

        hint = QLabel("Выделите текст и нажмите горячую клавишу")
        hint.setObjectName("hint")
        hint.setWordWrap(True)
        layout.addWidget(hint)
        divider = QFrame()
        divider.setObjectName("divider")
        layout.addWidget(divider)
        section = QLabel("ГОРЯЧИЕ КЛАВИШИ")
        section.setObjectName("section")
        layout.addWidget(section)

        self.shortcuts = {}
        for target, (label, _read, _presets) in HOTKEYS.items():
            button = QPushButton()
            button.setObjectName("hotkeyRow")
            button.setToolTip("Нажмите, чтобы изменить сочетание")
            button.setAccessibleName("Изменить комбинацию: " + label)
            button.setCursor(Qt.CursorShape.PointingHandCursor)
            button.setFocusPolicy(Qt.FocusPolicy.TabFocus)
            button.clicked.connect(lambda _checked=False, t=target: self.hotkey_requested.emit(t))
            self.shortcuts[target] = button
            layout.addWidget(button)

        divider = QFrame()
        divider.setObjectName("divider")
        layout.addWidget(divider)
        section = QLabel("ПАРАМЕТРЫ")
        section.setObjectName("section")
        layout.addWidget(section)

        self.autorun_toggle = ToggleSwitch()
        self.autorun_toggle.setAccessibleName("Автозапуск с Windows")
        self.autorun_toggle.clicked.connect(lambda: self.autorun_requested.emit())
        self.logging_toggle = ToggleSwitch()
        self.logging_toggle.setAccessibleName("Логирование в файл")
        self.logging_toggle.clicked.connect(lambda: self.logging_requested.emit())
        for label, switch in (("Автозапуск с Windows", self.autorun_toggle),
                              ("Логирование в файл", self.logging_toggle)):
            row = QHBoxLayout()
            row.addWidget(QLabel(label))
            row.addStretch()
            row.addWidget(switch)
            layout.addLayout(row)

        self._dismiss = QTimer(self)
        self._dismiss.setSingleShot(True)
        self._dismiss.setInterval(120)
        self._dismiss.timeout.connect(self._hide_if_inactive)
        self.refresh()

    def refresh(self):
        enabled = config.is_enabled()
        self.brand.load(config.ICON_ON_SVG if enabled else config.ICON_OFF_SVG)
        self.toggle.setChecked(enabled)
        self.autorun_toggle.setChecked(config.read_autorun_flag())
        self.logging_toggle.setChecked(config.read_file_logging_flag())
        self.status.setText("Включён" if enabled else "На паузе")
        for target, (label, read, _presets) in HOTKEYS.items():
            self.shortcuts[target].setText(label + "   ·   " + hotkey_text(read()) + "  ›")
        self.setStyleSheet(theme_stylesheet(theme_colors()))

    def open_near(self, tray_rect):
        self._dismiss.stop()
        self.refresh()
        self.adjustSize()
        anchor = tray_rect.center() if tray_rect.isValid() else QCursor.pos()
        screen = QApplication.screenAt(anchor) or self.screen()
        self.move(popup_position(tray_rect, screen.availableGeometry(), self.size(), anchor))
        self.show()
        self.raise_()
        self.activateWindow()

    def event(self, event):
        if event.type() == QEvent.Type.WindowDeactivate and self.isVisible():
            self._dismiss.start()
        return super().event(event)

    def _hide_if_inactive(self):
        if not self.isActiveWindow() and QApplication.activePopupWidget() is None:
            self.hide()

    def hideEvent(self, event):
        self._dismiss.stop()
        focused = self.focusWidget()
        if focused is not None:
            focused.clearFocus()
        super().hideEvent(event)

    def keyPressEvent(self, event):
        if event.key() == Qt.Key.Key_Escape:
            self.hide()
        else:
            super().keyPressEvent(event)


class TrayController:
    """Все виджеты живут в главном потоке, WinAPI передаёт выход через Event."""

    def __init__(self):
        self.popup = KeyFlipPopup()
        self.tray = QSystemTrayIcon()
        self.menu = QMenu()
        self._icon_state = None
        self.menu.addAction("Открыть", self.show_popup)
        self.enabled_action = self.menu.addAction("", self.toggle_enabled)
        self.menu.addSeparator()
        self.settings_menu = self.menu.addMenu("Настройки")
        self._hotkey_menus = {}
        self._preset_actions = {}
        for target, (label, _read, presets) in HOTKEYS.items():
            submenu = self.settings_menu.addMenu(label)
            group = QActionGroup(submenu)
            group.setExclusive(True)
            self._hotkey_menus[target] = submenu
            self._preset_actions[target] = []
            for index, (mods, key) in enumerate(presets):
                text = hotkey_text({"modifiers": mods, "key": key})
                action = submenu.addAction(text + (" (по умолчанию)" if index == 0 else ""))
                action.setCheckable(True)
                group.addAction(action)
                action.triggered.connect(lambda _checked=False, m=mods, k=key, t=target: self.set_preset(t, m, k))
                self._preset_actions[target].append((action, mods, key))
            submenu.addSeparator()
            submenu.addAction("Своя комбинация…", lambda _checked=False, t=target: self.edit_hotkey(t))
        self.menu.addSeparator()
        self.menu.addAction("Выход", actions.on_exit)
        self.menu.aboutToShow.connect(self.refresh)
        self.settings_menu.aboutToShow.connect(self.refresh)
        self.tray.setContextMenu(self.menu)
        self.tray.activated.connect(self._activated)
        self.popup.toggle_requested.connect(self.toggle_enabled)
        self.popup.autorun_requested.connect(self.toggle_autorun)
        self.popup.logging_requested.connect(self.toggle_logging)
        self.popup.hotkey_requested.connect(self.edit_hotkey)
        self._exit_timer = QTimer(self.tray)
        self._exit_timer.setInterval(100)
        self._exit_timer.timeout.connect(self._check_exit)
        QApplication.instance().styleHints().colorSchemeChanged.connect(self.refresh)
        self.refresh()

    def start(self):
        self.tray.show()
        self._exit_timer.start()

    def refresh(self, *_args):
        enabled = config.is_enabled()
        if self._icon_state != enabled:
            icon = QIcon(config.ICON_ON if enabled else config.ICON_OFF)
            self.tray.setIcon(icon)
            self._icon_state = enabled
        self.enabled_action.setText("Приостановить" if enabled else "Включить")
        self.tray.setToolTip("KeyFlip — " + ("включён" if enabled else "на паузе")
                             + "\nСмена раскладки: " + hotkey_text(config.read_translate_hotkey()))
        for target, (label, read, _presets) in HOTKEYS.items():
            value = read()
            self._hotkey_menus[target].setTitle(label + " · " + hotkey_text(value))
            for action, mods, key in self._preset_actions[target]:
                action.setChecked(actions.hotkey_lists_equal(list(mods), key, value.get("modifiers"), value.get("key", "")))
        self.menu.setStyleSheet(theme_stylesheet(theme_colors()))
        self.popup.refresh()

    def toggle_enabled(self):
        actions.toggle_enabled()
        self.refresh()

    def toggle_autorun(self):
        actions.toggle_autorun()
        self.refresh()

    def toggle_logging(self):
        actions.toggle_file_logging()
        self.refresh()

    def set_preset(self, target, mods, key):
        actions._set_hotkey_and_apply(list(mods), key, target)
        self.refresh()

    def edit_hotkey(self, target):
        self.popup.hide()
        capture_hotkey_and_apply_via_thread(target)

    def _activated(self, reason):
        if reason == QSystemTrayIcon.ActivationReason.Trigger:
            if self.popup.isVisible():
                self.popup.hide()
            else:
                self.show_popup()

    def show_popup(self):
        if not exit_event.is_set():
            self.popup.open_near(self.tray.geometry())

    def _check_exit(self):
        if exit_event.is_set():
            QApplication.instance().quit()

    def close(self):
        self._exit_timer.stop()
        self.menu.close()
        self.popup.hide()
        self.tray.hide()
