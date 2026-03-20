from __future__ import annotations

import json
import threading
import time
from datetime import datetime
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import pyautogui
import pygetwindow as gw
from pynput import keyboard, mouse
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QKeySequence, QShortcut
from PyQt6.QtWidgets import (
    QApplication,
    QButtonGroup,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QAbstractItemView,
    QFileDialog,
    QFormLayout,
    QGridLayout,
    QHBoxLayout,
    QInputDialog,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QRadioButton,
    QSpinBox,
    QStackedWidget,
    QStatusBar,
    QVBoxLayout,
    QWidget,
)

MACROS_FILE = Path("macros.json")
MOUSE_MOVE_DURATION_SEC = 0.2


@dataclass
class MacroEvent:
    t: float
    etype: str
    data: dict[str, Any]

    def to_dict(self) -> dict[str, Any]:
        return {"t": self.t, "etype": self.etype, "data": self.data}

    @staticmethod
    def from_dict(raw: dict[str, Any]) -> "MacroEvent":
        return MacroEvent(
            t=float(raw["t"]),
            etype=str(raw["etype"]),
            data=dict(raw["data"]),
        )


@dataclass
class MacroItem:
    name: str
    events: list[MacroEvent] = field(default_factory=list)
    play_mode: str = "once"
    repeat_count: int = 1
    repeat_until_at: str = ""
    repeat_interval_ms: int = 0
    start_delay: int = 1
    speed_percent: int = 100

    def to_dict(self) -> dict[str, Any]:
        return {
            "name": self.name,
            "play_mode": self.play_mode,
            "repeat_count": self.repeat_count,
            "repeat_until_at": self.repeat_until_at,
            "repeat_interval_ms": self.repeat_interval_ms,
            "start_delay": self.start_delay,
            "speed_percent": self.speed_percent,
            "events": [ev.to_dict() for ev in self.events],
        }

    @staticmethod
    def from_dict(raw: dict[str, Any]) -> "MacroItem":
        old_loops = int(raw.get("loops", 1))
        play_mode = str(raw.get("play_mode", "repeat_count" if old_loops > 1 else "once"))
        return MacroItem(
            name=str(raw["name"]),
            play_mode=play_mode,
            repeat_count=int(raw.get("repeat_count", old_loops)),
            repeat_until_at=str(raw.get("repeat_until_at", "")),
            repeat_interval_ms=int(raw.get("repeat_interval_ms", 0)),
            start_delay=int(raw.get("start_delay", 1)),
            speed_percent=int(raw.get("speed_percent", 100)),
            events=[MacroEvent.from_dict(item) for item in raw.get("events", [])],
        )


class MacroStore:
    def __init__(self, path: Path) -> None:
        self.path = path
        self.items: list[MacroItem] = []

    def load(self) -> None:
        if not self.path.exists():
            self.items = []
            return
        raw = json.loads(self.path.read_text(encoding="utf-8"))
        self.items = [MacroItem.from_dict(item) for item in raw]

    def save(self) -> None:
        payload = [item.to_dict() for item in self.items]
        self.path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")

    def unique_name(self, base: str = "新巨集") -> str:
        used = {item.name for item in self.items}
        if base not in used:
            return base
        idx = 2
        while True:
            candidate = f"{base} {idx}"
            if candidate not in used:
                return candidate
            idx += 1


class MacroEngine:
    def __init__(self) -> None:
        self.events: list[MacroEvent] = []
        self.is_recording = False
        self.is_playing = False
        self._record_start = 0.0
        self._record_lock = threading.Lock()
        self._play_stop_event = threading.Event()

        self._key_listener: keyboard.Listener | None = None
        self._mouse_listener: mouse.Listener | None = None

        self._kb_controller = keyboard.Controller()
        pyautogui.PAUSE = 0
        pyautogui.FAILSAFE = True

    def clear(self) -> None:
        with self._record_lock:
            self.events.clear()

    def set_events(self, events: list[MacroEvent]) -> None:
        with self._record_lock:
            self.events = [MacroEvent(ev.t, ev.etype, dict(ev.data)) for ev in events]

    def get_events_copy(self) -> list[MacroEvent]:
        with self._record_lock:
            return [MacroEvent(ev.t, ev.etype, dict(ev.data)) for ev in self.events]

    def start_recording(self) -> None:
        if self.is_recording:
            return
        self.is_recording = True
        self._record_start = time.perf_counter()
        self.events = []

        self._key_listener = keyboard.Listener(
            on_press=self._on_key_press,
            on_release=self._on_key_release,
        )
        self._mouse_listener = mouse.Listener(
            on_click=self._on_mouse_click,
            on_scroll=self._on_mouse_scroll,
            on_move=None,
        )
        self._key_listener.start()
        self._mouse_listener.start()

    def stop_recording(self) -> None:
        if not self.is_recording:
            return
        self.is_recording = False
        if self._key_listener:
            self._key_listener.stop()
            self._key_listener = None
        if self._mouse_listener:
            self._mouse_listener.stop()
            self._mouse_listener = None

    def _elapsed(self) -> float:
        return time.perf_counter() - self._record_start

    @staticmethod
    def _key_to_payload(k: keyboard.Key | keyboard.KeyCode) -> dict[str, Any]:
        if isinstance(k, keyboard.KeyCode):
            return {"kind": "code", "char": k.char, "vk": k.vk}
        return {"kind": "key", "name": k.name}

    @staticmethod
    def _payload_to_key(payload: dict[str, Any]) -> keyboard.Key | keyboard.KeyCode:
        if payload["kind"] == "key":
            return keyboard.Key[payload["name"]]
        return keyboard.KeyCode.from_vk(payload["vk"])

    def _on_key_press(self, k: keyboard.Key | keyboard.KeyCode) -> None:
        if not self.is_recording:
            return
        with self._record_lock:
            self.events.append(MacroEvent(self._elapsed(), "key_press", self._key_to_payload(k)))

    def _on_key_release(self, k: keyboard.Key | keyboard.KeyCode) -> None:
        if not self.is_recording:
            return
        with self._record_lock:
            self.events.append(MacroEvent(self._elapsed(), "key_release", self._key_to_payload(k)))

    def _on_mouse_click(self, x: int, y: int, button: mouse.Button, pressed: bool) -> None:
        if not self.is_recording:
            return
        with self._record_lock:
            self.events.append(
                MacroEvent(
                    self._elapsed(),
                    "mouse_click",
                    {"x": x, "y": y, "button": button.name, "pressed": pressed},
                )
            )

    def _on_mouse_scroll(self, x: int, y: int, dx: int, dy: int) -> None:
        if not self.is_recording:
            return
        with self._record_lock:
            self.events.append(
                MacroEvent(
                    self._elapsed(),
                    "mouse_scroll",
                    {"x": x, "y": y, "dx": dx, "dy": dy},
                )
            )

    def play_async(
        self,
        play_mode: str,
        repeat_count: int,
        repeat_until_at: str,
        repeat_interval_ms: int,
        start_delay: int,
        speed_percent: int,
    ) -> threading.Thread:
        if self.is_playing:
            raise RuntimeError("目前正在播放中")
        if not self.events:
            raise RuntimeError("尚未有任何錄製事件")

        self.is_playing = True
        self._play_stop_event.clear()

        def _runner() -> None:
            try:
                if start_delay > 0:
                    for _ in range(start_delay * 10):
                        if self._play_stop_event.is_set():
                            return
                        time.sleep(0.1)

                speed = max(1, speed_percent) / 100.0
                interval_sec = max(0, repeat_interval_ms) / 1000.0
                if play_mode == "once":
                    self._play_once(speed)
                elif play_mode == "repeat_count":
                    total = max(1, repeat_count)
                    for i in range(total):
                        if self._play_stop_event.is_set():
                            return
                        self._play_once(speed)
                        if i < total - 1 and not self._wait_interval(interval_sec):
                            return
                elif play_mode == "repeat_until":
                    try:
                        until_at = datetime.strptime(repeat_until_at.strip(), "%Y-%m-%d %H:%M:%S")
                    except ValueError as exc:
                        raise RuntimeError("重複至時間格式錯誤，請用 YYYY-MM-DD HH:MM:SS") from exc

                    while datetime.now() <= until_at:
                        if self._play_stop_event.is_set():
                            return
                        self._play_once(speed)
                        if datetime.now() <= until_at and not self._wait_interval(interval_sec):
                            return
                elif play_mode == "until_stopped":
                    while not self._play_stop_event.is_set():
                        self._play_once(speed)
                        if not self._wait_interval(interval_sec):
                            return
                else:
                    self._play_once(speed)
            finally:
                self.is_playing = False

        t = threading.Thread(target=_runner, daemon=True)
        t.start()
        return t

    def stop_play(self) -> None:
        self._play_stop_event.set()

    def _wait_interval(self, seconds: float) -> bool:
        if seconds <= 0:
            return True
        ticks = int(seconds / 0.1)
        remain = seconds - ticks * 0.1
        for _ in range(ticks):
            if self._play_stop_event.is_set():
                return False
            time.sleep(0.1)
        if remain > 0:
            if self._play_stop_event.is_set():
                return False
            time.sleep(remain)
        return not self._play_stop_event.is_set()

    def _play_once(self, speed: float) -> None:
        prev_t = 0.0
        for ev in self.events:
            if self._play_stop_event.is_set():
                return
            delta = max(0.0, ev.t - prev_t) / speed
            prev_t = ev.t
            if delta > 0:
                time.sleep(delta)

            if ev.etype == "key_press":
                self._kb_controller.press(self._payload_to_key(ev.data))
            elif ev.etype == "key_release":
                self._kb_controller.release(self._payload_to_key(ev.data))
            elif ev.etype == "mouse_click":
                button = self._map_mouse_button(ev.data["button"])
                x = ev.data.get("x")
                y = ev.data.get("y")
                has_xy = x is not None and y is not None
                if ev.data["pressed"]:
                    if has_xy:
                        pyautogui.mouseDown(x=x, y=y, button=button)
                    else:
                        pyautogui.mouseDown(button=button)
                else:
                    if has_xy:
                        pyautogui.mouseUp(x=x, y=y, button=button)
                    else:
                        pyautogui.mouseUp(button=button)
            elif ev.etype == "mouse_scroll":
                pyautogui.moveTo(ev.data["x"], ev.data["y"], duration=MOUSE_MOVE_DURATION_SEC)
                pyautogui.scroll(ev.data["dy"] * 120)
            elif ev.etype == "mouse_move":
                pyautogui.moveTo(ev.data["x"], ev.data["y"], duration=MOUSE_MOVE_DURATION_SEC)
            elif ev.etype == "text_input":
                pyautogui.typewrite(str(ev.data.get("text", "")), interval=0.02)
            elif ev.etype == "hotkey":
                keys = [str(k) for k in ev.data.get("keys", []) if str(k).strip()]
                if keys:
                    pyautogui.hotkey(*keys)
            elif ev.etype == "focus_window":
                title = str(ev.data.get("title", "")).strip()
                if not title:
                    continue
                wins = gw.getWindowsWithTitle(title)
                if not wins:
                    continue
                target = wins[0]
                if getattr(target, "isMinimized", False):
                    target.restore()
                target.activate()
            elif ev.etype == "wait":
                wait_seconds = max(0.0, float(ev.data.get("seconds", 0.0)))
                ticks = max(1, int(round(wait_seconds * 10)))
                for _ in range(ticks):
                    if self._play_stop_event.is_set():
                        return
                    time.sleep(0.1)

    @staticmethod
    def _map_mouse_button(raw: str) -> str:
        if raw in {"left", "right", "middle"}:
            return raw
        return "left"


class MouseMoveDialog(QDialog):
    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("新增滑鼠移動")
        self.resize(420, 160)
        layout = QVBoxLayout(self)
        grid = QGridLayout()

        self.x_box = QSpinBox()
        self.x_box.setRange(0, 10000)
        self.y_box = QSpinBox()
        self.y_box.setRange(0, 10000)
        self.live_label = QLabel("目前滑鼠座標：")
        self.btn_use_current = QPushButton("使用目前座標")

        grid.addWidget(QLabel("X"), 0, 0)
        grid.addWidget(self.x_box, 0, 1)
        grid.addWidget(QLabel("Y"), 1, 0)
        grid.addWidget(self.y_box, 1, 1)
        grid.addWidget(self.live_label, 2, 0, 1, 2)
        grid.addWidget(self.btn_use_current, 3, 0, 1, 2)
        layout.addLayout(grid)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

        self.btn_use_current.clicked.connect(self.use_current_position)
        self._timer_id = self.startTimer(500)
        self._refresh_cursor_label()

    def timerEvent(self, event) -> None:  # type: ignore[override]
        if event.timerId() == self._timer_id:
            self._refresh_cursor_label()

    def closeEvent(self, event) -> None:  # type: ignore[override]
        if self._timer_id:
            self.killTimer(self._timer_id)
            self._timer_id = 0
        super().closeEvent(event)

    def keyPressEvent(self, event) -> None:  # type: ignore[override]
        if event.key() in (Qt.Key.Key_Return, Qt.Key.Key_Enter):
            # Enter 直接把「目前滑鼠座標」帶入，不要直接觸發 Ok 並關閉視窗
            self.use_current_position()
            event.accept()
            return
        super().keyPressEvent(event)

    def _refresh_cursor_label(self) -> None:
        x, y = pyautogui.position()
        self.live_label.setText(f"目前滑鼠座標：({x}, {y})")

    def use_current_position(self) -> None:
        x, y = pyautogui.position()
        self.x_box.setValue(x)
        self.y_box.setValue(y)


class MouseClickDialog(QDialog):
    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("新增滑鼠點擊")
        self.resize(280, 120)
        layout = QVBoxLayout(self)
        self.left_radio = QRadioButton("左鍵")
        self.right_radio = QRadioButton("右鍵")
        self.left_radio.setChecked(True)
        self._btn_group = QButtonGroup(self)
        self._btn_group.addButton(self.left_radio)
        self._btn_group.addButton(self.right_radio)
        btn_row = QHBoxLayout()
        btn_row.addWidget(QLabel("按鍵"))
        btn_row.addWidget(self.left_radio)
        btn_row.addWidget(self.right_radio)
        btn_row.addStretch()
        layout.addLayout(btn_row)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def selected_button(self) -> str:
        return "right" if self.right_radio.isChecked() else "left"


class HotkeyCaptureDialog(QDialog):
    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("偵測組合鍵")
        self.resize(420, 140)
        self.keys: list[str] = []

        layout = QVBoxLayout(self)
        layout.addWidget(QLabel("請直接按下組合鍵，按 Enter 確認。"))
        self.capture_edit = QLineEdit()
        self.capture_edit.setReadOnly(True)
        self.capture_edit.setPlaceholderText("等待按鍵...")
        layout.addWidget(self.capture_edit)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)
        self.capture_edit.setFocus()

    def keyPressEvent(self, event) -> None:  # type: ignore[override]
        key = event.key()
        if key in (Qt.Key.Key_Return, Qt.Key.Key_Enter):
            if self.keys:
                self.accept()
            return

        keys: list[str] = []
        mods = event.modifiers()
        if mods & Qt.KeyboardModifier.ControlModifier:
            keys.append("ctrl")
        if mods & Qt.KeyboardModifier.AltModifier:
            keys.append("alt")
        if mods & Qt.KeyboardModifier.ShiftModifier:
            keys.append("shift")
        if mods & Qt.KeyboardModifier.MetaModifier:
            keys.append("win")

        text = event.text().lower().strip()
        if key not in (
            Qt.Key.Key_Control,
            Qt.Key.Key_Shift,
            Qt.Key.Key_Alt,
            Qt.Key.Key_Meta,
        ):
            if text and text.isprintable():
                keys.append(text)
            else:
                key_name = self._qt_key_to_name(key)
                if key_name:
                    keys.append(key_name)

        self.keys = keys
        self.capture_edit.setText("+".join(self.keys))

    @staticmethod
    def _qt_key_to_name(key: int) -> str:
        mapping = {
            Qt.Key.Key_F1: "f1",
            Qt.Key.Key_F2: "f2",
            Qt.Key.Key_F3: "f3",
            Qt.Key.Key_F4: "f4",
            Qt.Key.Key_F5: "f5",
            Qt.Key.Key_F6: "f6",
            Qt.Key.Key_F7: "f7",
            Qt.Key.Key_F8: "f8",
            Qt.Key.Key_F9: "f9",
            Qt.Key.Key_F10: "f10",
            Qt.Key.Key_F11: "f11",
            Qt.Key.Key_F12: "f12",
            Qt.Key.Key_Tab: "tab",
            Qt.Key.Key_Escape: "esc",
            Qt.Key.Key_Delete: "delete",
            Qt.Key.Key_Backspace: "backspace",
            Qt.Key.Key_Space: "space",
            Qt.Key.Key_Up: "up",
            Qt.Key.Key_Down: "down",
            Qt.Key.Key_Left: "left",
            Qt.Key.Key_Right: "right",
            Qt.Key.Key_Home: "home",
            Qt.Key.Key_End: "end",
            Qt.Key.Key_PageUp: "pageup",
            Qt.Key.Key_PageDown: "pagedown",
        }
        return mapping.get(key, "")


class WaitTimeDialog(QDialog):
    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("新增等待")
        self.resize(520, 140)
        layout = QFormLayout(self)
        self.hour_box = QSpinBox()
        self.hour_box.setRange(0, 999999)
        self.minute_box = QSpinBox()
        self.minute_box.setRange(0, 60)
        self.second_box = QSpinBox()
        self.second_box.setRange(0, 60)
        self.second_box.setValue(0)
        self.millisecond_box = QSpinBox()
        self.millisecond_box.setRange(0, 999)
        self.millisecond_box.setValue(500)
        for box in (self.hour_box, self.minute_box, self.second_box):
            box.setFixedWidth(72)
        self.millisecond_box.setFixedWidth(92)

        fields_row = QHBoxLayout()
        fields_row.addWidget(self.hour_box)
        fields_row.addWidget(QLabel("時"))
        fields_row.addWidget(self.minute_box)
        fields_row.addWidget(QLabel("分"))
        fields_row.addWidget(self.second_box)
        fields_row.addWidget(QLabel("秒"))
        fields_row.addWidget(self.millisecond_box)
        fields_row.addWidget(QLabel("毫秒"))
        fields_row.addStretch()
        layout.addRow("等待時間", fields_row)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addRow(buttons)

    def total_seconds(self) -> float:
        return (
            self.hour_box.value() * 3600
            + self.minute_box.value() * 60
            + self.second_box.value()
            + self.millisecond_box.value() / 1000.0
        )


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("Corvus 按鍵精靈")
        self.resize(860, 560)

        self.store = MacroStore(MACROS_FILE)
        self.engine = MacroEngine()
        self.current_index: int | None = None

        self._build_ui()
        self._load_store()

    def _build_ui(self) -> None:
        self.stack = QStackedWidget()
        self.list_page = QWidget()
        self.edit_page = QWidget()
        self.stack.addWidget(self.list_page)
        self.stack.addWidget(self.edit_page)
        self.setCentralWidget(self.stack)

        self._build_list_page()
        self._build_edit_page()

        self.setStatusBar(QStatusBar())
        self.statusBar().showMessage("就緒")
        self._set_editor_buttons()

    def _build_list_page(self) -> None:
        layout = QVBoxLayout(self.list_page)
        title = QLabel("巨集清單")
        title.setStyleSheet("font-size: 20px; font-weight: bold;")
        layout.addWidget(title)

        self.macro_list = QListWidget()
        self.macro_list.itemDoubleClicked.connect(lambda _: self.open_selected_macro())
        layout.addWidget(self.macro_list)

        row = QHBoxLayout()
        self.btn_new = QPushButton("新增巨集")
        self.btn_open = QPushButton("編輯巨集")
        self.btn_rename = QPushButton("重新命名")
        self.btn_delete = QPushButton("刪除巨集")
        self.btn_import = QPushButton("匯入單一巨集")
        self.btn_export = QPushButton("匯出單一巨集")
        self.btn_fixed_drag = QPushButton("執行固定拖曳巨集")
        row.addWidget(self.btn_new)
        row.addWidget(self.btn_open)
        row.addWidget(self.btn_rename)
        row.addWidget(self.btn_delete)
        row.addWidget(self.btn_import)
        row.addWidget(self.btn_export)
        row.addWidget(self.btn_fixed_drag)
        layout.addLayout(row)

        self.btn_new.clicked.connect(self.create_macro)
        self.btn_open.clicked.connect(self.open_selected_macro)
        self.btn_rename.clicked.connect(self.rename_selected_macro)
        self.btn_delete.clicked.connect(self.delete_selected_macro)
        self.btn_import.clicked.connect(self.import_macro)
        self.btn_export.clicked.connect(self.export_macro)
        self.btn_fixed_drag.clicked.connect(self.run_fixed_drag_macro)

    def _build_edit_page(self) -> None:
        layout = QVBoxLayout(self.edit_page)
        head = QHBoxLayout()
        self.btn_back = QPushButton("返回清單")
        self.edit_title = QLabel("編輯巨集")
        self.edit_title.setStyleSheet("font-size: 18px; font-weight: bold;")
        head.addWidget(self.btn_back)
        head.addWidget(self.edit_title)
        head.addStretch()
        layout.addLayout(head)

        self.name_edit = QLineEdit()
        self.play_mode_box = QComboBox()
        self.play_mode_box.addItems(["不重複", "重複次數", "重複至時間", "到手動停止"])
        self.repeat_count_box = QSpinBox()
        self.repeat_count_box.setRange(1, 9999)
        self.repeat_count_box.setValue(1)
        self.repeat_until_edit = QLineEdit()
        self.repeat_until_edit.setPlaceholderText("YYYY-MM-DD HH:MM:SS")
        self.repeat_interval_box = QSpinBox()
        self.repeat_interval_box.setRange(0, 600000)
        self.repeat_interval_box.setValue(0)
        self.repeat_interval_box.setSuffix(" ms")
        self.delay_box = QSpinBox()
        self.delay_box.setRange(0, 30)
        self.delay_box.setValue(1)
        self.delay_box.setSuffix(" 秒")
        self.speed_box = QSpinBox()
        self.speed_box.setRange(10, 500)
        self.speed_box.setSuffix(" %")

        row1 = QHBoxLayout()
        row1.addWidget(QLabel("巨集名稱"))
        row1.addWidget(self.name_edit)
        row1.addWidget(QLabel("開始延遲"))
        row1.addWidget(self.delay_box)
        row1.addWidget(QLabel("播放速度"))
        row1.addWidget(self.speed_box)
        row1.addStretch()
        layout.addLayout(row1)

        row2 = QHBoxLayout()
        row2.addWidget(QLabel("播放方式"))
        row2.addWidget(self.play_mode_box)
        row2.addWidget(QLabel("重複次數"))
        row2.addWidget(self.repeat_count_box)
        row2.addWidget(QLabel("重複至時間"))
        row2.addWidget(self.repeat_until_edit)
        row2.addWidget(QLabel("重複間隔"))
        row2.addWidget(self.repeat_interval_box)
        row2.addStretch()
        layout.addLayout(row2)

        add_action_row = QHBoxLayout()
        self.btn_focus_window = QPushButton("聚焦視窗")
        self.btn_add_wait = QPushButton("等待")
        self.btn_add_mouse_move = QPushButton("移動滑鼠")
        self.btn_add_mouse_click = QPushButton("滑鼠點擊")
        self.btn_add_text = QPushButton("鍵盤輸入")
        self.btn_add_hotkey = QPushButton("鍵盤組合鍵")
        add_action_row.addWidget(self.btn_focus_window)
        add_action_row.addWidget(self.btn_add_wait)
        add_action_row.addWidget(self.btn_add_mouse_move)
        add_action_row.addWidget(self.btn_add_mouse_click)
        add_action_row.addWidget(self.btn_add_text)
        add_action_row.addWidget(self.btn_add_hotkey)
        add_action_row.addStretch()
        layout.addLayout(add_action_row)

        control_row = QHBoxLayout()
        self.btn_record = QPushButton("開始錄製")
        self.btn_stop_record = QPushButton("停止錄製")
        self.btn_play = QPushButton("開始播放")
        self.btn_stop_play = QPushButton("停止播放")
        self.btn_clear = QPushButton("清空事件")
        self.btn_apply = QPushButton("儲存變更")
        control_row.addWidget(self.btn_record)
        control_row.addWidget(self.btn_stop_record)
        control_row.addWidget(self.btn_play)
        control_row.addWidget(self.btn_stop_play)
        control_row.addWidget(self.btn_clear)
        control_row.addWidget(self.btn_apply)
        layout.addLayout(control_row)

        reorder_row = QHBoxLayout()
        self.btn_move_up = QPushButton("上移")
        self.btn_move_down = QPushButton("下移")
        reorder_row.addWidget(self.btn_move_up)
        reorder_row.addWidget(self.btn_move_down)
        reorder_row.addStretch()
        layout.addLayout(reorder_row)

        self.event_list = QListWidget()
        self.event_list.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        layout.addWidget(self.event_list)

        self.hint = QLabel("提示：錄製為全域鍵盤/滑鼠事件，播放前請切到目標視窗。")
        self.hint.setWordWrap(True)
        layout.addWidget(self.hint)

        self.btn_back.clicked.connect(self.go_back_to_list)
        self.btn_record.clicked.connect(self.on_start_record)
        self.btn_stop_record.clicked.connect(self.on_stop_record)
        self.btn_play.clicked.connect(self.on_play)
        self.btn_stop_play.clicked.connect(self.on_stop_play)
        self.btn_clear.clicked.connect(self.on_clear_events)
        self.btn_apply.clicked.connect(self.apply_editor_changes)
        self.btn_focus_window.clicked.connect(self.pick_and_focus_window)
        self.btn_add_wait.clicked.connect(self.add_wait_action)
        self.btn_add_mouse_move.clicked.connect(self.add_mouse_move_action)
        self.btn_add_mouse_click.clicked.connect(self.add_mouse_click_action)
        self.btn_add_text.clicked.connect(self.add_text_input_action)
        self.btn_add_hotkey.clicked.connect(self.add_hotkey_action)

        self.shortcut_start_record = QShortcut(QKeySequence("F5"), self)
        self.shortcut_start_record.activated.connect(self.on_start_record)
        self.shortcut_stop_record = QShortcut(QKeySequence("F6"), self)
        self.shortcut_stop_record.activated.connect(self.on_stop_record)

        self.shortcut_delete_event = QShortcut(QKeySequence("Delete"), self)
        self.shortcut_delete_event.activated.connect(self.delete_selected_event)
        self.shortcut_copy_events = QShortcut(QKeySequence("Ctrl+C"), self)
        self.shortcut_copy_events.activated.connect(self.copy_selected_events)

        self.btn_move_up.clicked.connect(self.move_selected_event_up)
        self.btn_move_down.clicked.connect(self.move_selected_event_down)

        self.play_mode_box.currentIndexChanged.connect(self._update_play_mode_fields)
        self._update_play_mode_fields()

    def _load_store(self) -> None:
        try:
            self.store.load()
        except Exception as exc:
            QMessageBox.warning(self, "讀取失敗", f"讀取 macros.json 失敗：{exc}")
            self.store.items = []
        self._refresh_macro_list()

    def _refresh_macro_list(self) -> None:
        self.macro_list.clear()
        for item in self.store.items:
            self.macro_list.addItem(QListWidgetItem(f"{item.name}（{len(item.events)} 筆事件）"))

    def _selected_macro_index(self) -> int | None:
        row = self.macro_list.currentRow()
        if row < 0 or row >= len(self.store.items):
            return None
        return row

    def create_macro(self) -> None:
        self.current_index = None
        self.name_edit.setText(self.store.unique_name("新巨集"))
        self.play_mode_box.setCurrentIndex(self._play_mode_to_index("once"))
        self.repeat_count_box.setValue(1)
        self.repeat_until_edit.setText("")
        self.repeat_interval_box.setValue(0)
        self.delay_box.setValue(1)
        self.speed_box.setValue(100)
        self.engine.clear()
        self.edit_title.setText("編輯巨集：新巨集（尚未儲存）")
        self._refresh_event_list()
        self.stack.setCurrentWidget(self.edit_page)

    def open_selected_macro(self) -> None:
        idx = self._selected_macro_index()
        if idx is None:
            QMessageBox.information(self, "提示", "請先選擇一個巨集")
            return
        self.current_index = idx
        item = self.store.items[idx]
        self.name_edit.setText(item.name)
        self.play_mode_box.setCurrentIndex(self._play_mode_to_index(item.play_mode))
        self.repeat_count_box.setValue(item.repeat_count)
        self.repeat_until_edit.setText(item.repeat_until_at)
        self.repeat_interval_box.setValue(item.repeat_interval_ms)
        self.delay_box.setValue(item.start_delay)
        self.speed_box.setValue(item.speed_percent)
        self.engine.set_events(item.events)
        self.edit_title.setText(f"編輯巨集：{item.name}")
        self._refresh_event_list()
        self.stack.setCurrentWidget(self.edit_page)

    def rename_selected_macro(self) -> None:
        idx = self._selected_macro_index()
        if idx is None:
            QMessageBox.information(self, "提示", "請先選擇一個巨集")
            return
        current = self.store.items[idx]
        new_name = f"{current.name} (新版)"
        if any(item.name == new_name for i, item in enumerate(self.store.items) if i != idx):
            new_name = self.store.unique_name(current.name)
        current.name = new_name
        self._save_store()
        self._refresh_macro_list()
        self.macro_list.setCurrentRow(idx)
        self.statusBar().showMessage(f"已重新命名為：{new_name}")

    def delete_selected_macro(self) -> None:
        idx = self._selected_macro_index()
        if idx is None:
            QMessageBox.information(self, "提示", "請先選擇一個巨集")
            return
        name = self.store.items[idx].name
        confirm = QMessageBox.question(self, "確認刪除", f"確定要刪除巨集「{name}」嗎？")
        if confirm != QMessageBox.StandardButton.Yes:
            return
        self.store.items.pop(idx)
        self._save_store()
        self._refresh_macro_list()
        self.statusBar().showMessage(f"已刪除：{name}")

    def import_macro(self) -> None:
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "匯入單一巨集(JSON)",
            str(Path.cwd()),
            "JSON Files (*.json)",
        )
        if not file_path:
            return
        try:
            raw = json.loads(Path(file_path).read_text(encoding="utf-8"))
            events = [MacroEvent.from_dict(item) for item in raw]
            name = self.store.unique_name(Path(file_path).stem)
            self.store.items.append(MacroItem(name=name, events=events))
            self._save_store()
            self._refresh_macro_list()
            self.statusBar().showMessage(f"匯入成功：{name}")
        except Exception as exc:
            QMessageBox.critical(self, "匯入失敗", str(exc))

    def export_macro(self) -> None:
        idx = self._selected_macro_index()
        if idx is None:
            QMessageBox.information(self, "提示", "請先選擇要匯出的巨集")
            return
        item = self.store.items[idx]
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "匯出單一巨集(JSON)",
            str(Path.cwd() / f"{item.name}.json"),
            "JSON Files (*.json)",
        )
        if not file_path:
            return
        try:
            payload = [ev.to_dict() for ev in item.events]
            Path(file_path).write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
            self.statusBar().showMessage(f"匯出成功：{file_path}")
        except Exception as exc:
            QMessageBox.critical(self, "匯出失敗", str(exc))

    def run_fixed_drag_macro(self) -> None:
        def _runner() -> None:
            try:
                pyautogui.moveTo(1285, 495, duration=MOUSE_MOVE_DURATION_SEC)
                time.sleep(0.5)
                pyautogui.mouseDown(button="left")
                time.sleep(0.5)
                pyautogui.mouseUp(button="left")
                self.statusBar().showMessage("固定拖曳巨集執行完成")
            except Exception as exc:
                QMessageBox.critical(self, "執行失敗", str(exc))

        threading.Thread(target=_runner, daemon=True).start()

    def go_back_to_list(self) -> None:
        if self.engine.is_recording:
            self.engine.stop_recording()
        if self.engine.is_playing:
            self.engine.stop_play()
        # 新增中的草稿（current_index is None）不自動存，必須按「儲存變更」才建立。
        if self.current_index is not None:
            self.apply_editor_changes(silent=True)
        self.stack.setCurrentWidget(self.list_page)
        self._refresh_macro_list()

    def apply_editor_changes(self, silent: bool = False) -> None:
        name = self.name_edit.text().strip()
        if not name:
            QMessageBox.warning(self, "名稱不可空白", "巨集名稱不能空白")
            return
        for idx, item in enumerate(self.store.items):
            if idx != self.current_index and item.name == name:
                QMessageBox.warning(self, "名稱重複", "已有同名巨集，請改用其他名稱")
                return

        if self.current_index is None:
            target = MacroItem(
                name=name,
                play_mode=self._play_mode_from_index(self.play_mode_box.currentIndex()),
                repeat_count=self.repeat_count_box.value(),
                repeat_until_at=self.repeat_until_edit.text().strip(),
                repeat_interval_ms=self.repeat_interval_box.value(),
                start_delay=self.delay_box.value(),
                speed_percent=self.speed_box.value(),
                events=self.engine.get_events_copy(),
            )
            self.store.items.append(target)
            self.current_index = len(self.store.items) - 1
        else:
            target = self.store.items[self.current_index]
            target.name = name
            target.play_mode = self._play_mode_from_index(self.play_mode_box.currentIndex())
            target.repeat_count = self.repeat_count_box.value()
            target.repeat_until_at = self.repeat_until_edit.text().strip()
            target.repeat_interval_ms = self.repeat_interval_box.value()
            target.start_delay = self.delay_box.value()
            target.speed_percent = self.speed_box.value()
            target.events = self.engine.get_events_copy()

        self.edit_title.setText(f"編輯巨集：{target.name}")
        self._save_store()
        self._refresh_macro_list()
        if not silent:
            self.statusBar().showMessage("已儲存巨集變更")

    def _save_store(self) -> None:
        try:
            self.store.save()
        except Exception as exc:
            QMessageBox.critical(self, "儲存失敗", f"寫入 macros.json 失敗：{exc}")

    def _set_editor_buttons(self) -> None:
        self.btn_record.setEnabled(not self.engine.is_recording and not self.engine.is_playing)
        self.btn_stop_record.setEnabled(self.engine.is_recording)
        self.btn_play.setEnabled(not self.engine.is_recording and not self.engine.is_playing and bool(self.engine.events))
        self.btn_stop_play.setEnabled(self.engine.is_playing)
        self.btn_clear.setEnabled(bool(self.engine.events) and not self.engine.is_recording and not self.engine.is_playing)
        self.btn_apply.setEnabled(not self.engine.is_recording and not self.engine.is_playing)
        editable = not self.engine.is_recording and not self.engine.is_playing
        self.btn_focus_window.setEnabled(editable)
        self.btn_add_wait.setEnabled(editable)
        self.btn_add_mouse_move.setEnabled(editable)
        self.btn_add_mouse_click.setEnabled(editable)
        self.btn_add_text.setEnabled(editable)
        self.btn_add_hotkey.setEnabled(editable)
        self.btn_move_up.setEnabled(editable)
        self.btn_move_down.setEnabled(editable)

    def _refresh_event_list(self) -> None:
        self.event_list.clear()
        events = self.engine.events
        # 顯示列表中的每一項，對應到底層 events 的一個區塊：[start, end)
        self._event_blocks: list[tuple[int, int]] = []

        idx = 0
        while idx < len(events):
            ev = events[idx]

            # 滑鼠：按下/放開合併顯示成「點擊」
            if ev.etype == "mouse_click" and ev.data.get("pressed") and idx + 1 < len(events):
                nxt = events[idx + 1]
                if (
                    nxt.etype == "mouse_click"
                    and not nxt.data.get("pressed")
                    and ev.data.get("button") == nxt.data.get("button")
                ):
                    btn = (
                        "左鍵"
                        if ev.data.get("button") == "left"
                        else "右鍵"
                        if ev.data.get("button") == "right"
                        else "中鍵"
                    )
                    x = ev.data.get("x")
                    y = ev.data.get("y")
                    if x is None or y is None:
                        detail = f"滑鼠{btn}點擊"
                    else:
                        detail = f"滑鼠{btn}點擊 @ ({x}, {y})"
                    self._event_blocks.append((idx, idx + 2))
                    self.event_list.addItem(QListWidgetItem(f"{ev.t:7.3f}s | {detail}"))
                    idx += 2
                    continue

            # 鍵盤：按下/放開合併顯示成「按鍵」
            if ev.etype == "key_press" and idx + 1 < len(events):
                nxt = events[idx + 1]
                if nxt.etype == "key_release" and self._key_name(ev) == self._key_name(nxt):
                    detail = f"鍵盤按鍵：{self._key_name(ev)}"
                    self._event_blocks.append((idx, idx + 2))
                    self.event_list.addItem(QListWidgetItem(f"{ev.t:7.3f}s | {detail}"))
                    idx += 2
                    continue

            # 其他：單一事件
            detail = self._format_event_text(ev)
            self._event_blocks.append((idx, idx + 1))
            self.event_list.addItem(QListWidgetItem(f"{ev.t:7.3f}s | {detail}"))
            idx += 1

        self._set_editor_buttons()

    @staticmethod
    def _format_event_text(ev: MacroEvent) -> str:
        if ev.etype == "mouse_move":
            return f"滑鼠移動到 ({ev.data.get('x')}, {ev.data.get('y')})"
        if ev.etype == "mouse_click":
            btn = "左鍵" if ev.data.get("button") == "left" else "右鍵" if ev.data.get("button") == "right" else "中鍵"
            action = "按下" if ev.data.get("pressed") else "放開"
            x = ev.data.get("x")
            y = ev.data.get("y")
            if x is None or y is None:
                return f"滑鼠{btn}{action}"
            return f"滑鼠{btn}{action} @ ({x}, {y})"
        if ev.etype == "mouse_scroll":
            return f"滑鼠滾輪 dy={ev.data.get('dy')} @ ({ev.data.get('x')}, {ev.data.get('y')})"
        if ev.etype == "key_press":
            return f"鍵盤按下：{MainWindow._key_name(ev)}"
        if ev.etype == "key_release":
            return f"鍵盤放開：{MainWindow._key_name(ev)}"
        if ev.etype == "text_input":
            return f"輸入文字：{ev.data.get('text', '')}"
        if ev.etype == "hotkey":
            keys = ev.data.get("keys", [])
            return f"組合鍵：{'+'.join(str(k) for k in keys)}"
        if ev.etype == "focus_window":
            return f"切換視窗焦點：{ev.data.get('title', '')}"
        if ev.etype == "wait":
            total = max(0.0, float(ev.data.get("seconds", 0.0)))
            h = total // 3600
            m = (total % 3600) // 60
            s = int(total % 60)
            ms = int(round((total - int(total)) * 1000))
            return f"等待：{int(h):02d}:{int(m):02d}:{s:02d}.{ms:03d}"
        return f"{ev.etype} {ev.data}"

    @staticmethod
    def _key_name(ev: MacroEvent) -> str:
        return str(ev.data.get("name") or ev.data.get("char") or ev.data.get("vk"))

    def _selected_event_block_index(self) -> int | None:
        if self.stack.currentWidget() != self.edit_page:
            return None
        row = self.event_list.currentRow()
        if row < 0:
            return None
        blocks = getattr(self, "_event_blocks", None)
        if not blocks or row >= len(blocks):
            return None
        return row

    def _selected_event_block_indices(self) -> list[int]:
        if self.stack.currentWidget() != self.edit_page:
            return []
        blocks = getattr(self, "_event_blocks", None)
        if not blocks:
            return []
        rows = sorted({idx.row() for idx in self.event_list.selectedIndexes()})
        return [r for r in rows if 0 <= r < len(blocks)]

    def delete_selected_event(self) -> None:
        if self.engine.is_recording or self.engine.is_playing:
            return
        indices = self._selected_event_block_indices()
        if not indices:
            i = self._selected_event_block_index()
            if i is None:
                return
            indices = [i]
        events = self.engine.events
        blocks = self._event_blocks
        pieces = [events[s:e] for (s, e) in blocks]
        for i in sorted(indices, reverse=True):
            pieces.pop(i)
        new_events = [ev for piece in pieces for ev in piece]
        self.engine.set_events(new_events)
        self._refresh_event_list()
        new_row = min(indices[0], max(0, len(self._event_blocks) - 1))
        if self.event_list.count() > 0:
            self.event_list.setCurrentRow(new_row)

    def copy_selected_events(self) -> None:
        if self.engine.is_recording or self.engine.is_playing:
            return
        indices = self._selected_event_block_indices()
        if not indices:
            i = self._selected_event_block_index()
            if i is None:
                return
            indices = [i]

        events = self.engine.events
        blocks = self._event_blocks
        pieces = [events[s:e] for (s, e) in blocks]
        selected_pieces = [pieces[i] for i in indices]

        insert_pos = indices[-1] + 1
        for offset, piece in enumerate(selected_pieces):
            clone_piece = [MacroEvent(ev.t, ev.etype, dict(ev.data)) for ev in piece]
            pieces.insert(insert_pos + offset, clone_piece)

        new_events = [ev for piece in pieces for ev in piece]
        self.engine.set_events(new_events)
        self._refresh_event_list()

        for offset in range(len(selected_pieces)):
            row = insert_pos + offset
            if 0 <= row < self.event_list.count():
                self.event_list.item(row).setSelected(True)
        if 0 <= insert_pos < self.event_list.count():
            self.event_list.setCurrentRow(insert_pos)
        self.statusBar().showMessage(f"已複製 {len(selected_pieces)} 筆指令")

    def move_selected_event_up(self) -> None:
        if self.engine.is_recording or self.engine.is_playing:
            return
        i = self._selected_event_block_index()
        if i is None or i <= 0:
            return
        events = self.engine.events
        blocks = self._event_blocks
        pieces = [events[s:e] for (s, e) in blocks]
        pieces[i - 1], pieces[i] = pieces[i], pieces[i - 1]
        new_events = [ev for piece in pieces for ev in piece]
        self.engine.set_events(new_events)
        self._refresh_event_list()
        self.event_list.setCurrentRow(i - 1)

    def move_selected_event_down(self) -> None:
        if self.engine.is_recording or self.engine.is_playing:
            return
        i = self._selected_event_block_index()
        if i is None or i >= len(self._event_blocks) - 1:
            return
        events = self.engine.events
        blocks = self._event_blocks
        pieces = [events[s:e] for (s, e) in blocks]
        pieces[i + 1], pieces[i] = pieces[i], pieces[i + 1]
        new_events = [ev for piece in pieces for ev in piece]
        self.engine.set_events(new_events)
        self._refresh_event_list()
        self.event_list.setCurrentRow(i + 1)

    def on_start_record(self) -> None:
        if self.engine.is_playing:
            self.statusBar().showMessage("播放中，無法開始錄製")
            return
        self.engine.start_recording()
        self.statusBar().showMessage("錄製中...")
        self._refresh_event_list()

    def on_stop_record(self) -> None:
        self.engine.stop_recording()
        self.statusBar().showMessage(f"已停止錄製，共 {len(self.engine.events)} 筆事件")
        self._refresh_event_list()

    def on_play(self) -> None:
        try:
            self.engine.play_async(
                play_mode=self._play_mode_from_index(self.play_mode_box.currentIndex()),
                repeat_count=self.repeat_count_box.value(),
                repeat_until_at=self.repeat_until_edit.text().strip(),
                repeat_interval_ms=self.repeat_interval_box.value(),
                start_delay=self.delay_box.value(),
                speed_percent=self.speed_box.value(),
            )
        except RuntimeError as exc:
            QMessageBox.warning(self, "無法播放", str(exc))
            return
        self.statusBar().showMessage(f"播放中...（{self.delay_box.value()} 秒後開始）")
        self._set_editor_buttons()

        def _watch() -> None:
            while self.engine.is_playing:
                time.sleep(0.15)
            self.statusBar().showMessage("播放完成")
            self._set_editor_buttons()

        threading.Thread(target=_watch, daemon=True).start()

    def on_stop_play(self) -> None:
        self.engine.stop_play()
        self.statusBar().showMessage("停止播放中...")
        self._set_editor_buttons()

    def on_clear_events(self) -> None:
        self.engine.clear()
        self.statusBar().showMessage("已清空事件")
        self._refresh_event_list()

    @staticmethod
    def _play_mode_to_index(play_mode: str) -> int:
        mapping = {
            "once": 0,
            "repeat_count": 1,
            "repeat_until": 2,
            "until_stopped": 3,
        }
        return mapping.get(play_mode, 0)

    @staticmethod
    def _play_mode_from_index(index: int) -> str:
        mapping = {
            0: "once",
            1: "repeat_count",
            2: "repeat_until",
            3: "until_stopped",
        }
        return mapping.get(index, "once")

    def _update_play_mode_fields(self) -> None:
        mode = self._play_mode_from_index(self.play_mode_box.currentIndex())
        self.repeat_count_box.setEnabled(mode == "repeat_count")
        self.repeat_until_edit.setEnabled(mode == "repeat_until")
        self.repeat_interval_box.setEnabled(mode in {"repeat_count", "repeat_until", "until_stopped"})

    def _next_event_t(self, delta: float = 0.1) -> float:
        if not self.engine.events:
            return 0.0
        return self.engine.events[-1].t + delta

    def _append_event(self, etype: str, data: dict[str, Any], delta: float = 0.1) -> None:
        self.engine.events.append(MacroEvent(t=self._next_event_t(delta), etype=etype, data=data))

    def add_mouse_move_action(self) -> None:
        dialog = MouseMoveDialog(self)
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return
        self._append_event("mouse_move", {"x": dialog.x_box.value(), "y": dialog.y_box.value()}, delta=0.1)
        self._refresh_event_list()
        self.statusBar().showMessage("已新增：滑鼠移動")

    def add_mouse_click_action(self) -> None:
        dialog = MouseClickDialog(self)
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return
        button = dialog.selected_button()
        # 不需輸入座標：播放時直接點擊當下滑鼠位置
        self._append_event("mouse_click", {"button": button, "pressed": True}, delta=0.1)
        self._append_event("mouse_click", {"button": button, "pressed": False}, delta=0.5)
        self._refresh_event_list()
        self.statusBar().showMessage(f"已新增：滑鼠{('右' if button == 'right' else '左')}鍵點擊（目前位置）")

    def add_text_input_action(self) -> None:
        text, ok = QInputDialog.getText(self, "新增鍵盤輸入", "請輸入文字：")
        if not ok:
            return
        if not text.strip():
            QMessageBox.warning(self, "缺少文字", "請輸入至少 1 個字元")
            return
        self._append_event("text_input", {"text": text}, delta=0.1)
        self._refresh_event_list()
        self.statusBar().showMessage("已新增：鍵盤輸入")

    def add_hotkey_action(self) -> None:
        dialog = HotkeyCaptureDialog(self)
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return
        if not dialog.keys:
            QMessageBox.warning(self, "未偵測到按鍵", "請至少按一個按鍵")
            return
        self._append_event("hotkey", {"keys": dialog.keys}, delta=0.1)
        self._refresh_event_list()
        self.statusBar().showMessage(f"已新增：組合鍵 {'+'.join(dialog.keys)}")

    def add_wait_action(self) -> None:
        dialog = WaitTimeDialog(self)
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return
        seconds = dialog.total_seconds()
        if seconds <= 0.001:
            QMessageBox.warning(self, "等待時間無效", "等待時間必須超過 1 毫秒")
            return
        self._append_event("wait", {"seconds": seconds}, delta=0.1)
        self._refresh_event_list()
        self.statusBar().showMessage(f"已新增：等待 {seconds:.3f} 秒")

    def pick_and_focus_window(self) -> None:
        titles = [t.strip() for t in gw.getAllTitles() if t and t.strip()]
        unique_titles: list[str] = []
        seen: set[str] = set()
        for title in titles:
            if title not in seen:
                seen.add(title)
                unique_titles.append(title)

        if not unique_titles:
            QMessageBox.information(self, "沒有可用視窗", "目前沒有偵測到可切換的視窗。")
            return

        selected, ok = QInputDialog.getItem(
            self,
            "選擇視窗",
            "請選擇要切換焦點的視窗：",
            unique_titles,
            0,
            False,
        )
        if not ok or not selected:
            return

        self._append_event("focus_window", {"title": selected}, delta=0.1)
        self._refresh_event_list()
        self.statusBar().showMessage(f"已加入焦點步驟：{selected}")


def run() -> None:
    app = QApplication([])
    window = MainWindow()
    window.show()
    app.exec()


if __name__ == "__main__":
    run()
