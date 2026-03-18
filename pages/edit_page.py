import json
import main
from PyQt6.QtCore import Qt, QTimer
from PyQt6.QtWidgets import (
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QPushButton,
    QLabel,
    QListWidget,
    QListWidgetItem,
    QLineEdit,
    QMessageBox,
    QInputDialog,
    QDialog,
    QDialogButtonBox,
    QFormLayout,
    QComboBox,
    QSpinBox,
    QDoubleSpinBox,
    QFileDialog,
)

try:
    import pyautogui
except Exception:
    pyautogui = None


class MouseMoveDialog(QDialog):
    """單一對話框：固定座標 / 畫面比例 / 圖片比對；固定座標時即時顯示目前滑鼠座標"""
    MODE_ABSOLUTE = 0
    MODE_RATIO = 1
    MODE_IMAGE = 2

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("滑鼠移動")
        layout = QVBoxLayout(self)

        self.mode_combo = QComboBox()
        self.mode_combo.addItems(["固定座標 (x, y)", "畫面比例 (0~1)", "圖片比對"])
        self.mode_combo.currentIndexChanged.connect(self._on_mode_changed)
        form = QFormLayout()
        form.addRow("模式:", self.mode_combo)

        self.x_spin = QDoubleSpinBox()
        self.x_spin.setRange(0, 99999)
        self.x_spin.setDecimals(0)
        self.x_spin.setValue(500)
        self.y_spin = QDoubleSpinBox()
        self.y_spin.setRange(0, 99999)
        self.y_spin.setDecimals(0)
        self.y_spin.setValue(500)
        form.addRow("X:", self.x_spin)
        form.addRow("Y:", self.y_spin)

        self.lbl_mouse_pos = QLabel("目前滑鼠: (--, --)")
        form.addRow("", self.lbl_mouse_pos)

        self.image_path_edit = QLineEdit()
        self.image_path_edit.setPlaceholderText("圖片路徑（或按瀏覽選擇）")
        btn_browse = QPushButton("瀏覽…")
        btn_browse.clicked.connect(self._browse_image)
        row_img = QHBoxLayout()
        row_img.addWidget(self.image_path_edit)
        row_img.addWidget(btn_browse)
        self.image_row_widget = QWidget()
        self.image_row_widget.setLayout(row_img)
        form.addRow("圖片:", self.image_row_widget)
        self.image_row_widget.hide()

        self.duration_spin = QSpinBox()
        self.duration_spin.setRange(0, 600000)
        self.duration_spin.setSingleStep(10)
        self.duration_spin.setValue(200)
        self.duration_spin.setSuffix(" ms")
        form.addRow("移動耗時:", self.duration_spin)

        layout.addLayout(form)
        bb = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        bb.accepted.connect(self._on_accept)
        bb.rejected.connect(self.reject)
        layout.addWidget(bb)

        self._pos_timer = QTimer(self)
        self._pos_timer.timeout.connect(self._update_mouse_label)

    def _browse_image(self):
        path, _ = QFileDialog.getOpenFileName(
            self,
            "選擇比對用圖片",
            "",
            "圖片 (*.png *.jpg *.jpeg *.bmp *.gif);;全部 (*.*)",
        )
        if path:
            self.image_path_edit.setText(path)

    def _update_mouse_label(self):
        if self.mode_combo.currentIndex() != self.MODE_ABSOLUTE:
            return
        if pyautogui is None:
            self.lbl_mouse_pos.setText("目前滑鼠: (無法取得)")
            return
        x, y = pyautogui.position()
        self.lbl_mouse_pos.setText(f"目前滑鼠: ({x}, {y})")

    def _on_mode_changed(self):
        mode = self.mode_combo.currentIndex()
        self._pos_timer.stop()
        self.lbl_mouse_pos.setText("目前滑鼠: (--, --)")

        self.x_spin.setVisible(mode != self.MODE_IMAGE)
        self.y_spin.setVisible(mode != self.MODE_IMAGE)
        self.lbl_mouse_pos.setVisible(mode == self.MODE_ABSOLUTE)
        self.image_row_widget.setVisible(mode == self.MODE_IMAGE)

        if mode == self.MODE_RATIO:
            self.x_spin.setDecimals(2)
            self.y_spin.setDecimals(2)
            self.x_spin.setRange(0, 1)
            self.y_spin.setRange(0, 1)
            self.x_spin.setValue(0.5)
            self.y_spin.setValue(0.5)
        elif mode == self.MODE_ABSOLUTE:
            self.x_spin.setDecimals(0)
            self.y_spin.setDecimals(0)
            self.x_spin.setRange(0, 99999)
            self.y_spin.setRange(0, 99999)
            self.x_spin.setValue(500)
            self.y_spin.setValue(500)
            self._update_mouse_label()
            self._pos_timer.start(50)

    def showEvent(self, event):
        super().showEvent(event)
        if self.mode_combo.currentIndex() == self.MODE_ABSOLUTE:
            self._update_mouse_label()
            self._pos_timer.start(50)

    def _on_accept(self):
        if self.mode_combo.currentIndex() == self.MODE_IMAGE and not self.image_path_edit.text().strip():
            QMessageBox.warning(self, "滑鼠移動", "請選擇或輸入圖片路徑")
            return
        self.accept()

    def set_action(self, action_dict):
        """預填表單以編輯既有動作（在 exec 前呼叫）"""
        self.mode_combo.blockSignals(True)
        dur = int(action_dict.get("duration_ms", 200))
        self.duration_spin.setValue(dur)
        if action_dict.get("image"):
            self.mode_combo.setCurrentIndex(self.MODE_IMAGE)
        elif action_dict.get("ratio"):
            self.mode_combo.setCurrentIndex(self.MODE_RATIO)
        else:
            self.mode_combo.setCurrentIndex(self.MODE_ABSOLUTE)
        self.mode_combo.blockSignals(False)
        self._on_mode_changed()
        if action_dict.get("image"):
            self.image_path_edit.setText(action_dict.get("image", ""))
        elif action_dict.get("ratio"):
            self.x_spin.setValue(float(action_dict.get("x", 0.5)))
            self.y_spin.setValue(float(action_dict.get("y", 0.5)))
        else:
            self.x_spin.setValue(int(action_dict.get("x", 500)))
            self.y_spin.setValue(int(action_dict.get("y", 500)))

    def done(self, result):
        self._pos_timer.stop()
        super().done(result)

    def get_action(self):
        """回傳要插入的 action dict，或 None 表示取消"""
        if self.exec() != QDialog.DialogCode.Accepted:
            return None
        mode = self.mode_combo.currentIndex()
        duration_ms = self.duration_spin.value()
        action = {"action": "move", "duration_ms": duration_ms}
        if mode == self.MODE_IMAGE:
            action["image"] = self.image_path_edit.text().strip()
        elif mode == self.MODE_RATIO:
            action["ratio"] = True
            action["x"] = round(self.x_spin.value(), 4)
            action["y"] = round(self.y_spin.value(), 4)
        else:
            action["x"] = int(self.x_spin.value())
            action["y"] = int(self.y_spin.value())
        return action


class DragDialog(QDialog):
    """拖曳：選擇左/右鍵，從何處拖到何處（固定/比例/圖片比對）"""
    MODE_ABSOLUTE = 0
    MODE_RATIO = 1
    MODE_IMAGE = 2

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("滑鼠拖曳")
        layout = QVBoxLayout(self)
        form = QFormLayout()

        self.btn_combo = QComboBox()
        self.btn_combo.addItems(["left", "right"])
        form.addRow("按鍵:", self.btn_combo)

        self.from_mode = QComboBox()
        self.from_mode.addItems(["固定座標 (x, y)", "畫面比例 (0~1)", "圖片比對"])
        self.from_mode.currentIndexChanged.connect(self._on_mode_changed)
        form.addRow("起點模式:", self.from_mode)

        self.from_x = QDoubleSpinBox()
        self.from_x.setRange(0, 99999)
        self.from_x.setDecimals(0)
        self.from_x.setValue(100)
        self.from_y = QDoubleSpinBox()
        self.from_y.setRange(0, 99999)
        self.from_y.setDecimals(0)
        self.from_y.setValue(100)
        form.addRow("起點 X:", self.from_x)
        form.addRow("起點 Y:", self.from_y)

        self.from_image = QLineEdit()
        self.from_image.setPlaceholderText("起點圖片路徑（或按瀏覽選擇）")
        btn_from_browse = QPushButton("瀏覽…")
        btn_from_browse.clicked.connect(lambda: self._browse_image(self.from_image))
        from_row = QHBoxLayout()
        from_row.addWidget(self.from_image)
        from_row.addWidget(btn_from_browse)
        self.from_image_row = QWidget()
        self.from_image_row.setLayout(from_row)
        form.addRow("起點圖片:", self.from_image_row)

        self.to_mode = QComboBox()
        self.to_mode.addItems(["固定座標 (x, y)", "畫面比例 (0~1)", "圖片比對"])
        self.to_mode.currentIndexChanged.connect(self._on_mode_changed)
        form.addRow("終點模式:", self.to_mode)

        self.to_x = QDoubleSpinBox()
        self.to_x.setRange(0, 99999)
        self.to_x.setDecimals(0)
        self.to_x.setValue(500)
        self.to_y = QDoubleSpinBox()
        self.to_y.setRange(0, 99999)
        self.to_y.setDecimals(0)
        self.to_y.setValue(500)
        form.addRow("終點 X:", self.to_x)
        form.addRow("終點 Y:", self.to_y)

        self.to_image = QLineEdit()
        self.to_image.setPlaceholderText("終點圖片路徑（或按瀏覽選擇）")
        btn_to_browse = QPushButton("瀏覽…")
        btn_to_browse.clicked.connect(lambda: self._browse_image(self.to_image))
        to_row = QHBoxLayout()
        to_row.addWidget(self.to_image)
        to_row.addWidget(btn_to_browse)
        self.to_image_row = QWidget()
        self.to_image_row.setLayout(to_row)
        form.addRow("終點圖片:", self.to_image_row)

        self.lbl_mouse_pos = QLabel("目前滑鼠: (--, --)")
        form.addRow("", self.lbl_mouse_pos)

        self.duration_spin = QSpinBox()
        self.duration_spin.setRange(0, 600000)
        self.duration_spin.setSingleStep(10)
        self.duration_spin.setValue(200)
        self.duration_spin.setSuffix(" ms")
        form.addRow("拖曳耗時:", self.duration_spin)

        layout.addLayout(form)
        bb = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        bb.accepted.connect(self._on_accept)
        bb.rejected.connect(self.reject)
        layout.addWidget(bb)

        self._pos_timer = QTimer(self)
        self._pos_timer.timeout.connect(self._update_mouse_label)
        self._on_mode_changed()

    def _browse_image(self, target_edit: QLineEdit):
        path, _ = QFileDialog.getOpenFileName(
            self,
            "選擇比對用圖片",
            "",
            "圖片 (*.png *.jpg *.jpeg *.bmp *.gif);;全部 (*.*)",
        )
        if path:
            target_edit.setText(path)

    def _update_mouse_label(self):
        if pyautogui is None:
            self.lbl_mouse_pos.setText("目前滑鼠: (無法取得)")
            return
        x, y = pyautogui.position()
        self.lbl_mouse_pos.setText(f"目前滑鼠: ({x}, {y})")

    def _apply_mode(self, mode_combo, x_spin, y_spin, image_row):
        mode = mode_combo.currentIndex()
        is_image = mode == self.MODE_IMAGE
        x_spin.setVisible(not is_image)
        y_spin.setVisible(not is_image)
        image_row.setVisible(is_image)
        if mode == self.MODE_RATIO:
            x_spin.setDecimals(2)
            y_spin.setDecimals(2)
            x_spin.setRange(0, 1)
            y_spin.setRange(0, 1)
        else:
            x_spin.setDecimals(0)
            y_spin.setDecimals(0)
            x_spin.setRange(0, 99999)
            y_spin.setRange(0, 99999)

    def _on_mode_changed(self):
        self._apply_mode(self.from_mode, self.from_x, self.from_y, self.from_image_row)
        self._apply_mode(self.to_mode, self.to_x, self.to_y, self.to_image_row)
        # 只要任一端是固定座標，就即時顯示滑鼠座標
        show_mouse = (self.from_mode.currentIndex() == self.MODE_ABSOLUTE) or (self.to_mode.currentIndex() == self.MODE_ABSOLUTE)
        self.lbl_mouse_pos.setVisible(show_mouse)
        self._pos_timer.stop()
        if show_mouse:
            self._update_mouse_label()
            self._pos_timer.start(50)

    def showEvent(self, event):
        super().showEvent(event)
        self._on_mode_changed()

    def done(self, result):
        self._pos_timer.stop()
        super().done(result)

    def _on_accept(self):
        if self.from_mode.currentIndex() == self.MODE_IMAGE and not self.from_image.text().strip():
            QMessageBox.warning(self, "滑鼠拖曳", "請選擇或輸入起點圖片路徑")
            return
        if self.to_mode.currentIndex() == self.MODE_IMAGE and not self.to_image.text().strip():
            QMessageBox.warning(self, "滑鼠拖曳", "請選擇或輸入終點圖片路徑")
            return
        self.accept()

    def _endpoint_dict(self, mode_combo, x_spin, y_spin, image_edit):
        mode = mode_combo.currentIndex()
        if mode == self.MODE_IMAGE:
            return {"image": image_edit.text().strip()}
        if mode == self.MODE_RATIO:
            return {"ratio": True, "x": round(x_spin.value(), 4), "y": round(y_spin.value(), 4)}
        return {"x": int(x_spin.value()), "y": int(y_spin.value())}

    def set_action(self, action_dict):
        """預填拖曳設定"""
        self.btn_combo.setCurrentText(action_dict.get("btn", "left"))
        self.duration_spin.setValue(int(action_dict.get("duration_ms", 200)))

        frm = action_dict.get("from", {}) if isinstance(action_dict.get("from"), dict) else {}
        to = action_dict.get("to", {}) if isinstance(action_dict.get("to"), dict) else {}

        def set_endpoint(ep, mode_combo, x_spin, y_spin, image_edit):
            if ep.get("image"):
                mode_combo.setCurrentIndex(self.MODE_IMAGE)
                image_edit.setText(ep.get("image", ""))
            elif ep.get("ratio"):
                mode_combo.setCurrentIndex(self.MODE_RATIO)
                x_spin.setValue(float(ep.get("x", 0.5)))
                y_spin.setValue(float(ep.get("y", 0.5)))
            else:
                mode_combo.setCurrentIndex(self.MODE_ABSOLUTE)
                x_spin.setValue(int(ep.get("x", 100)))
                y_spin.setValue(int(ep.get("y", 100)))

        self.from_mode.blockSignals(True)
        self.to_mode.blockSignals(True)
        set_endpoint(frm, self.from_mode, self.from_x, self.from_y, self.from_image)
        set_endpoint(to, self.to_mode, self.to_x, self.to_y, self.to_image)
        self.from_mode.blockSignals(False)
        self.to_mode.blockSignals(False)
        self._on_mode_changed()

    def get_action(self):
        if self.exec() != QDialog.DialogCode.Accepted:
            return None
        return {
            "action": "drag",
            "btn": self.btn_combo.currentText(),
            "from": self._endpoint_dict(self.from_mode, self.from_x, self.from_y, self.from_image),
            "to": self._endpoint_dict(self.to_mode, self.to_x, self.to_y, self.to_image),
            "duration_ms": int(self.duration_spin.value()),
        }


class EditorWindow(QWidget):
    def __init__(self, on_back_callback):
        super().__init__()
        self.on_back_callback = on_back_callback
        self.actions = []
        
        root = QHBoxLayout()

        # 左側工具欄（插入動作）
        toolbar = QVBoxLayout()
        self.toolbar_title = QLabel("工具欄")
        self.toolbar_title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        toolbar.addWidget(self.toolbar_title)

        btn_focus = QPushButton("視窗 Focus")
        btn_focus.clicked.connect(self.add_action_focus)
        toolbar.addWidget(btn_focus)

        btn_wait = QPushButton("等待")
        btn_wait.clicked.connect(self.add_action_wait)
        toolbar.addWidget(btn_wait)

        toolbar.addWidget(QLabel("滑鼠動作"))
        btn_move = QPushButton("移動")
        btn_move.clicked.connect(self.add_action_mouse_move)
        toolbar.addWidget(btn_move)
        btn_click = QPushButton("點擊")
        btn_click.clicked.connect(self.add_action_mouse_click)
        toolbar.addWidget(btn_click)
        btn_drag = QPushButton("拖曳")
        btn_drag.clicked.connect(self.add_action_mouse_drag)
        toolbar.addWidget(btn_drag)

        toolbar.addWidget(QLabel("鍵盤輸入"))
        btn_type = QPushButton("一般輸入")
        btn_type.clicked.connect(self.add_action_keyboard_type)
        toolbar.addWidget(btn_type)
        btn_hotkey = QPushButton("組合鍵 (Hotkeys)")
        btn_hotkey.clicked.connect(self.add_action_keyboard_hotkey)
        toolbar.addWidget(btn_hotkey)

        toolbar.addWidget(QLabel("順序調整"))
        btn_up = QPushButton("上移一行")
        btn_up.clicked.connect(self.move_selected_up)
        toolbar.addWidget(btn_up)
        btn_down = QPushButton("下移一行")
        btn_down.clicked.connect(self.move_selected_down)
        toolbar.addWidget(btn_down)

        btn_remove = QPushButton("移除(單行)")
        btn_remove.clicked.connect(self.remove_selected_step)
        toolbar.addWidget(btn_remove)

        toolbar.addStretch(1)

        self.toolbar_widget = QWidget()
        self.toolbar_widget.setLayout(toolbar)
        self.toolbar_widget.setFixedWidth(160)
        self.toolbar_widget.setObjectName("toolbar")
        root.addWidget(self.toolbar_widget)

        # 右側主要編輯區
        layout = QVBoxLayout()

        # 檔名輸入
        self.lbl_name = QLabel("腳本名稱 (不需副檔名):")
        layout.addWidget(self.lbl_name)
        self.name_input = QLineEdit()
        layout.addWidget(self.name_input)

        # 指令清單（可讀模式）
        self.lbl_steps = QLabel("指令內容（可讀模式）:")
        layout.addWidget(self.lbl_steps)
        self.list_steps = QListWidget()
        self.list_steps.setSelectionMode(QListWidget.SelectionMode.SingleSelection)
        self.list_steps.itemDoubleClicked.connect(self._on_step_double_clicked)
        layout.addWidget(self.list_steps)

        # 按鈕
        btn_layout = QHBoxLayout()
        self.btn_cancel = QPushButton("取消 / 返回")
        self.btn_cancel.clicked.connect(self.on_back_callback)
        self.btn_cancel.setObjectName("secondaryButton")
        
        self.btn_save = QPushButton("儲存腳本")
        self.btn_save.setObjectName("primarySave")
        self.btn_save.clicked.connect(self.save_script)
        
        btn_layout.addWidget(self.btn_cancel)
        btn_layout.addWidget(self.btn_save)
        layout.addLayout(btn_layout)
        
        right = QWidget()
        right.setLayout(layout)
        root.addWidget(right, 1)
        self.setLayout(root)
        self.apply_theme()

    def apply_theme(self):
        """
        依據目前 Qt 調色盤的視窗背景亮度，套用深色/淺色樣式，
        避免在深色模式下按鈕與文字對比不足而「看不到」。
        """
        is_dark = False
        try:
            pal = self.palette()
            window = pal.color(self.backgroundRole())
            is_dark = window.lightness() < 128
        except Exception:
            is_dark = False

        if is_dark:
            self.setStyleSheet(
                """
                QLabel { color: #e8e8e8; }
                QLabel#toolbarTitle { font-weight: 700; }
                QWidget#toolbar { background-color: #232323; border-right: 1px solid #3a3a3a; }
                QLineEdit, QListWidget {
                    background-color: #2a2a2a;
                    color: #f2f2f2;
                    border: 1px solid #3a3a3a;
                    border-radius: 6px;
                    padding: 6px;
                    outline: none;
                    selection-background-color: #2d6cdf;
                    selection-color: #ffffff;
                }
                QListWidget:focus { outline: none; }
                QListWidget::item:focus { outline: none; }
                QListWidget::item:selected { background-color: #2d6cdf; color: #ffffff; }
                QPushButton {
                    background-color: #2f2f2f;
                    color: #f2f2f2;
                    border: 1px solid #444;
                    border-radius: 6px;
                    padding: 8px 10px;
                }
                QPushButton:hover { background-color: #3a3a3a; }
                QPushButton:pressed { background-color: #2a2a2a; }
                QPushButton:disabled { color: #888; border-color: #333; background-color: #262626; }
                QPushButton#primarySave {
                    background-color: #2ecc71;
                    color: white;
                    border: none;
                    font-weight: 700;
                }
                QPushButton#primarySave:hover { background-color: #27ae60; }
                """
            )
        else:
            self.setStyleSheet(
                """
                QLabel { color: #222; }
                QWidget#toolbar { background-color: #f6f6f6; border-right: 1px solid #ddd; }
                QLineEdit, QListWidget {
                    background-color: white;
                    color: #111;
                    border: 1px solid #cfcfcf;
                    border-radius: 6px;
                    padding: 6px;
                    outline: none;
                }
                QListWidget:focus { outline: none; }
                QListWidget::item:focus { outline: none; }
                /* 明確指定選取狀態文字顏色，避免淺色主題反白後看不到字 */
                QListWidget::item:selected { background-color: #cfe2ff; color: #111; }
                QPushButton {
                    background-color: #f2f2f2;
                    color: #111;
                    border: 1px solid #cfcfcf;
                    border-radius: 6px;
                    padding: 8px 10px;
                }
                QPushButton:hover { background-color: #eaeaea; }
                QPushButton:pressed { background-color: #dfdfdf; }
                QPushButton:disabled { color: #888; background-color: #f3f3f3; }
                QPushButton#primarySave {
                    background-color: #2ecc71;
                    color: white;
                    border: none;
                    font-weight: 700;
                }
                QPushButton#primarySave:hover { background-color: #27ae60; }
                """
            )

        self.toolbar_title.setObjectName("toolbarTitle")

    def prepare_new(self):
        """清空介面以新增腳本（無預設指令）"""
        self.name_input.setReadOnly(False)
        self.name_input.clear()
        self.actions = []
        self._loaded_filename = None
        self.refresh_steps()

    def load_existing(self, filename):
        """載入舊腳本內容"""
        file_path = main.SCRIPTS_DIR / filename
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
                if not isinstance(data, list):
                    raise ValueError("腳本內容必須是 JSON 陣列 (list)。")
                self.actions = data
                self.refresh_steps()
                self.name_input.setText(filename.replace(".json", ""))
                self.name_input.setReadOnly(False)  # 允許改名
                self._loaded_filename = filename
        except Exception as e:
            QMessageBox.critical(self, "錯誤", f"讀取失敗: {e}")

    def load_script(self, filename):
        """相容舊呼叫名稱"""
        self.load_existing(filename)

    def refresh_steps(self, select_row=None):
        self.list_steps.clear()
        for idx, act in enumerate(self.actions):
            item = QListWidgetItem(self.format_action(act))
            item.setData(Qt.ItemDataRole.UserRole, idx)
            self.list_steps.addItem(item)
        if select_row is not None and 0 <= select_row < self.list_steps.count():
            self.list_steps.setCurrentRow(select_row)
        elif self.list_steps.count() > 0 and self.list_steps.currentRow() < 0:
            self.list_steps.setCurrentRow(0)

    def format_action(self, act):
        t = act.get("action")
        if t == "focus":
            return f'Focus: title contains "{act.get("title", "")}"'
        if t == "move":
            dur = act.get("duration_ms", 200)
            if act.get("image"):
                return f"Mouse Move: 圖片比對 {act.get('image', '')} duration={dur}ms"
            if act.get("ratio"):
                return f"Mouse Move: 比例 x={act.get('x')} y={act.get('y')} duration={dur}ms"
            return f"Mouse Move: x={act.get('x')} y={act.get('y')} duration={dur}ms"
        if t == "click":
            btn = act.get("btn", "left")
            clicks = act.get("clicks", 1)
            interval = act.get("interval_ms", 0)
            extra = f" interval={interval}ms" if interval else ""
            return f"Mouse Click: btn={btn} clicks={clicks}{extra}"
        if t == "drag":
            btn = act.get("btn", "left")
            dur = act.get("duration_ms", 200)
            frm = act.get("from", {})
            to = act.get("to", {})
            def ep_str(ep):
                if isinstance(ep, dict):
                    if ep.get("image"):
                        return f"img:{ep.get('image')}"
                    if ep.get("ratio"):
                        return f"ratio({ep.get('x')},{ep.get('y')})"
                    return f"({ep.get('x')},{ep.get('y')})"
                return str(ep)
            return f"Mouse Drag: btn={btn} {ep_str(frm)} -> {ep_str(to)} duration={dur}ms"
        if t == "type":
            text = act.get("text", "")
            interval = act.get("interval_ms", 0)
            preview = text.replace("\r\n", "\\n").replace("\n", "\\n")
            if len(preview) > 60:
                preview = preview[:57] + "..."
            extra = f" interval={interval}ms" if interval else ""
            return f'Type: "{preview}"{extra}'
        if t == "hotkey":
            keys = act.get("keys", [])
            return f"Hotkey: {' + '.join(keys) if isinstance(keys, list) else keys}"
        if t == "wait":
            return f"Wait: {act.get('ms', 500)}ms"
        if t == "key":
            return f"Key: {act.get('key')}"
        return f"Unknown: {json.dumps(act, ensure_ascii=False)}"

    def _insert_action_after_cursor(self, action_dict):
        row = self.list_steps.currentRow()
        insert_at = row + 1 if row >= 0 else len(self.actions)
        self.actions.insert(insert_at, action_dict)
        self.refresh_steps(select_row=insert_at)

    def _on_step_double_clicked(self, item):
        """指令列快擊兩下 → 編輯該行指令"""
        row = self.list_steps.row(item)
        if row < 0 or row >= len(self.actions):
            return
        action = self.actions[row]
        act_type = action.get("action")
        updated = None
        if act_type == "focus":
            title, ok = QInputDialog.getText(
                self,
                "視窗 Focus",
                "視窗標題包含文字 (title contains):",
                text=action.get("title", ""),
            )
            if ok and title.strip():
                updated = {"action": "focus", "title": title.strip()}
        elif act_type == "move":
            dlg = MouseMoveDialog(self)
            dlg.set_action(action)
            updated = dlg.get_action()
        elif act_type == "click":
            btn = action.get("btn", "left")
            idx = ["left", "right", "middle"].index(btn) if btn in ("left", "right", "middle") else 0
            btn, ok = QInputDialog.getItem(self, "滑鼠點擊", "按鍵:", ["left", "right", "middle"], idx, False)
            if not ok:
                return
            clicks, ok = QInputDialog.getInt(self, "滑鼠點擊", "點擊次數:", action.get("clicks", 1), 1, 10, 1)
            if ok:
                updated = {"action": "click", "btn": btn, "clicks": int(clicks)}
        elif act_type == "drag":
            dlg = DragDialog(self)
            dlg.set_action(action)
            updated = dlg.get_action()
        elif act_type == "type":
            text, ok = QInputDialog.getMultiLineText(self, "鍵盤一般輸入", "輸入文字:", action.get("text", ""))
            if not ok:
                return
            interval_ms, ok = QInputDialog.getInt(
                self, "鍵盤一般輸入", "每字間隔 (ms):",
                action.get("interval_ms", 0), 0, 60000, 1,
            )
            if ok:
                updated = {"action": "type", "text": text}
                if interval_ms:
                    updated["interval_ms"] = int(interval_ms)
        elif act_type == "hotkey":
            keys = action.get("keys", [])
            keys_text = ",".join(keys) if isinstance(keys, list) else str(keys)
            keys_str, ok = QInputDialog.getText(
                self, "組合鍵 (Hotkeys)", "Keys (用逗號分隔，例如 ctrl,s):",
                text=keys_text,
            )
            if ok and keys_str.strip():
                key_list = [k.strip() for k in keys_str.split(",") if k.strip()]
                if key_list:
                    updated = {"action": "hotkey", "keys": key_list}
        elif act_type == "wait":
            ms, ok = QInputDialog.getInt(
                self, "等待", "等待時間 (ms):",
                action.get("ms", 500), 0, 3600000, 10,
            )
            if ok:
                updated = {"action": "wait", "ms": int(ms)}
        elif act_type == "key":
            key_val, ok = QInputDialog.getText(self, "按鍵", "Key:", text=action.get("key", ""))
            if ok and key_val.strip():
                updated = {"action": "key", "key": key_val.strip()}
        else:
            QMessageBox.information(self, "提示", "此指令類型不支援編輯")
            return
        if updated is not None:
            self.actions[row] = updated
            self.refresh_steps(select_row=row)

    def add_action_focus(self):
        title, ok = QInputDialog.getText(self, "視窗 Focus", "視窗標題包含文字 (title contains):")
        if not ok or not title.strip():
            return
        self._insert_action_after_cursor({"action": "focus", "title": title.strip()})

    def add_action_mouse_move(self):
        dlg = MouseMoveDialog(self)
        action = dlg.get_action()
        if action is not None:
            self._insert_action_after_cursor(action)

    def add_action_mouse_click(self):
        btn, ok = QInputDialog.getItem(self, "滑鼠點擊", "按鍵:", ["left", "right", "middle"], 0, False)
        if not ok:
            return
        clicks, ok = QInputDialog.getInt(self, "滑鼠點擊", "點擊次數:", 1, 1, 10, 1)
        if not ok:
            return
        action = {"action": "click", "btn": btn, "clicks": int(clicks)}
        self._insert_action_after_cursor(action)

    def add_action_mouse_drag(self):
        dlg = DragDialog(self)
        action = dlg.get_action()
        if action is not None:
            self._insert_action_after_cursor(action)

    def add_action_keyboard_type(self):
        text, ok = QInputDialog.getMultiLineText(self, "鍵盤一般輸入", "輸入文字:", "")
        if not ok:
            return
        interval_ms, ok = QInputDialog.getInt(self, "鍵盤一般輸入", "每字間隔 (ms):", 0, 0, 60000, 1)
        if not ok:
            return
        action = {"action": "type", "text": text}
        if interval_ms:
            action["interval_ms"] = int(interval_ms)
        self._insert_action_after_cursor(action)

    def add_action_keyboard_hotkey(self):
        keys_text, ok = QInputDialog.getText(self, "組合鍵 (Hotkeys)", "Keys (用逗號分隔，例如 ctrl,s):")
        if not ok or not keys_text.strip():
            return
        keys = [k.strip() for k in keys_text.split(",") if k.strip()]
        if not keys:
            return
        self._insert_action_after_cursor({"action": "hotkey", "keys": keys})

    def add_action_wait(self):
        ms, ok = QInputDialog.getInt(self, "等待", "等待時間 (ms):", 500, 0, 3600000, 10)
        if not ok:
            return
        self._insert_action_after_cursor({"action": "wait", "ms": int(ms)})

    def move_selected_up(self):
        row = self.list_steps.currentRow()
        if row <= 0 or row >= len(self.actions):
            return
        self.actions[row - 1], self.actions[row] = self.actions[row], self.actions[row - 1]
        self.refresh_steps(select_row=row - 1)

    def move_selected_down(self):
        row = self.list_steps.currentRow()
        if row < 0 or row >= len(self.actions) - 1:
            return
        self.actions[row + 1], self.actions[row] = self.actions[row], self.actions[row + 1]
        self.refresh_steps(select_row=row + 1)

    def remove_selected_step(self):
        row = self.list_steps.currentRow()
        if row < 0 or row >= len(self.actions):
            QMessageBox.information(self, "提示", "請先選取要移除的指令行")
            return
        del self.actions[row]
        next_row = min(row, len(self.actions) - 1) if self.actions else None
        self.refresh_steps(select_row=next_row if next_row is not None and next_row >= 0 else None)

    def save_script(self):
        name = self.name_input.text().strip()
        if not name:
            QMessageBox.warning(self, "提示", "請輸入名稱")
            return
        
        try:
            new_filename = f"{name}.json"
            new_path = main.SCRIPTS_DIR / new_filename
            old_filename = getattr(self, "_loaded_filename", None)
            old_path = (main.SCRIPTS_DIR / old_filename) if old_filename else None

            # 若為改名：處理衝突
            if old_path and old_path.exists() and old_path != new_path:
                if new_path.exists():
                    reply = QMessageBox.question(
                        self,
                        "檔名已存在",
                        f"腳本「{name}」已存在，要覆蓋它嗎？\n（原檔將被取代，無法復原）",
                        QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                        QMessageBox.StandardButton.No,
                    )
                    if reply != QMessageBox.StandardButton.Yes:
                        return
                try:
                    old_path.rename(new_path)
                except Exception:
                    # 若跨磁碟或其他原因 rename 失敗，改為先寫新檔再刪舊檔
                    with open(new_path, "w", encoding="utf-8") as f:
                        json.dump(self.actions, f, indent=4, ensure_ascii=False)
                    try:
                        old_path.unlink()
                    except Exception:
                        pass
            else:
                # 新增或同名覆寫
                with open(new_path, 'w', encoding='utf-8') as f:
                    json.dump(self.actions, f, indent=4, ensure_ascii=False)

            self._loaded_filename = new_filename
            
            QMessageBox.information(self, "成功", "腳本已存檔")
            self.on_back_callback()
        except Exception as e:
            QMessageBox.critical(self, "儲存失敗", str(e))