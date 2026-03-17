import json
import main
from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import (
    QWidget,
    QVBoxLayout,
    QPushButton,
    QLabel,
    QHBoxLayout,
    QListWidget,
    QListWidgetItem,
    QLineEdit,
    QMessageBox,
    QInputDialog,
)
                             
class EditorWindow(QWidget):
    def __init__(self, on_back_callback):
        super().__init__()
        self.on_back_callback = on_back_callback
        self.actions = []
        
        root = QHBoxLayout()

        # 左側工具欄（插入動作）
        toolbar = QVBoxLayout()
        toolbar_title = QLabel("工具欄")
        toolbar_title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        toolbar_title.setStyleSheet("font-weight: bold; padding: 6px;")
        toolbar.addWidget(toolbar_title)

        btn_focus = QPushButton("視窗 Focus")
        btn_focus.clicked.connect(self.add_action_focus)
        toolbar.addWidget(btn_focus)

        toolbar.addWidget(QLabel("滑鼠動作"))
        btn_move = QPushButton("移動")
        btn_move.clicked.connect(self.add_action_mouse_move)
        toolbar.addWidget(btn_move)
        btn_click = QPushButton("點擊")
        btn_click.clicked.connect(self.add_action_mouse_click)
        toolbar.addWidget(btn_click)

        toolbar.addWidget(QLabel("鍵盤輸入"))
        btn_type = QPushButton("一般輸入")
        btn_type.clicked.connect(self.add_action_keyboard_type)
        toolbar.addWidget(btn_type)
        btn_hotkey = QPushButton("組合鍵 (Hotkeys)")
        btn_hotkey.clicked.connect(self.add_action_keyboard_hotkey)
        toolbar.addWidget(btn_hotkey)

        btn_wait = QPushButton("等待")
        btn_wait.clicked.connect(self.add_action_wait)
        toolbar.addWidget(btn_wait)

        toolbar.addWidget(QLabel("順序調整"))
        btn_up = QPushButton("上移一行")
        btn_up.clicked.connect(self.move_selected_up)
        toolbar.addWidget(btn_up)
        btn_down = QPushButton("下移一行")
        btn_down.clicked.connect(self.move_selected_down)
        toolbar.addWidget(btn_down)

        toolbar.addStretch(1)

        toolbar_widget = QWidget()
        toolbar_widget.setLayout(toolbar)
        toolbar_widget.setFixedWidth(160)
        toolbar_widget.setStyleSheet("background-color: #f6f6f6; border-right: 1px solid #ddd;")
        root.addWidget(toolbar_widget)

        # 右側主要編輯區
        layout = QVBoxLayout()

        # 檔名輸入
        layout.addWidget(QLabel("腳本名稱 (不需副檔名):"))
        self.name_input = QLineEdit()
        layout.addWidget(self.name_input)

        # 指令清單（可讀模式）
        layout.addWidget(QLabel("指令內容（可讀模式）:"))
        self.list_steps = QListWidget()
        self.list_steps.setSelectionMode(QListWidget.SelectionMode.SingleSelection)
        layout.addWidget(self.list_steps)

        # 按鈕
        btn_layout = QHBoxLayout()
        self.btn_cancel = QPushButton("取消 / 返回")
        self.btn_cancel.clicked.connect(self.on_back_callback)
        
        self.btn_save = QPushButton("儲存腳本")
        self.btn_save.setStyleSheet("background-color: #2ecc71; color: white; font-weight: bold;")
        self.btn_save.clicked.connect(self.save_script)
        
        btn_layout.addWidget(self.btn_cancel)
        btn_layout.addWidget(self.btn_save)
        layout.addLayout(btn_layout)
        
        right = QWidget()
        right.setLayout(layout)
        root.addWidget(right, 1)
        self.setLayout(root)

    def prepare_new(self):
        """清空介面以新增腳本"""
        self.name_input.setReadOnly(False)
        self.name_input.clear()
        self.actions = [
            {"action": "focus", "title": "Notepad"},
            {"action": "move", "x": 500, "y": 500, "duration_ms": 200},
            {"action": "click", "btn": "left", "clicks": 1},
            {"action": "type", "text": "Hello", "interval_ms": 0},
            {"action": "hotkey", "keys": ["ctrl", "s"]},
            {"action": "wait", "ms": 500},
        ]
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
                self.name_input.setReadOnly(True) # 編輯時鎖定檔名
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
            return f"Mouse Move: x={act.get('x')} y={act.get('y')} duration={dur}ms"
        if t == "click":
            btn = act.get("btn", "left")
            clicks = act.get("clicks", 1)
            interval = act.get("interval_ms", 0)
            extra = f" interval={interval}ms" if interval else ""
            return f"Mouse Click: btn={btn} clicks={clicks}{extra}"
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

    def add_action_focus(self):
        title, ok = QInputDialog.getText(self, "視窗 Focus", "視窗標題包含文字 (title contains):")
        if not ok or not title.strip():
            return
        self._insert_action_after_cursor({"action": "focus", "title": title.strip()})

    def add_action_mouse_move(self):
        x, ok = QInputDialog.getInt(self, "滑鼠移動", "X:", 500, 0, 100000, 1)
        if not ok:
            return
        y, ok = QInputDialog.getInt(self, "滑鼠移動", "Y:", 500, 0, 100000, 1)
        if not ok:
            return
        duration_ms, ok = QInputDialog.getInt(self, "滑鼠移動", "移動耗時 (ms):", 200, 0, 600000, 10)
        if not ok:
            return
        action = {"action": "move", "x": int(x), "y": int(y)}
        if duration_ms:
            action["duration_ms"] = int(duration_ms)
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

    def save_script(self):
        name = self.name_input.text().strip()
        if not name:
            QMessageBox.warning(self, "提示", "請輸入名稱")
            return
        
        try:
            file_path = main.SCRIPTS_DIR / f"{name}.json"
            
            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(self.actions, f, indent=4, ensure_ascii=False)
            
            QMessageBox.information(self, "成功", "腳本已存檔")
            self.on_back_callback()
        except Exception as e:
            QMessageBox.critical(self, "儲存失敗", str(e))