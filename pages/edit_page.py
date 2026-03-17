import main
import json
from PyQt6.QtWidgets import (QApplication, QMainWindow, QStackedWidget, 
                             QWidget, QVBoxLayout, QPushButton, QLabel, 
                             QTextEdit, QHBoxLayout, QListWidget, QLineEdit, QMessageBox)
from PySide6.QtCore import Qt
                             
class EditorWindow(QWidget):
    def __init__(self, on_back_callback):
        super().__init__()
        self.on_back_callback = on_back_callback
        
        layout = QVBoxLayout()

        # 檔名輸入
        layout.addWidget(QLabel("腳本名稱 (不需副檔名):"))
        self.name_input = QLineEdit()
        layout.addWidget(self.name_input)

        # 內容編輯
        layout.addWidget(QLabel("指令內容 (JSON 格式):"))
        self.text_edit = QTextEdit()
        layout.addWidget(self.text_edit)

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
        
        self.setLayout(layout)

    def prepare_new(self):
        """清空介面以新增腳本"""
        self.name_input.setReadOnly(False)
        self.name_input.clear()
        default_data = [{"action": "move", "x": 500, "y": 500}, {"action": "click", "btn": "left"}]
        self.text_edit.setPlainText(json.dumps(default_data, indent=4, ensure_ascii=False))

    def load_existing(self, filename):
        """載入舊腳本內容"""
        file_path = main.SCRIPTS_DIR / filename
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
                self.text_edit.setPlainText(json.dumps(data, indent=4, ensure_ascii=False))
                self.name_input.setText(filename.replace(".json", ""))
                self.name_input.setReadOnly(True) # 編輯時鎖定檔名
        except Exception as e:
            QMessageBox.critical(self, "錯誤", f"讀取失敗: {e}")

    def save_script(self):
        name = self.name_input.text().strip()
        if not name:
            QMessageBox.warning(self, "提示", "請輸入名稱")
            return
        
        try:
            # 檢查 JSON 語法
            content = json.loads(self.text_edit.toPlainText())
            file_path = main.SCRIPTS_DIR / f"{name}.json"
            
            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(content, f, indent=4, ensure_ascii=False)
            
            QMessageBox.information(self, "成功", "腳本已存檔")
            self.on_back_callback()
        except json.JSONDecodeError:
            QMessageBox.critical(self, "格式錯誤", "請確保內容符合 JSON 語法")