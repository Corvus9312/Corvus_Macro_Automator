import os
import sys
import main
from PyQt6.QtWidgets import (QApplication, QMainWindow, QStackedWidget, 
                             QWidget, QVBoxLayout, QPushButton, QLabel, 
                             QTextEdit, QHBoxLayout, QListWidget, QMessageBox)

class ScriptListWindow(QWidget):
    def __init__(self, on_add_callback, on_edit_callback, on_run_callback, on_stop_callback):
        super().__init__()
        self.on_add_callback = on_add_callback
        self.on_edit_callback = on_edit_callback
        self.on_run_callback = on_run_callback
        self.on_stop_callback = on_stop_callback
        self.is_running = False
        
        layout = QVBoxLayout()
        
        # 標題
        title = QLabel("📋 Corvus 腳本管理員")
        title.setObjectName("pageTitle")
        layout.addWidget(title)

        # 列表區域
        self.list_widget = QListWidget()
        layout.addWidget(self.list_widget)

        # 按鈕區域
        btn_layout = QHBoxLayout()
        self.btn_add = QPushButton("新增腳本")
        self.btn_add.clicked.connect(self.on_add_callback)
        
        self.btn_edit = QPushButton("編輯選中項目")
        self.btn_edit.clicked.connect(self.handle_edit)
        
        self.btn_refresh = QPushButton("整理列表")
        self.btn_refresh.clicked.connect(self.refresh_list)

        self.btn_run = QPushButton("▶ 執行腳本")
        self.btn_run.setObjectName("primaryRun")
        self.btn_run.clicked.connect(self.handle_run)

        self.btn_stop = QPushButton("⏹ 停止")
        self.btn_stop.setEnabled(False) # 初始禁用
        self.btn_stop.setObjectName("dangerStop")
        self.btn_stop.clicked.connect(self.handle_stop)

        btn_layout.addWidget(self.btn_add)
        btn_layout.addWidget(self.btn_edit)
        btn_layout.addWidget(self.btn_refresh)
        btn_layout.addWidget(self.btn_run)
        btn_layout.addWidget(self.btn_stop)
        layout.addLayout(btn_layout)
        
        self.setLayout(layout)
        self.refresh_list()
        self.apply_theme()

    def apply_theme(self):
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
                QLabel#pageTitle { font-size: 18px; font-weight: 800; margin: 10px; }
                QListWidget {
                    background-color: #2a2a2a;
                    color: #f2f2f2;
                    border: 1px solid #3a3a3a;
                    border-radius: 6px;
                }
                QListWidget::item:selected { background-color: #3a3a3a; }
                QPushButton {
                    background-color: #2f2f2f;
                    color: #f2f2f2;
                    border: 1px solid #444;
                    border-radius: 6px;
                    padding: 8px 10px;
                }
                QPushButton:hover { background-color: #3a3a3a; }
                QPushButton:disabled { color: #888; border-color: #333; background-color: #262626; }
                QPushButton#primaryRun { background-color: #3498db; color: white; border: none; font-weight: 700; }
                QPushButton#primaryRun:hover { background-color: #2f89c6; }
                QPushButton#dangerStop { background-color: #e74c3c; color: white; border: none; font-weight: 700; }
                QPushButton#dangerStop:hover { background-color: #cf3f30; }
                """
            )
        else:
            self.setStyleSheet(
                """
                QLabel { color: #222; }
                QLabel#pageTitle { font-size: 18px; font-weight: 800; margin: 10px; }
                QListWidget {
                    background-color: white;
                    color: #111;
                    border: 1px solid #cfcfcf;
                    border-radius: 6px;
                }
                QListWidget::item:selected { background-color: #e8f0ff; }
                QPushButton {
                    background-color: #f2f2f2;
                    color: #111;
                    border: 1px solid #cfcfcf;
                    border-radius: 6px;
                    padding: 8px 10px;
                }
                QPushButton:hover { background-color: #eaeaea; }
                QPushButton:disabled { color: #888; background-color: #f3f3f3; }
                QPushButton#primaryRun { background-color: #3498db; color: white; border: none; font-weight: 700; }
                QPushButton#primaryRun:hover { background-color: #2f89c6; }
                QPushButton#dangerStop { background-color: #e74c3c; color: white; border: none; font-weight: 700; }
                QPushButton#dangerStop:hover { background-color: #cf3f30; }
                """
            )

    def refresh_list(self):
        """讀取資料夾內所有 JSON 檔案"""
        self.list_widget.clear()
        for file in main.SCRIPTS_DIR.glob("*.json"):
            self.list_widget.addItem(file.name)

    def handle_run(self):
        selected = self.list_widget.currentItem()
        if selected:
            self.btn_run.setEnabled(False)
            self.btn_stop.setEnabled(True)
            self.is_running = True
            self.on_run_callback(selected.text())
        else:
            QMessageBox.warning(self, "提示", "請先選擇一個腳本")

    def handle_stop(self):
        """處理停止按鈕點擊事件"""
        if not self.is_running:
            return
        self.is_running = False
        self.btn_run.setEnabled(True)
        self.btn_stop.setEnabled(False)
        self.on_stop_callback()

    def handle_edit(self):
        selected = self.list_widget.currentItem()
        if selected:
            self.on_edit_callback(selected.text())
        else:
            QMessageBox.warning(self, "提示", "請先選擇一個腳本")