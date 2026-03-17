import os
import sys
import main
from PyQt6.QtWidgets import (QApplication, QMainWindow, QStackedWidget, 
                             QWidget, QVBoxLayout, QPushButton, QLabel, 
                             QTextEdit, QHBoxLayout, QListWidget, QMessageBox)

class ScriptListWindow(QWidget):
    def __init__(self, on_add_callback, on_edit_callback):
        super().__init__()
        self.on_add_callback = on_add_callback
        self.on_edit_callback = on_edit_callback
        
        layout = QVBoxLayout()
        
        # 標題
        title = QLabel("📋 Corvus 腳本管理員")
        title.setStyleSheet("font-size: 18px; font-weight: bold; margin: 10px;")
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
        self.btn_run.setStyleSheet("background-color: #3498db; color: white;")
        self.btn_run.clicked.connect(self.handle_run)

        self.btn_stop = QPushButton("⏹ 停止")
        self.btn_stop.setEnabled(False) # 初始禁用
        self.btn_stop.clicked.connect(self.handle_stop)

        btn_layout.addWidget(self.btn_add)
        btn_layout.addWidget(self.btn_edit)
        btn_layout.addWidget(self.btn_refresh)
        layout.addLayout(btn_layout)
        
        self.setLayout(layout)
        self.refresh_list()

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
            self.on_run_callback(selected.text())

    def handle_stop(self):
        """處理停止按鈕點擊事件"""
        if self.is_running:
            self.is_running = False
            
            self.btn_run.setEnabled(True)
            self.btn_stop.setEnabled(False)
            
            print("Stop button clicked: Signal sent to terminate tasks.")
        else:
            print("No active task to stop.")

    def handle_edit(self):
        selected = self.list_widget.currentItem()
        if selected:
            self.on_edit_callback(selected.text())
        else:
            QMessageBox.warning(self, "提示", "請先選擇一個腳本")