import os
import sys
import json
from pathlib import Path
from pages import main_page, edit_page
from macro_engin import MacroEngine
from PyQt6.QtWidgets import (QApplication, QMainWindow, QStackedWidget, QMessageBox)
from PyQt6.QtGui import QPalette

try:
    import keyboard  # type: ignore
except Exception:
    keyboard = None

APP_DATA_ROOT = Path(os.environ['APPDATA']) / 'Corvus_Macro_Automator'
SCRIPTS_DIR = APP_DATA_ROOT / 'scripts'

SCRIPTS_DIR.mkdir(parents=True, exist_ok=True)

class CorvusMacroAutomator(QMainWindow):
    """主視窗控制中心"""
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Corvus Macro Automator")
        self.resize(700, 500)

        self.stack = QStackedWidget()
        self.setCentralWidget(self.stack)

        self.page_list = main_page.ScriptListWindow(
            self.go_to_add,
            self.go_to_edit,
            self.execute_script,
            self.stop_execution,
        )
        self.page_editor = edit_page.EditorWindow(self.back_to_list)

        self.stack.addWidget(self.page_list)   # Index 0
        self.stack.addWidget(self.page_editor) # Index 1

        self.worker = None

        self.init_hotkeys()
        self.apply_theme_to_pages()

    def apply_theme_to_pages(self):
        """
        依據目前 Qt 調色盤是否偏深色，讓各頁面套用一致的可視樣式。
        主要目的：避免深色模式下按鈕/文字對比不足。
        """
        try:
            pal = self.palette()
            window = pal.color(QPalette.ColorRole.Window)
            is_dark = window.lightness() < 128
        except Exception:
            is_dark = False

        for page in (self.page_list, self.page_editor):
            apply_fn = getattr(page, "apply_theme", None)
            if callable(apply_fn):
                apply_fn()

    def go_to_add(self):
        self.page_editor.prepare_new()
        self.stack.setCurrentIndex(1)

    def go_to_edit(self, filename):
        self.page_editor.load_existing(filename)
        self.stack.setCurrentIndex(1)

    def back_to_list(self):
        self.page_list.refresh_list()
        self.stack.setCurrentIndex(0)

    def execute_script(self, filename):
        file_path = SCRIPTS_DIR / filename
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                script_data = json.load(f)
            
            # 防止重複執行
            if self.worker and self.worker.isRunning():
                return

            # 切換按鈕狀態 (這裡建議透過訊號或直接存取頁面元件)
            self.page_list.btn_run.setEnabled(False)
            self.page_list.btn_stop.setEnabled(True)

            # 建立並啟動執行緒
            self.worker = MacroEngine(script_data)
            self.worker.status_update.connect(lambda msg: self.statusBar().showMessage(msg))
            self.worker.finished.connect(self.on_worker_finished)
            self.worker.start()

        except Exception as e:
            QMessageBox.critical(self, "錯誤", f"無法執行腳本: {e}")

    def on_worker_finished(self):
        self.page_list.btn_run.setEnabled(True)
        self.page_list.btn_stop.setEnabled(False)
        self.statusBar().showMessage("準備就緒")

    def stop_execution(self):
        if self.worker:
            self.worker.stop()

    def init_hotkeys(self):
        """設定全域熱鍵 F5 與 F6"""
        if keyboard is None:
            self.statusBar().showMessage("未安裝 keyboard：略過全域熱鍵 (F5/F6)")
            return
        keyboard.add_hotkey('f5', self.trigger_run_by_hotkey)
        keyboard.add_hotkey('f6', self.stop_execution)
        self.statusBar().showMessage("熱鍵已啟動: F5 執行 / F6 結束")

    def trigger_run_by_hotkey(self):
        """當按下 F5 時觸發"""
        selected = self.page_list.list_widget.currentItem()
        if selected:
            filename = selected.text()
            print(f"熱鍵啟動腳本: {filename}")
            self.execute_script(filename)
        else:
            self.statusBar().showMessage("錯誤: 請先用滑鼠選中一個腳本再按 F5")

    def stop_execution(self):
        """當按下 F6 或手動停止時觸發"""
        if self.worker and self.worker.isRunning():
            self.worker.stop()
            self.statusBar().showMessage("🛑 腳本已強制停止 (F6)")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = CorvusMacroAutomator()
    window.show()
    sys.exit(app.exec())