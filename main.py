import os
import sys

# 避免 Windows 上 Qt 呼叫 SetProcessDpiAwarenessContext 導致「存取被拒」
if sys.platform == "win32":
    os.environ.setdefault("QT_QPA_PLATFORM", "windows:dpiawareness=0")
    # 若系統/權限仍不允許設定 DPI awareness，Qt 會輸出警告訊息但通常不影響功能。
    # 這裡只關掉該分類的警告，避免一直刷 console。
    os.environ.setdefault("QT_LOGGING_RULES", "qt.qpa.window.warning=false")
    # 另外關閉自動縮放，避免 Qt 再嘗試調整 DPI 相關設定
    os.environ.setdefault("QT_ENABLE_HIGHDPI_SCALING", "0")
    os.environ.setdefault("QT_AUTO_SCREEN_SCALE_FACTOR", "0")

import json
from pathlib import Path
from pages import main_page, edit_page
from macro_engin import MacroEngine
from PyQt6.QtWidgets import (QApplication, QMainWindow, QStackedWidget, QMessageBox)
from PyQt6.QtGui import QPalette
from PyQt6.QtCore import Qt, QByteArray
from PyQt6.QtNetwork import QLocalServer, QLocalSocket

try:
    import keyboard  # type: ignore
except Exception:
    keyboard = None

APP_DATA_ROOT = Path(os.environ['APPDATA']) / 'Corvus_Macro_Automator'
SCRIPTS_DIR = APP_DATA_ROOT / 'scripts'
SCRIPTS_CONFIG_FILE = APP_DATA_ROOT / 'scripts_config.json'

SCRIPTS_DIR.mkdir(parents=True, exist_ok=True)

# 重複執行設定：repeat_mode = "once" | "count" | "duration" | "until_time" | "manual"
def get_script_config(script_stem):
    """取得腳本重複執行設定（檔名不含副檔名）"""
    if not SCRIPTS_CONFIG_FILE.exists():
        return {"repeat_mode": "once"}
    try:
        with open(SCRIPTS_CONFIG_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        return data.get(script_stem, {"repeat_mode": "once"})
    except Exception:
        return {"repeat_mode": "once"}

def set_script_config(script_stem, config):
    """儲存腳本重複執行設定"""
    data = {}
    if SCRIPTS_CONFIG_FILE.exists():
        try:
            with open(SCRIPTS_CONFIG_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
        except Exception:
            pass
    data[script_stem] = {k: v for k, v in config.items() if v is not None}
    with open(SCRIPTS_CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)

class CorvusMacroAutomator(QMainWindow):
    """主視窗控制中心"""
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Corvus Macro Automator")
        self.resize(875, 625)

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

            script_stem = file_path.stem
            config = get_script_config(script_stem)

            # 建立並啟動執行緒
            self.worker = MacroEngine(script_data, config)
            self.worker.status_update.connect(lambda msg: self.statusBar().showMessage(msg))
            self.worker.finished.connect(self.on_worker_finished)
            self.worker.start()

        except Exception as e:
            QMessageBox.critical(self, "錯誤", f"無法執行腳本: {e}")

    def on_worker_finished(self):
        self.page_list.btn_run.setEnabled(True)
        self.page_list.btn_stop.setEnabled(False)
        self.statusBar().showMessage("準備就緒")
        # 腳本若曾 focus 其他視窗，結束時把焦點帶回本程式
        self.activateWindow()
        self.raise_()

    def bring_to_front(self):
        """被第二次啟動時呼叫：把視窗顯示到最前面"""
        try:
            if self.isMinimized():
                self.showNormal()
            else:
                self.show()
            self.raise_()
            self.activateWindow()
        except Exception:
            pass

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
            stem = selected.data(Qt.ItemDataRole.UserRole) or selected.text()
            filename = stem + ".json"
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

    # 單一執行個體：若已在執行，通知既有實例顯示並結束本次啟動
    server_name = "Corvus_Macro_Automator_single_instance"
    socket = QLocalSocket()
    socket.connectToServer(server_name)
    if socket.waitForConnected(200):
        socket.write(QByteArray(b"raise"))
        socket.flush()
        socket.waitForBytesWritten(200)
        socket.disconnectFromServer()
        sys.exit(0)

    # 沒有既有實例：建立 server 供後續啟動通知
    server = QLocalServer()
    try:
        QLocalServer.removeServer(server_name)  # 避免上次非正常退出留下的殘留
    except Exception:
        pass
    server.listen(server_name)

    window = CorvusMacroAutomator()
    window.show()

    def _on_new_connection():
        conn = server.nextPendingConnection()
        if conn:
            conn.readyRead.connect(lambda: (conn.readAll(), window.bring_to_front(), conn.disconnectFromServer()))

    server.newConnection.connect(_on_new_connection)
    sys.exit(app.exec())