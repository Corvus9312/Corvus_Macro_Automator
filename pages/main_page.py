import main
from PyQt6.QtCore import QTime, Qt
from PyQt6.QtWidgets import (
    QWidget,
    QVBoxLayout,
    QPushButton,
    QLabel,
    QHBoxLayout,
    QListWidget,
    QListWidgetItem,
    QMessageBox,
    QDialog,
    QDialogButtonBox,
    QFormLayout,
    QComboBox,
    QSpinBox,
    QTimeEdit,
)

REPEAT_MODES = [
    ("執行一次", "once"),
    ("按次數", "count"),
    ("按時長（秒）", "duration"),
    ("執行到指定時間", "until_time"),
    ("執行到自行取消", "manual"),
]


class ScriptConfigDialog(QDialog):
    """腳本重複執行設定對話框"""
    def __init__(self, script_stem, parent=None):
        super().__init__(parent)
        self.script_stem = script_stem
        self.setWindowTitle(f"執行設定：{script_stem}")
        layout = QVBoxLayout(self)
        form = QFormLayout()

        self.mode_combo = QComboBox()
        for label, _ in REPEAT_MODES:
            self.mode_combo.addItem(label)
        self.mode_combo.currentIndexChanged.connect(self._on_mode_changed)
        form.addRow("執行模式:", self.mode_combo)

        self.count_spin = QSpinBox()
        self.count_spin.setRange(1, 999999)
        self.count_spin.setValue(5)
        self.count_spin.setSuffix(" 次")
        form.addRow("次數:", self.count_spin)

        self.duration_spin = QSpinBox()
        self.duration_spin.setRange(1, 86400)
        self.duration_spin.setValue(3600)
        self.duration_spin.setSuffix(" 秒")
        form.addRow("時長:", self.duration_spin)

        self.until_time_edit = QTimeEdit()
        self.until_time_edit.setDisplayFormat("HH:mm")
        self.until_time_edit.setTime(QTime(17, 0))
        form.addRow("執行到時間:", self.until_time_edit)

        layout.addLayout(form)
        bb = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        bb.accepted.connect(self.accept)
        bb.rejected.connect(self.reject)
        layout.addWidget(bb)

        self._on_mode_changed()

    def _on_mode_changed(self):
        idx = self.mode_combo.currentIndex()
        _, mode = REPEAT_MODES[idx]
        self.count_spin.setVisible(mode == "count")
        self.duration_spin.setVisible(mode == "duration")
        self.until_time_edit.setVisible(mode == "until_time")

    def get_config(self):
        mode = REPEAT_MODES[self.mode_combo.currentIndex()][1]
        cfg = {"repeat_mode": mode}
        if mode == "count":
            cfg["count"] = self.count_spin.value()
        elif mode == "duration":
            cfg["duration_seconds"] = self.duration_spin.value()
        elif mode == "until_time":
            cfg["until_time"] = self.until_time_edit.time().toString("HH:mm")
        return cfg

    def set_config(self, cfg):
        mode = cfg.get("repeat_mode", "once")
        for i, (_, m) in enumerate(REPEAT_MODES):
            if m == mode:
                self.mode_combo.setCurrentIndex(i)
                break
        self.count_spin.setValue(int(cfg.get("count", 5)))
        self.duration_spin.setValue(int(cfg.get("duration_seconds", 3600)))
        until = cfg.get("until_time", "17:00")
        try:
            parts = until.split(":")
            h, m = int(parts[0]), int(parts[1]) if len(parts) > 1 else 0
            self.until_time_edit.setTime(QTime(h, m))
        except Exception:
            self.until_time_edit.setTime(QTime(17, 0))
        self._on_mode_changed()

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

        # 執行 / 停止（清單上方）
        run_stop_layout = QHBoxLayout()
        self.btn_run = QPushButton("▶ 執行腳本")
        self.btn_run.setObjectName("primaryRun")
        self.btn_run.clicked.connect(self.handle_run)
        self.btn_stop = QPushButton("⏹ 停止")
        self.btn_stop.setEnabled(False)
        self.btn_stop.setObjectName("dangerStop")
        self.btn_stop.clicked.connect(self.handle_stop)
        run_stop_layout.addWidget(self.btn_run)
        run_stop_layout.addWidget(self.btn_stop)
        run_stop_layout.addStretch(1)
        layout.addLayout(run_stop_layout)

        # 列表區域（快擊兩下可編輯）
        self.list_widget = QListWidget()
        self.list_widget.itemDoubleClicked.connect(self._on_script_double_clicked)
        layout.addWidget(self.list_widget)

        # 下方按鈕區域
        btn_layout = QHBoxLayout()
        self.btn_add = QPushButton("新增腳本")
        self.btn_add.clicked.connect(self.on_add_callback)
        self.btn_edit = QPushButton("編輯選中項目")
        self.btn_edit.clicked.connect(self.handle_edit)
        self.btn_refresh = QPushButton("整理列表")
        self.btn_refresh.clicked.connect(self.refresh_list)
        self.btn_settings = QPushButton("設定")
        self.btn_settings.clicked.connect(self.handle_settings)
        self.btn_delete = QPushButton("刪除腳本")
        self.btn_delete.setObjectName("dangerDelete")
        self.btn_delete.clicked.connect(self.handle_delete)

        btn_layout.addWidget(self.btn_add)
        btn_layout.addWidget(self.btn_edit)
        btn_layout.addWidget(self.btn_settings)
        btn_layout.addWidget(self.btn_refresh)
        btn_layout.addWidget(self.btn_delete)
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
                    outline: none;
                }
                QListWidget:focus { outline: none; }
                QListWidget::item:focus { outline: none; }
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
                QPushButton#dangerDelete { background-color: #c0392b; color: white; border: none; }
                QPushButton#dangerDelete:hover { background-color: #a93226; }
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
                    outline: none;
                }
                QListWidget:focus { outline: none; }
                QListWidget::item:focus { outline: none; }
                /* 明確指定選取狀態的文字顏色，避免系統主題導致反白後看不到字 */
                QListWidget::item:selected { background-color: #cfe2ff; color: #111; }
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
                QPushButton#dangerDelete { background-color: #c0392b; color: white; border: none; }
                QPushButton#dangerDelete:hover { background-color: #a93226; }
                """
            )

    def refresh_list(self):
        """讀取資料夾內所有 JSON 檔案（每個腳本各自設定，並顯示設定摘要）"""
        selected_stem = self._selected_stem()
        self.list_widget.clear()
        for file in sorted(main.SCRIPTS_DIR.glob("*.json")):
            stem = file.stem
            cfg = main.get_script_config(stem)
            summary = self._format_repeat_summary(cfg)
            item = QListWidgetItem(f"{stem}  [{summary}]")
            item.setData(Qt.ItemDataRole.UserRole, stem)
            self.list_widget.addItem(item)
            if selected_stem == stem:
                self.list_widget.setCurrentItem(item)

    def _selected_stem(self):
        item = self.list_widget.currentItem()
        if not item:
            return None
        stem = item.data(Qt.ItemDataRole.UserRole)
        return stem or item.text()

    def _format_repeat_summary(self, cfg):
        mode = (cfg or {}).get("repeat_mode", "once")
        if mode == "count":
            return f"次數 x{int(cfg.get('count', 1))}"
        if mode == "duration":
            return f"時長 {int(cfg.get('duration_seconds', 0))}s"
        if mode == "until_time":
            return f"到 {cfg.get('until_time', '17:00')}"
        if mode == "manual":
            return "手動停止"
        return "一次"

    def _selected_filename(self):
        """目前選取項目的完整檔名（含 .json），若無選取則回傳 None"""
        stem = self._selected_stem()
        return (stem + ".json") if stem else None

    def _on_script_double_clicked(self, item):
        """腳本清單快擊兩下 → 編輯該腳本"""
        if item:
            stem = item.data(Qt.ItemDataRole.UserRole) or item.text()
            self.on_edit_callback(stem + ".json")

    def handle_run(self):
        filename = self._selected_filename()
        if filename:
            self.btn_run.setEnabled(False)
            self.btn_stop.setEnabled(True)
            self.is_running = True
            self.on_run_callback(filename)
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
        filename = self._selected_filename()
        if filename:
            self.on_edit_callback(filename)
        else:
            QMessageBox.warning(self, "提示", "請先選擇一個腳本")

    def handle_settings(self):
        item = self.list_widget.currentItem()
        if not item:
            QMessageBox.warning(self, "提示", "請先選擇一個腳本")
            return
        script_stem = item.data(Qt.ItemDataRole.UserRole) or item.text()
        dlg = ScriptConfigDialog(script_stem, self)
        cfg = main.get_script_config(script_stem)
        dlg.set_config(cfg)
        if dlg.exec() == QDialog.DialogCode.Accepted:
            main.set_script_config(script_stem, dlg.get_config())
            QMessageBox.information(self, "完成", "執行設定已儲存")
            self.refresh_list()

    def handle_delete(self):
        filename = self._selected_filename()
        if not filename:
            QMessageBox.warning(self, "提示", "請先選擇要刪除的腳本")
            return
        display_name = filename.replace(".json", "", 1)
        reply = QMessageBox.question(
            self,
            "確認刪除",
            f"確定要刪除腳本「{display_name}」嗎？\n此操作無法復原。",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if reply != QMessageBox.StandardButton.Yes:
            return
        try:
            file_path = main.SCRIPTS_DIR / filename
            if file_path.exists():
                file_path.unlink()
                self.refresh_list()
                QMessageBox.information(self, "完成", "腳本已刪除")
            else:
                QMessageBox.warning(self, "提示", "檔案不存在，請整理列表後再試")
        except Exception as e:
            QMessageBox.critical(self, "錯誤", f"刪除失敗：{e}")