import time
import pyautogui
from PyQt6.QtCore import QThread, pyqtSignal

try:
    import pygetwindow as gw
except Exception:  # pragma: no cover
    gw = None

class MacroEngine(QThread):
    finished = pyqtSignal()
    status_update = pyqtSignal(str)

    def __init__(self, script_data):
        super().__init__()
        self.script_data = script_data
        self.is_running = True
        pyautogui.FAILSAFE = True # 啟用安全機制

    def run(self):
        try:
            self.status_update.emit("腳本開始執行...")
            # 這裡可以根據需求決定是否要迴圈執行
            for action in self.script_data:
                if not self.is_running:
                    break
                
                act_type = action.get("action")
                
                if act_type == "focus":
                    title = action.get("title", "")
                    if not title:
                        raise ValueError("focus action requires 'title'")
                    if gw is None:
                        raise RuntimeError("pygetwindow 未安裝，無法使用 focus")
                    wins = gw.getWindowsWithTitle(title)
                    if not wins:
                        raise RuntimeError(f"找不到視窗: {title}")
                    wins[0].activate()
                elif act_type == "move":
                    duration_ms = int(action.get("duration_ms", 200))
                    pyautogui.moveTo(action["x"], action["y"], duration=max(duration_ms, 0) / 1000)
                elif act_type == "click":
                    button = action.get("btn", "left")
                    clicks = int(action.get("clicks", 1))
                    interval_ms = int(action.get("interval_ms", 0))
                    pyautogui.click(button=button, clicks=clicks, interval=max(interval_ms, 0) / 1000)
                elif act_type == "wait":
                    time.sleep(action.get("ms", 500) / 1000)
                elif act_type == "type":
                    text = action.get("text", "")
                    interval_ms = int(action.get("interval_ms", 0))
                    pyautogui.write(text, interval=max(interval_ms, 0) / 1000)
                elif act_type == "hotkey":
                    keys = action.get("keys")
                    if not isinstance(keys, list) or not keys:
                        raise ValueError("hotkey action requires 'keys' as non-empty list")
                    pyautogui.hotkey(*keys)
                elif act_type == "key":
                    pyautogui.press(action["key"])
                
                self.status_update.emit(f"執行動作: {act_type}")
            
            self.status_update.emit("腳本執行完畢")
        except pyautogui.FailSafeException:
            self.status_update.emit("安全機制觸發：停止執行")
        except Exception as e:
            self.status_update.emit(f"錯誤: {str(e)}")
        finally:
            self.finished.emit()

    def stop(self):
        self.is_running = False