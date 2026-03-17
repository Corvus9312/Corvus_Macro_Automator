import time
import pyautogui
from PyQt6.QtCore import QThread, pyqtSignal

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
                
                if act_type == "move":
                    pyautogui.moveTo(action['x'], action['y'], duration=0.2)
                elif act_type == "click":
                    button = action.get("btn", "left")
                    pyautogui.click(button=button)
                elif act_type == "wait":
                    time.sleep(action.get("ms", 500) / 1000)
                elif act_type == "key":
                    pyautogui.press(action['key'])
                
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