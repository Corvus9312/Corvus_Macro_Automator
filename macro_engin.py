import time
from datetime import datetime, time as dtime

import pyautogui
from PyQt6.QtCore import QThread, pyqtSignal

try:
    import pygetwindow as gw
except Exception:  # pragma: no cover
    gw = None

class MacroEngine(QThread):
    finished = pyqtSignal()
    status_update = pyqtSignal(str)

    def __init__(self, script_data, config=None):
        super().__init__()
        self.script_data = script_data
        self.config = config or {}
        self.is_running = True
        pyautogui.FAILSAFE = True  # 啟用安全機制

    def _run_one_round(self):
        """執行一輪腳本，回傳 True 表示正常跑完，False 表示被中斷"""
        def resolve_point(ep):
            """支援固定/比例/圖片比對，回傳 (x, y)"""
            if not isinstance(ep, dict):
                raise ValueError("endpoint must be dict")
            if "image" in ep:
                image_path = ep["image"]
                try:
                    pos = pyautogui.locateCenterOnScreen(image_path)
                except Exception:
                    pos = None
                if pos is None:
                    raise RuntimeError(f"圖片比對失敗，找不到: {image_path}")
                return int(pos.x), int(pos.y)
            if ep.get("ratio"):
                w, h = pyautogui.size()
                return int(round(ep["x"] * w)), int(round(ep["y"] * h))
            return int(ep["x"]), int(ep["y"])

        for action in self.script_data:
            if not self.is_running:
                return False
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
                duration = max(duration_ms, 0) / 1000
                if "image" in action:
                    image_path = action["image"]
                    try:
                        pos = pyautogui.locateCenterOnScreen(image_path)
                    except Exception:
                        pos = None
                    if pos is None:
                        raise RuntimeError(f"圖片比對失敗，找不到: {image_path}")
                    pyautogui.moveTo(pos.x, pos.y, duration=duration)
                elif action.get("ratio"):
                    w, h = pyautogui.size()
                    x = int(round(action["x"] * w))
                    y = int(round(action["y"] * h))
                    pyautogui.moveTo(x, y, duration=duration)
                else:
                    pyautogui.moveTo(action["x"], action["y"], duration=duration)
            elif act_type == "click":
                button = action.get("btn", "left")
                clicks = int(action.get("clicks", 1))
                interval_ms = int(action.get("interval_ms", 0))
                pyautogui.click(button=button, clicks=clicks, interval=max(interval_ms, 0) / 1000)
            elif act_type == "drag":
                button = action.get("btn", "left")
                duration_ms = int(action.get("duration_ms", 200))
                duration = max(duration_ms, 0) / 1000
                frm = action.get("from", {})
                to = action.get("to", {})
                fx, fy = resolve_point(frm)
                tx, ty = resolve_point(to)
                pyautogui.moveTo(fx, fy, duration=0)
                pyautogui.mouseDown(button=button)
                try:
                    pyautogui.moveTo(tx, ty, duration=duration)
                finally:
                    pyautogui.mouseUp(button=button)
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
        return True

    def run(self):
        mode = self.config.get("repeat_mode", "once")
        count = int(self.config.get("count", 1))
        duration_seconds = int(self.config.get("duration_seconds", 0))
        until_time_str = self.config.get("until_time")  # "HH:MM"

        try:
            self.status_update.emit("腳本開始執行...")
            run_count = 0

            if mode == "once":
                self._run_one_round()
            elif mode == "count":
                for _ in range(count):
                    if not self.is_running:
                        break
                    run_count += 1
                    self.status_update.emit(f"第 {run_count}/{count} 輪")
                    if not self._run_one_round():
                        break
            elif mode == "duration":
                end_at = time.time() + duration_seconds
                while time.time() < end_at and self.is_running:
                    run_count += 1
                    self.status_update.emit(f"第 {run_count} 輪（剩餘 {int(end_at - time.time())} 秒）")
                    if not self._run_one_round():
                        break
            elif mode == "until_time":
                target = None
                if until_time_str:
                    try:
                        parts = until_time_str.strip().split(":")
                        h, m = int(parts[0]), int(parts[1]) if len(parts) > 1 else 0
                        target = dtime(hour=h, minute=m, second=0)
                    except Exception:
                        target = dtime(hour=17, minute=0, second=0)
                if target is not None:
                    while self.is_running:
                        now = datetime.now().time()
                        if now >= target:
                            break
                        run_count += 1
                        self.status_update.emit(f"第 {run_count} 輪（執行到 {until_time_str}）")
                        if not self._run_one_round():
                            break
                        time.sleep(1)
            else:  # manual
                while self.is_running:
                    run_count += 1
                    self.status_update.emit(f"第 {run_count} 輪（手動停止為止）")
                    if not self._run_one_round():
                        break

            self.status_update.emit("腳本執行完畢")
        except pyautogui.FailSafeException:
            self.status_update.emit("安全機制觸發：停止執行")
        except Exception as e:
            self.status_update.emit(f"錯誤: {str(e)}")
        finally:
            self.finished.emit()

    def stop(self):
        self.is_running = False