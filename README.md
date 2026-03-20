# Corvus Macro Automator

使用 `uv` 建立的 `PyQt6` 按鍵精靈專案，可錄製與播放鍵盤/滑鼠事件。

## 功能

- 錄製全域鍵盤事件（按下/放開）
- 錄製滑鼠點擊與滾輪事件
- 以 JSON 儲存與載入巨集
- 可設定播放次數、開始延遲與播放速度

## 使用方式

1. 安裝依賴：

   ```bash
   uv sync
   ```

2. 執行程式：

   ```bash
   uv run python main.py
   ```

   或使用 script：

   ```bash
   uv run corvus-macro
   ```

## 注意事項

- 錄製與播放為全域輸入控制，請在安全環境下使用。
- 播放前請先切到目標視窗，避免誤操作。
