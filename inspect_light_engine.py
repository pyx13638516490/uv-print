# inspect_light_engine.py
import time
from pywinauto.application import Application

print("--- 光機軟體 UI 結構偵測工具 ---")

try:
    # 1. 請先手動打開 "Full-HD UV LE Controller v2.1.exe"
    print("請確保光機控制軟體已經手動打開...")
    time.sleep(2)

    # 2. 連接到軟體
    window_title = "Full-HD UV LE Controller v2.1"
    print(f"正在嘗試連接到視窗: '{window_title}'")
    app = Application(backend="uia").connect(title=window_title, timeout=20)
    main_win = app.window(title=window_title)
    print("連接成功！")

    # 3. 打印出所有可用的控制項資訊
    print("\n--- 視窗內所有控制項的詳細資訊如下 ---")
    main_win.print_control_identifiers(depth=4)
    print("\n--- 偵測完畢 ---")

except Exception as e:
    print(f"\n發生錯誤: {e}")
    print("請確認軟體是否已打開，且視窗標題完全符合。")

input("\n按 Enter 鍵結束...")