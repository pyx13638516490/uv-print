# main_controller.py - PC端主控制程式 (最終混合模式穩定版)
# 版本日期: 2025-09-14
# 流程: 腳本啟動軟體 -> 用戶手動設定光機 -> 腳本接管打印

import os
import time
import zipfile
import tkinter as tk
from PIL import Image, ImageTk
import socket
from pywinauto.application import Application
from screeninfo import get_monitors
import subprocess


# --- 1. 使用者設定區 ---
class PrintConfig:
    # 軟體與文件路徑
    CONTROLLER_EXE_PATH = "Full-HD UV LE Controller v2.1.exe"
    ZIP_FILE_PATH = "layers.zip"
    TEMP_EXTRACT_DIR = "temp_layers"

    # Z軸剝離運動參數
    PEEL_LIFT_DISTANCE = 5.05
    PEEL_RETURN_DISTANCE = 5.0

    # 曝光參數 (僅用於計算打印時間)
    NORMAL_EXPOSURE_TIME_S = 2.5
    FIRST_LAYER_EXPOSURE_TIME_S = 5
    TRANSITION_LAYERS = 5

    # 硬體連接設定
    ESP32_IP_ADDRESS = "10.10.17.187"  # 請修改為您 ESP32 的實際 IP
    ESP32_PORT = 8899

    # 投影儀螢幕索引 (0=主螢幕, 1=第二個螢幕, ...)
    PROJECTOR_MONITOR_INDEX = 1


# --- 2. 光機 GUI 自動化控制模組 (簡化版) ---
class LightEngineGUIControl:
    def __init__(self):
        """僅連接到已手動打開的軟體"""
        self.app = None
        self.main_win = None
        try:
            window_title = "Full-HD UV LE Controller v2.1"
            print(f"正在連接到已手動設定好的視窗: '{window_title}'...")
            self.app = Application(backend="uia").connect(title=window_title, timeout=60)
            self.main_win = self.app.window(title=window_title)
            self.main_win.wait('ready', timeout=30)
            self.led_combo = self.main_win.child_window(auto_id="ComboBoxLedEnable")
            self.set_led_onoff_button = self.main_win.child_window(auto_id="ButtonSetLedOnOff")
            print("成功連接到光機軟體，自動化已準備就緒。")
        except Exception as e:
            print(f"錯誤: 連接到控制軟體失敗。請確認您已手動打開並設定好軟體。 {e}")
            exit()

    def led_on(self):
        print("指令: 開啟曝光 (LED ON)")
        self.led_combo.select("On")
        self.set_led_onoff_button.click()
        time.sleep(0.1)

    def led_off(self):
        print("指令: 關閉曝光 (LED OFF)")
        self.led_combo.select("Off")
        self.set_led_onoff_button.click()
        time.sleep(0.1)

    def close(self):
        print("自動化打印流程已結束，請手動關閉光機控制軟體。")


# --- 3. 投影儀HDMI顯示模組 (螢幕索引版) ---
class ProjectorDisplay:
    def __init__(self, monitor_index):
        self.root = tk.Tk()
        monitors = sorted(get_monitors(), key=lambda m: m.x)
        print(f"偵測到 {len(monitors)} 個螢幕，並已排序。")
        target_monitor = None
        if len(monitors) > monitor_index:
            target_monitor = monitors[monitor_index]
            print(f"已選擇索引為 {monitor_index} 的螢幕作為投影目標。")
        else:
            print(f"警告: 找不到索引為 {monitor_index} 的螢幕，將在主螢幕上顯示。")
            target_monitor = monitors[0]
        print(
            f"將在螢幕上顯示: {target_monitor.width}x{target_monitor.height} at ({target_monitor.x}, {target_monitor.y})")
        is_primary = (target_monitor.x == 0 and target_monitor.y == 0)
        if is_primary and len(monitors) > 1:
            print("警告: 選擇的投影目標為主螢幕，將以小視窗模式顯示以避免鎖定操作。")
            self.root.geometry(f"800x600+{target_monitor.x + 100}+{target_monitor.y + 100}")
            self.root.title("投影預覽 (主螢幕模式)")
        else:
            geometry_str = f"{target_monitor.width}x{target_monitor.height}+{target_monitor.x}+{target_monitor.y}"
            self.root.geometry(geometry_str)
            self.root.overrideredirect(True)
        self.root.configure(bg='black', cursor="none")
        self.label = tk.Label(self.root, bg='black')
        self.label.pack(expand=True, fill=tk.BOTH)
        self.root.update_idletasks()

    def show_image(self, image_path):
        try:
            img = Image.open(image_path)
            win_width = self.root.winfo_width()
            win_height = self.root.winfo_height()
            if win_width > 1 and win_height > 1:
                if img.size != (win_width, win_height):
                    img = img.resize((win_width, win_height), Image.Resampling.LANCZOS)
            self.tk_image = ImageTk.PhotoImage(img)
            self.label.config(image=self.tk_image)
            self.root.update()
        except Exception as e:
            print(f"顯示圖片錯誤: {e}")

    def blank_screen(self):
        self.label.config(image='')
        self.root.update()

    def close(self):
        self.root.destroy()
        print("顯示視窗已關閉。")


# --- 4. Z軸TCP通訊模組 (同步通訊版) ---
class ZAxisControl:
    def __init__(self, host, port, timeout=120):
        self.host = host
        self.port = port
        self.sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        self.sock.settimeout(timeout)
        try:
            print(f"正在通過 TCP 連接到 ESP32 於 {host}:{port}...")
            self.sock.connect((host, port))
            self.reader = self.sock.makefile('r')
            print(f"成功連接到 ESP32。")
        except Exception as e:
            print(f"錯誤: 無法通過 TCP 連接到 ESP32。 {e}")
            exit()

    def _send_cmd_and_wait_response(self, cmd):
        try:
            full_cmd = cmd + "\n"
            self.sock.sendall(full_cmd.encode())
            response = self.reader.readline().strip()
            return response
        except Exception as e:
            print(f"通訊錯誤: {e}")
            return "ERROR"

    def send_config(self, lift_dist, return_dist):
        """發送配置指令，並等待 ESP32 的回覆"""
        cmd = f"CONFIG,{lift_dist},{return_dist}"
        print(f"發送Z軸配置: {cmd}")
        response = self._send_cmd_and_wait_response(cmd)
        print(f" -> ESP32響應: {response}")
        return "OK" in response

    def move_to_next_layer(self):
        print("發送Z軸運動指令...")
        response = self._send_cmd_and_wait_response("NEXT_LAYER")
        if "DONE" in response:
            print("Z軸運動完成。")
            return True
        else:
            print(f"Z軸運動錯誤或超時！響應: {response}")
            return False

    def move_relative(self, distance_mm):
        print(f"發送相對移動指令: {distance_mm} mm...")
        cmd = f"MOVE_REL,{distance_mm}"
        response = self._send_cmd_and_wait_response(cmd)
        if "DONE" in response:
            print("相對移動完成。")
            return True
        else:
            print(f"相對移動錯誤！響應: {response}")
            return False

    def close(self):
        self.sock.close()
        print("TCP 連接已關閉。")


# --- 5. 主流程控制 (混合模式版) ---
def main():
    config = PrintConfig()
    display = None
    z_axis = None
    light_engine = None
    print_completed_successfully = False
    total_layers = 0
    try:
        print(f"正在從 {config.ZIP_FILE_PATH} 解壓縮文件...")
        if not os.path.exists(config.TEMP_EXTRACT_DIR):
            os.makedirs(config.TEMP_EXTRACT_DIR)
        with zipfile.ZipFile(config.ZIP_FILE_PATH, 'r') as zip_ref:
            zip_ref.extractall(config.TEMP_EXTRACT_DIR)
        image_files = sorted([f for f in os.listdir(config.TEMP_EXTRACT_DIR) if f.endswith('.png')],
                             key=lambda x: int(os.path.splitext(x)[0]))
        total_layers = len(image_files)
        if total_layers == 0:
            raise FileNotFoundError("錯誤: 壓縮包中未找到任何PNG文件。")
        image_paths = [os.path.join(config.TEMP_EXTRACT_DIR, f) for f in image_files]
        print(f"找到 {total_layers} 個切片文件。")

        exe_path = os.path.abspath(config.CONTROLLER_EXE_PATH)
        exe_dir = os.path.dirname(exe_path)
        print(f"正在從目錄 '{exe_dir}' 啟動軟體: {os.path.basename(exe_path)}")
        subprocess.Popen(exe_path, cwd=exe_dir)
        z_axis = ZAxisControl(config.ESP32_IP_ADDRESS, config.ESP32_PORT)
        if not z_axis.send_config(config.PEEL_LIFT_DISTANCE, config.PEEL_RETURN_DISTANCE):
            raise RuntimeError("下位機配置失敗，程式終止。")
        while True:
            user_command = input(
                "\n>>> 軟體已啟動。請手動完成設定（Projector ON -> 點擊彈窗 -> 選HDMI -> 設電流），完成後在此處輸入 'print' 並按 Enter 鍵繼續：")
            if user_command.strip().lower() == 'print':
                break
        light_engine = LightEngineGUIControl()
        print("正在創建投影顯示視窗...")
        display = ProjectorDisplay(config.PROJECTOR_MONITOR_INDEX)
        display.blank_screen()
        print("\n--- 所有硬體已初始化，準備開始打印 ---")
        start_time = time.time()
        for i, image_path in enumerate(image_paths):
            layer_num = i + 1
            print(f"\n--- 正在打印第 {layer_num} / {total_layers} 層 ---")
            if layer_num == 1:
                exposure_time = config.FIRST_LAYER_EXPOSURE_TIME_S
            elif layer_num <= config.TRANSITION_LAYERS:
                progress = (layer_num - 1) / (config.TRANSITION_LAYERS - 1)
                exposure_time = config.FIRST_LAYER_EXPOSURE_TIME_S - \
                                (config.FIRST_LAYER_EXPOSURE_TIME_S - config.NORMAL_EXPOSURE_TIME_S) * progress
            else:
                exposure_time = config.NORMAL_EXPOSURE_TIME_S
            print(f"曝光時間: {exposure_time:.2f} 秒")
            display.show_image(image_path)
            light_engine.led_on()
            time.sleep(exposure_time)
            light_engine.led_off()
            display.blank_screen()
            if layer_num < total_layers:
                if not z_axis.move_to_next_layer():
                    print("Z軸運動失敗，打印終止！")
                    break
        else:
            print_completed_successfully = True
        end_time = time.time()
        if print_completed_successfully:
            print(f"\n打印完成！總耗時: {(end_time - start_time) / 60:.2f} 分鐘。")

    except Exception as e:
        print(f"\n程式運行時發生錯誤: {e}")

    finally:
        if print_completed_successfully and z_axis and total_layers > 1:
            print("\n正在執行打印結束後的回位程序...")
            layer_height = config.PEEL_LIFT_DISTANCE - config.PEEL_RETURN_DISTANCE
            total_print_height = (total_layers - 1) * layer_height
            if total_print_height > 0:
                z_axis.move_relative(-total_print_height)
            z_axis.move_relative(2)
            print("回位程序完成。")
        print("\n正在關閉所有設備...")
        if light_engine:
            light_engine.close()
        if z_axis:
            z_axis.close()
        if display:
            display.close()


if __name__ == "__main__":
    main()