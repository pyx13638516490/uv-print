# main_controller_hybrid.py - PC端主控制程式 (I2C精準曝光 + GUI設定電流)
# 版本日期: 2025-09-14
# 流程: 腳本啟動軟體 -> 腳本用GUI設定電流 -> 用戶手動確認 -> 腳本用I2C接管打印

import os
import time
import zipfile
import tkinter as tk
from PIL import Image, ImageTk
import socket
from screeninfo import get_monitors
import subprocess
import ctypes  # 用於I2C控制


# --- 1. 使用者設定區 ---
class PrintConfig:
    # 軟體與文件路徑
    CONTROLLER_EXE_PATH = "Full-HD UV LE Controller v2.1.exe"
    ZIP_FILE_PATH = "layers.zip"
    TEMP_EXTRACT_DIR = "temp_layers"

    # Z軸剝離運動參數
    PEEL_LIFT_DISTANCE = 5.05
    PEEL_RETURN_DISTANCE = 5.0

    # 曝光參數
    NORMAL_EXPOSURE_TIME_S = 2.5
    FIRST_LAYER_EXPOSURE_TIME_S = 5
    TRANSITION_LAYERS = 5

    # !!! 新增：預設LED電流值 !!!
    # 根據手冊，NVM+的電流值範圍是91-810，對應11.7186mA/digit
    # 5000mA / 5.8593mA ≈ 853 (v1.8手冊算法)
    # v2.1的UI顯示0-1023，這裡我們直接輸入UI顯示的值
    LED_CURRENT_VALUE = 853  # 請根據您的樹脂需求修改此值

    # 硬體連接設定
    ESP32_IP_ADDRESS = "10.10.17.187"
    ESP32_PORT = 8899

    # 投影儀螢幕索引 (0=主螢幕, 1=第二個螢幕, ...)
    PROJECTOR_MONITOR_INDEX = 1


# --- 2. 混合光機控制模組 (I2C + GUI) ---
class HybridLightEngineControl:
    """
    混合控制器：
    - 使用 pywinauto 連接GUI，用於設定電流。
    - 使用 ctypes 呼叫DLL，用於透過I2C精準控制LED開關。
    """

    def __init__(self):
        # I2C 相關初始化
        self.CY_SUCCESS = 0
        self.I2C_SLAVE_ADDRESS = 0x1B
        self.I2C_SPEED = 100  # 100kbit/s
        self.cy_handle = ctypes.c_void_p()
        self.dll = None

        # GUI 相關初始化
        self.app = None
        self.main_win = None
        self.current_textbox = None
        self.set_current_button = None

        try:
            # 載入Cypress DLL
            self.dll = ctypes.windll.LoadLibrary("cyusbserial.dll")
            print("成功載入 cyusbserial.dll")
            self._initialize_i2c_device()
        except Exception as e:
            print(f"I2C初始化失敗: {e}")
            raise

        try:
            # 連接到GUI應用程式
            from pywinauto.application import Application  # 延後導入
            window_title = "Full-HD UV LE Controller v2.1"
            print(f"正在連接到GUI視窗: '{window_title}'...")
            self.app = Application(backend="uia").connect(title=window_title, timeout=60)
            self.main_win = self.app.window(title=window_title)
            self.main_win.wait('ready', timeout=30)

            # 獲取設定電流所需的GUI元件
            self.current_textbox = self.main_win.child_window(auto_id="TextBoxCurrent")
            self.set_current_button = self.main_win.child_window(auto_id="ButtonSetLedCurrent")
            print("成功連接到光機GUI軟體。")
        except Exception as e:
            print(f"連接到GUI失敗: {e}")
            self.close()  # 如果GUI連接失敗，也關閉I2C
            raise

    # --- I2C控制相關方法 ---
    def _initialize_i2c_device(self):
        device_id = ctypes.c_ubyte(0)
        num_devices = ctypes.c_uint(0)
        status = self.dll.CyGetListofDevices(ctypes.byref(num_devices))
        if status != self.CY_SUCCESS or num_devices.value == 0:
            raise ConnectionError("找不到任何Cypress USB-Serial設備。")
        status = self.dll.CyOpen(device_id, 0, ctypes.byref(self.cy_handle))
        if status != self.CY_SUCCESS:
            raise ConnectionError(f"開啟Cypress設備失敗，錯誤碼: {status}")

        class I2C_CONFIG(ctypes.Structure):
            _fields_ = [("frequency", ctypes.c_ulong), ("slaveAddress", ctypes.c_ubyte),
                        ("isMaster", ctypes.c_bool), ("isClockStreching", ctypes.c_bool)]

        i2c_config = I2C_CONFIG()
        self.dll.CyGetI2cConfig(self.cy_handle, ctypes.byref(i2c_config))
        i2c_config.frequency = self.I2C_SPEED * 1000
        i2c_config.isMaster = True
        i2c_config.slaveAddress = self.I2C_SLAVE_ADDRESS
        status = self.dll.CySetI2cConfig(self.cy_handle, ctypes.byref(i2c_config))
        if status != self.CY_SUCCESS:
            self.close()
            raise RuntimeError(f"設定I2C參數失敗，錯誤碼: {status}")
        print("Cypress設備初始化成功，I2C通訊已準備就緒。")

    def _send_i2c_command(self, command, data_list):
        buffer_list = [command] + data_list
        buffer_size = len(buffer_list)

        class I2C_DATA_XFER(ctypes.Structure):
            _fields_ = [("slaveAddress", ctypes.c_ubyte), ("buffer", ctypes.POINTER(ctypes.c_ubyte)),
                        ("length", ctypes.c_ulong), ("isStopBit", ctypes.c_bool), ("isNakBit", ctypes.c_bool)]

        write_buffer = (ctypes.c_ubyte * buffer_size)(*buffer_list)
        xfer_params = I2C_DATA_XFER(slaveAddress=self.I2C_SLAVE_ADDRESS, buffer=write_buffer,
                                    length=buffer_size, isStopBit=True)
        status = self.dll.CyI2cWrite(self.cy_handle, ctypes.byref(xfer_params), 500)
        return status == self.CY_SUCCESS

    def led_on(self):
        """[I2C] 開啟LED"""
        if not self._send_i2c_command(0x52, [0x02]):
            print("警告: 發送 I2C 'LED ON' 指令失敗！")

    def led_off(self):
        """[I2C] 關閉LED"""
        if not self._send_i2c_command(0x52, [0x00]):
            print("警告: 發送 I2C 'LED OFF' 指令失敗！")

    # --- GUI控制相關方法 ---
    def set_current_via_gui(self, current_value):
        """[GUI] 設定LED電流"""
        try:
            print(f"指令: 透過GUI設定電流值為 {current_value}...")
            self.current_textbox.set_edit_text(str(current_value))
            time.sleep(0.1)  # 等待UI反應
            self.set_current_button.click()
            print(" -> 電流設定指令已發送。")
            return True
        except Exception as e:
            print(f"錯誤: 透過GUI設定電流失敗: {e}")
            return False

    # --- 關閉與清理 ---
    def close(self):
        """關閉所有連接"""
        if self.cy_handle:
            self.dll.CyClose(self.cy_handle)
            print("Cypress I2C 連接已關閉。")
        print("自動化流程已結束，請手動關閉光機控制軟體。")


# --- (ProjectorDisplay 和 ZAxisControl 類別保持不變) ---
class ProjectorDisplay:
    def __init__(self, monitor_index):
        self.root = tk.Tk()
        try:
            monitors = sorted(get_monitors(), key=lambda m: m.x)
        except Exception:
            monitors = [get_monitors()[0]] if get_monitors() else []

        target_monitor = monitors[monitor_index] if len(monitors) > monitor_index else (
            monitors[0] if monitors else None)
        if not target_monitor: raise RuntimeError("找不到任何螢幕。")

        print(f"投影目標: {target_monitor.width}x{target_monitor.height} at ({target_monitor.x}, {target_monitor.y})")

        is_primary = (target_monitor.x == 0 and target_monitor.y == 0) and len(monitors) > 1
        if is_primary:
            self.root.geometry(f"800x600+{target_monitor.x + 100}+{target_monitor.y + 100}")
            self.root.title("投影預覽")
        else:
            self.root.geometry(f"{target_monitor.width}x{target_monitor.height}+{target_monitor.x}+{target_monitor.y}")
            self.root.overrideredirect(True)

        self.root.configure(bg='black', cursor="none")
        self.label = tk.Label(self.root, bg='black')
        self.label.pack(expand=True, fill=tk.BOTH)
        self.root.update_idletasks()
        self.target_size = (self.root.winfo_width(), self.root.winfo_height())

    def show_image(self, image_path):
        try:
            img = Image.open(image_path).resize(self.target_size, Image.Resampling.LANCZOS)
            self.tk_image = ImageTk.PhotoImage(img)
            self.label.config(image=self.tk_image)
            self.root.update()
        except Exception as e:
            print(f"顯示圖片錯誤: {e}")

    def blank_screen(self):
        self.label.config(image=''); self.root.update()

    def close(self):
        self.root.destroy(); print("顯示視窗已關閉。")


class ZAxisControl:
    def __init__(self, host, port, timeout=120):
        try:
            print(f"正在連接到ESP32於 {host}:{port}...")
            self.sock = socket.create_connection((host, port), timeout)
            self.reader = self.sock.makefile('r', encoding='utf-8')
            print("成功連接到 ESP32。")
        except Exception as e:
            raise ConnectionError(f"無法連接到 ESP32: {e}")

    def _send_cmd_and_wait_response(self, cmd):
        try:
            self.sock.sendall((cmd + "\n").encode('utf-8'))
            return self.reader.readline().strip()
        except (socket.timeout, ConnectionResetError) as e:
            print(f"通訊錯誤: {e}"); return "ERROR"

    def send_config(self, lift_dist, return_dist):
        res = self._send_cmd_and_wait_response(f"CONFIG,{lift_dist},{return_dist}")
        print(f"Z軸配置響應: {res}")
        return "OK" in res

    def move_to_next_layer(self):
        print("發送Z軸運動指令...");
        return "DONE" in self._send_cmd_and_wait_response("NEXT_LAYER")

    def move_relative(self, distance_mm):
        print(f"發送相對移動指令: {distance_mm} mm...");
        return "DONE" in self._send_cmd_and_wait_response(f"MOVE_REL,{distance_mm}")

    def close(self):
        self.sock.close(); print("TCP 連接已關閉。")


# --- 5. 主流程控制 (混合模式版) ---
def main():
    config = PrintConfig()
    display = None
    z_axis = None
    light_engine = None

    try:
        # ... (解壓縮檔案的程式碼不變)
        if not os.path.exists(config.TEMP_EXTRACT_DIR): os.makedirs(config.TEMP_EXTRACT_DIR)
        with zipfile.ZipFile(config.ZIP_FILE_PATH, 'r') as z:
            z.extractall(config.TEMP_EXTRACT_DIR)
        image_files = sorted([f for f in os.listdir(config.TEMP_EXTRACT_DIR) if f.endswith('.png')],
                             key=lambda x: int(os.path.splitext(x)[0]))
        total_layers = len(image_files)
        if total_layers == 0: raise FileNotFoundError("壓縮包中未找到任何PNG文件。")
        image_paths = [os.path.join(config.TEMP_EXTRACT_DIR, f) for f in image_files]
        print(f"找到 {total_layers} 個切片文件。")

        exe_path = os.path.abspath(config.CONTROLLER_EXE_PATH)
        subprocess.Popen(exe_path, cwd=os.path.dirname(exe_path))

        z_axis = ZAxisControl(config.ESP32_IP_ADDRESS, config.ESP32_PORT)
        if not z_axis.send_config(config.PEEL_LIFT_DISTANCE, config.PEEL_RETURN_DISTANCE):
            raise RuntimeError("下位機配置失敗，程式終止。")

        input("\n>>> 軟體已啟動。請手動操作GUI，點擊 Projector ON 並在彈窗中點擊OK。\n"
              "    完成後，請按 Enter 鍵讓程式繼續...")

        # !!!!!!! 核心改變 !!!!!!!
        light_engine = HybridLightEngineControl()

        # 透過GUI設定電流
        if not light_engine.set_current_via_gui(config.LED_CURRENT_VALUE):
            raise RuntimeError("設定電流失敗，程式終止。")

        input(f">>> 電流已設定為 {config.LED_CURRENT_VALUE}。請在GUI上確認HDMI為影像來源。\n"
              "    一切就緒後，請按 Enter 鍵開始打印...")

        display = ProjectorDisplay(config.PROJECTOR_MONITOR_INDEX)

        print("\n--- 所有硬體已初始化，準備開始打印 ---")
        start_time = time.time()

        for i, image_path in enumerate(image_paths):
            layer_num = i + 1
            print(f"\n--- 正在打印第 {layer_num} / {total_layers} 層 ---")

            # (計算曝光時間的邏輯不變)
            if layer_num == 1:
                exposure_time = config.FIRST_LAYER_EXPOSURE_TIME_S
            elif layer_num <= config.TRANSITION_LAYERS:
                p = (layer_num - 1) / (config.TRANSITION_LAYERS - 1)
                exposure_time = config.FIRST_LAYER_EXPOSURE_TIME_S - (
                            config.FIRST_LAYER_EXPOSURE_TIME_S - config.NORMAL_EXPOSURE_TIME_S) * p
            else:
                exposure_time = config.NORMAL_EXPOSURE_TIME_S
            print(f"曝光時間: {exposure_time:.2f} 秒")

            # 使用精準的I2C控制曝光
            display.show_image(image_path)
            light_engine.led_on()
            time.sleep(exposure_time)
            light_engine.led_off()
            display.blank_screen()

            if layer_num < total_layers and not z_axis.move_to_next_layer():
                print("Z軸運動失敗，打印終止！");
                break
        else:
            print(f"\n打印完成！總耗時: {(time.time() - start_time) / 60:.2f} 分鐘。")

    except Exception as e:
        print(f"\n程式運行時發生嚴重錯誤: {e}")

    finally:
        print("\n--- 正在執行清理程序 ---")
        # (清理邏輯不變)
        if light_engine: light_engine.close()
        if z_axis: z_axis.close()
        if display: display.close()
        if os.path.exists(config.TEMP_EXTRACT_DIR):
            import shutil
            shutil.rmtree(config.TEMP_EXTRACT_DIR)
            print("臨時文件夾已刪除。")
        print("所有設備已關閉，程序結束。")


if __name__ == "__main__":
    main()