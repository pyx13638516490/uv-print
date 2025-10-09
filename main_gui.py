# main_gui.py - 三軸穩定版 (v3.1 - A軸改為限位開關控制)

import sys, os, time, zipfile, socket, subprocess
from multiprocessing.connection import Client
from PIL import Image

from PyQt5.QtWidgets import (QApplication, QWidget, QVBoxLayout, QHBoxLayout, QGridLayout, QGroupBox,
                             QLabel, QLineEdit, QPushButton, QPlainTextEdit, QDoubleSpinBox,
                             QFileDialog)
from PyQt5.QtCore import QThread, QObject, pyqtSignal, pyqtSlot

from pywinauto.application import Application

# --- 後端邏輯 ---
class LightEngineGUIControl:
    def __init__(self):
        self.app = None; self.main_win = None
        try:
            window_title = "Full-HD UV LE Controller v2.1"; self.app = Application(backend="uia").connect(title=window_title, timeout=60); self.main_win = self.app.window(title=window_title); self.main_win.wait('ready', timeout=30)
            self.led_combo = self.main_win.child_window(auto_id="ComboBoxLedEnable"); self.set_led_onoff_button = self.main_win.child_window(auto_id="ButtonSetLedOnOff")
        except Exception as e: raise RuntimeError(f"連接到控制軟體失敗: {e}")
    def led_on(self):
        try: self.main_win.set_focus(); self.led_combo.select("On"); time.sleep(0.1); self.set_led_onoff_button.click(); time.sleep(0.1)
        except Exception as e: raise RuntimeError(f"自動化控制 'LED On' 失敗: {e}")
    def led_off(self):
        try: self.main_win.set_focus(); self.led_combo.select("Off"); time.sleep(0.1); self.set_led_onoff_button.click(); time.sleep(0.1)
        except Exception as e: raise RuntimeError(f"自動化控制 'LED Off' 失敗: {e}")
    def close(self): pass

class MotionController:
    def __init__(self, host, port, timeout=60):
        self.sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM); self.sock.settimeout(timeout); self.sock.connect((host, port)); self.reader = self.sock.makefile('r')
    def _send_cmd_and_wait_response(self, cmd):
        full_cmd = cmd + "\n"; self.sock.sendall(full_cmd.encode()); response = self.reader.readline().strip(); return response
    def close(self): self.sock.close()
    def config_axis(self, axis, pulse_per_rev, lead): return "OK" in self._send_cmd_and_wait_response(f"CONFIG_AXIS,{axis},{pulse_per_rev},{lead}")
    def config_z_peel(self, params): return "OK" in self._send_cmd_and_wait_response(f"CONFIG_Z_PEEL,{params['peel_lift_z1']},{params['peel_return_z2']},{params['z_speed_down']},{params['z_speed_up']}")
    def config_a_wipe(self, params):
        return "OK" in self._send_cmd_and_wait_response(f"CONFIG_A_WIPE,{params['a_fast_speed']},{params['a_slow_speed']}")
    def move_to_next_layer(self): return "DONE" in self._send_cmd_and_wait_response("NEXT_LAYER")
    def move_relative(self, axis, distance, speed): accel = speed * 2; return "DONE" in self._send_cmd_and_wait_response(f"MOVE_REL,{axis},{distance},{speed},{accel}")

class PrintWorker(QObject):
    log = pyqtSignal(str); finished = pyqtSignal(); error = pyqtSignal(str)
    def __init__(self, params): super().__init__(); self.params = params; self.is_running = True
    @pyqtSlot()
    def run(self):
        motion_controller = None; light_engine = None; projector_process = None; projector_conn = None; light_engine_process = None
        try:
            black_image_path = self.params['black_image_path']; self.log.emit("--- 打印任務開始 ---")
            exe_path = self.params['controller_exe_path']; self.log.emit(f"正在檢查光機控制軟體路徑: {exe_path}...")
            if not os.path.exists(exe_path): raise RuntimeError(f"光機控制軟體未找到，請檢查路徑: {exe_path}")
            self.log.emit("正在啟動光機控制軟體..."); light_engine_process = subprocess.Popen([exe_path]); time.sleep(3)
            self.log.emit("正在啟動獨立投影視窗進程..."); address = ('localhost', 6000); authkey = b'secret-key-for-projector'
            python_exe = sys.executable; projector_script = os.path.join(os.path.dirname(__file__), 'projector_view.py')
            if not os.path.exists(projector_script): raise RuntimeError(f"投影腳本 projector_view.py 未找到！")
            cmd = [python_exe, projector_script, str(self.params['monitor_index']), address[0], str(address[1]), authkey.decode()]
            startupinfo = None
            if os.name == 'nt': startupinfo = subprocess.STARTUPINFO(); startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            projector_process = subprocess.Popen(cmd, startupinfo=startupinfo); time.sleep(2)
            projector_conn = Client(address, authkey=authkey); self.log.emit("投影視窗進程已連接。")
            self.log.emit(f"正在從 {self.params['zip_path']} 解壓縮文件...")
            if not os.path.exists(self.params['temp_dir']): os.makedirs(self.params['temp_dir'])
            with zipfile.ZipFile(self.params['zip_path'], 'r') as zip_ref: zip_ref.extractall(self.params['temp_dir'])
            image_files = sorted([f for f in os.listdir(self.params['temp_dir']) if f.endswith('.png') and os.path.splitext(f)[0].isdigit()], key=lambda x: int(os.path.splitext(x)[0]))
            total_layers = len(image_files); image_paths = [os.path.join(self.params['temp_dir'], f) for f in image_files]; self.log.emit(f"找到 {total_layers} 個切片文件。")
            self.log.emit("正在連接到 ESP32..."); motion_controller = MotionController(self.params['esp32_ip'], self.params['esp32_port']); self.log.emit("ESP32 連接成功。")
            self.log.emit("正在發送所有配置..."); motion_controller.config_axis('z', self.params['z_pulse_rev'], self.params['z_lead']); motion_controller.config_axis('a', self.params['a_pulse_rev'], self.params['a_lead']); motion_controller.config_axis('c', self.params['c_pulse_rev'], self.params['c_lead'])
            motion_controller.config_z_peel(self.params); motion_controller.config_a_wipe(self.params); self.log.emit("配置發送完成。")
            self.log.emit("正在連接到光機控制軟體..."); light_engine = LightEngineGUIControl(); self.log.emit("光機軟體連接成功。")
            projector_conn.send({'command': 'show', 'path': black_image_path})
            self.log.emit("--- 所有硬體已初始化，打印循環開始 ---")
            for i, image_path in enumerate(image_paths):
                if not self.is_running: self.log.emit("打印任務被用戶終止。"); break
                layer_num = i + 1; self.log.emit(f"\n--- 正在打印第 {layer_num} / {total_layers} 層 ---")
                if layer_num == 1: exposure_time = self.params['first_layer_expo']
                elif layer_num <= self.params['transition_layers']: progress = (layer_num - 1) / (self.params['transition_layers'] - 1); exposure_time = self.params['first_layer_expo'] - (self.params['first_layer_expo'] - self.params['normal_expo']) * progress
                else: exposure_time = self.params['normal_expo']
                self.log.emit(f"曝光時間: {exposure_time:.2f} 秒")
                projector_conn.send({'command': 'show', 'path': image_path}); light_engine.led_on(); time.sleep(exposure_time)
                projector_conn.send({'command': 'show', 'path': black_image_path}); light_engine.led_off()
                if layer_num < total_layers:
                    if not motion_controller.move_to_next_layer(): raise RuntimeError("層間運動失敗，打印終止！")
            else: self.log.emit("\n打印完成！")
        except Exception as e: self.error.emit(f"打印過程中發生錯誤: {e}")
        finally:
            self.log.emit("正在關閉所有設備...")
            if projector_conn: projector_conn.send({'command': 'close'}); projector_conn.close()
            if projector_process: projector_process.terminate()
            if motion_controller: motion_controller.close()
            if light_engine_process: light_engine_process.terminate()
            self.finished.emit()
    def stop(self): self.is_running = False

class PrintConfig:
    ZIP_FILE_PATH = "layers.zip"; CONTROLLER_EXE_PATH = "Full-HD UV LE Controller v2.1.exe"; TEMP_EXTRACT_DIR = "temp_layers"
    BLACK_IMAGE_PATH = os.path.join(TEMP_EXTRACT_DIR, "black.png"); PROJECTOR_MONITOR_INDEX = 1
    ESP32_IP_ADDRESS = "10.10.17.187"; ESP32_PORT = 8899
    Z_PULSE_PER_REV = 12800.0; Z_LEAD = 5.0; Z_PEEL_SPEED = 20.0; Z_JOG_SPEED = 10.0
    A_PULSE_PER_REV = 12800.0; A_LEAD = 75.0
    A_WIPE_SPEED_FAST = 80.0; A_WIPE_SPEED_SLOW = 10.0; A_JOG_SPEED = 40.0
    C_PULSE_PER_REV = 12800.0; C_LEAD = 5.0; C_JOG_DISTANCE = 10.0; C_JOG_SPEED = 20.0
    NORMAL_EXPOSURE_TIME_S = 2.5; FIRST_LAYER_EXPOSURE_TIME_S = 5.0; TRANSITION_LAYERS = 5

class MainWindow(QWidget):
    def __init__(self):
        super().__init__(); self.worker_thread = None; self.print_worker = None; self.motion_controller = None; self.initUI()
    def initUI(self):
        self.setWindowTitle('三軸 DLP 打印機控制器')
        main_layout = QVBoxLayout()
        conn_group = QGroupBox("連接設定"); conn_layout = QHBoxLayout(); conn_layout.addWidget(QLabel("ESP32 IP:")); self.esp32_ip_edit = QLineEdit(PrintConfig.ESP32_IP_ADDRESS); conn_layout.addWidget(self.esp32_ip_edit); self.connect_button = QPushButton("連接 & 初始化 ESP32"); conn_layout.addWidget(self.connect_button); conn_group.setLayout(conn_layout); main_layout.addWidget(conn_group)
        params_group = QGroupBox("打印參數設定"); params_layout = QGridLayout(); params_layout.addWidget(QLabel("層高 (mm):"), 0, 0); self.layer_height_edit = QDoubleSpinBox(); self.layer_height_edit.setDecimals(3); self.layer_height_edit.setValue(0.05); params_layout.addWidget(self.layer_height_edit, 0, 1); params_layout.addWidget(QLabel("Z 軸剝離距離 (mm):"), 0, 2); self.peel_base_dist_edit = QDoubleSpinBox(); self.peel_base_dist_edit.setValue(5.0); params_layout.addWidget(self.peel_base_dist_edit, 0, 3); params_layout.addWidget(QLabel("底層曝光 (s):"), 1, 0); self.first_expo_edit = QDoubleSpinBox(); self.first_expo_edit.setValue(PrintConfig.FIRST_LAYER_EXPOSURE_TIME_S); params_layout.addWidget(self.first_expo_edit, 1, 1); params_layout.addWidget(QLabel("正常曝光 (s):"), 1, 2); self.normal_expo_edit = QDoubleSpinBox(); self.normal_expo_edit.setValue(PrintConfig.NORMAL_EXPOSURE_TIME_S); params_layout.addWidget(self.normal_expo_edit, 1, 3); params_group.setLayout(params_layout); main_layout.addWidget(params_group)
        speed_group = QGroupBox("速度設定 (mm/s)"); speed_layout = QGridLayout()
        speed_layout.addWidget(QLabel("Z 軸下移速度:"), 0, 0); self.z_speed_down_edit = QDoubleSpinBox(); self.z_speed_down_edit.setValue(PrintConfig.Z_PEEL_SPEED); speed_layout.addWidget(self.z_speed_down_edit, 0, 1)
        speed_layout.addWidget(QLabel("Z 軸上移速度:"), 0, 2); self.z_speed_up_edit = QDoubleSpinBox(); self.z_speed_up_edit.setValue(PrintConfig.Z_PEEL_SPEED); speed_layout.addWidget(self.z_speed_up_edit, 0, 3)
        speed_layout.addWidget(QLabel("A 軸擦拭速度 (快):"), 1, 0); self.a_speed_fast_edit = QDoubleSpinBox(); self.a_speed_fast_edit.setValue(PrintConfig.A_WIPE_SPEED_FAST); speed_layout.addWidget(self.a_speed_fast_edit, 1, 1)
        speed_layout.addWidget(QLabel("A 軸擦拭速度 (慢):"), 1, 2); self.a_speed_slow_edit = QDoubleSpinBox(); self.a_speed_slow_edit.setValue(PrintConfig.A_WIPE_SPEED_SLOW); speed_layout.addWidget(self.a_speed_slow_edit, 1, 3)
        speed_layout.addWidget(QLabel("C 軸恆定速度:"), 2, 0); self.c_jog_speed_edit = QDoubleSpinBox(); self.c_jog_speed_edit.setValue(PrintConfig.C_JOG_SPEED); speed_layout.addWidget(self.c_jog_speed_edit, 2, 1)
        speed_group.setLayout(speed_layout); main_layout.addWidget(speed_group)
        self.jog_group = QGroupBox("手動控制"); jog_layout = QGridLayout(); jog_layout.addWidget(QLabel("Z 軸距離(mm):"), 0, 0); self.z_jog_dist_edit = QDoubleSpinBox(); self.z_jog_dist_edit.setValue(10.0); jog_layout.addWidget(self.z_jog_dist_edit, 0, 1); self.z_up_button = QPushButton("Z 軸向上"); jog_layout.addWidget(self.z_up_button, 0, 2); self.z_down_button = QPushButton("Z 軸向下"); jog_layout.addWidget(self.z_down_button, 0, 3); jog_layout.addWidget(QLabel("A 軸距離(mm):"), 1, 0); self.a_jog_dist_edit = QDoubleSpinBox(); self.a_jog_dist_edit.setValue(10.0); jog_layout.addWidget(self.a_jog_dist_edit, 1, 1); self.a_fwd_button = QPushButton("A 軸向前"); jog_layout.addWidget(self.a_fwd_button, 1, 2); self.a_back_button = QPushButton("A 軸向後"); jog_layout.addWidget(self.a_back_button, 1, 3); jog_layout.addWidget(QLabel("C 軸距離(mm):"), 2, 0); self.c_jog_dist_edit = QDoubleSpinBox(); self.c_jog_dist_edit.setValue(PrintConfig.C_JOG_DISTANCE); jog_layout.addWidget(self.c_jog_dist_edit, 2, 1); self.c_up_button = QPushButton("C 軸向上"); jog_layout.addWidget(self.c_up_button, 2, 2); self.c_down_button = QPushButton("C 軸向下"); jog_layout.addWidget(self.c_down_button, 2, 3); self.jog_group.setLayout(jog_layout); main_layout.addWidget(self.jog_group)
        control_layout = QHBoxLayout(); self.start_button = QPushButton("開始打印"); self.stop_button = QPushButton("終止打印"); control_layout.addWidget(self.start_button); control_layout.addWidget(self.stop_button); main_layout.addLayout(control_layout); self.log_widget = QPlainTextEdit(); self.log_widget.setReadOnly(True); main_layout.addWidget(self.log_widget); self.setLayout(main_layout)
        self.connect_button.clicked.connect(self.connect_esp32); self.start_button.clicked.connect(self.start_print); self.stop_button.clicked.connect(self.stop_print)
        self.z_up_button.clicked.connect(lambda: self.jog_axis('z', 1)); self.z_down_button.clicked.connect(lambda: self.jog_axis('z', -1)); self.a_fwd_button.clicked.connect(lambda: self.jog_axis('a', 1)); self.a_back_button.clicked.connect(lambda: self.jog_axis('a', -1)); self.c_up_button.clicked.connect(lambda: self.jog_axis('c', 1)); self.c_down_button.clicked.connect(lambda: self.jog_axis('c', -1))
        self.set_controls_enabled(False)
    def set_controls_enabled(self, enabled): self.jog_group.setEnabled(enabled); self.start_button.setEnabled(enabled); self.stop_button.setEnabled(False)
    def get_params(self):
        peel_base = self.peel_base_dist_edit.value(); layer_height = self.layer_height_edit.value()
        return {
            'esp32_ip': self.esp32_ip_edit.text(), 'esp32_port': PrintConfig.ESP32_PORT, 'zip_path': PrintConfig.ZIP_FILE_PATH,
            'temp_dir': PrintConfig.TEMP_EXTRACT_DIR, 'monitor_index': PrintConfig.PROJECTOR_MONITOR_INDEX, 'controller_exe_path': PrintConfig.CONTROLLER_EXE_PATH,
            'black_image_path': PrintConfig.BLACK_IMAGE_PATH,
            'first_layer_expo': self.first_expo_edit.value(), 'normal_expo': self.normal_expo_edit.value(), 'transition_layers': PrintConfig.TRANSITION_LAYERS,
            'z_pulse_rev': PrintConfig.Z_PULSE_PER_REV, 'z_lead': PrintConfig.Z_LEAD, 'a_pulse_rev': PrintConfig.A_PULSE_PER_REV, 'a_lead': PrintConfig.A_LEAD, 'c_pulse_rev': PrintConfig.C_PULSE_PER_REV, 'c_lead': PrintConfig.C_LEAD,
            'peel_lift_z1': peel_base + layer_height, 'peel_return_z2': peel_base, 'z_speed_down': self.z_speed_down_edit.value(), 'z_speed_up': self.z_speed_up_edit.value(),
            'a_fast_speed': self.a_speed_fast_edit.value(),
            'a_slow_speed': self.a_speed_slow_edit.value(), 'c_jog_speed': self.c_jog_speed_edit.value(), 'z_jog_speed': PrintConfig.Z_JOG_SPEED, 'a_jog_speed': PrintConfig.A_JOG_SPEED,
        }
    def log(self, message): self.log_widget.appendPlainText(message)
    @pyqtSlot()
    def connect_esp32(self):
        if self.motion_controller: self.motion_controller.close(); self.motion_controller = None
        try:
            params = self.get_params(); self.log(f"正在連接並初始化 ESP32 於 {params['esp32_ip']}..."); self.motion_controller = MotionController(params['esp32_ip'], params['esp32_port'])
            self.motion_controller.config_axis('z', params['z_pulse_rev'], params['z_lead']); self.motion_controller.config_axis('a', params['a_pulse_rev'], params['a_lead']); self.motion_controller.config_axis('c', params['c_pulse_rev'], params['c_lead'])
            self.motion_controller.config_z_peel(params); self.motion_controller.config_a_wipe(params); self.log("軸配置與參數發送成功。")
            self.set_controls_enabled(True); self.connect_button.setText("重新連接 & 初始化"); self.log("ESP32 已連接並初始化。")
        except Exception as e: self.log(f"錯誤: 無法連接或初始化 ESP32: {e}"); self.set_controls_enabled(False)
    def start_print(self):
        self.set_controls_enabled(False); self.stop_button.setEnabled(True); self.log_widget.clear(); params = self.get_params()
        self.worker_thread = QThread(); self.print_worker = PrintWorker(params); self.print_worker.moveToThread(self.worker_thread); self.worker_thread.started.connect(self.print_worker.run); self.print_worker.finished.connect(self.on_task_finished); self.print_worker.log.connect(self.log); self.print_worker.error.connect(self.on_task_error); self.worker_thread.start()
    def stop_print(self):
        if self.print_worker: self.print_worker.stop(); self.log("正在發送終止信號...")
    def on_task_finished(self):
        self.log("任務執行緒已結束。"); self.worker_thread.quit(); self.worker_thread.wait(); self.set_controls_enabled(True)
    def on_task_error(self, err_msg): self.log(f"錯誤: {err_msg}"); self.on_task_finished()
    def jog_axis(self, axis, direction):
        if not self.motion_controller: self.log("錯誤: 請先連接到 ESP32。"); return
        dist_edit_map = {'z': self.z_jog_dist_edit, 'a': self.a_jog_dist_edit, 'c': self.c_jog_dist_edit}
        try:
            params = self.get_params(); dist = dist_edit_map[axis].value() * direction; speed = params.get(f'{axis}_jog_speed', 20.0); self.log(f"手動控制: {axis} 軸移動 {dist} mm..."); self.motion_controller.move_relative(axis, dist, speed); self.log("手動控制完成。")
        except Exception as e: self.log(f"手動控制出錯: {e}")
    def closeEvent(self, event):
        if self.motion_controller: self.motion_controller.close()
        if self.worker_thread and self.worker_thread.isRunning(): self.stop_print(); self.worker_thread.quit(); self.worker_thread.wait()
        event.accept()

if __name__ == '__main__':
    temp_dir = PrintConfig.TEMP_EXTRACT_DIR; black_image_path = PrintConfig.BLACK_IMAGE_PATH
    if not os.path.exists(temp_dir): os.makedirs(temp_dir)
    if not os.path.exists(black_image_path):
        print(f"'{black_image_path}' not found, creating a new one...")
        black_img = Image.new('RGB', (1920, 1080), 'black'); black_img.save(black_image_path)
    app = QApplication(sys.argv); ex = MainWindow(); ex.show(); sys.exit(app.exec_())