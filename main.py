# main.py - 四軸 TCP 控制版 (支援動態參數配置)
import machine
import time
import uasyncio

# --- 1. 自定義異步隊列類 (保持不變) ---
class AsyncQueue:
    def __init__(self): self.items = []; self.event = uasyncio.Event()
    async def put(self, item): self.items.append(item); self.event.set()
    async def get(self):
        while not self.items: await self.event.wait()
        item = self.items.pop(0)
        if not self.items: self.event.clear()
        return item

# --- 2. 硬體設定區 (保持不變) ---
Z_STEP_PIN, Z_DIR_PIN, Z_ENA_PIN = 26, 25, 27
A_STEP_PIN, A_DIR_PIN, A_ENA_PIN = 23, 22, 21
B_STEP_PIN, B_DIR_PIN, B_ENA_PIN = 19, 18, 5
C_STEP_PIN, C_DIR_PIN, C_ENA_PIN = 17, 16, 4
LEVEL_SENSOR_PIN = 34

# --- 3. 步進馬達驅動類 (保持不變) ---
class Stepper:
    def __init__(self, step_pin, dir_pin, ena_pin, is_dm_driver=False):
        self.step_pin_num = step_pin
        self.step = machine.Pin(self.step_pin_num, machine.Pin.OUT)
        self.dir = machine.Pin(dir_pin, machine.Pin.OUT)
        self.use_ena = not is_dm_driver
        if self.use_ena:
            self.ena = machine.Pin(ena_pin, machine.Pin.OUT)
        self.steps_per_mm = 200.0
        self.pwm = None
        self.dir.value(0)
        self.step.value(0)
        self.disable()

    def enable(self):
        if self.use_ena: self.ena.value(0)
    def disable(self):
        if self.use_ena: self.ena.value(1)

    async def move_rel(self, distance_mm, speed_mm_s, accel_mm_s2):
        if self.steps_per_mm == 0: return
        total_steps = int(abs(distance_mm) * self.steps_per_mm)
        if total_steps == 0: return
        
        self.enable()
        self.dir.value(1 if distance_mm < 0 else 0)
        
        max_speed_steps_s = speed_mm_s * self.steps_per_mm
        accel_steps_s2 = accel_mm_s2 * self.steps_per_mm
        
        delays = []
        accel_steps = int(0.5 * (max_speed_steps_s**2) / accel_steps_s2) if accel_steps_s2 > 0 else 0
        if total_steps <= 2 * accel_steps: accel_steps = total_steps // 2
        decel_start_step = total_steps - accel_steps
        
        for i in range(total_steps):
            step_count = i + 1
            if step_count <= accel_steps:
                speed = (max_speed_steps_s / accel_steps) * step_count if accel_steps > 0 else max_speed_steps_s
            elif step_count > decel_start_step:
                speed = max_speed_steps_s - (max_speed_steps_s / accel_steps) * (step_count - decel_start_step) if accel_steps > 0 else max_speed_steps_s
            else:
                speed = max_speed_steps_s
            
            if speed > 0:
                delays.append(1_000_000 // int(speed))
            else:
                delays.append(1_000_000)
        
        print(f"INFO: Moving {distance_mm}mm with acceleration...")
        for i, delay in enumerate(delays):
            self.step.value(1)
            time.sleep_us(2)
            self.step.value(0)
            time.sleep_us(max(2, delay))
            if i % 100 == 0:
                await uasyncio.sleep_ms(0)

# --- 4. 全域變數 ---
command_queue = AsyncQueue()
steppers = { 'z': Stepper(Z_STEP_PIN, Z_DIR_PIN, Z_ENA_PIN, is_dm_driver=True), 'a': Stepper(A_STEP_PIN, A_DIR_PIN, A_ENA_PIN, is_dm_driver=True), 'b': Stepper(B_STEP_PIN, B_DIR_PIN, B_ENA_PIN, is_dm_driver=False), 'c': Stepper(C_STEP_PIN, C_DIR_PIN, C_ENA_PIN, is_dm_driver=True) }
adc = machine.ADC(machine.Pin(LEVEL_SENSOR_PIN)); adc.atten(machine.ADC.ATTN_11DB)
LEVEL_LOW_THRESHOLD = 1000; LEVEL_HIGH_THRESHOLD = 3000
level_compensation_enabled = True

# --- 5. 異步任務 ---
async def tcp_server(host, port):
    print(f"TCP 伺服器啟動於 {host}:{port}")
    async def handle_client(reader, writer):
        print("客戶端已連接")
        while True:
            try:
                data = await reader.readline();
                if data: await command_queue.put((data.decode().strip(), writer))
                else: print("客戶端斷開連接"); break
            except Exception as e: print(f"讀取錯誤: {e}"); break
        writer.close(); await writer.wait_closed()
    await uasyncio.start_server(handle_client, host, port)

async def level_compensator():
    print("液位補償任務已啟動。")
    b_move_step = 0.05
    b_speed_down, b_speed_up = 2.0, 2.0 # 將由 command_processor 更新
    while True:
        if level_compensation_enabled:
            current_level_adc = adc.read()
            if current_level_adc < LEVEL_LOW_THRESHOLD:
                print(f"檢測到液位過低 (ADC: {current_level_adc})，向下補償...")
                await steppers['b'].move_rel(-b_move_step, b_speed_down, b_speed_down * 2)
            elif current_level_adc > LEVEL_HIGH_THRESHOLD:
                print(f"檢測到液位過高 (ADC: {current_level_adc})，向上補償...")
                await steppers['b'].move_rel(b_move_step, b_speed_up, b_speed_up * 2)
        await uasyncio.sleep_ms(1000)

async def command_processor():
    print("指令處理器已啟動")
    # 參數預設值
    params = {
        'peel_lift_z1': 5.05, 'peel_return_z2': 5.0,
        'z_speed_down': 20.0, 'z_speed_up': 20.0,
        'wipe_dist': 50.0, 'wipe_speed_fast': 80.0, 'wipe_speed_slow': 10.0,
        'b_speed_down': 2.0, 'b_speed_up': 2.0,
    }
    
    while True:
        cmd, writer = await command_queue.get()
        print(f"收到指令: {cmd}")
        response = ""; parts = cmd.split(','); command = parts[0].upper()
        try:
            if command == "CONFIG_AXIS":
                axis, pulse_per_rev, lead = parts[1].lower(), float(parts[2]), float(parts[3])
                if axis in steppers: steppers[axis].steps_per_mm = pulse_per_rev / lead; response = f"OK: Axis {axis} configured.\n"
                else: response = "ERROR: Invalid axis.\n"
            elif command == "CONFIG_Z_PEEL": # Z軸剝離參數
                params['peel_lift_z1'], params['peel_return_z2'], params['z_speed_down'], params['z_speed_up'] = map(float, parts[1:])
                response = "OK: Z peel params configured.\n"
            elif command == "CONFIG_A_WIPE": # A軸擦拭參數
                params['wipe_dist'], params['wipe_speed_fast'], params['wipe_speed_slow'] = map(float, parts[1:])
                response = "OK: A wipe params configured.\n"
            elif command == "CONFIG_B_LEVEL": # B軸液位補償速度
                params['b_speed_down'], params['b_speed_up'] = map(float, parts[1:])
                # 更新 level_compensator 任務中的變數 (如果需要)
                global b_speed_down, b_speed_up
                b_speed_down, b_speed_up = params['b_speed_down'], params['b_speed_up']
                response = "OK: B level params configured.\n"
            elif command == "NEXT_LAYER":
                # 使用動態配置的參數
                await steppers['z'].move_rel(-params['peel_lift_z1'], params['z_speed_down'], params['z_speed_down'] * 2)
                await steppers['a'].move_rel(params['wipe_dist'], params['wipe_speed_fast'], params['wipe_speed_fast'] * 2)
                await steppers['z'].move_rel(params['peel_return_z2'], params['z_speed_up'], params['z_speed_up'] * 2)
                await steppers['a'].move_rel(-params['wipe_dist'], params['wipe_speed_slow'], params['wipe_speed_slow'] * 2)
                response = "DONE\n"
            elif command == "MOVE_REL":
                axis, distance, speed, accel = parts[1].lower(), float(parts[2]), float(parts[3]), float(parts[4])
                if axis in steppers: await steppers[axis].move_rel(distance, speed, accel); response = "DONE\n"
                else: response = "ERROR: Invalid axis.\n"
            elif command == "ENABLE_LEVEL_COMP":
                global level_compensation_enabled
                is_enabled = int(parts[1]); level_compensation_enabled = (is_enabled == 1)
                status = "enabled" if level_compensation_enabled else "disabled"
                response = f"OK: Level compensation {status}.\n"
            else: response = "ERROR: Unknown command.\n"
        except Exception as e: response = f"ERROR: Processing command failed: {e}\n"
        if response and writer: writer.write(response.encode()); await writer.drain()

async def main():
    import network
    host_ip = network.WLAN(network.STA_IF).ifconfig()[0]
    server_task = uasyncio.create_task(tcp_server(host_ip, 8899))
    processor_task = uasyncio.create_task(command_processor())
    level_task = uasyncio.create_task(level_compensator())
    print("ESP32 4-Axis Controller Ready.")
    await uasyncio.gather(server_task, processor_task, level_task)

if __name__ == "__main__":
    try: uasyncio.run(main())
    except KeyboardInterrupt: print("Program stopped.")
