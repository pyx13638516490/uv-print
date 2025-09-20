# main.py - TCP 通訊版 (修正了 Stepper bug)
import machine
import time
import uasyncio

# --- 1. 自定義異步隊列類 ---
class AsyncQueue:
    def __init__(self):
        self.items = []
        self.event = uasyncio.Event()
    async def put(self, item):
        self.items.append(item)
        self.event.set()
    async def get(self):
        while not self.items:
            await self.event.wait()
        item = self.items.pop(0)
        if not self.items:
            self.event.clear()
        return item

# --- 2. 使用者設定區 ---
DIR_PIN = 25
STEP_PIN = 26
STEPS_PER_MM = 3200.0
MAX_SPEED_MM_S = 10.0
ACCELERATION_MM_S2 = 20.0

# --- 3. 步進馬達加減速驅動類 ---
class Stepper:
    def __init__(self, dir_pin, step_pin, steps_per_mm):
        self.dir = machine.Pin(dir_pin, machine.Pin.OUT)
        self.step = machine.Pin(step_pin, machine.Pin.OUT)
        self.steps_per_mm = steps_per_mm
        self.step.value(0)
        self.dir.value(0)
    async def move_rel(self, distance_mm, max_speed, accel):
        total_steps = int(abs(distance_mm) * self.steps_per_mm)
        if total_steps == 0:
            return
        self.dir.value(1 if distance_mm < 0 else 0)
        max_speed_steps_s = max_speed * self.steps_per_mm
        accel_steps_s2 = accel * self.steps_per_mm
        accel_steps = int(0.5 * (max_speed_steps_s**2) / accel_steps_s2)
        if total_steps <= 2 * accel_steps:
            accel_steps = total_steps // 2
        
        decel_start_step = total_steps - accel_steps
        
        print(f"INFO: Moving {distance_mm}mm, {total_steps} steps.")
        for i in range(total_steps):
            step_count = i + 1
            if step_count <= accel_steps:
                speed = (max_speed_steps_s / accel_steps) * step_count
            elif step_count > decel_start_step:
                speed = max_speed_steps_s - (max_speed_steps_s / accel_steps) * (step_count - decel_start_step)
            else:
                speed = max_speed_steps_s
            if speed > 0:
                delay = 1_000_000 // int(speed)
            else:
                delay = 1_000_000
            self.step.value(1)
            time.sleep_us(2)
            self.step.value(0)
            time.sleep_us(max(2, delay))
            
            if i % 50 == 0:
                await uasyncio.sleep_ms(0)

# --- 4. 全域變數 ---
command_queue = AsyncQueue()
stepper = Stepper(DIR_PIN, STEP_PIN, STEPS_PER_MM)
peel_lift_dist_mm = 5.0
peel_return_dist_mm = 5.05

# --- 5. 異步任務 ---
async def tcp_server(host, port):
    print(f"TCP 伺服器啟動於 {host}:{port}")
    async def handle_client(reader, writer):
        print("客戶端已連接")
        while True:
            try:
                data = await reader.readline()
                if data:
                    cmd = data.decode().strip()
                    await command_queue.put((cmd, writer))
                else:
                    print("客戶端斷開連接")
                    break
            except Exception as e:
                print(f"讀取錯誤: {e}")
                break
        writer.close()
        await writer.wait_closed()
    await uasyncio.start_server(handle_client, host, port)

async def command_processor():
    global peel_lift_dist_mm, peel_return_dist_mm
    print("指令處理器已啟動")
    while True:
        cmd, writer = await command_queue.get()
        print(f"收到指令: {cmd}")
        response = ""
        if cmd.startswith("CONFIG"):
            try:
                parts = cmd.split(',')
                peel_lift_dist_mm = float(parts[1])
                peel_return_dist_mm = float(parts[2])
                response = f"OK: Config received.\n"
            except (IndexError, ValueError):
                response = "ERROR: Invalid CONFIG format.\n"
        elif cmd == "NEXT_LAYER":
            await stepper.move_rel(peel_lift_dist_mm, MAX_SPEED_MM_S, ACCELERATION_MM_S2)
            await stepper.move_rel(-peel_return_dist_mm, MAX_SPEED_MM_S, ACCELERATION_MM_S2)
            response = "DONE\n"
        elif cmd.startswith("MOVE_REL"):
            try:
                parts = cmd.split(',')
                distance = float(parts[1])
                await stepper.move_rel(distance, MAX_SPEED_MM_S / 2, ACCELERATION_MM_S2)
                response = "DONE\n"
            except (IndexError, ValueError):
                response = "ERROR: Invalid MOVE_REL format.\n"

        if response and writer:
            writer.write(response.encode())
            await writer.drain()

async def main():
    import network
    host_ip = network.WLAN(network.STA_IF).ifconfig()[0]
    server_task = uasyncio.create_task(tcp_server(host_ip, 8899))
    processor_task = uasyncio.create_task(command_processor())
    print("ESP32 Z-Axis Controller Ready.")
    await uasyncio.gather(server_task, processor_task)

# --- 6. 程式入口 ---
if __name__ == "__main__":
    try:
        uasyncio.run(main())
    except KeyboardInterrupt:
        print("Program stopped.")