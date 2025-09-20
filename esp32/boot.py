# boot.py - 僅用於連接 Wi-Fi
import network

# --- 請修改這裡的 Wi-Fi 名稱和密碼 ---
ssid = 'YOUR_WIFI_SSID'
password = 'YOUR_WIFI_PASSWORD'
# ------------------------------------

station = network.WLAN(network.STA_IF)

if not station.isconnected():
    print(f"正在連接到 Wi-Fi: {ssid}...")
    station.active(True)
    station.connect(ssid, password)
    while not station.isconnected():
        pass

print("Wi-Fi 連接成功！")
print('網路配置:', station.ifconfig())