# projector_view.py
# 功能：在指定螢幕上全螢幕顯示圖像，並透過網路監聽指令。

import sys
from multiprocessing.connection import Listener
from PyQt5.QtWidgets import QApplication, QWidget, QLabel, QVBoxLayout
from PyQt5.QtGui import QPixmap, QColor
from PyQt5.QtCore import Qt, QObject, pyqtSignal, QThread


# --- 1. 用於在背景接收指令的監聽器執行緒 ---
class CommandListener(QObject):
    # 定義信號，用於安全地從背景執行緒向主 GUI 執行緒傳遞指令
    command_received = pyqtSignal(dict)

    def __init__(self, address, authkey):
        super().__init__()
        self.address = address
        self.authkey = authkey
        self.is_running = True

    def run(self):
        """監聽網路連線並接收指令"""
        print(f"[Projector] Listening on {self.address}")
        # 使用 Listener 來接收來自 Client (main_gui.py) 的連線
        with Listener(self.address, authkey=self.authkey) as listener:
            with listener.accept() as conn:
                print(f"[Projector] Connection accepted from {listener.last_accepted}")
                while self.is_running:
                    try:
                        # 等待並接收指令
                        msg = conn.recv()
                        print(f"[Projector] Received command: {msg}")
                        # 透過信號發送指令到主執行緒
                        self.command_received.emit(msg)
                        if msg.get('command') == 'close':
                            self.is_running = False
                    except EOFError:
                        print("[Projector] Connection closed by main GUI.")
                        self.is_running = False
                    except Exception as e:
                        print(f"[Projector] Error receiving command: {e}")
                        self.is_running = False
        print("[Projector] Listener thread finished.")


# --- 2. 用於顯示影像的全螢幕視窗 ---
class ProjectorWindow(QWidget):
    def __init__(self):
        super().__init__()
        # 設定 UI
        self.setWindowTitle('Projector View')
        # 設定無邊框屬性
        self.setWindowFlags(Qt.FramelessWindowHint)

        # 使用佈局來管理 QLabel
        layout = QVBoxLayout()
        layout.setContentsMargins(0, 0, 0, 0)
        self.setLayout(layout)

        # 用於顯示圖片的 QLabel
        self.image_label = QLabel()
        self.image_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.image_label)

        # 初始為黑畫面
        self.show_blank()

    def show_image(self, image_path):
        """載入並顯示指定的圖片"""
        pixmap = QPixmap(image_path)
        self.image_label.setPixmap(pixmap)
        print(f"[Projector] Displaying image: {image_path}")

    def show_blank(self):
        """顯示黑畫面"""
        # 清除圖片即可，因為背景是黑的
        self.image_label.clear()
        print("[Projector] Displaying blank screen.")


# --- 3. 主程式邏輯 ---
if __name__ == '__main__':
    # 從命令列讀取參數
    if len(sys.argv) != 5:
        print("Usage: python projector_view.py <monitor_index> <host> <port> <authkey>")
        sys.exit(1)

    monitor_index = int(sys.argv[1])
    host = sys.argv[2]
    port = int(sys.argv[3])
    authkey = sys.argv[4].encode()

    app = QApplication(sys.argv)

    # 檢查顯示器是否存在
    screens = app.screens()
    if monitor_index >= len(screens):
        print(f"Error: Monitor index {monitor_index} is out of range. Available monitors: {len(screens)}")
        sys.exit(1)

    # 創建視窗
    window = ProjectorWindow()

    # 將視窗移動到指定的顯示器並全螢幕顯示
    screen = screens[monitor_index]
    window.move(screen.geometry().x(), screen.geometry().y())
    window.showFullScreen()

    # 創建並啟動背景監聽執行緒
    listener_thread = QThread()
    command_listener = CommandListener(address=(host, port), authkey=authkey)
    command_listener.moveToThread(listener_thread)

    # 連接信號與槽
    listener_thread.started.connect(command_listener.run)
    command_listener.command_received.connect(
        lambda msg: {
            'show': lambda: window.show_image(msg['path']),
            'blank': window.show_blank,
            'close': app.quit
        }.get(msg.get('command'), lambda: print(f"Unknown command: {msg}"))()
    )
    # 監聽執行緒結束後也退出程式
    listener_thread.finished.connect(app.quit)

    listener_thread.start()

    print(f"[Projector] GUI started on monitor {monitor_index}. Waiting for commands...")
    sys.exit(app.exec_())