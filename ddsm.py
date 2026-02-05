import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox
import serial
import serial.tools.list_ports
import json
import threading
import time
from datetime import datetime

class DDSMControlApp:
    def __init__(self, root):
        self.root = root
        self.root.title("DDSM115 馬達控制介面 (v5 - 按鈕版)")
        self.root.geometry("600x750") 

        # 序列埠設定
        self.ser = None
        self.is_connected = False
        self.baud_rate = 115200 

        # 建立 UI 元件
        self.create_widgets()
        
        # 啟動自動刷新序列埠
        self.refresh_ports()
        
        # 啟動背景監控 (USB 拔除偵測)
        self.monitor_connection()

    def create_widgets(self):
        # --- 區域 1: 連線設定 ---
        connection_frame = ttk.LabelFrame(self.root, text="1. 連線設定 (Connection)")
        connection_frame.pack(pady=5, padx=10, fill="x")

        ttk.Label(connection_frame, text="序列埠:").pack(side="left", padx=5)
        
        self.port_combobox = ttk.Combobox(connection_frame, state="readonly", width=25)
        self.port_combobox.pack(side="left", padx=5)
        self.port_combobox.bind("<<ComboboxSelected>>", self.on_port_select)

        self.btn_refresh = ttk.Button(connection_frame, text="重整", command=self.refresh_ports, width=6)
        self.btn_refresh.pack(side="left", padx=5)

        self.btn_connect = ttk.Button(connection_frame, text="連線", command=self.toggle_connection)
        self.btn_connect.pack(side="left", padx=5)

        # --- 區域 2: 安全設定 (Heartbeat) ---
        safety_frame = ttk.LabelFrame(self.root, text="2. 安全設定 (Safety)")
        safety_frame.pack(pady=5, padx=10, fill="x")

        btn_hb = ttk.Button(
            safety_frame, 
            text="設定 Heartbeat 為 1秒", 
            command=self.set_heartbeat_1s,
            width=25
        )
        btn_hb.pack(side="left", padx=10, pady=5)
        
        ttk.Label(safety_frame, text="(無訊號1秒後自動停機)", foreground="blue").pack(side="left")

        # --- 區域 3: 馬達 ID 設定 ---
        id_frame = ttk.LabelFrame(self.root, text="3. 修改馬達 ID (Write ID)")
        id_frame.pack(pady=5, padx=10, fill="x")

        # 讀取 ID
        btn_read_id = ttk.Button(id_frame, text="讀取當前 ID", command=self.read_motor_id)
        btn_read_id.pack(pady=(5, 5))

        ttk.Separator(id_frame, orient='horizontal').pack(fill='x', pady=5)

        # 快速寫入 ID 按鈕列
        write_frame = ttk.Frame(id_frame)
        write_frame.pack(pady=(0, 5))
        ttk.Label(write_frame, text="將 ID 修改為:").pack(side="left", padx=5)
        
        for i in range(1, 5):
            btn = ttk.Button(
                write_frame, 
                text=f"ID {i}", 
                width=6,
                command=lambda target=i: self.write_motor_id(target)
            )
            btn.pack(side="left", padx=5)

        # --- 區域 4: 試轉與停止 (改為按鈕式) ---
        motion_frame = ttk.LabelFrame(self.root, text="4. 試轉操作 (Motion Control)")
        motion_frame.pack(pady=5, padx=10, fill="x")

        # 速度設定
        speed_frame = ttk.Frame(motion_frame)
        speed_frame.pack(pady=5)
        ttk.Label(speed_frame, text="目標速度 (RPM):").pack(side="left", padx=5)
        self.entry_speed = ttk.Entry(speed_frame, width=8)
        self.entry_speed.insert(0, "50")
        self.entry_speed.pack(side="left", padx=5)

        ttk.Separator(motion_frame, orient='horizontal').pack(fill='x', pady=5)

        # 啟動按鈕列
        run_frame = ttk.Frame(motion_frame)
        run_frame.pack(pady=5)
        ttk.Label(run_frame, text="啟動 (Run):", foreground="green", width=12).pack(side="left", padx=5)
        
        for i in range(1, 5):
            btn = ttk.Button(
                run_frame, 
                text=f"Run {i}", 
                width=6,
                command=lambda target=i: self.run_motor(target)
            )
            btn.pack(side="left", padx=5)

        # 四馬達同時 80 RPM
        all_run_frame = ttk.Frame(motion_frame)
        all_run_frame.pack(pady=5)
        ttk.Label(all_run_frame, text="群控 (All):", foreground="green", width=12).pack(side="left", padx=5)
        btn_all_80 = ttk.Button(
            all_run_frame,
            text="Run 1-4 @80",
            width=15,
            command=self.run_all_80rpm
        )
        btn_all_80.pack(side="left", padx=5)

        # 停止按鈕列
        stop_frame = ttk.Frame(motion_frame)
        stop_frame.pack(pady=5)
        ttk.Label(stop_frame, text="停止 (Stop):", foreground="red", width=12).pack(side="left", padx=5)
        
        for i in range(1, 5):
            btn = ttk.Button(
                stop_frame, 
                text=f"Stop {i}", 
                width=6,
                command=lambda target=i: self.stop_motor(target)
            )
            btn.pack(side="left", padx=5)

        # --- 區域 5: 訊息日誌 ---
        log_frame = ttk.LabelFrame(self.root, text="通訊日誌 (Log)")
        log_frame.pack(pady=5, padx=10, fill="both", expand=True)

        self.log_text = scrolledtext.ScrolledText(log_frame, state='disabled', height=10)
        self.log_text.pack(fill="both", expand=True, padx=5, pady=5)

    # --- 核心通訊功能 ---
    def refresh_ports(self):
        ports = serial.tools.list_ports.comports()
        port_list = [port.device for port in ports]
        self.port_combobox['values'] = port_list
        if port_list:
            if self.port_combobox.get() not in port_list:
                self.port_combobox.current(0)
        else:
            self.port_combobox.set("")

    def on_port_select(self, event):
        if self.is_connected:
            self.disconnect_serial()

    def toggle_connection(self):
        if not self.is_connected:
            self.connect_serial()
        else:
            self.disconnect_serial()

    def connect_serial(self):
        port = self.port_combobox.get()
        if not port:
            messagebox.showwarning("警告", "請先選擇序列埠")
            return
        try:
            self.ser = serial.Serial(port, self.baud_rate, timeout=1)
            self.is_connected = True
            self.btn_connect.config(text="斷線")
            self.log_message(f"已連線至 {port}")
            self.read_thread = threading.Thread(target=self.read_serial_loop, daemon=True)
            self.read_thread.start()
        except serial.SerialException as e:
            messagebox.showerror("錯誤", f"無法連線: {e}")

    def disconnect_serial(self):
        if self.ser and self.ser.is_open:
            try:
                self.ser.close()
            except:
                pass
        self.ser = None
        self.is_connected = False
        self.btn_connect.config(text="連線")
        self.log_message("已斷線")

    def monitor_connection(self):
        """監控 USB 是否被拔除"""
        if self.is_connected and self.ser:
            try:
                current_port = self.ser.port
                ports = [p.device for p in serial.tools.list_ports.comports()]
                if current_port not in ports:
                    self.log_message("USB 斷線，重置狀態...")
                    self.disconnect_serial()
            except Exception:
                self.disconnect_serial()
        self.root.after(1000, self.monitor_connection)

    def read_serial_loop(self):
        while self.is_connected and self.ser and self.ser.is_open:
            try:
                if self.ser.in_waiting:
                    line = self.ser.readline().decode('utf-8', errors='ignore').strip()
                    if line:
                        self.log_message(f"[RX] {line}")
            except serial.SerialException:
                self.root.after(0, self.disconnect_serial)
                break
            except Exception:
                break
            time.sleep(0.01)

    def send_json(self, command_dict):
        if not self.is_connected or not self.ser:
            messagebox.showwarning("警告", "尚未連線")
            return
        try:
            json_str = json.dumps(command_dict)
            cmd_bytes = (json_str + '\n').encode('utf-8')
            self.ser.write(cmd_bytes)
            self.log_message(f"[TX] {json_str}")
            time.sleep(0.01)  # 小延遲確保傳送完成
        except serial.SerialException:
            self.disconnect_serial()
            messagebox.showerror("錯誤", "發送失敗 (裝置斷開)")

    # --- 控制指令區 ---
    
    def set_heartbeat_1s(self):
        cmd = {"T": 11001, "time": 1000}
        self.send_json(cmd)
        self.log_message("設定: Heartbeat = 1秒")

    def read_motor_id(self):
        self.send_json({"T": 10031})

    def write_motor_id(self, target_id):
        self.send_json({"T": 10011, "id": int(target_id)})

    def run_motor(self, target_id):
        """試轉指定 ID"""
        try:
            speed = int(self.entry_speed.get())
            # DDSM115 cmd=RPM
            cmd = {"T": 10010, "id": int(target_id), "cmd": speed, "act": 0}
            self.send_json(cmd)
        except ValueError:
            messagebox.showerror("錯誤", "速度必須是數字")

    def run_all_80rpm(self):
        """同時讓 1~4 號馬達以 80 RPM 運轉"""
        for target_id in range(1, 5):
            cmd = {"T": 10010, "id": int(target_id), "cmd": 100, "act": 0}
            self.send_json(cmd)

    def stop_motor(self, target_id):
        """停止指定 ID"""
        cmd = {"T": 10010, "id": int(target_id), "cmd": 0, "act": 0}
        self.send_json(cmd)

    def log_message(self, msg):
        timestamp = datetime.now().strftime("%H:%M:%S.%f")[:-3]
        formatted_msg = f"[{timestamp}] {msg}\n"
        self.root.after(0, self._append_log, formatted_msg)

    def _append_log(self, msg):
        self.log_text.config(state='normal')
        self.log_text.insert(tk.END, msg)
        self.log_text.see(tk.END)
        self.log_text.config(state='disabled')

if __name__ == "__main__":
    root = tk.Tk()
    app = DDSMControlApp(root)
    root.mainloop()