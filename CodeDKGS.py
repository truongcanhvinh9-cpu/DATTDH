import cv2
import customtkinter as ctk
from PIL import Image, ImageTk
from tkinter import messagebox, ttk
import threading
import time
import sqlite3
from datetime import datetime
from pyzbar.pyzbar import decode
from openpyxl import Workbook
from tkinter import filedialog
from snap7.util import set_bool
try:
    import snap7
    from snap7.util import set_int
    SNAP7_AVAILABLE = True
except ImportError:
    SNAP7_AVAILABLE = False

DB_FILE = "product_monitor.db"
SAME_PRODUCT_DELAY = 4.0
DIFF_PRODUCT_DELAY = 1.5
PRODUCT_MAPPING = {
    "tayho": 1,
    "caugiay": 2,
    "bacninh": 3
}

def init_db():
    try:
        conn = sqlite3.connect(DB_FILE)
        cur = conn.cursor()
        cur.execute("""CREATE TABLE IF NOT EXISTS scan_history (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            product_code TEXT,
            plc_value INTEGER,
            timestamp TEXT
        )""")
        conn.commit()
    except Exception as e:
        print(f"DB Init Error: {e}")
    finally:
        if 'conn' in locals(): conn.close()

def log_scan_db(product_code, plc_value):
    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    cur.execute("INSERT INTO scan_history (product_code, plc_value, timestamp) VALUES (?, ?, ?)",
                (product_code, plc_value, ts))
    conn.commit()
    conn.close()

def fetch_history_data():
    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()
    cur.execute("SELECT timestamp, product_code, plc_value FROM scan_history ORDER BY id DESC LIMIT 50")
    rows = cur.fetchall()
    conn.close()
    return rows

def get_product_counts_db():
    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()
    cur.execute("SELECT product_code, COUNT(*) FROM scan_history GROUP BY product_code")
    rows = cur.fetchall()
    conn.close()
    counts = {"tayho": 0, "caugiay": 0, "bacninh": 0}
    for row in rows:
        if row[0] in counts:
            counts[row[0]] = row[1]
    return counts

def clear_history_db():
    try:
        conn = sqlite3.connect(DB_FILE)
        cur = conn.cursor()
        cur.execute("DELETE FROM scan_history")
        conn.commit()
        conn.close()
        return True
    except Exception as e:
        print(e)
        return False

class ProductScannerApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("HỆ THỐNG ĐIỀU KHIỂN GIÁM SÁT VÀ PHÂN LOẠI BƯU KIỆN QR CODE")
        self.geometry("1200x720")
        ctk.set_appearance_mode("Dark")
        init_db()
        self.stop_event = threading.Event()
        self.is_camera_running = False
        self.thread_cam = None
        self.last_scan_time = 0
        self.last_scanned_code = None
        self.plc_connected = False
        if SNAP7_AVAILABLE:
            self.plc = snap7.client.Client()
        else:
            self.plc = None
        self.setup_ui()

    def setup_ui(self):
        self.grid_rowconfigure(0, weight=0)
        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(0, weight=3)
        self.grid_columnconfigure(1, weight=2)

        header_frame = ctk.CTkFrame(self, corner_radius=0, fg_color="#1E1E1E")
        header_frame.grid(row=0, column=0, columnspan=2, sticky="ew", ipadx=10, ipady=10)
        ctk.CTkLabel(header_frame, text="ĐỒ ÁN", font=("Segoe UI", 16, "bold"), text_color="#F39C12").pack(pady=(5, 0))
        ctk.CTkLabel(header_frame, text="HỆ THỐNG PHÂN LOẠI BƯU KIỆN THEO MÃ QR", font=("Segoe UI", 24, "bold"), text_color="#00BFFF").pack(pady=(0, 5))

        left_frame = ctk.CTkFrame(self, corner_radius=10)
        left_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=10)
        ctk.CTkLabel(left_frame, text="MÃ SẢN PHẨM VỪA QUÉT", font=("Segoe UI", 16, "bold")).pack(pady=(15, 0))
        self.lbl_product = ctk.CTkLabel(left_frame, text="SẴN SÀNG", font=("Consolas", 48, "bold"), text_color="#F39C12")
        self.lbl_product.pack(pady=5)
        self.lbl_plc_sent = ctk.CTkLabel(left_frame, text="Trạng thái PLC: Chờ...", font=("Arial", 14), text_color="gray")
        self.lbl_plc_sent.pack(pady=(0, 10))

        plc_frame = ctk.CTkFrame(left_frame, fg_color="#333", corner_radius=8)
        plc_frame.pack(fill="x", padx=20, pady=5, ipady=5)
        ctk.CTkLabel(plc_frame, text="IP PLC S7-1200:").pack(side="left", padx=10)
        self.entry_plc_ip = ctk.CTkEntry(plc_frame, width=120)
        self.entry_plc_ip.insert(0, "192.168.0.1")
        self.entry_plc_ip.pack(side="left", padx=5)
        self.btn_plc_connect = ctk.CTkButton(plc_frame, text="KẾT NỐI", width=80, fg_color="#1F6AA5", command=self.toggle_plc)
        self.btn_plc_connect.pack(side="left", padx=5)
        self.lbl_plc_status = ctk.CTkLabel(plc_frame, text="🔴 Ngắt kết nối", text_color="#FF4C4C", font=("Arial", 12, "bold"))
        self.lbl_plc_status.pack(side="right", padx=15)

        cam_frame = ctk.CTkFrame(left_frame, fg_color="transparent")
        cam_frame.pack(pady=10)
        ctk.CTkLabel(cam_frame, text="Camera ID:").pack(side="left", padx=5)
        self.cbo_camera = ctk.CTkComboBox(cam_frame, values=["0", "1", "2"], width=70)
        self.cbo_camera.set("0")
        self.cbo_camera.pack(side="left", padx=5)
        self.btn_toggle_cam = ctk.CTkButton(cam_frame, text="BẬT CAMERA", width=120, fg_color="#2E8B57", command=self.toggle_camera)
        self.btn_toggle_cam.pack(side="left", padx=5)
        control_frame = ctk.CTkFrame(left_frame, fg_color="transparent")
        control_frame.pack(pady=10)

        self.btn_start = ctk.CTkButton(
            control_frame,
            text="START",
            width=120,
            fg_color="#27AE60",
            command=self.start_system
        )

        self.btn_start.pack(side="left", padx=10)

        self.btn_stop = ctk.CTkButton(
            control_frame,
            text="STOP",
            width=120,
            fg_color="#C0392B",
            command=self.stop_system
        )

        self.btn_stop.pack(side="left", padx=10)
        self.video_label = ctk.CTkLabel(left_frame, text="[CAMERA OFF]", fg_color="#111", corner_radius=8)
        self.video_label.pack(fill="both", expand=True, padx=20, pady=10)

        right_frame = ctk.CTkFrame(self, corner_radius=10)
        right_frame.grid(row=1, column=1, sticky="nsew", padx=(0, 10), pady=10)
        ctk.CTkLabel(right_frame, text="THỐNG KÊ SẢN PHẨM", font=("Segoe UI", 16, "bold")).pack(pady=(15, 10))

        counters_frame = ctk.CTkFrame(right_frame, fg_color="transparent")
        counters_frame.pack(fill="x", padx=10, pady=5)
        counters_frame.grid_columnconfigure((0, 1, 2), weight=1)
        self.lbl_tayho_count = self.create_counter_box(counters_frame, "Tây Hồ", "#27AE60", 0)
        self.lbl_caugiay_count = self.create_counter_box(counters_frame, "Cầu Giấy", "#2980B9", 1)
        self.lbl_bacninh_count = self.create_counter_box(counters_frame, "Bắc Ninh", "#D35400", 2)

        ctk.CTkFrame(right_frame, height=2, fg_color="#444").pack(fill="x", padx=20, pady=15)
        ctk.CTkLabel(right_frame, text="LỊCH SỬ PHÂN LOẠI", font=("Segoe UI", 16, "bold")).pack(pady=(0, 5))

        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Treeview", background="#2B2B2B", foreground="white", fieldbackground="#2B2B2B", rowheight=28, font=("Arial", 11))
        style.map("Treeview", background=[('selected', '#1F6AA5')])
        style.configure("Treeview.Heading", background="#444", foreground="white", font=("Arial", 11, "bold"))

        self.tree = ttk.Treeview(right_frame, columns=("time", "product", "plc_val"), show="headings")
        self.tree.heading("time", text="Thời gian")
        self.tree.column("time", width=130, anchor="center")
        self.tree.heading("product", text="Mã SP")
        self.tree.column("product", width=80, anchor="center")
        self.tree.heading("plc_val", text="Giá trị PLC")
        self.tree.column("plc_val", width=80, anchor="center")
        self.tree.pack(fill="both", expand=True, padx=15, pady=5)
        button_frame = ctk.CTkFrame(right_frame, fg_color="transparent")
        button_frame.pack(pady=15)

        
        ctk.CTkButton(
            button_frame,
            text="Xuất Excel",
            fg_color="white",
            text_color="black",
            width=150,
            command=self.export_to_excel
        ).pack(side="left", padx=10)

        
        ctk.CTkButton(
            button_frame,
            text="Xóa lịch sử",
            fg_color="#C0392B",
            width=150,
            command=self.action_clear_history
        ).pack(side="left", padx=10)
        self.refresh_ui_data()
       
    def create_counter_box(self, parent, title, color, col):
        box = ctk.CTkFrame(parent, fg_color=color, corner_radius=8)
        box.grid(row=0, column=col, padx=5, sticky="nsew")
        ctk.CTkLabel(box, text=title, font=("Arial", 14, "bold"), text_color="white").pack(pady=(5, 0))
        lbl_count = ctk.CTkLabel(box, text="0", font=("Consolas", 24, "bold"), text_color="white")
        lbl_count.pack(pady=(0, 5))
        return lbl_count

    def toggle_plc(self):
        if not SNAP7_AVAILABLE:
            messagebox.showerror("Lỗi", "Chưa cài đặt thư viện python-snap7!")
            return
        if self.plc_connected:
            try:
                self.plc.disconnect()
            except:
                pass
            self.plc_connected = False
            self.btn_plc_connect.configure(text="KẾT NỐI", fg_color="#1F6AA5")
            self.lbl_plc_status.configure(text="🔴 Ngắt kết nối", text_color="#FF4C4C")
        else:
            ip = self.entry_plc_ip.get()
            try:
                self.plc.connect(ip, 0, 1)
                self.plc_connected = True
                self.btn_plc_connect.configure(text="NGẮT KN", fg_color="#C0392B")
                self.lbl_plc_status.configure(text="🟢 Đã kết nối", text_color="#00FF00")
            except Exception as e:
                messagebox.showerror("Lỗi PLC", f"Không thể kết nối S7-1200 tại {ip}.\nLý do: {e}")

    def write_int_to_plc(self, db_number, byte_offset, value):
        if not self.plc_connected or self.plc is None: return False
        try:
            data = bytearray(2)
            set_int(data, 0, value)
            self.plc.db_write(db_number, byte_offset, data)
            return True
        except Exception:
            self.lbl_plc_status.configure(text="🔴 Lỗi truyền thông", text_color="#FF4C4C")
            return False
    def write_bool_to_plc(self, db_number, byte_offset, bit_offset, value):
        if not self.plc_connected or self.plc is None:
           return False

        try:
            data = self.plc.db_read(db_number, byte_offset, 1)

            set_bool(data, 0, bit_offset, value)

            self.plc.db_write(db_number, byte_offset, data)

            return True

        except Exception as e:
            print(e)
            return False
    def start_system(self):
        success = self.write_bool_to_plc(
            db_number=1,
            byte_offset=0,
            bit_offset=0,
            value=True
        )

        if success:
            self.lbl_plc_sent.configure(
                text="HỆ THỐNG ĐANG CHẠY",
                text_color="#00FF00"
            )
            self.after(
                1000,
                lambda: self.write_bool_to_plc(
                    db_number=1,
                    byte_offset=0,
                    bit_offset=0,
                    value=False
                )
            )
    def stop_system(self):
        success = self.write_bool_to_plc(
            db_number=1,
            byte_offset=0,
            bit_offset=1,
            value=True
        )

        if success:
            self.lbl_plc_sent.configure(
                text="HỆ THỐNG ĐÃ DỪNG",
                text_color="#FF4C4C"
           )

            self.after(
               1000,
               lambda: self.write_bool_to_plc(
                   db_number=1,
                   byte_offset=0,
                   bit_offset=1,
                   value=False
               )
           )
           
    def toggle_camera(self):
        if self.is_camera_running:
            self.is_camera_running = False
            self.stop_event.set()
            self.btn_toggle_cam.configure(state="disabled", text="⏳ ĐANG TẮT...", fg_color="#F39C12")
            threading.Thread(target=self._wait_for_cam_stop, daemon=True).start()
        else:
            self.is_camera_running = True
            self.stop_event.clear()
            self.btn_toggle_cam.configure(state="disabled", text="⏳ ĐANG BẬT...", fg_color="#F39C12")
            self.thread_cam = threading.Thread(target=self.video_loop, daemon=True)
            self.thread_cam.start()

    def _wait_for_cam_stop(self):
        if self.thread_cam is not None:
            self.thread_cam.join(timeout=3.0)
        self.after(0, self._reset_cam_ui)

    def _reset_cam_ui(self):
        self.video_label.configure(image="", text="[CAMERA OFF]")
        self.btn_toggle_cam.configure(state="normal", text="BẬT CAMERA", fg_color="#2E8B57")

    def video_loop(self):
        cam_id = int(self.cbo_camera.get())
        cap = cv2.VideoCapture(cam_id, cv2.CAP_DSHOW)
        if not cap.isOpened():
            self.is_camera_running = False
            self.after(0, lambda: messagebox.showerror("Lỗi", f"Không thể mở Camera ID: {cam_id}!"))
            self.after(0, self._reset_cam_ui)
            return
        self.after(0, lambda: self.btn_toggle_cam.configure(state="normal", text="TẮT CAMERA", fg_color="#C0392B"))
        while not self.stop_event.is_set():
            ret, frame = cap.read()
            if not ret: break
            decoded_objects = decode(frame)
            for obj in decoded_objects:
                points = obj.polygon
                if len(points) == 4:
                    pts = [(p.x, p.y) for p in points]
                    for i in range(4):
                        cv2.line(frame, pts[i], pts[(i + 1) % 4], (0, 255, 0), 3)
                qr_text = obj.data.decode('utf-8').strip()
                if qr_text in PRODUCT_MAPPING:
                    self.process_product(qr_text)
                else:
                     self.after(0, lambda:
                      self.lbl_product.configure(
                        text="QR KHÔNG HỢP LỆ",
                        text_color="#FF0000"
                        )
                    )
            img_rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
            img = Image.fromarray(img_rgb)
            label_width = self.video_label.winfo_width()
            label_height = self.video_label.winfo_height()
            if label_width > 10 and label_height > 10:
                img.thumbnail((label_width, label_height))
            imgtk = ImageTk.PhotoImage(image=img)
            if not self.stop_event.is_set():
                self.after(0, self._update_video_label, imgtk)
            time.sleep(0.015)
        cap.release()

    def _update_video_label(self, imgtk):
        if self.winfo_exists() and self.is_camera_running:
            self.video_label.configure(image=imgtk, text="")
            self.video_label.image = imgtk

    def process_product(self, product_code):
        current_time = time.time()
        if product_code == self.last_scanned_code:
            if current_time - self.last_scan_time < SAME_PRODUCT_DELAY:
                return
        else:
            if current_time - self.last_scan_time < DIFF_PRODUCT_DELAY:
                return
        self.last_scan_time = current_time
        self.last_scanned_code = product_code
        plc_value = PRODUCT_MAPPING[product_code]
        write_success = self.write_int_to_plc(db_number=2, byte_offset=0, value=plc_value)
        log_scan_db(product_code, plc_value)
        if self.winfo_exists():
            self.after(0, lambda: self.lbl_product.configure(text=f"ĐÃ QUÉT: {product_code}", text_color="#00FF00"))
            status_text = f"Đã ghi DB2 = {plc_value}" if write_success else f"Chưa nối PLC (Giá trị ảo: {plc_value})"
            color = "#00FF00" if write_success else "#F39C12"
            self.after(0, lambda: self.lbl_plc_sent.configure(text=status_text, text_color=color))
            self.after(0, self.refresh_ui_data)
            self.after(int(DIFF_PRODUCT_DELAY * 1000), self._reset_scan_ui_state)
    def export_to_excel(self):
        try:
            data = fetch_history_data()  # lấy dữ liệu từ DB

            if not data:
                messagebox.showwarning("Thông báo", "Không có dữ liệu để xuất!")
                return

            # Chọn nơi lưu file
            file_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel file", "*.xlsx")],
                title="Lưu file Excel"
             )

            if not file_path:
                return

            # Tạo file Excel
            wb = Workbook()
            ws = wb.active
            ws.title = "Scan History"

            # Header
            ws.append(["Thời gian", "Mã sản phẩm", "Giá trị PLC"])

            # Ghi dữ liệu
            for row in data:
                ws.append(row)

            # Thêm thống kê
            counts = get_product_counts_db()
            ws.append([])
            ws.append(["THỐNG KÊ"])
            for k, v in counts.items():
                ws.append([k, v])

            # Lưu file
            wb.save(file_path)

            messagebox.showinfo("Thành công", f"Đã xuất file:\n{file_path}")

        except Exception as e:
            messagebox.showerror("Lỗi", f"Xuất Excel thất bại!\n{e}")
    def _reset_scan_ui_state(self):
        if self.winfo_exists():
            self.lbl_product.configure(text="SẴN SÀNG", text_color="#F39C12")

    def action_clear_history(self):
        if messagebox.askyesno("Xác nhận", "Bạn có chắc chắn muốn xóa toàn bộ dữ liệu thống kê và lịch sử?"):
            if clear_history_db():
                self.refresh_ui_data()

    def refresh_ui_data(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        data = fetch_history_data()
        for row in data:
            self.tree.insert("", "end", values=(row[0], row[1], row[2]))
        counts = get_product_counts_db()
        self.lbl_tayho_count.configure(text=str(counts.get("tayho", 0)))
        self.lbl_caugiay_count.configure(text=str(counts.get("caugiay", 0)))
        self.lbl_bacninh_count.configure(text=str(counts.get("bacninh", 0)))

    def on_closing(self):
        self.stop_event.set()
        if self.plc_connected:
            try:
                self.plc.disconnect()
            except:
                pass
        self.destroy()

if __name__ == "__main__":
    app = ProductScannerApp()
    app.protocol("WM_DELETE_WINDOW", app.on_closing)
    app.mainloop()
