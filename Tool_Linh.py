import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import qrcode
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
import os
import threading
import datetime
import socket
from PIL import Image, ImageTk
import logging
import csv

# Tắt log mặc định của Flask
log = logging.getLogger('werkzeug')
log.setLevel(logging.ERROR)

from flask import Flask, request, jsonify

app_flask = Flask(__name__)
gui_app = None

# Giao diện Web Mobile
HTML_TEMPLATE = """
<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>Trạm Quét QR</title>
    <script src="https://unpkg.com/html5-qrcode"></script>
    <style>
        body { font-family: Arial, sans-serif; text-align: center; margin: 0; padding: 15px; background: #f4f4f9;}
        #reader { width: 100%; max-width: 500px; margin: 0 auto; border-radius: 8px; overflow: hidden; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }
        #result { margin-top: 20px; padding: 15px; border-radius: 5px; font-weight: bold; font-size: 1.2em; }
        .success { background-color: #d4edda; color: #155724; border: 1px solid #c3e6cb; }
        .error { background-color: #f8d7da; color: #721c24; border: 1px solid #f5c6cb; }
        .warning { background-color: #fff3cd; color: #856404; border: 1px solid #ffeeba; }
    </style>
</head>
<body>
    <h2>ĐIỂM DANH SỰ KIỆN</h2>
    <div id="reader"></div>
    <div id="result">Đang chờ quét mã...</div>

    <script>
        const html5QrCode = new Html5Qrcode("reader");
        let isScanning = true;

        function onScanSuccess(decodedText, decodedResult) {
            if (!isScanning) return;
            isScanning = false; 
            
            let resDiv = document.getElementById('result');
            resDiv.innerHTML = "Đang xử lý...";
            resDiv.className = "warning";

            fetch('/api/checkin', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ qr_data: decodedText }) // Gửi nguyên chuỗi QR
            })
            .then(response => response.json())
            .then(data => {
                if(data.status === 'success'){
                    resDiv.innerHTML = "✅ Thành công:<br>" + data.name;
                    resDiv.className = "success";
                } else if (data.status === 'exists') {
                    resDiv.innerHTML = "⚠️ Đã điểm danh trước đó:<br>" + data.name;
                    resDiv.className = "warning";
                } else {
                    resDiv.innerHTML = "❌ Lỗi: " + data.message;
                    resDiv.className = "error";
                }
                setTimeout(() => { isScanning = true; }, 1500); 
            })
            .catch(err => {
                resDiv.innerHTML = "❌ Mất kết nối máy chủ!";
                resDiv.className = "error";
                setTimeout(() => { isScanning = true; }, 2000);
            });
        }

        html5QrCode.start(
            { facingMode: "environment" },
            { fps: 15, qrbox: { width: 250, height: 250 } },
            onScanSuccess
        ).catch(err => {
            document.getElementById('result').innerHTML = "Lỗi Camera: Hãy dùng HTTPS và cấp quyền Camera.";
            document.getElementById('result').className = "error";
        });
    </script>
</body>
</html>
"""

@app_flask.route('/')
def index():
    return HTML_TEMPLATE

@app_flask.route('/api/checkin', methods=['POST'])
def api_checkin():
    data = request.json
    qr_data = data.get('qr_data', '').strip()
    if not qr_data:
        return jsonify({'status': 'error', 'message': 'Mã QR rỗng'})
    
    if gui_app:
        result = gui_app.xu_ly_quet_mobile(qr_data)
        return jsonify(result)
    return jsonify({'status': 'error', 'message': 'Tool chưa sẵn sàng'})

# ================= CLASS TKINTER CHÍNH =================
class QREventApp:
    def __init__(self, root):
        global gui_app
        gui_app = self
        
        self.root = root
        self.root.title("Công Cụ Quản Lý Sự Kiện QR (Pro Version)")
        self.root.geometry("850x750")
        
        self.df_dangki = pd.DataFrame()
        self.danh_sach_tham_gia = []
        self.server_running = False
        
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(expand=True, fill='both', padx=10, pady=10)
        
        self.tab1 = ttk.Frame(self.notebook)
        self.tab2 = ttk.Frame(self.notebook)
        self.tab3 = ttk.Frame(self.notebook)
        
        self.notebook.add(self.tab1, text='1. Tạo & Gửi Mail Nhóm')
        self.notebook.add(self.tab2, text='2. Điểm Danh (Real-time Backup)')
        self.notebook.add(self.tab3, text='3. Đối Soát Kép')
        
        self.setup_tab1()
        self.setup_tab2()
        self.setup_tab3()

    # --- TAB 1 (TẠO VÀ GỬI EMAIL NHÓM KÈM HTML) ---
    def setup_tab1(self):
        frame = ttk.LabelFrame(self.tab1, text="Cấu hình dữ liệu đầu vào", padding=10)
        frame.pack(fill='x', padx=10, pady=10)

        tk.Button(frame, text="Chọn File Excel Đăng Ký", command=self.tai_file_excel).grid(row=0, column=0, pady=5, sticky='w')
        self.lbl_file_status = tk.Label(frame, text="Chưa chọn file", fg="red")
        self.lbl_file_status.grid(row=0, column=1, padx=10)

        tk.Label(frame, text="Cột Họ Tên:").grid(row=1, column=0, sticky='w', pady=5)
        self.cb_hoten = ttk.Combobox(frame, state="readonly")
        self.cb_hoten.grid(row=1, column=1, pady=5)

        tk.Label(frame, text="Cột Email (Khóa Nhóm):").grid(row=2, column=0, sticky='w', pady=5)
        self.cb_email = ttk.Combobox(frame, state="readonly")
        self.cb_email.grid(row=2, column=1, pady=5)

        tk.Label(frame, text="Cột Số Điện Thoại:").grid(row=3, column=0, sticky='w', pady=5)
        self.cb_sdt = ttk.Combobox(frame, state="readonly")
        self.cb_sdt.grid(row=3, column=1, pady=5)

        frame_email = ttk.LabelFrame(self.tab1, text="Tài khoản gửi Email (Gmail)", padding=10)
        frame_email.pack(fill='x', padx=10, pady=5)

        tk.Label(frame_email, text="Email gửi:").grid(row=0, column=0, sticky='w')
        self.entry_email_gui = tk.Entry(frame_email, width=40)
        self.entry_email_gui.grid(row=0, column=1, padx=5, pady=5)

        tk.Label(frame_email, text="App Password:").grid(row=1, column=0, sticky='w')
        self.entry_pass = tk.Entry(frame_email, width=40, show="*")
        self.entry_pass.grid(row=1, column=1, padx=5, pady=5)

        self.btn_send = tk.Button(self.tab1, text="Bắt đầu tạo và gửi QR (Nhóm)", bg="green", fg="white", font=("Arial", 10, "bold"), command=self.bat_dau_gui_email)
        self.btn_send.pack(pady=15)

    def tai_file_excel(self):
        filepath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if filepath:
            try:
                self.df_dangki = pd.read_excel(filepath)
                columns = list(self.df_dangki.columns)
                self.cb_hoten['values'] = columns
                self.cb_email['values'] = columns
                self.cb_sdt['values'] = columns
                
                if columns: self.cb_hoten.set(columns[0])
                if len(columns) > 1: self.cb_email.set(columns[1])
                if len(columns) > 2: self.cb_sdt.set(columns[2])
                
                self.lbl_file_status.config(text=f"Đã tải: {os.path.basename(filepath)}", fg="green")
            except Exception as e:
                messagebox.showerror("Lỗi", f"Không thể đọc file: {e}")

    def bat_dau_gui_email(self):
        col_hoten = self.cb_hoten.get()
        col_email = self.cb_email.get()
        col_sdt = self.cb_sdt.get()
        email_gui = self.entry_email_gui.get()
        password = self.entry_pass.get()

        if self.df_dangki.empty or not col_hoten or not col_email or not email_gui or not password:
            messagebox.showwarning("Cảnh báo", "Vui lòng điền đủ thông tin.")
            return

        self.btn_send.config(state="disabled", text="Đang xử lý Form HTML...")
        threading.Thread(target=self.tien_hanh_gui, args=(col_hoten, col_email, col_sdt, email_gui, password), daemon=True).start()

    def tien_hanh_gui(self, col_hoten, col_email, col_sdt, email_gui, password):
        try:
            server = smtplib.SMTP('smtp.gmail.com', 587)
            server.starttls()
            server.login(email_gui, password)
            os.makedirs('qr_codes', exist_ok=True)
            
            # GOM NHÓM THEO EMAIL
            grouped = self.df_dangki.groupby(col_email)
            
            for email_nhan, group in grouped:
                msg = MIMEMultipart('related')
                msg['Subject'] = 'Vé Điện Tử Tham Gia Sự Kiện'
                msg['From'] = email_gui
                msg['To'] = str(email_nhan)

                msg_alt = MIMEMultipart('alternative')
                msg.attach(msg_alt)

                # Form Email HTML Chuẩn Chỉnh
                html_content = f"""
                <div style="font-family: Arial, sans-serif; max-width: 600px; margin: auto; border: 1px solid #ddd; border-radius: 8px; overflow: hidden; box-shadow: 0 4px 10px rgba(0,0,0,0.1);">
                    <div style="background-color: #2c3e50; color: white; padding: 20px; text-align: center;">
                        <h2 style="margin: 0; text-transform: uppercase;">Xác nhận đăng ký tham gia</h2>
                    </div>
                    <div style="padding: 20px; color: #333;">
                        <p style="font-size: 16px;">Kính gửi Khách hàng,</p>
                        <p style="font-size: 16px; line-height: 1.5;">Cảm ơn bạn đã đăng ký tham gia sự kiện. Hệ thống ghi nhận bạn đã đăng ký cho <b>{len(group)}</b> người. Dưới đây là các Mã QR dùng để check-in độc lập cho từng người:</p>
                """

                # Duyệt từng người trong nhóm chung 1 Email
                for index, row in group.iterrows():
                    hoten = str(row[col_hoten])
                    sdt = str(row[col_sdt]) if col_sdt else "N/A"
                    
                    # Dữ liệu QR chứa đủ thông tin phân tách bằng |||
                    qr_data = f"{hoten}|||{email_nhan}|||{sdt}"
                    qr = qrcode.make(qr_data)
                    qr_filename = f"qr_codes/qr_{index}.png"
                    qr.save(qr_filename)
                    
                    # Chèn QR nội tuyến
                    html_content += f"""
                        <div style="margin-top: 15px; padding: 15px; border: 2px dashed #3498db; border-radius: 8px; text-align: center; background-color: #f8f9fa;">
                            <h3 style="margin: 0 0 5px 0; color: #2980b9;">Khách mời: {hoten}</h3>
                            <p style="margin: 0 0 10px 0; font-size: 14px; color: #7f8c8d;">SĐT: {sdt}</p>
                            <img src="cid:qr_img_{index}" alt="QR Code" style="width: 200px; height: 200px; border-radius: 5px;">
                        </div>
                    """
                
                html_content += """
                        <p style="margin-top: 20px; font-size: 13px; color: #95a5a6; text-align: center;">* Vui lòng xuất trình mã này tại quầy lễ tân. Xin cảm ơn!</p>
                    </div>
                </div>
                """
                
                msg_alt.attach(MIMEText(html_content, 'html'))
                
                # Đính kèm hình ảnh để HTML có thể gọi ra bằng Content-ID
                for index, row in group.iterrows():
                    with open(f"qr_codes/qr_{index}.png", 'rb') as img_file:
                        img = MIMEImage(img_file.read())
                        img.add_header('Content-ID', f'<qr_img_{index}>')
                        img.add_header('Content-Disposition', 'inline')
                        msg.attach(img)

                server.send_message(msg)
                
            server.quit()
            messagebox.showinfo("Thành công", "Đã gửi Email nhóm kèm Form HTML chuyên nghiệp thành công!")
        except Exception as e:
            messagebox.showerror("Lỗi", f"Gửi thất bại: {e}")
        finally:
            self.btn_send.config(state="normal", text="Bắt đầu tạo và gửi QR (Nhóm)")

    # --- TAB 2 (ĐIỂM DANH AUTO-BACKUP & CHỐNG TRÙNG) ---
    def setup_tab2(self):
        lbl_cbt = tk.Label(self.tab2, text="* Dữ liệu sẽ tự động Backup ra file 'backup_diemdanh_realtime.csv' sau mỗi lượt quét.", fg="green")
        lbl_cbt.pack(pady=5)

        frame_mobile = ttk.LabelFrame(self.tab2, text="Cách 1: Biến Điện thoại thành Máy quét", padding=10)
        frame_mobile.pack(fill='x', padx=10, pady=5)
        
        self.btn_start_server = tk.Button(frame_mobile, text="Khởi động Server Nội Bộ", bg="#17a2b8", fg="white", command=self.khoi_dong_server)
        self.btn_start_server.pack(pady=5)
        
        self.lbl_ip = tk.Label(frame_mobile, text="Nhấn Khởi động để lấy link", font=("Arial", 10, "bold"), fg="blue")
        self.lbl_ip.pack()

        self.lbl_server_qr = tk.Label(frame_mobile)
        self.lbl_server_qr.pack(pady=5)

        frame_usb = ttk.LabelFrame(self.tab2, text="Cách 2: Máy quét USB", padding=10)
        frame_usb.pack(fill='x', padx=10, pady=5)
        
        self.entry_scan = tk.Entry(frame_usb, width=40, font=("Arial", 12))
        self.entry_scan.pack(pady=5)
        self.entry_scan.bind('<Return>', self.xu_ly_ma_quet_usb)

        columns = ('Họ Tên', 'Email', 'Số Điện Thoại', 'Thời Gian')
        self.tree = ttk.Treeview(self.tab2, columns=columns, show='headings', height=8)
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=150)
        self.tree.pack(expand=True, fill='both', padx=10, pady=5)

        tk.Button(self.tab2, text="Xuất file Điểm Danh Hiện Tại", bg="blue", fg="white", command=self.xuat_file_diem_danh).pack(pady=5)

    def get_local_ip(self):
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        try:
            s.connect(('10.255.255.255', 1))
            ip = s.getsockname()[0]
        except: ip = '127.0.0.1'
        finally: s.close()
        return ip

    def khoi_dong_server(self):
        if self.server_running: return
        ip = self.get_local_ip()
        port = 5000
        url = f"https://{ip}:{port}" 
        
        self.lbl_ip.config(text=f"Truy cập web bằng điện thoại: {url}")
        qr_img = qrcode.make(url).resize((150, 150))
        self.qr_photo = ImageTk.PhotoImage(qr_img)
        self.lbl_server_qr.config(image=self.qr_photo)
        
        self.btn_start_server.config(state="disabled", text="Server Đang Chạy")
        self.server_running = True
        threading.Thread(target=lambda: app_flask.run(host='0.0.0.0', port=port, ssl_context='adhoc'), daemon=True).start()

    def xu_ly_quet_mobile(self, qr_text):
        thoi_gian = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        # Bóc tách dữ liệu từ QR
        try:
            parts = qr_text.split('|||')
            if len(parts) >= 3:
                hoten, email, sdt = parts[0], parts[1], parts[2]
            else:
                return {'status': 'error', 'message': 'Mã QR cũ hoặc sai định dạng'}
        except:
            return {'status': 'error', 'message': 'Lỗi đọc mã'}

        # Kiểm tra trùng lặp dựa trên combo (Họ Tên + Email)
        da_quet = [(item['Email'], item['Họ Tên']) for item in self.danh_sach_tham_gia]
        
        if (email, hoten) in da_quet:
            return {'status': 'exists', 'email': email, 'name': hoten}
            
        record = {
            'Họ Tên': hoten,
            'Email': email,
            'Số Điện Thoại': sdt,
            'Thời Gian CheckIn': thoi_gian
        }
        self.danh_sach_tham_gia.append(record)
        
        # AUTO BACKUP: Ghi trực tiếp 1 dòng vào file CSV chống mất dữ liệu
        backup_file = "backup_diemdanh_realtime.csv"
        file_exists = os.path.exists(backup_file)
        with open(backup_file, 'a', newline='', encoding='utf-8-sig') as f:
            writer = csv.writer(f)
            if not file_exists:
                writer.writerow(['Họ Tên', 'Email', 'Số Điện Thoại', 'Thời Gian CheckIn'])
            writer.writerow([hoten, email, sdt, thoi_gian])

        self.root.after(0, self._cap_nhat_treeview, hoten, email, sdt, thoi_gian)
        return {'status': 'success', 'email': email, 'name': hoten}

    def _cap_nhat_treeview(self, hoten, email, sdt, thoi_gian):
        self.tree.insert('', 0, values=(hoten, email, sdt, thoi_gian)) 

    def xu_ly_ma_quet_usb(self, event):
        qr_text = self.entry_scan.get().strip()
        self.entry_scan.delete(0, tk.END)
        if qr_text:
            res = self.xu_ly_quet_mobile(qr_text)
            if res['status'] == 'exists':
                messagebox.showwarning("Cảnh báo", f"Khách mời đã điểm danh rồi!\n({res['name']} - {res['email']})")

    def xuat_file_diem_danh(self):
        # Lấy bản sao list hiện tại để không xung đột thread nếu có người đang quét
        data_to_export = list(self.danh_sach_tham_gia)
        if not data_to_export:
            return messagebox.showinfo("Thông báo", "Chưa có dữ liệu.")
        filepath = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if filepath:
            pd.DataFrame(data_to_export).to_excel(filepath, index=False)
            messagebox.showinfo("Thành công", f"Đã lưu tại:\n{filepath}\n(Dữ liệu quét ngầm lúc lưu đã được đưa vào Auto-backup)")

    # --- TAB 3 (ĐỐI SOÁT KÉP: HỌ TÊN + EMAIL) ---
    def setup_tab3(self):
        frame = ttk.LabelFrame(self.tab3, text="Cấu hình File Đăng Ký Gốc", padding=10)
        frame.pack(fill='x', padx=10, pady=5)

        self.file_dangki_path, self.file_diemdanh_path = "", ""

        tk.Button(frame, text="1. Chọn file Đăng Ký gốc", command=lambda: self.chon_file('dangki')).grid(row=0, column=0, pady=5, sticky='w')
        self.lbl_file1 = tk.Label(frame, text="Chưa chọn")
        self.lbl_file1.grid(row=0, column=1, padx=10)

        tk.Label(frame, text="Khóa 1 (Họ Tên):").grid(row=1, column=0, sticky='w', pady=5)
        self.cb_khoa_hoten = ttk.Combobox(frame, state="readonly")
        self.cb_khoa_hoten.grid(row=1, column=1, pady=5)

        tk.Label(frame, text="Khóa 2 (Email):").grid(row=2, column=0, sticky='w', pady=5)
        self.cb_khoa_email = ttk.Combobox(frame, state="readonly")
        self.cb_khoa_email.grid(row=2, column=1, pady=5)

        frame2 = ttk.LabelFrame(self.tab3, text="Cấu hình File Điểm Danh", padding=10)
        frame2.pack(fill='x', padx=10, pady=5)

        tk.Button(frame2, text="2. Chọn file Điểm Danh", command=lambda: self.chon_file('diemdanh')).grid(row=0, column=0, pady=5, sticky='w')
        self.lbl_file2 = tk.Label(frame2, text="Chưa chọn")
        self.lbl_file2.grid(row=0, column=1, padx=10)

        tk.Button(self.tab3, text="Thực Hiện Đối Soát", bg="purple", fg="white", font=("Arial", 11, "bold"), command=self.thuc_hien_doi_soat).pack(pady=20)

    def chon_file(self, loai):
        filepath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls *.csv")])
        if filepath:
            if loai == 'dangki':
                self.file_dangki_path = filepath
                self.lbl_file1.config(text=os.path.basename(filepath))
                df = pd.read_excel(filepath) if filepath.endswith('.xls') or filepath.endswith('.xlsx') else pd.read_csv(filepath)
                cols = list(df.columns)
                self.cb_khoa_hoten['values'] = cols
                self.cb_khoa_email['values'] = cols
                if cols: self.cb_khoa_hoten.set(cols[0])
                if len(cols) > 1: self.cb_khoa_email.set(cols[1])
            else:
                self.file_diemdanh_path = filepath
                self.lbl_file2.config(text=os.path.basename(filepath))

    def thuc_hien_doi_soat(self):
        col_hoten = self.cb_khoa_hoten.get()
        col_email = self.cb_khoa_email.get()

        if not self.file_dangki_path or not self.file_diemdanh_path or not col_hoten or not col_email:
            return messagebox.showwarning("Cảnh báo", "Vui lòng chọn đủ file và Cột khóa.")
            
        try:
            df_dangki = pd.read_excel(self.file_dangki_path) if self.file_dangki_path.endswith('.xlsx') else pd.read_csv(self.file_dangki_path)
            df_diemdanh = pd.read_excel(self.file_diemdanh_path) if self.file_diemdanh_path.endswith('.xlsx') else pd.read_csv(self.file_diemdanh_path)
            
            # Gộp dựa trên cả Họ Tên VÀ Email để đảm bảo chính xác tuyệt đối
            df_diemdanh_subset = df_diemdanh[['Họ Tên', 'Email', 'Thời Gian CheckIn']].copy()
            df_diemdanh_subset.rename(columns={'Họ Tên': col_hoten, 'Email': col_email}, inplace=True)

            df_ketqua = pd.merge(df_dangki, df_diemdanh_subset, on=[col_hoten, col_email], how='left')
            df_ketqua['Trạng Thái'] = df_ketqua['Thời Gian CheckIn'].apply(lambda x: 'Đã tham gia' if pd.notna(x) else 'Vắng mặt')

            save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile="Bao_Cao_Doi_Soat_Kep.xlsx")
            if save_path:
                df_ketqua.to_excel(save_path, index=False)
                messagebox.showinfo("Thành công", "Đã xuất file đối soát!")
        except Exception as e:
            messagebox.showerror("Lỗi", f"Lỗi đối soát: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = QREventApp(root)
    root.mainloop()