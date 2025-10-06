import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime

class DiemDanhApp:
    def __init__(self, root):
        self.root = root
        self.root.title("📋 Hệ Thống Điểm Danh Sinh Viên")
        self.root.geometry("900x650")
        self.root.configure(bg="#f0f0f0")
        
        # Variables
        self.file_danh_sach = tk.StringVar(value="config.xlsx")
        self.file_diem_danh = tk.StringVar(value="input (1).xlsx")
        
        self.setup_ui()
    
    def setup_ui(self):
        # Header
        header_frame = tk.Frame(self.root, bg="#2196F3", height=80)
        header_frame.pack(fill="x")
        header_frame.pack_propagate(False)
        
        title_label = tk.Label(
            header_frame,
            text="📋 HỆ THỐNG ĐIỂM DANH SINH VIÊN",
            font=("Arial", 20, "bold"),
            bg="#2196F3",
            fg="white"
        )
        title_label.pack(pady=20)
        
        # File selection frame
        file_frame = tk.LabelFrame(
            self.root,
            text="📁 Chọn File",
            font=("Arial", 12, "bold"),
            bg="#f0f0f0",
            fg="#333"
        )
        file_frame.pack(pady=15, padx=20, fill="x")
        
        # File danh sách
        tk.Label(file_frame, text="File danh sách:", font=("Arial", 10), bg="#f0f0f0").grid(row=0, column=0, sticky="w", padx=10, pady=5)
        tk.Entry(file_frame, textvariable=self.file_danh_sach, width=50, font=("Arial", 10)).grid(row=0, column=1, padx=10, pady=5)
        tk.Button(file_frame, text="Chọn", command=lambda: self.browse_file(self.file_danh_sach), bg="#4CAF50", fg="white").grid(row=0, column=2, padx=10, pady=5)
        
        # File điểm danh
        tk.Label(file_frame, text="File điểm danh:", font=("Arial", 10), bg="#f0f0f0").grid(row=1, column=0, sticky="w", padx=10, pady=5)
        tk.Entry(file_frame, textvariable=self.file_diem_danh, width=50, font=("Arial", 10)).grid(row=1, column=1, padx=10, pady=5)
        tk.Button(file_frame, text="Chọn", command=lambda: self.browse_file(self.file_diem_danh), bg="#4CAF50", fg="white").grid(row=1, column=2, padx=10, pady=5)
        
        # Button điểm danh
        btn_frame = tk.Frame(self.root, bg="#f0f0f0")
        btn_frame.pack(pady=10)
        
        self.btn_diem_danh = tk.Button(
            btn_frame,
            text="🚀 BẮT ĐẦU ĐIỂM DANH",
            command=self.diem_danh,
            font=("Arial", 14, "bold"),
            bg="#FF5722",
            fg="white",
            width=25,
            height=2,
            cursor="hand2"
        )
        self.btn_diem_danh.pack()
        
        # Result frame
        result_frame = tk.LabelFrame(
            self.root,
            text="📊 Kết Quả",
            font=("Arial", 12, "bold"),
            bg="#f0f0f0",
            fg="#333"
        )
        result_frame.pack(pady=10, padx=20, fill="both", expand=True)
        
        # Treeview for results
        columns = ("STT", "MSSV", "Họ và Tên")
        self.tree = ttk.Treeview(result_frame, columns=columns, show="headings", height=15)
        
        # Define headings
        self.tree.heading("STT", text="STT")
        self.tree.heading("MSSV", text="Mã Số Sinh Viên")
        self.tree.heading("Họ và Tên", text="Họ và Tên")
        
        # Define columns
        self.tree.column("STT", width=50, anchor="center")
        self.tree.column("MSSV", width=150, anchor="center")
        self.tree.column("Họ và Tên", width=400, anchor="w")
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(result_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        self.tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Status bar
        self.status_label = tk.Label(
            self.root,
            text="⏰ Sẵn sàng điểm danh",
            font=("Arial", 10),
            bg="#263238",
            fg="white",
            anchor="w",
            padx=10
        )
        self.status_label.pack(side="bottom", fill="x")
    
    def browse_file(self, var):
        filename = filedialog.askopenfilename(
            title="Chọn file Excel",
            filetypes=(("Excel files", "*.xlsx *.xls"), ("All files", "*.*"))
        )
        if filename:
            var.set(filename)
    
    def diem_danh(self):
        try:
            # Clear previous results
            for item in self.tree.get_children():
                self.tree.delete(item)
            
            self.status_label.config(text="🔄 Đang xử lý...")
            self.root.update()
            
            # Read files
            df_danh_sach = pd.read_excel(self.file_danh_sach.get())
            df_diem_danh = pd.read_excel(self.file_diem_danh.get())
            
            # Find MSSV columns
            col_mssv_danh_sach = None
            col_ho_lot = None
            col_ten = None
            
            for col in df_danh_sach.columns:
                col_lower = str(col).lower()
                if 'mã' in col_lower or 'mssv' in col_lower:
                    col_mssv_danh_sach = col
                elif 'họ' in col_lower:
                    col_ho_lot = col
                elif 'tên' in col_lower and 'họ' not in col_lower:
                    col_ten = col
            
            col_mssv_diem_danh = None
            for col in df_diem_danh.columns:
                col_lower = str(col).lower()
                if 'mã' in col_lower or 'mssv' in col_lower:
                    col_mssv_diem_danh = col
                    break
            
            if col_mssv_danh_sach is None or col_mssv_diem_danh is None:
                messagebox.showerror("Lỗi", "Không tìm thấy cột MSSV trong file!")
                self.status_label.config(text="❌ Lỗi: Không tìm thấy cột MSSV")
                return
            
            # Process data
            df_danh_sach[col_mssv_danh_sach] = df_danh_sach[col_mssv_danh_sach].astype(str).str.strip()
            df_diem_danh[col_mssv_diem_danh] = df_diem_danh[col_mssv_diem_danh].astype(str).str.strip()
            
            mssv_co_mat = set(df_diem_danh[col_mssv_diem_danh].dropna())
            
            sinh_vien_vang = []
            
            for idx, row in df_danh_sach.iterrows():
                mssv = str(row[col_mssv_danh_sach]).strip()
                ho_ten = ""
                
                if col_ho_lot and col_ten:
                    ho_ten = f"{row[col_ho_lot]} {row[col_ten]}"
                elif col_ho_lot:
                    ho_ten = str(row[col_ho_lot])
                else:
                    for col in df_danh_sach.columns:
                        if col != col_mssv_danh_sach:
                            ho_ten += str(row[col]) + " "
                    ho_ten = ho_ten.strip()
                
                if mssv not in mssv_co_mat:
                    sinh_vien_vang.append((mssv, ho_ten))
            
            # Display results
            if sinh_vien_vang:
                for i, (mssv, ho_ten) in enumerate(sinh_vien_vang, 1):
                    self.tree.insert("", "end", values=(i, mssv, ho_ten), tags=("vang",))
                
                # Style for absent students
                self.tree.tag_configure("vang", background="#ffebee", foreground="#c62828")
                
                so_vang = len(sinh_vien_vang)
                tong_sv = len(df_danh_sach)
                ty_le = (so_vang / tong_sv) * 100 if tong_sv > 0 else 0
                
                self.status_label.config(
                    text=f"❌ Có {so_vang}/{tong_sv} sinh viên vắng ({ty_le:.1f}%) - {datetime.now().strftime('%H:%M:%S %d/%m/%Y')}"
                )
                
                messagebox.showwarning(
                    "Kết quả điểm danh",
                    f"⚠️ CÓ SINH VIÊN VẮNG!\n\n"
                    f"• Tổng số: {tong_sv} sinh viên\n"
                    f"• Vắng: {so_vang} sinh viên ({ty_le:.1f}%)\n"
                    f"• Có mặt: {tong_sv - so_vang} sinh viên ({100-ty_le:.1f}%)"
                )
            else:
                tong_sv = len(df_danh_sach)
                self.status_label.config(
                    text=f"✅ ĐỦ - Không có sinh viên nào vắng! ({tong_sv}/{tong_sv}) - {datetime.now().strftime('%H:%M:%S %d/%m/%Y')}"
                )
                
                messagebox.showinfo(
                    "Kết quả điểm danh",
                    f"🎉 ĐIỂM DANH ĐỦ!\n\n"
                    f"Tất cả {tong_sv} sinh viên đều có mặt.\n"
                    f"Tỷ lệ: 100% ✓"
                )
        
        except FileNotFoundError:
            messagebox.showerror("Lỗi", "Không tìm thấy file! Vui lòng kiểm tra đường dẫn.")
            self.status_label.config(text="❌ Lỗi: Không tìm thấy file")
        except Exception as e:
            messagebox.showerror("Lỗi", f"Đã xảy ra lỗi:\n{str(e)}")
            self.status_label.config(text=f"❌ Lỗi: {str(e)}")


def main():
    root = tk.Tk()
    app = DiemDanhApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()