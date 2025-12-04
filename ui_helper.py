import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from threading import Thread


class AppUI:
    def __init__(self):
        self.window = tk.Tk()
        self.window.title('App Chấm Kênh - New Version')
        self.window.geometry("1000x750")
        
        self.file_path = ''
        self.target_date = ''
        
        self._create_menu()
        self._create_widgets()
    
    def _create_menu(self):
        """Tạo menu bar"""
        menu = tk.Menu(self.window)
        file_menu = tk.Menu(menu, tearoff=0)
        file_menu.add_command(label='Mở file', command=self._open_file)
        file_menu.add_separator()
        file_menu.add_command(label='Thoát', command=self.window.quit)
        menu.add_cascade(label='File', menu=file_menu)
        self.window.config(menu=menu)
    
    def _create_widgets(self):
        """Tạo các widget"""
        # === HEADER ===
        lbl_title = tk.Label(
            self.window, 
            text="ỨNG DỤNG CHẤM KÊNH DINH DƯỠNG", 
            fg="blue", 
            font=("Times New Roman", 18, "bold")
        )
        lbl_title.pack(pady=10)
        
        # === FRAME CHỌN FILE ===
        frame_file = ttk.LabelFrame(self.window, text="Chọn file Excel", padding=10)
        frame_file.pack(fill="x", padx=20, pady=10)
        
        self.lbl_file = ttk.Label(frame_file, text="Chưa chọn file...", font=("Arial", 10))
        self.lbl_file.pack(side="left", fill="x", expand=True)
        
        btn_browse = ttk.Button(frame_file, text="Duyệt...", command=self._open_file)
        btn_browse.pack(side="right")
        
        # === FRAME NGÀY TARGET ===
        frame_date = ttk.LabelFrame(self.window, text="Ngày tính tháng tuổi", padding=10)
        frame_date.pack(fill="x", padx=20, pady=5)
        
        ttk.Label(frame_date, text="Ngày target (dd/mm/yyyy):").pack(side="left")
        self.entry_date = ttk.Entry(frame_date, width=15)
        self.entry_date.pack(side="left", padx=10)
        self.entry_date.insert(0, "30/06/2021")
        
        # === FRAME CHỨC NĂNG ===
        frame_functions = ttk.LabelFrame(self.window, text="Chức năng", padding=10)
        frame_functions.pack(fill="x", padx=20, pady=10)
        
        # Checkbox variables
        self.var_calc_month = tk.BooleanVar()
        self.var_sdd = tk.BooleanVar(value=True)
        self.var_cncc = tk.BooleanVar(value=True)
        self.var_adj_height = tk.BooleanVar()
        self.var_adj_weight = tk.BooleanVar()
        self.var_export = tk.BooleanVar()
        self.var_overwrite = tk.BooleanVar(value=True)
        self.var_summary_2 = tk.BooleanVar()
        self.var_summary_5 = tk.BooleanVar()
        
        # Row 1
        row1 = ttk.Frame(frame_functions)
        row1.pack(fill="x", pady=5)
        ttk.Checkbutton(row1, text="Tính tháng tuổi", variable=self.var_calc_month).pack(side="left", padx=20)
        ttk.Checkbutton(row1, text="Chấm SDD (CN/tuổi, CC/tuổi)", variable=self.var_sdd).pack(side="left", padx=20)
        ttk.Checkbutton(row1, text="Chấm CN/CC", variable=self.var_cncc).pack(side="left", padx=20)
        
        # Row 2 - Adj
        row2 = ttk.Frame(frame_functions)
        row2.pack(fill="x", pady=5)
        ttk.Checkbutton(row2, text="Adj chiều cao", variable=self.var_adj_height).pack(side="left", padx=20)
        ttk.Checkbutton(row2, text="Adj cân nặng", variable=self.var_adj_weight).pack(side="left", padx=20)
        
        # Row 3 - Export
        row3 = ttk.Frame(frame_functions)
        row3.pack(fill="x", pady=5)
        ttk.Checkbutton(row3, text="Ghi đè vào file gốc", variable=self.var_overwrite).pack(side="left", padx=20)
        ttk.Checkbutton(row3, text="Xuất ra file mới", variable=self.var_export).pack(side="left", padx=20)
        
        # Row 4 - Thống kê
        row4 = ttk.Frame(frame_functions)
        row4.pack(fill="x", pady=5)
        ttk.Checkbutton(row4, text="Thống kê trẻ dưới 2 tuổi", variable=self.var_summary_2).pack(side="left", padx=20)
        ttk.Checkbutton(row4, text="Thống kê trẻ dưới 5 tuổi", variable=self.var_summary_5).pack(side="left", padx=20)
        
        # === BUTTON THỰC HIỆN ===
        btn_execute = ttk.Button(
            self.window, 
            text="THỰC HIỆN", 
            command=self._on_execute,
            style="Accent.TButton"
        )
        btn_execute.pack(pady=15)
        
        # === FRAME KẾT QUẢ ===
        frame_result = ttk.LabelFrame(self.window, text="Kết quả", padding=10)
        frame_result.pack(fill="both", expand=True, padx=20, pady=10)
        
        # Scrolled text for results
        self.txt_result = tk.Text(frame_result, height=12, font=("Consolas", 9))
        scrollbar = ttk.Scrollbar(frame_result, orient="vertical", command=self.txt_result.yview)
        self.txt_result.configure(yscrollcommand=scrollbar.set)
        
        self.txt_result.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # === FOOTER ===
        lbl_footer = tk.Label(self.window, text="© 2025 - App Chấm Kênh v2.0", fg="gray")
        lbl_footer.pack(side="bottom", pady=5)
    
    def _open_file(self):
        """Mở dialog chọn file"""
        filetypes = [("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        file_path = filedialog.askopenfilename(filetypes=filetypes)
        if file_path:
            self.file_path = file_path
            self.lbl_file.configure(text=file_path)
            self._log(f"Đã chọn file: {file_path}")
            self.load_target_file()
    
    def _log(self, message: str):
        """Ghi log vào text box"""
        self.txt_result.insert(tk.END, message + "\n")
        self.txt_result.see(tk.END)
    
    def _clear_log(self):
        """Xóa log"""
        self.txt_result.delete(1.0, tk.END)
    
    def _print_summary(self, text: str):
        """Hiển thị popup cửa sổ với nội dung text"""
        # Tạo cửa sổ popup
        popup = tk.Toplevel(self.window)
        popup.title("Thống kê dinh dưỡng")
        popup.geometry("500x400")
        popup.resizable(True, True)
        
        # Text widget với scrollbar
        frame = ttk.Frame(popup, padding=10)
        frame.pack(fill="both", expand=True)
        
        txt = tk.Text(frame, font=("Consolas", 10), wrap="word")
        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=txt.yview)
        txt.configure(yscrollcommand=scrollbar.set)
        
        txt.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Chèn nội dung
        txt.insert(tk.END, text)
        txt.config(state="disabled")  # Không cho chỉnh sửa
        
        # Nút đóng
        btn_close = ttk.Button(popup, text="Đóng", command=popup.destroy)
        btn_close.pack(pady=10)
        
        # Focus vào popup
        popup.focus_set()
        popup.grab_set()
    
    def _on_execute(self):
        """Xử lý khi nhấn nút Thực hiện"""
        if not self.file_path:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn file Excel trước!")
            return
        
        # Chạy trong thread riêng để không block UI
        thread = Thread(target=self._execute_tasks)
        thread.start()
    
    def _execute_tasks(self):
        """Thực hiện các tác vụ đã chọn"""
        self._clear_log()
        self._log("=" * 50)
        self._log("BẮT ĐẦU XỬ LÝ...")
        self._log("=" * 50)
        
        try:
            target_date = self.entry_date.get()
            
            if self.var_calc_month.get():
                self._log("\n[1] Đang tính tháng tuổi...")
                self.execute_month_age()
                self._log("    ✓ Hoàn thành tính tháng tuổi")
            
            if self.var_sdd.get():
                self._log("\n[2] Đang chấm SDD (CN/tuổi, CC/tuổi)...")
                self.execute_weight_by_age()
                self.execute_height_by_age()
                self._log("    ✓ Hoàn thành chấm SDD")
            
            if self.var_cncc.get():
                self._log("\n[3] Đang chấm CN/CC...")
                self.execute_weight_by_height()
                self._log("    ✓ Hoàn thành chấm CN/CC")
            
            if self.var_adj_height.get():
                self._log("\n[4] Đang adj chiều cao...")
                self.adjust_height()
                self._log("    ✓ Hoàn thành adj chiều cao")
            
            if self.var_adj_weight.get():
                self._log("\n[5] Đang adj cân nặng...")
                self.adjust_weight()
                self._log("    ✓ Hoàn thành adj cân nặng")
            
            if self.var_export.get():
                self._log("\n[6] Đang xuất ra file mới...")
                self.export_to_excel()
                self._log("    ✓ Hoàn thành xuất file")
            
            if self.var_summary_2.get():
                self._log("\n[7] Thống kê trẻ dưới 2 tuổi:")
                summary = self.summary_statistics(max_months=24)
                text = self.format_summary(summary, "TRẺ DƯỚI 2 TUỔI (≤24 tháng)")
                self._print_summary(text)
                self._log("    ✓ Đã hiển thị thống kê")
            
            if self.var_summary_5.get():
                self._log("\n[8] Thống kê trẻ dưới 5 tuổi:")
                summary = self.summary_statistics(max_months=60)
                text = self.format_summary(summary, "TRẺ DƯỚI 5 TUỔI (≤60 tháng)")
                self._print_summary(text)
                self._log("    ✓ Đã hiển thị thống kê")
            
            self._log("\n" + "=" * 50)
            self._log("HOÀN THÀNH TẤT CẢ!")
            self._log("=" * 50)
            
            messagebox.showinfo("Thành công", "Đã hoàn thành xử lý!")
            
        except Exception as e:
            import traceback
            traceback.print_exc()  # In chi tiết lỗi ra terminal
            self._log(f"\n❌ LỖI: {str(e)}")
            messagebox.showerror("Lỗi", f"Đã xảy ra lỗi: {str(e)}")

    def load_target_file(self):
        raise NotImplementedError("Not implemented yet")
    
    def execute_month_age(self):
        raise NotImplementedError("Not implemented yet")
    
    def execute_weight_by_age(self):
        raise NotImplementedError("Not implemented yet")
    
    def execute_height_by_age(self):
        raise NotImplementedError("Not implemented yet")
    
    def execute_weight_by_height(self):
        raise NotImplementedError("Not implemented yet")
    
    def adjust_height(self):
        raise NotImplementedError("Not implemented yet")
    
    def adjust_weight(self):
        raise NotImplementedError("Not implemented yet")
    
    def export_to_excel(self):
        raise NotImplementedError("Not implemented yet")
    
    def summary_statistics(self, max_months):
        raise NotImplementedError("Not implemented yet")
    
    def format_summary(self, summary: dict, title: str) -> str:
        raise NotImplementedError("Not implemented yet")
    
    def run(self):
        """Chạy ứng dụng"""
        self.window.mainloop()


def create_app() -> AppUI:
    """Tạo và trả về instance của AppUI"""
    return AppUI()


if __name__ == "__main__":
    app = create_app()
    app.run()
