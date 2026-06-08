from datetime import datetime
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from threading import Thread
from process_helper import summary_table, combine_summaries
from excel_file_helper import build_under5_statistics_table, convert_xls_to_xlsx


class AppUI:
    def __init__(self):
        self.window = tk.Tk()
        self.window.title('App Chấm Kênh - New Version')
        self.window.geometry("1000x750")
        
        self.file_path = ''
        self.folder_path = ''
        self.is_folder_mode = False
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
        frame_file = ttk.LabelFrame(self.window, text="Chọn file Excel hoặc Folder", padding=10)
        frame_file.pack(fill="x", padx=20, pady=10)
        
        self.lbl_file = ttk.Label(frame_file, text="Chưa chọn file/folder...", font=("Arial", 10))
        self.lbl_file.pack(side="left", fill="x", expand=True)
        
        btn_browse_folder = ttk.Button(frame_file, text="Chọn Folder", command=self._open_folder)
        btn_browse_folder.pack(side="right", padx=5)
        
        btn_browse = ttk.Button(frame_file, text="Chọn File", command=self._open_file)
        btn_browse.pack(side="right")
        
        # === FRAME NGÀY TARGET ===
        frame_date = ttk.LabelFrame(self.window, text="Ngày tính tháng tuổi", padding=10)
        frame_date.pack(fill="x", padx=20, pady=5)
        
        ttk.Label(frame_date, text="Ngày target (dd/mm/yyyy):").pack(side="left")
        self.entry_date = ttk.Entry(frame_date, width=15)
        self.entry_date.pack(side="left", padx=10)
        self.entry_date.insert(0, datetime.now().strftime("%d/%m/%Y"))
        
        # === FRAME CHỨC NĂNG ===
        frame_functions = ttk.LabelFrame(self.window, text="Chức năng", padding=10)
        frame_functions.pack(fill="x", padx=20, pady=10)
        
        # Checkbox variables
        self.var_calc_month = tk.BooleanVar()
        self.var_round_height = tk.BooleanVar()
        self.var_sdd = tk.BooleanVar()
        self.var_cncc = tk.BooleanVar()
        self.var_adj_height = tk.BooleanVar()
        self.var_adj_weight = tk.BooleanVar()
        self.var_export = tk.BooleanVar()
        self.var_overwrite = tk.BooleanVar()
        self.var_summary_2 = tk.BooleanVar()
        self.var_summary_5 = tk.BooleanVar()
        self.var_under5_table = tk.BooleanVar()
        
        # Row 1
        row1 = ttk.Frame(frame_functions)
        row1.pack(fill="x", pady=5)
        ttk.Checkbutton(row1, text="Tính tháng tuổi", variable=self.var_calc_month).pack(side="left", padx=20)
        ttk.Checkbutton(row1, text="Chấm SDD (CN/tuổi, CC/tuổi)", variable=self.var_sdd).pack(side="left", padx=20)
        ttk.Checkbutton(row1, text="Chấm CN/CC", variable=self.var_cncc).pack(side="left", padx=20)
        
        # Row 2 - Adj
        row2 = ttk.Frame(frame_functions)
        row2.pack(fill="x", pady=5)
        ttk.Checkbutton(row2, text="Làm tròn CC 0.5", variable=self.var_round_height).pack(side="left", padx=20)
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
        ttk.Checkbutton(row4, text="Xuất bảng tổng hợp <5T", variable=self.var_under5_table).pack(side="left", padx=20)
        
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
            # File .xls (định dạng cũ) -> convert sang .xlsx trước khi xử lý
            if file_path.lower().endswith('.xls'):
                try:
                    self._log(f"Phát hiện file .xls, đang convert sang .xlsx...")
                    file_path = convert_xls_to_xlsx(file_path)
                    self._log(f"✓ Đã convert: {file_path}")
                except Exception as e:
                    self._log(f"❌ Không convert được .xls: {e}")
                    messagebox.showerror("Lỗi", f"Không convert được file .xls:\n{e}")
                    return
            self.file_path = file_path
            self.folder_path = ''
            self.is_folder_mode = False
            self.lbl_file.configure(text=file_path)
            self._log(f"Đã chọn file: {file_path}")
            self.load_target_file()
    
    def _open_folder(self):
        """Mở dialog chọn folder"""
        folder_path = filedialog.askdirectory()
        if folder_path:
            self.folder_path = folder_path
            self.file_path = ''
            self.is_folder_mode = True
            self.lbl_file.configure(text=f"📁 {folder_path}")
            self._log(f"Đã chọn folder: {folder_path}")
    
    def _log(self, message: str):
        """Ghi log vào text box"""
        self.txt_result.insert(tk.END, message + "\n")
        self.txt_result.see(tk.END)
    
    def _clear_log(self):
        """Xóa log"""
        self.txt_result.delete(1.0, tk.END)
    
    def _print_summary(self, text: str, file_path: str = ""):
        """Hiển thị popup cửa sổ với nội dung text, kèm tên file đang thống kê."""
        popup = tk.Toplevel(self.window)
        popup.title("Thống kê dinh dưỡng")
        popup.geometry("520x420")
        popup.resizable(True, True)

        frame = ttk.Frame(popup, padding=10)
        frame.pack(fill="both", expand=True)

        if file_path:
            ttk.Label(frame, text=f"File: {file_path}", foreground="blue").pack(anchor="w", pady=(0, 6))

        txt = tk.Text(frame, font=("Consolas", 10), wrap="word")
        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=txt.yview)
        txt.configure(yscrollcommand=scrollbar.set)

        txt.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        txt.insert(tk.END, text)
        txt.config(state="disabled")

        ttk.Button(popup, text="Đóng", command=popup.destroy).pack(pady=10)

    def _print_summary_df(self, df, file_path: str = ""):
        """Hiển thị DataFrame thống kê trong popup, kèm tên file."""
        popup = tk.Toplevel(self.window)
        popup.title("Bảng thống kê")
        popup.geometry("920x520")
        popup.resizable(True, True)

        frame = ttk.Frame(popup, padding=10)
        frame.pack(fill="both", expand=True)

        if file_path:
            ttk.Label(frame, text=f"File: {file_path}", foreground="blue").pack(anchor="w", pady=(0, 6))

        txt = tk.Text(frame, font=("Consolas", 10), wrap="none")
        scrollbar_y = ttk.Scrollbar(frame, orient="vertical", command=txt.yview)
        scrollbar_x = ttk.Scrollbar(frame, orient="horizontal", command=txt.xview)
        txt.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)

        txt.pack(side="left", fill="both", expand=True)
        scrollbar_y.pack(side="right", fill="y")
        scrollbar_x.pack(side="bottom", fill="x")

        txt.insert(tk.END, df.to_string(index=False))
        txt.config(state="disabled")

        ttk.Button(popup, text="Đóng", command=popup.destroy).pack(pady=8)
    
    def _on_execute(self):
        """Xử lý khi nhấn nút Thực hiện"""
        if not self.file_path and not self.folder_path:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn file Excel hoặc folder trước!")
            return
        
        # Chạy trong thread riêng để không block UI
        if self.is_folder_mode:
            thread = Thread(target=self._execute_folder_tasks)
        else:
            thread = Thread(target=self._execute_tasks)
        thread.start()
    
    def _execute_folder_tasks(self):
        """Xử lý tất cả file trong folder, kèm thống kê đa tầng (file và folder)."""
        from pathlib import Path
        import pandas as pd
        
        self._clear_log()
        self._log("=" * 50)
        self._log(f"BẮT ĐẦU XỬ LÝ FOLDER: {self.folder_path}")
        self._log("=" * 50)
        
        root = Path(self.folder_path).resolve()
        # Bỏ qua file khóa tạm của Excel (tên bắt đầu bằng ~$) và file kết quả (_ketqua)
        all_xlsx_files = [f for f in root.rglob('*.xlsx') if not f.name.startswith('~$')]
        xlsx_files = [f for f in all_xlsx_files if not f.stem.endswith('_ketqua')]

        skipped_count = len(all_xlsx_files) - len(xlsx_files)
        if skipped_count > 0:
            self._log(f"\n⏭ Bỏ qua {skipped_count} file kết quả (_ketqua)")

        # Convert các file .xls (định dạng cũ) sang .xlsx, bỏ qua nếu đã có .xlsx cùng tên
        existing_xlsx_paths = {f.with_suffix('') for f in xlsx_files}
        xls_files = [f for f in root.rglob('*.xls') if not f.name.startswith('~$')]
        for xls_file in xls_files:
            if xls_file.with_suffix('') in existing_xlsx_paths:
                self._log(f"\n⏭ Bỏ qua {xls_file.name} (đã có bản .xlsx)")
                continue
            try:
                self._log(f"\n🔄 Convert {xls_file.name} sang .xlsx...")
                new_path = convert_xls_to_xlsx(str(xls_file))
                xlsx_files.append(Path(new_path))
                self._log(f"    ✓ {Path(new_path).name}")
            except Exception as e:
                self._log(f"    ❌ Không convert được {xls_file.name}: {e}")
        
        if not xlsx_files:
            self._log("\n❌ Không tìm thấy file Excel nào trong folder!")
            messagebox.showwarning("Cảnh báo", "Không tìm thấy file Excel nào trong folder!")
            return
        
        self._log(f"\nTìm thấy {len(xlsx_files)} file Excel")
        self._log("-" * 50)
        
        success_count = 0
        failed_files = []
        summaries_2 = []  # (path, summary)
        summaries_5 = []
        
        for i, xlsx_file in enumerate(xlsx_files, 1):
            self._log(f"\n[{i}/{len(xlsx_files)}] Đang xử lý: {xlsx_file.name}")
            try:
                self.file_path = str(xlsx_file)
                self.load_target_file()
                self._process_single_file()

                if self.var_summary_2.get():
                    s2 = self.summary_statistics(max_months=24)
                    summaries_2.append((self.file_path, s2))
                    self._log("    ✓ Thống kê ≤24 tháng")

                if self.var_summary_5.get():
                    s5 = self.summary_statistics(max_months=60)
                    summaries_5.append((self.file_path, s5))
                    self._log("    ✓ Thống kê ≤60 tháng")

                success_count += 1
                self._log("    ✓ Hoàn thành")
            except Exception as e:
                failed_files.append((xlsx_file.name, str(e)))
                self._log(f"    ❌ Lỗi: {str(e)}")
        
        def build_file_df(summaries):
            rows = []
            for path_str, summary in summaries:
                df_row = summary_table(summary)
                df_row.insert(0, 'Đường dẫn', path_str)
                df_row.insert(0, 'Loại', 'File')
                df_row['depth'] = None
                rows.append(df_row)
            if not rows:
                return None
            return pd.concat(rows, ignore_index=True)

        def build_folder_df(summaries):
            if not summaries:
                return None
            folder_groups = {}
            for path_str, summary in summaries:
                p = Path(path_str).resolve()
                # Chỉ cộng vào folder chứa trực tiếp file (p.parent)
                parent_folder = p.parent
                if parent_folder >= root:  # Chỉ xử lý trong root
                    folder_groups.setdefault(str(parent_folder), []).append(summary)
            
            # Thêm tổng kết cho root folder (tất cả files)
            if summaries:
                folder_groups[str(root)] = [s for _, s in summaries]
            
            rows = []
            for folder_path, lst in folder_groups.items():
                combined = combine_summaries(lst)
                if combined is None:
                    continue
                df_row = summary_table(combined)
                df_row.insert(0, 'Đường dẫn', folder_path)
                df_row.insert(0, 'Loại', 'Tổng folder')
                relative = Path(folder_path).resolve().relative_to(root)
                depth = len(relative.parts)
                df_row['depth'] = depth
                rows.append(df_row)
            return pd.concat(rows, ignore_index=True) if rows else None

        def combine_file_folder(df_files, df_folders):
            frames = []
            if df_files is not None:
                frames.append(df_files)
            if df_folders is not None:
                frames.append(df_folders)
            if not frames:
                return None
            combined = pd.concat(frames, ignore_index=True)
            if 'Loại' in combined.columns:
                order_map = {'File': 0, 'Tổng folder': 1}
                combined['__order'] = combined['Loại'].map(order_map).fillna(99)
                if 'depth' in combined.columns:
                    combined['depth'] = pd.to_numeric(combined['depth'], errors='coerce')
                combined = combined.sort_values(by=['__order', 'depth'], ascending=[True, False], ignore_index=True)
                combined = combined.drop(columns=['__order', 'depth'], errors='ignore')
            return combined

        def export_summary_excels(df_under_2, df_under_5):
            under5_table = None
            if summaries_5 and self.var_under5_table.get():
                combined_5 = combine_summaries([s for _, s in summaries_5])
                if combined_5:
                    under5_table = build_under5_statistics_table(combined_5)
            if df_under_2 is None and df_under_5 is None:
                return None
            out_path = root / "Tổng hợp thống kê.xlsx"
            with pd.ExcelWriter(out_path) as writer:
                if df_under_2 is not None:
                    df_under_2.to_excel(writer, sheet_name='Dưới 2 tuổi', index=False)
                if df_under_5 is not None:
                    df_under_5.to_excel(writer, sheet_name='Dưới 5 tuổi', index=False)
                if under5_table is not None:
                    under5_table.to_excel(writer, sheet_name='Thong ke <5T', index=False)
            return out_path

        df_files_2 = build_file_df(summaries_2) if self.var_summary_2.get() else None
        df_files_5 = build_file_df(summaries_5) if self.var_summary_5.get() else None
        df_folders_2 = build_folder_df(summaries_2) if self.var_summary_2.get() else None
        df_folders_5 = build_folder_df(summaries_5) if self.var_summary_5.get() else None

        df_under_2 = combine_file_folder(df_files_2, df_folders_2) if self.var_summary_2.get() else None
        df_under_5 = combine_file_folder(df_files_5, df_folders_5) if self.var_summary_5.get() else None

        out_path = export_summary_excels(df_under_2, df_under_5)

        # Summary
        self._log("\n" + "=" * 50)
        self._log(f"HOÀN THÀNH: {success_count}/{len(xlsx_files)} files")
        if failed_files:
            self._log(f"\nCác file lỗi:")
            for name, error in failed_files:
                self._log(f"  - {name}: {error}")
        if out_path:
            self._log(f"\nĐã lưu thống kê: {out_path}")
        self._log("=" * 50)
        
        messagebox.showinfo("Hoàn thành", f"Đã xử lý {success_count}/{len(xlsx_files)} files")
    
    def _process_single_file(self):
        """Xử lý 1 file - được gọi từ cả single mode và folder mode"""
        if self.var_calc_month.get():
            self.execute_month_age()
        
        if self.var_round_height.get():
            self.round_height_cells()
        
        if self.var_sdd.get():
            self.execute_weight_by_age()
            self.execute_height_by_age()
        
        if self.var_cncc.get():
            self.execute_weight_by_height()
        
        if self.var_adj_height.get():
            self.adjust_height()
        
        if self.var_adj_weight.get():
            self.adjust_weight()
        
        if self.var_export.get():
            self.export_to_excel()
    
    def _execute_tasks(self):
        """Thuc hien cac tac vu da chon (single file mode)"""
        self._clear_log()
        self._log("=" * 50)
        self._log("BAT DAU XU LY...")
        self._log("=" * 50)
        
        try:
            step = 1
            if self.var_calc_month.get():
                self._log(f"\n[{step}] Dang tinh thang tuoi...")
                self.execute_month_age()
                self._log("    OK Hoan thanh tinh thang tuoi")
                step += 1
            
            if self.var_round_height.get():
                self._log(f"\n[{step}] Dang lam tron chieu cao ve 0.5...")
                self.round_height_cells()
                self._log("    OK Hoan thanh lam tron chieu cao")
                step += 1
            
            if self.var_sdd.get():
                self._log(f"\n[{step}] Dang cham SDD (CN/tuoi, CC/tuoi)...")
                self.execute_weight_by_age()
                self.execute_height_by_age()
                self._log("    OK Hoan thanh cham SDD")
                step += 1
            
            if self.var_cncc.get():
                self._log(f"\n[{step}] Dang cham CN/CC...")
                self.execute_weight_by_height()
                self._log("    OK Hoan thanh cham CN/CC")
                step += 1
            
            if self.var_adj_height.get():
                self._log(f"\n[{step}] Dang adj chieu cao...")
                self.adjust_height()
                self._log("    OK Hoan thanh adj chieu cao")
                step += 1
            
            if self.var_adj_weight.get():
                self._log(f"\n[{step}] Dang adj can nang...")
                self.adjust_weight()
                self._log("    OK Hoan thanh adj can nang")
                step += 1
            
            if self.var_export.get():
                self._log(f"\n[{step}] Dang xuat ra file moi...")
                self.export_to_excel()
                self._log("    OK Hoan thanh xuat file")
                step += 1
            
            if self.var_summary_2.get():
                self._log(f"\n[{step}] Thong ke tre duoi 2 tuoi:")
                summary = self.summary_statistics(max_months=24)
                text_sum = self.format_summary(summary, "TRE DUOI 2 TUOI (<=24 thang)")
                self._print_summary(text_sum, file_path=self.file_path)
                self._log("    OK Da hien thi thong ke")
                step += 1
            
            if self.var_summary_5.get():
                self._log(f"\n[{step}] Thong ke tre duoi 5 tuoi:")
                summary = self.summary_statistics(max_months=60)
                text_sum = self.format_summary(summary, "TRE DUOI 5 TUOI (<=60 thang)")
                self._print_summary(text_sum, file_path=self.file_path)
                self._log("    OK Da hien thi thong ke")
                step += 1
            
            self._log("\n" + "=" * 50)
            self._log("HOAN THANH TAT CA!")
            self._log("=" * 50)
            
            messagebox.showinfo("Thanh cong", "Da hoan thanh xu ly!")
            
        except Exception as e:
            import traceback
            traceback.print_exc()  # In chi tiet loi ra terminal
            self._log(f"\nLOI: {str(e)}")
            messagebox.showerror("Loi", f"Da xay ra loi: {str(e)}")

    
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
    
    def round_height_cells(self):
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
