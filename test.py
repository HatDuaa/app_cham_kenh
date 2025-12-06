from excel_file_helper import *
from process_helper import *
from ui_helper import *
import numpy as np


class mainApp(AppUI):
    def __init__(self):
        self.df_children = pd.DataFrame()
        self.df_weight_by_age = pd.DataFrame()
        self.df_height_by_age = pd.DataFrame()
        self.df_weight_by_height_0_2 = pd.DataFrame()
        self.df_weight_by_height_2_5 = pd.DataFrame()
        super().__init__()
        self.load_data()

    def load_data(self):
        self.df_weight_by_age = load_cn_per_old()
        self.df_height_by_age = load_cc_per_old()
        self.df_weight_by_height_0_2 = load_cn_cc('CN_CC_0-2_tuoi.xlsx')
        self.df_weight_by_height_2_5 = load_cn_cc('CN_CC_2-5_tuoi.xlsx')

    def load_target_file(self):
        self.df_children = load_danh_sach_can_do(self.file_path)

    def execute_month_age(self):
        target_date_str = self.entry_date.get()
        target_date = datetime.strptime(target_date_str, "%d/%m/%Y")
        self.df_children = apply_month_age(self.df_children, birth_col='ngay_sinh', target_date=target_date)
        if self.var_overwrite.get():
            write_column_to_excel(self.df_children, 'thang_tuoi', self.file_path, 'F', start_row=7)
    
    def execute_weight_by_age(self):
        self.df_children = execute_weight_by_age(self.df_children, self.df_weight_by_age)
        if self.var_overwrite.get():
            write_column_to_excel(self.df_children, 'execute_cn_tuoi', self.file_path, 'I', start_row=7)
    
    def execute_height_by_age(self):
        self.df_children = execute_height_by_age(self.df_children, self.df_height_by_age)
        if self.var_overwrite.get():
            write_column_to_excel(self.df_children, 'execute_cc_tuoi', self.file_path, 'K', start_row=7)
    
    def execute_weight_by_height(self):
        self.df_children = execute_weight_by_height(self.df_children, self.df_weight_by_height_0_2, self.df_weight_by_height_2_5)
        if self.var_overwrite.get():
            write_column_to_excel(self.df_children, 'execute_cn_cc', self.file_path, 'L', start_row=7)
    
    def adjust_height(self):
        self.df_children = adjust_height_by_age(self.df_children, self.df_height_by_age)
        if self.var_overwrite.get():
            write_column_to_excel(self.df_children, 'chieu_cao', self.file_path, 'J', start_row=7)
    
    def adjust_weight(self):
        self.df_children = adjust_weight_by_height_and_age(
            self.df_children, 
            self.df_weight_by_height_0_2, 
            self.df_weight_by_height_2_5,
            self.df_weight_by_age
        )
        if self.var_overwrite.get():
            write_column_to_excel(self.df_children, 'can_nang', self.file_path, 'H', start_row=7)
    
    def summary_statistics(self, max_months: Optional[int] = None):
        return summary_statistics(self.df_children, max_months=max_months)
    
    def format_summary(self, summary: dict, title: str = "TRẺ DƯỚI 2 TUỔI") -> str:
        """Chuyển dict thống kê thành text"""
        s = summary
        
        lines = [
            f"{'='*60}",
            f"TỔNG KẾT DINH DƯỠNG - {title}",
            f"{'='*60}",
            f"",
            f"1. TỔNG SỐ TRẺ: {s['tong_so_tre']['tong']} (Trai: {s['tong_so_tre']['trai']}, Gái: {s['tong_so_tre']['gai']})",
            f"",
            f"2. TRẺ ĐƯỢC CÂN: {s['tre_duoc_can']['tong']} (Trai: {s['tre_duoc_can']['trai']}, Gái: {s['tre_duoc_can']['gai']})",
            f"   Tỷ lệ: {s['tre_duoc_can']['ty_le']}%",
            f"",
            f"3. TRẺ ĐƯỢC ĐO: {s['tre_duoc_do']['tong']} (Trai: {s['tre_duoc_do']['trai']}, Gái: {s['tre_duoc_do']['gai']})",
            f"   Tỷ lệ: {s['tre_duoc_do']['ty_le']}%",
            f"",
            f"4. SDD CÂN NẶNG/TUỔI: {s['sdd_cn_tuoi']['tong']} (Trai: {s['sdd_cn_tuoi']['trai']}, Gái: {s['sdd_cn_tuoi']['gai']})",
            f"   - Mức -2SD: {s['sdd_cn_tuoi']['muc_2sd']}, Mức -3SD: {s['sdd_cn_tuoi']['muc_3sd']}",
            f"   Tỷ lệ SDD: {s['sdd_cn_tuoi']['ty_le']}%",
            f"",
            f"5. SDD CHIỀU CAO/TUỔI: {s['sdd_cc_tuoi']['tong']} (Trai: {s['sdd_cc_tuoi']['trai']}, Gái: {s['sdd_cc_tuoi']['gai']})",
            f"   - Mức -2SD: {s['sdd_cc_tuoi']['muc_2sd']}, Mức -3SD: {s['sdd_cc_tuoi']['muc_3sd']}",
            f"   Tỷ lệ SDD: {s['sdd_cc_tuoi']['ty_le']}%",
            f"",
            f"6. SDD CÂN NẶNG/CHIỀU CAO: {s['sdd_cn_cc']['tong']} (Trai: {s['sdd_cn_cc']['trai']}, Gái: {s['sdd_cn_cc']['gai']})",
            f"   - Mức -2SD: {s['sdd_cn_cc']['muc_2sd']}, Mức -3SD: {s['sdd_cn_cc']['muc_3sd']}",
            f"",
            f"7. THỪA CÂN, BÉO PHÌ (CN/Tuổi +2SD, +3SD): {s['thua_can_beo_phi']['tong']} (Trai: {s['thua_can_beo_phi']['trai']}, Gái: {s['thua_can_beo_phi']['gai']})",
            f"",
            f"{'='*60}",
        ]
        return "\n".join(lines)
    
    def export_to_excel(self):
        output_file = self.file_path.replace('.xlsx', '_ketqua.xlsx')
        under5_table = None
        if self.var_under5_table.get():
            summary_5 = self.summary_statistics(max_months=60)
            under5_table = build_under5_statistics_table(summary_5)
        export_to_excel(
            self.df_children,
            input_file=self.file_path,
            output_file=output_file,
            under5_table=under5_table,
            under5_sheet_name='Thong ke <5T'
        )

        

def main():
    app = mainApp()
    app.run()


if __name__ == "__main__":
    main()
