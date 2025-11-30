from excel_file_helper import *
from process_helper import *
from ui_helper import *
import numpy as np



# # Load all data
# df_weight_by_age = load_cn_per_old()                        # Cân nặng theo tuổi
# df_height_by_age = load_cc_per_old()                        # Chiều cao theo tuổi
# df_weight_by_height_0_2 = load_cn_cc('CN_CC_0-2_tuoi.xlsx') # Cân nặng theo chiều cao (0-24 tháng)
# df_weight_by_height_2_5 = load_cn_cc('CN_CC_2-5_tuoi.xlsx') # Cân nặng theo chiều cao (24-60 tháng)
# df_children_records = load_danh_sach_can_do()               # Danh sách cân đo trẻ em



# df_children_records = execute_weight_by_age(df_children_records, df_weight_by_age)
# df_children_records = execute_height_by_age(df_children_records, df_height_by_age)
# df_children_records = execute_weight_by_height(df_children_records, df_weight_by_height_0_2, df_weight_by_height_2_5)

# # Tổng kết trẻ dưới 2 tuổi (≤24 tháng)
# summary_under_2 = summary_statistics(df_children_records, max_months=24)
# print_summary(summary_under_2, "TRẺ DƯỚI 2 TUỔI (≤24 tháng)")

# # Tổng kết trẻ dưới 5 tuổi (≤60 tháng) - tất cả
# summary_under_5 = summary_statistics(df_children_records, max_months=60)
# print_summary(summary_under_5, "TRẺ DƯỚI 5 TUỔI (≤60 tháng)")

# # Ghi đè kết quả vào file Excel
# export_to_excel(df_children_records, file_path='C HA.xlsx')

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
        self.df_children = apply_month_age(self.df_children, birth_col='ngay_sinh', target_date='ngay_do')
    
    def execute_weight_by_age(self):
        self.df_children = execute_weight_by_age(self.df_children, self.df_weight_by_age)
        if self.var_overwrite.get():
            write_column_to_excel(self.df_children, 'execute_cn_tuoi', self.file_path, 'I')
    
    def execute_height_by_age(self):
        self.df_children = execute_height_by_age(self.df_children, self.df_height_by_age)
        if self.var_overwrite.get():
            write_column_to_excel(self.df_children, 'execute_cc_tuoi', self.file_path, 'K')
    
    def execute_weight_by_height(self):
        self.df_children = execute_weight_by_height(self.df_children, self.df_weight_by_height_0_2, self.df_weight_by_height_2_5)
        if self.var_overwrite.get():
            write_column_to_excel(self.df_children, 'execute_cn_cc', self.file_path, 'L')
    
    def summary_statistics(self, df, max_months):
        return summary_statistics(df, max_months=max_months)
    
    def format_summary(self, summary: dict, title: str = "TRẺ DƯỚI 2 TUỔI") -> str:
        """Chuyển dict thống kê thành text"""
        s = summary
        tc = s['thua_can_beo_phi']
        
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
            f"   - Thừa cân (+2SD, +3SD): {s['sdd_cn_tuoi']['thua_2sd_3sd']}",
            f"   Tỷ lệ SDD: {s['sdd_cn_tuoi']['ty_le']}%",
            f"",
            f"5. SDD CHIỀU CAO/TUỔI: {s['sdd_cc_tuoi']['tong']} (Trai: {s['sdd_cc_tuoi']['trai']}, Gái: {s['sdd_cc_tuoi']['gai']})",
            f"   - Mức -2SD: {s['sdd_cc_tuoi']['muc_2sd']}, Mức -3SD: {s['sdd_cc_tuoi']['muc_3sd']}",
            f"   Tỷ lệ SDD: {s['sdd_cc_tuoi']['ty_le']}%",
            f"",
            f"6. SDD CÂN NẶNG/CHIỀU CAO: {s['sdd_cn_cc']['tong']} (Trai: {s['sdd_cn_cc']['trai']}, Gái: {s['sdd_cn_cc']['gai']})",
            f"   - Mức -2SD: {s['sdd_cn_cc']['muc_2sd']}, Mức -3SD: {s['sdd_cn_cc']['muc_3sd']}",
            f"",
            f"7. THỪA CÂN, BÉO PHÌ:",
            f"   CC/Tuổi: {tc['cc_tuoi']['tong']} (Trai: {tc['cc_tuoi']['trai']}, Gái: {tc['cc_tuoi']['gai']})",
            f"      - Mức +2SD: {tc['cc_tuoi']['muc_2sd']}, Mức +3SD: {tc['cc_tuoi']['muc_3sd']}",
            f"   CN/CC: {tc['cn_cc']['tong']} (Trai: {tc['cn_cc']['trai']}, Gái: {tc['cn_cc']['gai']})",
            f"      - Mức +2SD: {tc['cn_cc']['muc_2sd']}, Mức +3SD: {tc['cn_cc']['muc_3sd']}",
            f"",
            f"{'='*60}",
        ]
        return "\n".join(lines)
    
    def export_to_excel(self):
        output_file = self.file_path.replace('.xlsx', '_ketqua.xlsx')
        export_to_excel(self.df_children, input_file=self.file_path, output_file=output_file)

        

def main():
    app = mainApp()
    app.run()


if __name__ == "__main__":
    main()
