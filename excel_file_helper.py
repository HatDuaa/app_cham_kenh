import openpyxl
import pandas as pd
import os
import shutil




def load_cc_per_old(file_path: str = 'CC_tuoi.xlsx') -> pd.DataFrame:

    df = pd.read_excel(file_path, skiprows=1)

    cols = ['thang_tuoi', 'minus_3sd', 'minus_2sd', 'trung_vi', 'plus_2sd', 'plus_3sd']

    df_trai = df.iloc[:, 0:6].copy()
    df_trai.columns = cols
    df_trai['gioi_tinh'] = 'trai'

    df_gai = df.iloc[:, 7:13].copy()
    df_gai.columns = cols
    df_gai['gioi_tinh'] = 'gai'

    result = pd.concat([df_trai, df_gai], ignore_index=True)
    result = result.dropna()

    return result


def load_cn_per_old(file_path: str = 'CN_tuoi.xlsx') -> pd.DataFrame:

    df = pd.read_excel(file_path, skiprows=1)

    cols = ['thang_tuoi', 'minus_3sd', 'minus_2sd', 'trung_vi', 'plus_2sd', 'plus_3sd']

    df_trai = df.iloc[:, 0:6].copy()
    df_trai.columns = cols
    df_trai['gioi_tinh'] = 'trai'

    df_gai = df.iloc[:, 7:13].copy()
    df_gai.columns = cols
    df_gai['gioi_tinh'] = 'gai'

    result = pd.concat([df_trai, df_gai], ignore_index=True)
    result = result.dropna()

    return result


def load_cn_cc(file_path: str) -> pd.DataFrame:
    """Load cân nặng theo chiều cao"""
    df = pd.read_excel(file_path, skiprows=1)

    cols = ['chieu_cao', 'minus_3sd', 'minus_2sd', 'trung_vi', 'plus_2sd', 'plus_3sd']

    df_trai = df.iloc[:, 0:6].copy()
    df_trai.columns = cols
    df_trai['gioi_tinh'] = 'trai'

    df_gai = df.iloc[:, 7:13].copy()
    df_gai.columns = cols
    df_gai['gioi_tinh'] = 'gai'

    result = pd.concat([df_trai, df_gai], ignore_index=True)
    result = result.dropna()

    return result


def load_danh_sach_can_do(file_path: str, sheet_name: str = 'Sheet1') -> pd.DataFrame:
    """Load danh sách cân đo trẻ em (format mới 2025)"""
    df = pd.read_excel(file_path, sheet_name=sheet_name, skiprows=6, header=None)
    
    # Cột: 0-STT, 1-Họ tên, 2-Nam, 3-Nữ, 4-Ngày sinh, 5-Tháng tuổi, 6-Họ tên mẹ, 
    #      7-Cân nặng, 8-TT CN/Tuổi, 9-Chiều cao, 10-TT CC/Tuổi, 11-CN/CC
    df = df.iloc[:, [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11]]
    
    df.columns = [
        'stt', 'ho_ten', 'nam', 'nu', 'ngay_sinh', 'thang_tuoi',
        'ho_ten_me', 'can_nang', 'execute_cn_tuoi', 'chieu_cao', 'execute_cc_tuoi', 'execute_cn_cc'
    ]
    
    # Loại bỏ các dòng không phải data (stt không phải số)
    df = df[pd.to_numeric(df['stt'], errors='coerce').notna()]
    
    # Tạo cột note để lưu giá trị không hợp lệ
    df['note'] = ''
    
    # Lưu giá trị không hợp lệ vào cột note trước khi chuyển đổi
    for col in ['can_nang', 'chieu_cao', 'thang_tuoi']:
        # Tìm các giá trị không phải số
        invalid_mask = pd.to_numeric(df[col], errors='coerce').isna() & df[col].notna()
        # Copy giá trị không hợp lệ vào note
        df.loc[invalid_mask, 'note'] = df.loc[invalid_mask, 'note'] + df.loc[invalid_mask, col].astype(str) + '; '
    
    # Chuyển các cột số về đúng kiểu dữ liệu (giá trị không hợp lệ thành NaN)
    df['stt'] = pd.to_numeric(df['stt'], errors='coerce')
    df['thang_tuoi'] = pd.to_numeric(df['thang_tuoi'], errors='coerce')
    df['can_nang'] = pd.to_numeric(df['can_nang'], errors='coerce')
    df['chieu_cao'] = pd.to_numeric(df['chieu_cao'], errors='coerce')
    
    # Xóa dấu ; thừa ở cuối note
    df['note'] = df['note'].str.rstrip('; ')
    
    # Thêm cột giới tính dựa trên cột nam/nu
    df['gioi_tinh'] = df.apply(
        lambda row: 'trai' if pd.notna(row['nam']) and row['nam'] == 'x' else 'gai', axis=1
    )
    
    # Reset index
    df = df.reset_index(drop=True)
    
    return df


def export_to_excel(
    df_result: pd.DataFrame,
    input_file: str = None,
    output_file: str = None,
    data_start_row: int = 7,
    sheet_name: str = None
) -> None:
    """
    Copy file Excel gốc và ghi kết quả vào, giữ nguyên định dạng.
    
    Args:
        df_result: DataFrame chứa kết quả cần ghi
        input_file: File Excel gốc làm template
        output_file: File Excel đầu ra (None = ghi đè file gốc)
        data_start_row: Hàng bắt đầu ghi data (mặc định là 8)
        sheet_name: Tên sheet (None = sheet đầu tiên)
    """
    # Nếu không có output_file, ghi đè file gốc
    if output_file is None:
        output_file = input_file
    
    # Chuẩn hóa đường dẫn
    input_file = os.path.normpath(input_file)
    output_file = os.path.normpath(output_file)
    
    # Copy file gốc sang file mới (giữ nguyên định dạng)
    if input_file != output_file:
        shutil.copy2(input_file, output_file)
    
    # Mở file bằng openpyxl
    wb = openpyxl.load_workbook(output_file)
    ws = wb.active if sheet_name is None else wb[sheet_name]
    
    # Mapping cột DataFrame -> cột Excel (format mới 2025)
    column_mapping = {
        'stt': 'A',
        'ho_ten': 'B',
        'nam': 'C',
        'nu': 'D',
        'ngay_sinh': 'E',
        'thang_tuoi': 'F',
        'ho_ten_me': 'G',
        'can_nang': 'H',
        'execute_cn_tuoi': 'I',
        'chieu_cao': 'J',
        'execute_cc_tuoi': 'K',
        'execute_cn_cc': 'L',
        'note': 'M',
    }
    
    # Ghi từng dòng data
    for idx, row in df_result.iterrows():
        excel_row = data_start_row + idx
        
        for col_name, excel_col in column_mapping.items():
            if col_name in row.index:
                value = row[col_name]
                # Xử lý giá trị NaN
                if pd.isna(value):
                    cell_value = None
                elif col_name in ['stt', 'thang_tuoi'] and pd.notna(value):
                    cell_value = int(value)
                else:
                    cell_value = value
                
                ws[f"{excel_col}{excel_row}"] = cell_value
    
    # Lưu file
    wb.save(output_file)
    wb.close()
        