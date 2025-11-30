import openpyxl
import pandas as pd





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


def load_danh_sach_can_do(file_path: str = 'C HA.xlsx') -> pd.DataFrame:
    """Load danh sách cân đo trẻ em"""
    df = pd.read_excel(file_path, skiprows=6, header=None)
    
    df = df.iloc[:, [0, 1, 2, 3, 4, 5, 7, 10]]
    
    df.columns = [
        'stt', 'ho_ten', 'nam', 'nu', 'ngay_sinh', 'thang_tuoi',
        'can_nang', 'chieu_cao'
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
    header_rows: int = 5,
    footer_start_keyword: str = 'Tổng hợp trang'
) -> None:
    """
    Ghi kết quả vào file Excel mới dựa trên template file gốc.
    Thay thế cột A-O (15 cột) bằng 12 cột mới, giữ nguyên cột P trở đi.
    
    Args:
        df_result: DataFrame chứa kết quả cần ghi
        input_file: File Excel gốc làm template
        output_file: File Excel đầu ra (None = ghi đè file gốc)
        header_rows: Số dòng header title
        footer_start_keyword: Từ khóa bắt đầu footer
    """
    # Nếu không có output_file, ghi đè file gốc
    if output_file is None:
        output_file = input_file
    
    # Đọc file gốc
    df_original = pd.read_excel(input_file, header=None)
    
    # Tìm dòng bắt đầu footer
    footer_mask = df_original.iloc[:, 0].astype(str).str.contains(footer_start_keyword, na=False)
    footer_start = footer_mask.idxmax() if footer_mask.any() else len(df_original)
    
    # Data bắt đầu từ dòng 7 (5 title + 2 header cột)
    data_start = header_rows + 2
    
    # Lấy các cột từ P trở đi (index 15+) của phần data
    df_extra_cols = df_original.iloc[data_start:footer_start, 15:].copy().reset_index(drop=True)
    
    # Chuẩn bị data mới (12 cột - thêm cột Note)
    data_rows = []
    for _, row in df_result.iterrows():
        data_row = [
            int(row['stt']) if pd.notna(row['stt']) else '',
            row['ho_ten'] if pd.notna(row.get('ho_ten')) else '',
            row['nam'] if pd.notna(row.get('nam')) else '',
            row['nu'] if pd.notna(row.get('nu')) else '',
            row['ngay_sinh'] if pd.notna(row.get('ngay_sinh')) else '',
            int(row['thang_tuoi']) if pd.notna(row.get('thang_tuoi')) else '',
            row['can_nang'] if pd.notna(row.get('can_nang')) else '',
            row.get('execute_cn_tuoi', '') if pd.notna(row.get('execute_cn_tuoi')) else '',
            row['chieu_cao'] if pd.notna(row.get('chieu_cao')) else '',
            row.get('execute_cc_tuoi', '') if pd.notna(row.get('execute_cc_tuoi')) else '',
            row.get('execute_cn_cc', '') if pd.notna(row.get('execute_cn_cc')) else '',
            row.get('note', '') if pd.notna(row.get('note')) else '',  # Cột Note
        ]
        data_rows.append(data_row)
    
    df_data = pd.DataFrame(data_rows)
    
    # Ghép data với các cột extra (P trở đi)
    df_data_full = pd.concat([df_data, df_extra_cols], axis=1, ignore_index=True)
    
    # Header cột mới (12 cột - thêm Note)
    new_header = ['Stt', 'Họ tên trẻ', 'Nam', 'Nữ', 'Ngày sinh', 'Tháng tuổi', 'Cân nặng', 'CN/Tuổi', 'Chiều cao', 'CC/Tuổi', 'CN/CC', 'Ghi chú']
    extra_header = df_original.iloc[5, 15:].tolist()  # Lấy header của cột P trở đi
    full_header = new_header + extra_header
    df_new_header = pd.DataFrame([full_header])
    
    # Header title (5 dòng đầu)
    df_header_title = df_original.iloc[:header_rows].copy()
    
    # Footer
    df_footer = df_original.iloc[footer_start:].copy()
    
    # Ghép tất cả
    df_output = pd.concat([df_header_title, df_new_header, df_data_full, df_footer], ignore_index=True)
    
    # Ghi file
    df_output.to_excel(output_file, index=False, header=False)
    print(f"Đã xuất kết quả ra file: {output_file}")
        