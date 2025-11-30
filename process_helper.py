import pandas as pd
import numpy as np
from datetime import datetime
from openpyxl import load_workbook


def calculate_month_age(birth_date, target_date) -> int:
    if pd.isna(birth_date):
        return None
    
    birth = pd.to_datetime(birth_date)
    target = pd.to_datetime(target_date)
    
    # Tính số tháng (chỉ quan tâm năm và tháng)
    months = (target.year - birth.year) * 12 + (target.month - birth.month)
    
    return months


def apply_month_age(df: pd.DataFrame, target_date, birth_col: str = 'ngay_sinh', result_col: str = 'thang_tuoi') -> pd.DataFrame:
    df = df.copy()
    df[result_col] = df[birth_col].apply(lambda x: calculate_month_age(x, target_date))
    return df


def execute_weight_by_age(df_children: pd.DataFrame, df_weight_by_age: pd.DataFrame) -> pd.DataFrame:
    df = pd.merge(df_children, df_weight_by_age, how='left', left_on=['thang_tuoi', 'gioi_tinh'], right_on=['thang_tuoi', 'gioi_tinh'])
    
    conditions = [
        df['can_nang'] < df['minus_3sd'],
        df['can_nang'] < df['minus_2sd'],
        df['can_nang'] <= df['plus_2sd'],
        df['can_nang'] <= df['plus_3sd'],
        df['can_nang'] > df['plus_3sd'],
    ]
    choices = ['-3 SD', '-2 SD', 'BT', '+2 SD', '+3 SD']
    
    df['execute_cn_tuoi'] = np.select(conditions, choices, default=None)
    
    df = df.drop(columns=df_weight_by_age.columns.difference(['thang_tuoi', 'gioi_tinh']))
    
    return df


def execute_height_by_age(df_children: pd.DataFrame, df_height_by_age: pd.DataFrame) -> pd.DataFrame:
    df = pd.merge(df_children, df_height_by_age, how='left', left_on=['thang_tuoi', 'gioi_tinh'], right_on=['thang_tuoi', 'gioi_tinh'])
    
    conditions = [
        df['chieu_cao'] < df['minus_3sd'],
        df['chieu_cao'] < df['minus_2sd'],
        df['chieu_cao'] <= df['plus_2sd'],
        df['chieu_cao'] <= df['plus_3sd'],
        df['chieu_cao'] > df['plus_3sd'],
    ]
    choices = ['-3 SD', '-2 SD', 'BT', '+2 SD', '+3 SD']
    
    df['execute_cc_tuoi'] = np.select(conditions, choices, default=None)
    
    # Bỏ các cột tham chiếu
    df = df.drop(columns=df_height_by_age.columns.difference(['thang_tuoi', 'gioi_tinh']))
    
    return df


def execute_weight_by_height(df_children: pd.DataFrame, df_weight_by_height_0_2: pd.DataFrame, df_weight_by_height_2_5: pd.DataFrame) -> pd.DataFrame:
    # Tách những hàng có chieu_cao và không có
    df_with_height = df_children[df_children['chieu_cao'].notna()].copy()
    df_without_height = df_children[df_children['chieu_cao'].isna()].copy()
    
    # Sắp xếp theo chieu_cao để dùng merge_asof
    df_with_height = df_with_height.sort_values('chieu_cao')
    
    # Tách trẻ 0-24 tháng và trẻ > 24 tháng
    df_0_24 = df_with_height[df_with_height['thang_tuoi'] <= 24]
    df_25_60 = df_with_height[df_with_height['thang_tuoi'] > 24]
    
    results = []
    
    # Merge_asof cho từng giới tính (0-24 tháng)
    for gioi_tinh in ['trai', 'gai']:
        df_child = df_0_24[df_0_24['gioi_tinh'] == gioi_tinh]
        df_ref = df_weight_by_height_0_2[df_weight_by_height_0_2['gioi_tinh'] == gioi_tinh].sort_values('chieu_cao')
        if not df_child.empty:
            merged = pd.merge_asof(df_child, df_ref, on='chieu_cao', direction='backward', suffixes=('', '_ref'))
            merged = merged.drop(columns=['gioi_tinh_ref'], errors='ignore')
            results.append(merged)
    
    # Merge_asof cho từng giới tính (25-60 tháng)
    for gioi_tinh in ['trai', 'gai']:
        df_child = df_25_60[df_25_60['gioi_tinh'] == gioi_tinh]
        df_ref = df_weight_by_height_2_5[df_weight_by_height_2_5['gioi_tinh'] == gioi_tinh].sort_values('chieu_cao')
        if not df_child.empty:
            merged = pd.merge_asof(df_child, df_ref, on='chieu_cao', direction='backward', suffixes=('', '_ref'))
            merged = merged.drop(columns=['gioi_tinh_ref'], errors='ignore')
            results.append(merged)
    
    # Gộp lại với những hàng không có chiều cao
    results.append(df_without_height)
    df = pd.concat(results, ignore_index=True)
    
    conditions = [
        df['can_nang'] < df['minus_3sd'],
        df['can_nang'] < df['minus_2sd'],
        df['can_nang'] <= df['plus_2sd'],
        df['can_nang'] <= df['plus_3sd'],
        df['can_nang'] > df['plus_3sd'],
    ]
    choices = ['-3 SD', '-2 SD', 'BT', '+2 SD', '+3 SD']
    
    df['execute_cn_cc'] = np.select(conditions, choices, default=None)
    
    # Bỏ các cột tham chiếu
    df = df.drop(columns=df_weight_by_height_0_2.columns.difference(['chieu_cao', 'gioi_tinh']))
    
    # Sắp xếp lại theo stt và reset index
    df = df.sort_values('stt').reset_index(drop=True)
    
    return df


def summary_statistics(df: pd.DataFrame, max_months: int = None) -> dict:
    """
    Tổng kết dinh dưỡng
    Args:
        df: DataFrame chứa dữ liệu trẻ em
        max_months: Giới hạn tháng tuổi (None = tất cả, 24 = dưới 2 tuổi, 60 = dưới 5 tuổi)
    """
    # Lọc theo tháng tuổi nếu có
    if max_months:
        df_filtered = df[df['thang_tuoi'] <= max_months].copy()
    else:
        df_filtered = df.copy()
    
    # === TỔNG SỐ TRẺ ===
    total = len(df_filtered)
    total_trai = len(df_filtered[df_filtered['gioi_tinh'] == 'trai'])
    total_gai = len(df_filtered[df_filtered['gioi_tinh'] == 'gai'])
    
    # === TRẺ ĐƯỢC CÂN ===
    df_can = df_filtered[df_filtered['can_nang'].notna()]
    can_trai = len(df_can[df_can['gioi_tinh'] == 'trai'])
    can_gai = len(df_can[df_can['gioi_tinh'] == 'gai'])
    ty_le_can = len(df_can) / total * 100 if total > 0 else 0
    
    # === TRẺ ĐƯỢC ĐO ===
    df_do = df_filtered[df_filtered['chieu_cao'].notna()]
    do_trai = len(df_do[df_do['gioi_tinh'] == 'trai'])
    do_gai = len(df_do[df_do['gioi_tinh'] == 'gai'])
    ty_le_do = len(df_do) / total * 100 if total > 0 else 0
    
    # === SDD CÂN NẶNG/TUỔI ===
    # SDD: -2SD và -3SD
    sdd_cn_tuoi = df_can[df_can['execute_cn_tuoi'].isin(['-2 SD', '-3 SD'])]
    sdd_cn_tuoi_trai = len(sdd_cn_tuoi[sdd_cn_tuoi['gioi_tinh'] == 'trai'])
    sdd_cn_tuoi_gai = len(sdd_cn_tuoi[sdd_cn_tuoi['gioi_tinh'] == 'gai'])
    sdd_cn_tuoi_2sd = len(df_can[df_can['execute_cn_tuoi'] == '-2 SD'])
    sdd_cn_tuoi_3sd = len(df_can[df_can['execute_cn_tuoi'] == '-3 SD'])
    # Thừa cân: +2SD, +3SD
    thua_cn_tuoi = df_can[df_can['execute_cn_tuoi'].isin(['+2 SD', '+3 SD'])]
    thua_cn_tuoi_2sd_3sd = len(thua_cn_tuoi)
    ty_le_sdd_cn_tuoi = len(sdd_cn_tuoi) / len(df_can) * 100 if len(df_can) > 0 else 0
    
    # === SDD CHIỀU CAO/TUỔI ===
    sdd_cc_tuoi = df_do[df_do['execute_cc_tuoi'].isin(['-2 SD', '-3 SD'])]
    sdd_cc_tuoi_trai = len(sdd_cc_tuoi[sdd_cc_tuoi['gioi_tinh'] == 'trai'])
    sdd_cc_tuoi_gai = len(sdd_cc_tuoi[sdd_cc_tuoi['gioi_tinh'] == 'gai'])
    sdd_cc_tuoi_2sd = len(df_do[df_do['execute_cc_tuoi'] == '-2 SD'])
    sdd_cc_tuoi_3sd = len(df_do[df_do['execute_cc_tuoi'] == '-3 SD'])
    ty_le_sdd_cc_tuoi = len(sdd_cc_tuoi) / len(df_do) * 100 if len(df_do) > 0 else 0
    
    # === SDD CÂN NẶNG/CHIỀU CAO ===
    df_cn_cc = df_filtered[df_filtered['execute_cn_cc'].notna()]
    sdd_cn_cc = df_cn_cc[df_cn_cc['execute_cn_cc'].isin(['-2 SD', '-3 SD'])]
    sdd_cn_cc_trai = len(sdd_cn_cc[sdd_cn_cc['gioi_tinh'] == 'trai'])
    sdd_cn_cc_gai = len(sdd_cn_cc[sdd_cn_cc['gioi_tinh'] == 'gai'])
    sdd_cn_cc_2sd = len(df_cn_cc[df_cn_cc['execute_cn_cc'] == '-2 SD'])
    sdd_cn_cc_3sd = len(df_cn_cc[df_cn_cc['execute_cn_cc'] == '-3 SD'])
    
    # === THỪA CÂN, BÉO PHÌ (CC/Tuổi và CN/CC) ===
    # Thừa cân CC/tuổi
    thua_cc_tuoi = df_do[df_do['execute_cc_tuoi'].isin(['+2 SD', '+3 SD'])]
    thua_cc_tuoi_trai = len(thua_cc_tuoi[thua_cc_tuoi['gioi_tinh'] == 'trai'])
    thua_cc_tuoi_gai = len(thua_cc_tuoi[thua_cc_tuoi['gioi_tinh'] == 'gai'])
    thua_cc_tuoi_2sd = len(df_do[df_do['execute_cc_tuoi'] == '+2 SD'])
    thua_cc_tuoi_3sd = len(df_do[df_do['execute_cc_tuoi'] == '+3 SD'])
    
    # Thừa cân CN/CC
    thua_cn_cc = df_cn_cc[df_cn_cc['execute_cn_cc'].isin(['+2 SD', '+3 SD'])]
    thua_cn_cc_trai = len(thua_cn_cc[thua_cn_cc['gioi_tinh'] == 'trai'])
    thua_cn_cc_gai = len(thua_cn_cc[thua_cn_cc['gioi_tinh'] == 'gai'])
    thua_cn_cc_2sd = len(df_cn_cc[df_cn_cc['execute_cn_cc'] == '+2 SD'])
    thua_cn_cc_3sd = len(df_cn_cc[df_cn_cc['execute_cn_cc'] == '+3 SD'])
    
    result = {
        'tong_so_tre': {'tong': total, 'trai': total_trai, 'gai': total_gai},
        'tre_duoc_can': {
            'tong': len(df_can), 'trai': can_trai, 'gai': can_gai,
            'ty_le': round(ty_le_can, 2)
        },
        'tre_duoc_do': {
            'tong': len(df_do), 'trai': do_trai, 'gai': do_gai,
            'ty_le': round(ty_le_do, 2)
        },
        'sdd_cn_tuoi': {
            'tong': len(sdd_cn_tuoi), 'trai': sdd_cn_tuoi_trai, 'gai': sdd_cn_tuoi_gai,
            'muc_2sd': sdd_cn_tuoi_2sd, 'muc_3sd': sdd_cn_tuoi_3sd,
            'thua_2sd_3sd': thua_cn_tuoi_2sd_3sd,
            'ty_le': round(ty_le_sdd_cn_tuoi, 2)
        },
        'sdd_cc_tuoi': {
            'tong': len(sdd_cc_tuoi), 'trai': sdd_cc_tuoi_trai, 'gai': sdd_cc_tuoi_gai,
            'muc_2sd': sdd_cc_tuoi_2sd, 'muc_3sd': sdd_cc_tuoi_3sd,
            'ty_le': round(ty_le_sdd_cc_tuoi, 2)
        },
        'sdd_cn_cc': {
            'tong': len(sdd_cn_cc), 'trai': sdd_cn_cc_trai, 'gai': sdd_cn_cc_gai,
            'muc_2sd': sdd_cn_cc_2sd, 'muc_3sd': sdd_cn_cc_3sd
        },
        'thua_can_beo_phi': {
            'cc_tuoi': {
                'tong': len(thua_cc_tuoi), 'trai': thua_cc_tuoi_trai, 'gai': thua_cc_tuoi_gai,
                'muc_2sd': thua_cc_tuoi_2sd, 'muc_3sd': thua_cc_tuoi_3sd
            },
            'cn_cc': {
                'tong': len(thua_cn_cc), 'trai': thua_cn_cc_trai, 'gai': thua_cn_cc_gai,
                'muc_2sd': thua_cn_cc_2sd, 'muc_3sd': thua_cn_cc_3sd
            }
        }
    }
    
    return result


def print_summary(summary: dict, title: str = "TRẺ DƯỚI 2 TUỔI"):
    """In báo cáo tổng kết"""
    print(f"\n{'='*60}")
    print(f"TỔNG KẾT DINH DƯỠNG - {title}")
    print(f"{'='*60}")
    
    s = summary
    print(f"\n1. TỔNG SỐ TRẺ: {s['tong_so_tre']['tong']} (Trai: {s['tong_so_tre']['trai']}, Gái: {s['tong_so_tre']['gai']})")
    
    print(f"\n2. TRẺ ĐƯỢC CÂN: {s['tre_duoc_can']['tong']} (Trai: {s['tre_duoc_can']['trai']}, Gái: {s['tre_duoc_can']['gai']})")
    print(f"   Tỷ lệ: {s['tre_duoc_can']['ty_le']}%")
    
    print(f"\n3. TRẺ ĐƯỢC ĐO: {s['tre_duoc_do']['tong']} (Trai: {s['tre_duoc_do']['trai']}, Gái: {s['tre_duoc_do']['gai']})")
    print(f"   Tỷ lệ: {s['tre_duoc_do']['ty_le']}%")
    
    print(f"\n4. SDD CÂN NẶNG/TUỔI: {s['sdd_cn_tuoi']['tong']} (Trai: {s['sdd_cn_tuoi']['trai']}, Gái: {s['sdd_cn_tuoi']['gai']})")
    print(f"   - Mức -2SD: {s['sdd_cn_tuoi']['muc_2sd']}, Mức -3SD: {s['sdd_cn_tuoi']['muc_3sd']}")
    print(f"   - Thừa cân (+2SD, +3SD): {s['sdd_cn_tuoi']['thua_2sd_3sd']}")
    print(f"   Tỷ lệ SDD: {s['sdd_cn_tuoi']['ty_le']}%")
    
    print(f"\n5. SDD CHIỀU CAO/TUỔI: {s['sdd_cc_tuoi']['tong']} (Trai: {s['sdd_cc_tuoi']['trai']}, Gái: {s['sdd_cc_tuoi']['gai']})")
    print(f"   - Mức -2SD: {s['sdd_cc_tuoi']['muc_2sd']}, Mức -3SD: {s['sdd_cc_tuoi']['muc_3sd']}")
    print(f"   Tỷ lệ SDD: {s['sdd_cc_tuoi']['ty_le']}%")
    
    print(f"\n6. SDD CÂN NẶNG/CHIỀU CAO: {s['sdd_cn_cc']['tong']} (Trai: {s['sdd_cn_cc']['trai']}, Gái: {s['sdd_cn_cc']['gai']})")
    print(f"   - Mức -2SD: {s['sdd_cn_cc']['muc_2sd']}, Mức -3SD: {s['sdd_cn_cc']['muc_3sd']}")
    
    print(f"\n7. THỪA CÂN, BÉO PHÌ:")
    tc = s['thua_can_beo_phi']
    print(f"   CC/Tuổi: {tc['cc_tuoi']['tong']} (Trai: {tc['cc_tuoi']['trai']}, Gái: {tc['cc_tuoi']['gai']})")
    print(f"      - Mức +2SD: {tc['cc_tuoi']['muc_2sd']}, Mức +3SD: {tc['cc_tuoi']['muc_3sd']}")
    print(f"   CN/CC: {tc['cn_cc']['tong']} (Trai: {tc['cn_cc']['trai']}, Gái: {tc['cn_cc']['gai']})")
    print(f"      - Mức +2SD: {tc['cn_cc']['muc_2sd']}, Mức +3SD: {tc['cn_cc']['muc_3sd']}")
    
    print(f"\n{'='*60}\n")


def write_column_to_excel(
    df: pd.DataFrame,
    df_column: str,
    excel_path: str,
    excel_column: str,
    start_row: int = 7,
    sheet_name: str = None
) -> None:
    """
    Ghi đè 1 cột từ DataFrame vào 1 cột trong file Excel.
    Chỉ ghi đè phần nội dung, giữ nguyên header.
    
    Args:
        df: DataFrame chứa dữ liệu cần ghi
        df_column: Tên cột trong DataFrame cần lấy dữ liệu
        excel_path: Đường dẫn file Excel
        excel_column: Tên cột trong Excel (A, B, C, ... hoặc AA, AB, ...)
        start_row: Hàng bắt đầu ghi (mặc định là 8)
        sheet_name: Tên sheet (None = sheet đầu tiên)
    """
    # Load workbook
    wb = load_workbook(excel_path)
    ws = wb.active if sheet_name is None else wb[sheet_name]
    
    # Lấy dữ liệu từ DataFrame
    data = df[df_column].tolist()
    
    # Ghi từng giá trị vào Excel
    for i, value in enumerate(data):
        row = start_row + i
        cell = f"{excel_column}{row}"
        
        # Xử lý giá trị NaN/None
        if pd.isna(value):
            ws[cell] = None
        else:
            ws[cell] = value
    
    # Lưu file
    wb.save(excel_path)
    wb.close()
    
    print(f"Đã ghi {len(data)} giá trị từ cột '{df_column}' vào cột {excel_column} (từ hàng {start_row})")