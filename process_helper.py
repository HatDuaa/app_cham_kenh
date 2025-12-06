from typing import Optional
import pandas as pd
import numpy as np
from datetime import datetime
from openpyxl import load_workbook


def calculate_month_age(birth_date: datetime, target_date: datetime) -> Optional[int]:
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
    
    df['execute_cn_tuoi'] = np.select(conditions, choices, default='')
    
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
    
    df['execute_cc_tuoi'] = np.select(conditions, choices, default='')
    
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
    
    df['execute_cn_cc'] = np.select(conditions, choices, default='')
    
    # Bỏ các cột tham chiếu
    df = df.drop(columns=df_weight_by_height_0_2.columns.difference(['chieu_cao', 'gioi_tinh']))
    
    # Sắp xếp lại theo stt và reset index
    df = df.sort_values('stt').reset_index(drop=True)
    
    return df


def adjust_height_by_age(
    df_children: pd.DataFrame, 
    df_height_by_age: pd.DataFrame,
    increase_range: tuple = (0.5, 2.0)
) -> pd.DataFrame:
    """
    Điều chỉnh chiều cao: tăng random một khoảng nhưng không vượt ngưỡng SD (giữ nguyên trạng thái).
    
    Args:
        df_children: DataFrame trẻ em (cần có cột thang_tuoi, gioi_tinh, execute_cc_tuoi, chieu_cao)
        df_height_by_age: DataFrame chuẩn chiều cao/tuổi
        increase_range: Tuple (min, max) khoảng tăng thêm (cm), mặc định (0.5, 2.0)
    
    Returns:
        DataFrame với cột 'chieu_cao_adj' đã được điều chỉnh
    """
    df = df_children.copy()
    
    # Merge với bảng chuẩn
    df = pd.merge(df, df_height_by_age, how='left', on=['thang_tuoi', 'gioi_tinh'], suffixes=('', '_ref'))
    
    # Tạo random increase cho mỗi hàng
    n = len(df)
    random_increases = np.random.uniform(increase_range[0], increase_range[1], n)
    
    def calc_height_adj(row, increase):
        """Tăng chiều cao, giữ trong khoảng (lower, upper) để không thay đổi trạng thái"""
        status = row['execute_cc_tuoi']
        current_height = row['chieu_cao']
        
        if pd.isna(current_height):
            return np.nan
        
        # Xác định ngưỡng dưới và ngưỡng trên dựa vào trạng thái
        if status == '-3 SD':
            lower_limit = current_height  # Giữ nguyên hoặc tăng
            upper_limit = row['minus_3sd']
        elif status == '-2 SD':
            lower_limit = row['minus_3sd']
            upper_limit = row['minus_2sd']
        elif status == 'BT':
            lower_limit = row['minus_2sd']
            upper_limit = row['plus_2sd']
        elif status == '+2 SD':
            lower_limit = row['plus_2sd']
            upper_limit = row['plus_3sd']
        elif status == '+3 SD':
            lower_limit = row['plus_3sd']
            upper_limit = current_height + 10  # Không giới hạn trên
        else:
            return current_height
        
        # Tăng lên và giữ trong khoảng (lower, upper)
        new_height = max(current_height, lower_limit)  # Không nhỏ hơn lower
        new_height = new_height + increase
        new_height = min(new_height, upper_limit)  # Không lớn hơn upper
        return new_height
    
    # Ghi đè cột chieu_cao
    df['chieu_cao_tmp'] = [calc_height_adj(row, random_increases[idx]) for idx, (_, row) in enumerate(df.iterrows())]
    df['chieu_cao_tmp'] = df['chieu_cao_tmp'].round(1)

    df_children['chieu_cao'] = df['chieu_cao_tmp']
    
    return df_children


def adjust_weight_by_height_and_age(
    df_children: pd.DataFrame,
    df_weight_by_height_0_2: pd.DataFrame,
    df_weight_by_height_2_5: pd.DataFrame,
    df_weight_by_age: pd.DataFrame,
    height_col: str = 'chieu_cao',
    increase_range: tuple = (0.1, 0.5)
) -> pd.DataFrame:
    """
    Điều chỉnh cân nặng: tăng random một khoảng nhưng không vượt ngưỡng SD (giữ nguyên trạng thái CN/CC và CN/Tuổi).
    
    Args:
        df_children: DataFrame trẻ em
        df_weight_by_height_0_2: DataFrame chuẩn CN/CC cho trẻ 0-24 tháng
        df_weight_by_height_2_5: DataFrame chuẩn CN/CC cho trẻ 25-60 tháng
        df_weight_by_age: DataFrame chuẩn CN/Tuổi
        height_col: Tên cột chiều cao để dùng (mặc định 'chieu_cao')
        increase_range: Tuple (min, max) khoảng tăng thêm (kg), mặc định (0.2, 0.8)
    
    Returns:
        DataFrame với cột 'can_nang' đã được điều chỉnh
    """
    df = df_children.copy()
    
    # Kiểm tra cột chiều cao
    if height_col not in df.columns:
        height_col = 'chieu_cao'
    
    # Tách những hàng có chiều cao và không có
    df_with_height = df[df[height_col].notna()].copy()
    
    if df_with_height.empty:
        return df_children
    
    # Merge với bảng chuẩn CN/Tuổi để lấy ngưỡng
    df_weight_by_age_renamed = df_weight_by_age.rename(columns={
        'minus_3sd': 'minus_3sd_tuoi',
        'minus_2sd': 'minus_2sd_tuoi',
        'plus_2sd': 'plus_2sd_tuoi',
        'plus_3sd': 'plus_3sd_tuoi',
        'median': 'median_tuoi'
    })
    df_with_height = pd.merge(df_with_height, df_weight_by_age_renamed, 
                               how='left', on=['thang_tuoi', 'gioi_tinh'])
    
    # Sắp xếp theo chiều cao để dùng merge_asof
    df_with_height = df_with_height.sort_values(height_col)
    
    # Đổi tên cột tạm thời để merge_asof
    df_with_height = df_with_height.rename(columns={height_col: 'chieu_cao_temp'})
    
    # Tách trẻ 0-24 tháng và trẻ > 24 tháng
    df_0_24 = df_with_height[df_with_height['thang_tuoi'] <= 24].copy()
    df_25_60 = df_with_height[df_with_height['thang_tuoi'] > 24].copy()
    
    results = []
    
    # Merge_asof cho từng giới tính (0-24 tháng)
    for gioi_tinh in ['trai', 'gai']:
        df_child = df_0_24[df_0_24['gioi_tinh'] == gioi_tinh].copy()
        df_ref = df_weight_by_height_0_2[df_weight_by_height_0_2['gioi_tinh'] == gioi_tinh].sort_values('chieu_cao')
        if not df_child.empty and not df_ref.empty:
            merged = pd.merge_asof(
                df_child, df_ref, 
                left_on='chieu_cao_temp', right_on='chieu_cao',
                direction='nearest', suffixes=('', '_ref')
            )
            results.append(merged)
    
    # Merge_asof cho từng giới tính (25-60 tháng)
    for gioi_tinh in ['trai', 'gai']:
        df_child = df_25_60[df_25_60['gioi_tinh'] == gioi_tinh].copy()
        df_ref = df_weight_by_height_2_5[df_weight_by_height_2_5['gioi_tinh'] == gioi_tinh].sort_values('chieu_cao')
        if not df_child.empty and not df_ref.empty:
            merged = pd.merge_asof(
                df_child, df_ref,
                left_on='chieu_cao_temp', right_on='chieu_cao',
                direction='nearest', suffixes=('', '_ref')
            )
            results.append(merged)
    
    if not results:
        df['can_nang_adj'] = np.nan
        return df
    
    df_merged = pd.concat(results, ignore_index=True)
    
    # Tạo random increase cho mỗi hàng
    n = len(df_merged)
    random_increases = np.random.uniform(increase_range[0], increase_range[1], n)
    
    def calc_weight_adj(row, increase):
        """Tăng cân nặng, giữ trong khoảng (lower, upper) để không thay đổi trạng thái (cả CN/CC và CN/Tuổi)"""
        status_cn_cc = row['execute_cn_cc']
        status_cn_tuoi = row.get('execute_cn_tuoi', '')
        current_weight = row['can_nang']
        
        if pd.isna(current_weight):
            return np.nan
        
        # Xác định ngưỡng dưới và trên dựa vào trạng thái CN/CC
        if status_cn_cc == '-3 SD':
            lower_limit_cc = current_weight  # Giữ nguyên hoặc tăng
            upper_limit_cc = row['minus_3sd']
        elif status_cn_cc == '-2 SD':
            lower_limit_cc = row['minus_3sd']
            upper_limit_cc = row['minus_2sd']
        elif status_cn_cc == 'BT':
            lower_limit_cc = row['minus_2sd']
            upper_limit_cc = row['plus_2sd']
        elif status_cn_cc == '+2 SD':
            lower_limit_cc = row['plus_2sd']
            upper_limit_cc = row['plus_3sd']
        elif status_cn_cc == '+3 SD':
            lower_limit_cc = row['plus_3sd']
            upper_limit_cc = current_weight + 5  # Không giới hạn trên
        else:
            lower_limit_cc = current_weight
            upper_limit_cc = current_weight + 5
        
        # Xác định ngưỡng dưới và trên dựa vào trạng thái CN/Tuổi (nếu có)
        if status_cn_tuoi == '-3 SD':
            lower_limit_tuoi = current_weight
            upper_limit_tuoi = row.get('minus_3sd_tuoi', upper_limit_cc)
        elif status_cn_tuoi == '-2 SD':
            lower_limit_tuoi = row.get('minus_3sd_tuoi', lower_limit_cc)
            upper_limit_tuoi = row.get('minus_2sd_tuoi', upper_limit_cc)
        elif status_cn_tuoi == 'BT':
            lower_limit_tuoi = row.get('minus_2sd_tuoi', lower_limit_cc)
            upper_limit_tuoi = row.get('plus_2sd_tuoi', upper_limit_cc)
        elif status_cn_tuoi == '+2 SD':
            lower_limit_tuoi = row.get('plus_2sd_tuoi', lower_limit_cc)
            upper_limit_tuoi = row.get('plus_3sd_tuoi', upper_limit_cc)
        elif status_cn_tuoi == '+3 SD':
            lower_limit_tuoi = row.get('plus_3sd_tuoi', lower_limit_cc)
            upper_limit_tuoi = current_weight + 5
        else:
            lower_limit_tuoi = current_weight
            upper_limit_tuoi = current_weight + 5
        
        # Lấy max của lower và min của upper để đảm bảo nằm trong cả 2 khoảng
        lower_limit = max(lower_limit_cc, lower_limit_tuoi)
        upper_limit = min(upper_limit_cc, upper_limit_tuoi)
        
        # Tăng lên và giữ trong khoảng (lower, upper)
        new_weight = max(current_weight, lower_limit)  # Đảm bảo >= lower
        new_weight = new_weight + increase              # Cộng random
        new_weight = min(new_weight, upper_limit)       # Đảm bảo <= upper
        return new_weight
    
    df_merged['can_nang_tmp'] = [calc_weight_adj(row, random_increases[idx]) for idx, (_, row) in enumerate(df_merged.iterrows())]
    df_merged['can_nang_tmp'] = df_merged['can_nang_tmp'].round(1)
    
    # Tạo mapping stt -> can_nang_tmp
    adj_map = dict(zip(df_merged['stt'], df_merged['can_nang_tmp']))
    
    df_children['can_nang'] = df_children['stt'].map(adj_map).fillna(df_children['can_nang'])
    
    return df_children


def summary_statistics(df: pd.DataFrame, max_months: Optional[int] = None) -> dict:
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
    thua_cn_tuoi_trai = len(thua_cn_tuoi[thua_cn_tuoi['gioi_tinh'] == 'trai'])
    thua_cn_tuoi_gai = len(thua_cn_tuoi[thua_cn_tuoi['gioi_tinh'] == 'gai'])
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
    ty_le_sdd_cn_cc = len(sdd_cn_cc) / len(df_cn_cc) * 100 if len(df_cn_cc) > 0 else 0
    
    # === THỪA CÂN VÀ BÉO PHÌ (CN/Tuổi) ===
    # Tổng thừa cân + béo phì (+2SD và +3SD)
    thua_can_beo_phi = df_can[df_can['execute_cn_tuoi'].isin(['+2 SD', '+3 SD'])]
    thua_can_beo_phi_tong = len(thua_can_beo_phi)
    thua_can_beo_phi_trai = len(thua_can_beo_phi[thua_can_beo_phi['gioi_tinh'] == 'trai'])
    thua_can_beo_phi_gai = len(thua_can_beo_phi[thua_can_beo_phi['gioi_tinh'] == 'gai'])
    
    # Đếm riêng mức +2SD và +3SD
    thua_can_beo_phi_2sd = len(df_can[df_can['execute_cn_tuoi'] == '+2 SD'])
    thua_can_beo_phi_3sd = len(df_can[df_can['execute_cn_tuoi'] == '+3 SD'])
    
    # Thừa cân: +2SD (giữ lại để tương thích)
    thua_can_cn_tuoi = df_can[df_can['execute_cn_tuoi'] == '+2 SD']
    thua_can_cn_tuoi_tong = len(thua_can_cn_tuoi)
    thua_can_cn_tuoi_trai = len(thua_can_cn_tuoi[thua_can_cn_tuoi['gioi_tinh'] == 'trai'])
    thua_can_cn_tuoi_gai = len(thua_can_cn_tuoi[thua_can_cn_tuoi['gioi_tinh'] == 'gai'])
    
    # Béo phì: +3SD (giữ lại để tương thích)
    beo_phi_cn_tuoi = df_can[df_can['execute_cn_tuoi'] == '+3 SD']
    beo_phi_cn_tuoi_tong = len(beo_phi_cn_tuoi)
    beo_phi_cn_tuoi_trai = len(beo_phi_cn_tuoi[beo_phi_cn_tuoi['gioi_tinh'] == 'trai'])
    beo_phi_cn_tuoi_gai = len(beo_phi_cn_tuoi[beo_phi_cn_tuoi['gioi_tinh'] == 'gai'])
    
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
            'ty_le': round(ty_le_sdd_cn_tuoi, 2)
        },
        'sdd_cc_tuoi': {
            'tong': len(sdd_cc_tuoi), 'trai': sdd_cc_tuoi_trai, 'gai': sdd_cc_tuoi_gai,
            'muc_2sd': sdd_cc_tuoi_2sd, 'muc_3sd': sdd_cc_tuoi_3sd,
            'ty_le': round(ty_le_sdd_cc_tuoi, 2)
        },
        'sdd_cn_cc': {
            'tong': len(sdd_cn_cc), 'trai': sdd_cn_cc_trai, 'gai': sdd_cn_cc_gai,
            'muc_2sd': sdd_cn_cc_2sd, 'muc_3sd': sdd_cn_cc_3sd,
            'ty_le': round(ty_le_sdd_cn_cc, 2)
        },
        'thua_can_cn_tuoi': {
            'tong': thua_can_cn_tuoi_tong, 'trai': thua_can_cn_tuoi_trai, 'gai': thua_can_cn_tuoi_gai
        },
        'beo_phi_cn_tuoi': {
            'tong': beo_phi_cn_tuoi_tong, 'trai': beo_phi_cn_tuoi_trai, 'gai': beo_phi_cn_tuoi_gai
        },
        'thua_can_beo_phi': {
            'tong': thua_can_beo_phi_tong, 'trai': thua_can_beo_phi_trai, 'gai': thua_can_beo_phi_gai,
            'muc_2sd': thua_can_beo_phi_2sd, 'muc_3sd': thua_can_beo_phi_3sd
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
    print(f"   Tỷ lệ SDD: {s['sdd_cn_tuoi']['ty_le']}%")
    
    print(f"\n5. SDD CHIỀU CAO/TUỔI: {s['sdd_cc_tuoi']['tong']} (Trai: {s['sdd_cc_tuoi']['trai']}, Gái: {s['sdd_cc_tuoi']['gai']})")
    print(f"   - Mức -2SD: {s['sdd_cc_tuoi']['muc_2sd']}, Mức -3SD: {s['sdd_cc_tuoi']['muc_3sd']}")
    print(f"   Tỷ lệ SDD: {s['sdd_cc_tuoi']['ty_le']}%")
    
    print(f"\n6. SDD CÂN NẶNG/CHIỀU CAO: {s['sdd_cn_cc']['tong']} (Trai: {s['sdd_cn_cc']['trai']}, Gái: {s['sdd_cn_cc']['gai']})")
    print(f"   - Mức -2SD: {s['sdd_cn_cc']['muc_2sd']}, Mức -3SD: {s['sdd_cn_cc']['muc_3sd']}")
    
    print(f"\n7. THỪA CÂN, BÉO PHÌ (CN/Tuổi +2SD, +3SD):")
    tc = s['thua_can_beo_phi']
    print(f"   Tổng: {tc['tong']} (Trai: {tc['trai']}, Gái: {tc['gai']})")
    
    print(f"\n{'='*60}\n")


def summary_table(summary: dict) -> pd.DataFrame:
    """Chuyển summary dict thành DataFrame dạng bảng phẳng để in/excel - theo đúng thứ tự bảng mẫu."""
    s = summary
    total = s['tong_so_tre']['tong'] or 1  # tránh chia 0
    len_can = s['tre_duoc_can']['tong'] or 1
    len_do = s['tre_duoc_do']['tong'] or 1
    
    # Tính tỷ lệ được cân đo (trên tổng số trẻ)
    ty_le_duoc_can_do = round(len_can / total * 100, 2)

    row = {
        # 1. TỔNG SỐ TRẺ
        'Tổng số trẻ': s['tong_so_tre']['tong'],
        'TS Trai': s['tong_so_tre']['trai'],
        'TS Gái': s['tong_so_tre']['gai'],
        
        # 2. SỐ TRẺ ĐƯỢC CÂN, ĐO
        'Số trẻ được cân, đo': len_can,
        'Cân đo Trai': s['tre_duoc_can']['trai'],
        'Cân đo Gái': s['tre_duoc_can']['gai'],
        'Tỷ lệ được cân, đo (%)': ty_le_duoc_can_do,
        
        # 3. CẦN XÁC ĐỊNH TRẺ DƯỚI 5 TUỔI CÓ THỪA CÂN HOẶC BÉO PHÌ
        'Thừa cân/béo phì': s['thua_can_beo_phi']['tong'],
        'Thừa cân/béo phì Trai': s['thua_can_beo_phi']['trai'],
        'Thừa cân/béo phì Gái': s['thua_can_beo_phi']['gai'],
        'Tỷ lệ thừa cân/béo phì (%)': round(s['thua_can_beo_phi']['tong'] / len_can * 100, 2),
        
        # 3.1 Trong đó: Thừa cân (+2SD)
        'Thừa cân (+2SD)': s['thua_can_beo_phi']['muc_2sd'],
        
        # 3.2 Trong đó: Béo phì (+3SD)
        'Béo phì (+3SD)': s['thua_can_beo_phi']['muc_3sd'],
        
        # 1. SỐ TRẺ BỊ SDD CN/TUỔI (THỂ NHẸ CÂN)
        'SDD CN/tuổi (nhẹ cân)': s['sdd_cn_tuoi']['tong'],
        'SDD CN/tuổi Trai': s['sdd_cn_tuoi']['trai'],
        'SDD CN/tuổi Gái': s['sdd_cn_tuoi']['gai'],
        'Tỷ lệ SDD CN/tuổi (%)': s['sdd_cn_tuoi']['ty_le'],
        
        # 1.1 Trong đó: SDD nhẹ cân (-2SD)
        'SDD nhẹ cân (-2SD)': s['sdd_cn_tuoi']['muc_2sd'],
        
        # 1.2 Trong đó: SDD nặng cân (-3SD) 
        'SDD nặng cân (-3SD)': s['sdd_cn_tuoi']['muc_3sd'],
        
        # 4. SỐ TRẺ BỊ SDD CC/TUỔI (THỂ THẤP CÒI)
        'SDD CC/tuổi (thấp còi)': s['sdd_cc_tuoi']['tong'],
        'SDD CC/tuổi Trai': s['sdd_cc_tuoi']['trai'],
        'SDD CC/tuổi Gái': s['sdd_cc_tuoi']['gai'],
        'Tỷ lệ SDD CC/tuổi (%)': s['sdd_cc_tuoi']['ty_le'],
        
        # 4.1 Trong đó: SDD thấp còi (-2SD)
        'SDD thấp còi (-2SD)': s['sdd_cc_tuoi']['muc_2sd'],
        
        # 4.2 Trong đó: SDD thấp còi nặng (-3SD)
        'SDD thấp còi nặng (-3SD)': s['sdd_cc_tuoi']['muc_3sd'],
        
        # 5. SỐ TRẺ BỊ SDD CN/CC (THỂ GẦY CÒM)
        'SDD CN/CC (gầy còm)': s['sdd_cn_cc']['tong'],
        'SDD CN/CC Trai': s['sdd_cn_cc']['trai'],
        'SDD CN/CC Gái': s['sdd_cn_cc']['gai'],
        'Tỷ lệ SDD CN/CC (%)': s['sdd_cn_cc']['ty_le'],
        
        # 5.1 Trong đó: SDD gầy còm (-2SD)
        'SDD gầy còm (-2SD)': s['sdd_cn_cc']['muc_2sd'],
        
        # 5.2 Trong đó: SDD gầy còm nặng (-3SD)
        'SDD gầy còm nặng (-3SD)': s['sdd_cn_cc']['muc_3sd'],
    }

    return pd.DataFrame([row])


def combine_summaries(summaries: list[dict]) -> Optional[dict]:
    """Gộp nhiều summary (đã cùng cấu trúc) thành một summary cộng gộp.

    Cộng tất cả các chỉ số đếm, sau đó tính lại tỷ lệ dựa trên tổng gộp, không trung bình cộng.
    """
    if not summaries:
        return None

    def add(a: dict, b: dict, keys: list[str]):
        for k in keys:
            a[k] = a.get(k, 0) + b.get(k, 0)

    agg = {
        'tong_so_tre': {'tong': 0, 'trai': 0, 'gai': 0},
        'tre_duoc_can': {'tong': 0, 'trai': 0, 'gai': 0},
        'tre_duoc_do': {'tong': 0, 'trai': 0, 'gai': 0},
        'sdd_cn_tuoi': {'tong': 0, 'trai': 0, 'gai': 0, 'muc_2sd': 0, 'muc_3sd': 0},
        'sdd_cc_tuoi': {'tong': 0, 'trai': 0, 'gai': 0, 'muc_2sd': 0, 'muc_3sd': 0},
        'sdd_cn_cc': {'tong': 0, 'trai': 0, 'gai': 0, 'muc_2sd': 0, 'muc_3sd': 0},
        'thua_can_cn_tuoi': {'tong': 0, 'trai': 0, 'gai': 0},
        'beo_phi_cn_tuoi': {'tong': 0, 'trai': 0, 'gai': 0},
        'thua_can_beo_phi': {'tong': 0, 'trai': 0, 'gai': 0, 'muc_2sd': 0, 'muc_3sd': 0}
    }

    for s in summaries:
        add(agg['tong_so_tre'], s['tong_so_tre'], ['tong', 'trai', 'gai'])
        add(agg['tre_duoc_can'], s['tre_duoc_can'], ['tong', 'trai', 'gai'])
        add(agg['tre_duoc_do'], s['tre_duoc_do'], ['tong', 'trai', 'gai'])

        add(agg['sdd_cn_tuoi'], s['sdd_cn_tuoi'], ['tong', 'trai', 'gai', 'muc_2sd', 'muc_3sd'])
        add(agg['sdd_cc_tuoi'], s['sdd_cc_tuoi'], ['tong', 'trai', 'gai', 'muc_2sd', 'muc_3sd'])
        add(agg['sdd_cn_cc'], s['sdd_cn_cc'], ['tong', 'trai', 'gai', 'muc_2sd', 'muc_3sd'])

        add(agg['thua_can_cn_tuoi'], s['thua_can_cn_tuoi'], ['tong', 'trai', 'gai'])
        add(agg['beo_phi_cn_tuoi'], s['beo_phi_cn_tuoi'], ['tong', 'trai', 'gai'])
        add(agg['thua_can_beo_phi'], s['thua_can_beo_phi'], ['tong', 'trai', 'gai', 'muc_2sd', 'muc_3sd'])

    total = agg['tong_so_tre']['tong']
    len_can = agg['tre_duoc_can']['tong']
    len_do = agg['tre_duoc_do']['tong']
    len_cn_cc = agg['sdd_cn_cc']['tong'] + (agg['sdd_cn_cc']['muc_2sd'] + agg['sdd_cn_cc']['muc_3sd'] - agg['sdd_cn_cc']['tong']) if agg['sdd_cn_cc']['tong'] else agg['tre_duoc_can']['tong']

    agg['tre_duoc_can']['ty_le'] = round(len_can / total * 100, 2) if total > 0 else 0
    agg['tre_duoc_do']['ty_le'] = round(len_do / total * 100, 2) if total > 0 else 0

    agg['sdd_cn_tuoi']['ty_le'] = round(agg['sdd_cn_tuoi']['tong'] / len_can * 100, 2) if len_can > 0 else 0
    agg['sdd_cc_tuoi']['ty_le'] = round(agg['sdd_cc_tuoi']['tong'] / len_do * 100, 2) if len_do > 0 else 0
    # len_cn_cc lấy theo số dòng có execute_cn_cc, ở đây dùng len_can như xấp xỉ nếu thiếu
    denom_cn_cc = len_cn_cc if len_cn_cc > 0 else len_can
    agg['sdd_cn_cc']['ty_le'] = round(agg['sdd_cn_cc']['tong'] / denom_cn_cc * 100, 2) if denom_cn_cc > 0 else 0

    return agg


def export_summary_table_to_excel(summary: dict, excel_path: str, sheet_name: str = 'Summary') -> None:
    """Ghi bảng tổng hợp ra Excel (tạo file mới hoặc ghi đè)."""
    df_table = summary_table(summary)
    df_table.to_excel(excel_path, sheet_name=sheet_name, index=False)


def write_column_to_excel(
    df: pd.DataFrame,
    df_column: str,
    excel_path: str,
    excel_column: str,
    start_row: int = 7,
    sheet_name: str = 'Sheet1'
) -> None:
    """
    Ghi đè 1 cột từ DataFrame vào 1 cột trong file Excel.
    Chỉ ghi đè phần nội dung, giữ nguyên header và format.
    
    Args:
        df: DataFrame chứa dữ liệu cần ghi
        df_column: Tên cột trong DataFrame cần lấy dữ liệu
        excel_path: Đường dẫn file Excel
        excel_column: Tên cột trong Excel (A, B, C, ... hoặc AA, AB, ...)
        start_row: Hàng bắt đầu ghi (mặc định là 7)
        sheet_name: Tên sheet (mặc định là 'Sheet1')
    """
    # Load workbook - giữ nguyên format
    wb = load_workbook(excel_path)
    ws = wb[sheet_name]
    
    if ws is None:
        wb.close()
        raise ValueError("Not found sheet in Excel file.")
    
    # Lấy dữ liệu từ DataFrame
    data = df[df_column].tolist()
    
    # Ghi từng giá trị vào Excel - chỉ thay đổi value, giữ nguyên format của cell
    for i, value in enumerate(data):
        row = start_row + i
        cell = ws[f"{excel_column}{row}"]
        
        # Xử lý giá trị NaN/None - chỉ ghi value, không thay đổi style
        if pd.isna(value):
            cell.value = None
        else:
            cell.value = value
    
    # Lưu file
    wb.save(excel_path)
    wb.close()
    
    print(f"Đã ghi {len(data)} giá trị từ cột '{df_column}' vào cột {excel_column} (từ hàng {start_row})")