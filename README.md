# App Chấm Kênh Dinh Dưỡng

Ứng dụng desktop (GUI tkinter) hỗ trợ chấm tình trạng dinh dưỡng trẻ em theo chuẩn WHO. Nhập file Excel danh sách cân/đo trẻ → tự động tính tháng tuổi, chấm suy dinh dưỡng (SDD) theo các chỉ số CN/tuổi, CC/tuổi, CN/CC, điều chỉnh số liệu và xuất thống kê. Hỗ trợ xử lý từng file hoặc cả thư mục (đệ quy).

---

## Tính năng chính

- **Tính tháng tuổi**: tính số tháng tuổi của trẻ tại một ngày target (dd/mm/yyyy) dựa trên ngày sinh.
- **Chấm SDD theo chuẩn WHO** (phân loại `-3 SD`, `-2 SD`, `BT`, `+2 SD`, `+3 SD`):
  - **CN/tuổi** (cân nặng theo tuổi) — thể nhẹ cân.
  - **CC/tuổi** (chiều cao theo tuổi) — thể thấp còi.
  - **CN/CC** (cân nặng theo chiều cao) — thể gầy còm. Dùng bảng chuẩn riêng cho trẻ 0–24 tháng và 25–60 tháng, khớp theo chiều cao bằng `merge_asof`.
- **Làm tròn chiều cao về bước 0.5 cm** để đồng bộ dữ liệu.
- **Điền số đo vào ô còn trống** (luôn **giữ nguyên số đo đã có**, **bỏ qua trẻ "vắng"** — xem dưới):
  - *Điền CC/CN theo kết quả*: với ô trống **đã có sẵn trạng thái** (CC/tuổi hoặc CN/tuổi), điền số nằm trong band của trạng thái đó (re-chấm ra đúng trạng thái cũ).
  - *Điền CC/CN bình thường*: với ô trống **không có trạng thái**, điền số sát đầu thấp khoảng bình thường (BT).
  - *Điền CN/CC*: điền chiều cao và/hoặc cân nặng để đạt **chỉ tiêu CN/CC**, đồng thời cố gắng giữ chỉ tiêu CN/tuổi & CC/tuổi (uốn trong band, ưu tiên CN/CC). Tự xử lý 3 trường hợp: thiếu CN, thiếu CC, hoặc thiếu cả hai.
  - **Trẻ "vắng"**: dòng có chữ `vắng` (mọi biến thể dấu/hoa thường, ở bất kỳ cột nào) bị coi là vắng mặt → **không điền** số nào.
  - Chỉ tiêu/trạng thái để trống → mặc định coi là **BT**.
- **Thống kê dinh dưỡng** cho trẻ dưới 2 tuổi (≤24 tháng) và dưới 5 tuổi (≤60 tháng): tổng số trẻ, số trẻ được cân/đo, số ca SDD từng thể (chi tiết -2SD/-3SD), thừa cân/béo phì, kèm tỷ lệ %.
- **Xuất kết quả Excel**: ghi đè vào file gốc hoặc xuất ra file mới `*_ketqua.xlsx` (giữ nguyên định dạng template). Có thể kèm sheet bảng tổng hợp `<5T` (Tổng/Nam/Nữ).
- **Xử lý cả thư mục**: quét đệ quy mọi file `.xlsx` (bỏ qua file `_ketqua`), xử lý hàng loạt và tổng hợp thống kê đa tầng (theo từng file và theo từng folder) ra file `Tổng hợp thống kê.xlsx`.

---

## Tech stack / Yêu cầu

- **Python**: 3.10+ (code dùng cú pháp `str | None`, `list[dict]`).
- **Thư viện**:
  - `pandas` — xử lý dữ liệu, merge bảng chuẩn.
  - `numpy` — phân loại theo điều kiện (`np.select`), làm tròn, random điều chỉnh.
  - `openpyxl` (3.1.5) — đọc/ghi Excel, giữ nguyên định dạng cell.
  - `tkinter` — giao diện GUI (đi kèm Python chuẩn, không cần cài qua pip).
  - `pyinstaller` — đóng gói file `.exe`.

---

## Cài đặt

```bash
# (khuyến nghị) tạo virtual environment
python -m venv venv
venv\Scripts\activate        # Windows

# cài dependencies từ requirements.txt
pip install -r requirements.txt

# hoặc cài thủ công
pip install pandas numpy openpyxl pyinstaller
```

> Lưu ý: `requirements.txt` không liệt kê `numpy` riêng (nó được kéo theo `pandas`), nhưng code import trực tiếp `numpy`. Bản cài `pandas` đã bao gồm sẵn.

---

## Cách chạy app

```bash
python main.py
```

Cửa sổ GUI "ỨNG DỤNG CHẤM KÊNH DINH DƯỠNG" sẽ hiện ra.

> Quan trọng: app đọc các file bảng chuẩn WHO ở **thư mục làm việc hiện tại**: `CN_tuoi.xlsx`, `CC_tuoi.xlsx`, `CN_CC_0-2_tuoi.xlsx`, `CN_CC_2-5_tuoi.xlsx`. Phải chạy app tại thư mục chứa các file này (hoặc đặt chúng cạnh file thực thi).

---

## Đóng gói file .exe

Lệnh build (lấy từ `note.txt`):

```bash
pyinstaller --onefile --noconsole --icon=icon.ico --name=AppChamKenh main.py
```

File `.exe` sinh ra trong thư mục `dist/`. Nhớ đặt 4 file Excel bảng chuẩn WHO (`CN_tuoi.xlsx`, `CC_tuoi.xlsx`, `CN_CC_0-2_tuoi.xlsx`, `CN_CC_2-5_tuoi.xlsx`) cùng thư mục với file `.exe` khi chạy.

---

## Cấu trúc thư mục

```
app_cham_kenh/
├── main.py                  # Entry point: class mainApp kế thừa AppUI, gắn logic vào nút bấm
├── ui_helper.py             # Lớp AppUI - dựng giao diện tkinter, xử lý chọn file/folder, chạy tác vụ trong thread
├── process_helper.py        # Logic xử lý: tính tháng tuổi, chấm SDD, adj, thống kê, ghi cột vào Excel
├── excel_file_helper.py     # Đọc/ghi Excel: load bảng chuẩn WHO, load danh sách cân đo, export kết quả
├── requirements.txt         # Dependencies
├── note.txt                 # Ghi chú + lệnh build pyinstaller
├── icon.ico                 # Icon cho file .exe
│
├── CN_tuoi.xlsx             # Bảng chuẩn WHO: cân nặng theo tuổi
├── CC_tuoi.xlsx             # Bảng chuẩn WHO: chiều cao theo tuổi
├── CN_CC_0-2_tuoi.xlsx      # Bảng chuẩn WHO: cân nặng theo chiều cao (0-24 tháng)
└── CN_CC_2-5_tuoi.xlsx      # Bảng chuẩn WHO: cân nặng theo chiều cao (25-60 tháng)
```

### Vai trò file chính

| File | Vai trò |
|------|---------|
| `main.py` | Lớp `mainApp(AppUI)` triển khai các hành động (tính tháng tuổi, chấm SDD, adj, export). `main()` khởi tạo và chạy app. |
| `ui_helper.py` | Lớp cơ sở `AppUI` dựng toàn bộ giao diện (menu, checkbox chức năng, ô log, popup thống kê), điều phối luồng xử lý single-file và folder. |
| `process_helper.py` | Hàm thuần xử lý dữ liệu pandas/numpy: `apply_month_age`, `execute_weight_by_age`, `execute_height_by_age`, `execute_weight_by_height`, `fill_height_by_status`, `fill_weight_by_status`, `fill_missing_height_normal`, `fill_missing_weight_normal`, `fill_cn_cc`, `summary_statistics`, `combine_summaries`, `write_column_to_excel`... |
| `excel_file_helper.py` | I/O Excel: `load_cn_per_old`, `load_cc_per_old`, `load_cn_cc`, `load_danh_sach_can_do`, `export_to_excel`, `build_under5_statistics_table`. |

---

## Luồng xử lý

1. **Chọn nguồn dữ liệu**: bấm "Chọn File" (1 file Excel) hoặc "Chọn Folder" (xử lý hàng loạt). Khi chọn file, app load ngay danh sách cân đo vào DataFrame.
2. **Nhập ngày target** (dd/mm/yyyy) để tính tháng tuổi (mặc định là ngày hiện tại).
3. **Tick các chức năng** muốn chạy: tính tháng tuổi / làm tròn CC / điền CC-CN theo kết quả / điền CC-CN bình thường / điền CN/CC / chấm SDD / chấm CN/CC / xuất file / thống kê... (các chức năng "điền" chạy **trước** khi chấm SDD).
4. **Chọn cách ghi kết quả**:
   - *Ghi đè vào file gốc*: mỗi chức năng ghi thẳng cột tương ứng vào file Excel gốc (giữ nguyên format).
   - *Xuất ra file mới*: tạo bản sao `<tên>_ketqua.xlsx` rồi ghi kết quả vào đó.
5. **Bấm "THỰC HIỆN"**: tác vụ chạy trong thread riêng (không treo UI). Tiến trình hiện ở ô log; thống kê hiện trong popup.
6. **Folder mode**: quét đệ quy mọi `.xlsx`, xử lý từng file, rồi xuất `Tổng hợp thống kê.xlsx` (sheet "Dưới 2 tuổi", "Dưới 5 tuổi", và "Thong ke <5T" nếu bật).

---

## Định dạng file Excel đầu vào

App đọc sheet **`Sheet1`**, **bỏ qua 6 dòng đầu** (`skiprows=6`, không dùng header) — tức **dữ liệu bắt đầu từ dòng 7**. Phần header báo cáo nằm ở các dòng 1–6.

Thứ tự cột (12 cột đầu, A→L):

| Cột Excel | Trường | Ý nghĩa |
|-----------|--------|---------|
| A | `stt` | Số thứ tự |
| B | `ho_ten` | Họ tên trẻ |
| C | `nam` | Đánh dấu `x` nếu là nam |
| D | `nu` | Đánh dấu nếu là nữ |
| E | `ngay_sinh` | Ngày sinh |
| F | `thang_tuoi` | Tháng tuổi (app tính, ghi vào cột F) |
| G | `ho_ten_me` | Họ tên mẹ |
| H | `can_nang` | Cân nặng (kg) |
| I | `execute_cn_tuoi` | Kết quả chấm CN/tuổi |
| J | `chieu_cao` | Chiều cao (cm) |
| K | `execute_cc_tuoi` | Kết quả chấm CC/tuổi |
| L | `execute_cn_cc` | Kết quả chấm CN/CC |
| M | `note` | (app ghi thêm) giá trị nhập sai/không hợp lệ |

Quy ước dữ liệu:
- **Giới tính** xác định qua cột C (`nam == 'x'` → trai, ngược lại → gái).
- Khi ghi đè, app ghi từng cột riêng bắt đầu từ **dòng 7** (`start_row=7`); khi export ra file mới, ghi toàn bộ data từ dòng 7 (`data_start_row=7`).
- Các dòng có `stt` chứa chữ (không phải số, không trống) bị loại bỏ.
- Giá trị `can_nang` / `chieu_cao` / `thang_tuoi` nhập sai (không phải số) được chuyển thành `NaN` và lưu nguyên bản vào cột **note** (M).

### Định dạng bảng chuẩn WHO

Các file `CN_tuoi.xlsx`, `CC_tuoi.xlsx`, `CN_CC_*.xlsx` được đọc với `skiprows=1`. Cấu trúc: cột 0–5 là dữ liệu **trai**, cột 7–12 là dữ liệu **gái**, mỗi nhóm gồm `[mốc (tháng tuổi hoặc chiều cao), -3SD, -2SD, trung vị, +2SD, +3SD]`.

---

## Ghi chú / Lưu ý

- App **bắt buộc** phải có 4 file bảng chuẩn WHO trong thư mục làm việc; thiếu sẽ lỗi khi khởi động (`load_data` chạy ngay trong `__init__`).
- Folder mode tự bỏ qua các file kết thúc bằng `_ketqua` để tránh xử lý lại file đã xuất.
- Các chức năng **điền** chỉ tác động **ô còn trống** (không bao giờ sửa số đo đã có), dùng số ngẫu nhiên (`np.random.uniform`) → khác nhau mỗi lần chạy nhưng luôn rơi đúng band trạng thái/chỉ tiêu. Trẻ "vắng" bị bỏ qua hoàn toàn.
- Logic chấm CN/CC tách theo độ tuổi (≤24 tháng dùng bảng 0-2, >24 tháng dùng bảng 2-5) và khớp chiều cao gần nhất theo hướng `backward`.
- Thống kê gộp (theo thôn/folder) tính lại tỷ lệ từ tổng số liệu cộng dồn; mẫu số CN/CC = số trẻ **được đánh giá** CN/CC (`danh_gia`).
- Ghi chú nghiệp vụ từ `note.txt`: trẻ vắng dưới 6 tháng (lập nghĩa 1, lập vinh 1); trẻ vắng trên 6 tháng (6 trẻ, tuỳ thôn).
