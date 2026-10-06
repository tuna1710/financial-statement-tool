# Financial Statement Tool (Thông tư 200/2014/TT-BTC)

[![Open In Colab](https://colab.research.google.com/assets/colab-badge.svg)](https://colab.research.google.com/github/tuna1710/financial-statement-tool/blob/main/Bao_Cao_Tai_Chinh_Colab.ipynb)
[![Excel XLSM](https://img.shields.io/badge/Excel-Macro%20.XLSM-green.svg)](#-2-file-excel-macro-xlsm-1-click-chọn-thư-mục-chạy-offline)
[![Python 3.9+](https://img.shields.io/badge/python-3.9+-blue.svg)](https://www.python.org/downloads/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

Công cụ tự động hóa đọc Bảng cân đối tài khoản (Trial Balance) và lập **Bộ Báo cáo Tài chính hoàn chỉnh** theo chuẩn mực Kế toán Doanh nghiệp Việt Nam (**Thông tư 200/2014/TT-BTC**).

Dự án cung cấp **3 hình thức triển khai hoàn hảo** đáp ứng mọi đối tượng người dùng:
1. 📑 **File Excel Macro (`Bo_Bao_Cao_Tai_Chinh_TT200.xlsm`):** Dành riêng cho Kế toán dùng trực tiếp trong Excel. Mở file, bấm 1 nút chọn thư mục chứa dữ liệu là báo cáo tự động hình thành!
2. ☁️ **Google Colab Cloud (1-Click Run):** Không cần cài đặt phần mềm trên máy tính, trực quan hóa biểu đồ tài chính và tải báo cáo ngay trên trình duyệt web.
3. 💻 **Python CLI (Đa kỳ nâng cao):** Tự động hóa hàng loạt không giới hạn quý (Q1 -> Q4 cả năm), tốc độ tính toán tức thì, cam kết cân đối tuyệt đối 100% (Chênh lệch = 0 VNĐ).

---

## 📑 1. File Excel Macro (`.xlsm`) - 1-Click Chọn thư mục & Chạy Offline

Dành cho người dùng muốn trải nghiệm đơn giản và trực quan nhất ngay trong Microsoft Excel:

- **File có sẵn trong repo:** [`Bo_Bao_Cao_Tai_Chinh_TT200.xlsm`](Bo_Bao_Cao_Tai_Chinh_TT200.xlsm)
- **Cách sử dụng:**
  1. Tải file `Bo_Bao_Cao_Tai_Chinh_TT200.xlsm` về máy tính và mở bằng Microsoft Excel (chọn *Enable Content / Enable Macros* nếu được hỏi).
  2. Tại trang **`TRANG_CHU`**, bấm vào nút to màu xanh Navy:  
     👉 **"📁 BẤM VÀO ĐÂY ĐỂ CHỌN THƯ MỤC INPUT & LẬP BCTC"**
  3. Chọn thư mục chứa các file Excel Bảng cân đối tài khoản các quý (ví dụ thư mục `data/` chứa Q1, Q2...).
  4. Hệ thống VBA Macro sẽ tự động:
     - Quét toàn bộ file trong thư mục, tự nhận diện Quý và Năm.
     - Lọc bỏ sạch các dòng tiêu đề trang in lặp lại.
     - Cập nhật số liệu chuẩn xác vào các Sheet: **`B01-DN (CDKT)`**, **`B02-DN (KQKD)`**, **`KPI_Dashboard`**.
     - Đảm bảo **Tổng Tài sản = Tổng Nguồn vốn (Chênh lệch 0 VNĐ)**.

---

## ⚡ 2. Chạy trực tiếp trên Google Colab (Không cần cài đặt)

Chỉ với 1 cú click, bạn có thể chạy toàn bộ công cụ, trực quan hóa biểu đồ và xuất báo cáo ngay trên trình duyệt web:

[![Open In Colab](https://colab.research.google.com/assets/colab-badge.svg)](https://colab.research.google.com/github/tuna1710/financial-statement-tool/blob/main/Bao_Cao_Tai_Chinh_Colab.ipynb)

**Quy trình 5 bước trên Colab:**
1. **Bước 1:** Khởi tạo môi trường & clone mã nguồn tự động.
2. **Bước 2:** Chọn phương thức nạp dữ liệu: dùng file mẫu có sẵn hoặc tải file `.xlsx` trực tiếp từ máy tính của bạn.
3. **Bước 3:** Chạy tự động lập Báo cáo Tài chính chuẩn TT200 (B01-DN, B02-DN, B03-DN, KPI Dashboard).
4. **Bước 4:** Khảo sát Dashboard biểu đồ trực quan (Doanh thu & Lợi nhuận, Biên lợi nhuận gộp/ròng, Quy mô Tài sản - Vốn chủ sở hữu, Khả năng thanh toán hiện hành & tức thời).
5. **Bước 5:** Tải file Excel Báo cáo hoàn chỉnh (`Bo_Bao_Cao_Tai_Chinh.xlsx`) về máy tính.

---

## 🚀 3. Tính năng nổi bật của Bộ công cụ

1. **Hỗ trợ xử lý đa kỳ linh hoạt (Multi-period support):**
   - Tự động nhận diện kỳ kế toán từ nội dung file Excel hoặc tên file (ví dụ: `Quý 1 năm 2025`, `Quý 2 năm 2025`, `Quý 3 năm 2025`, `Quý 4 năm 2025`...).
   - Hỗ trợ bất kỳ số lượng kỳ: 1 quý, 2 quý, 3 quý, 4 quý (cả năm) hoặc chuỗi nhiều năm.
   - Tự động sắp xếp các kỳ theo trình tự thời gian và tự động mở rộng các cột báo cáo tương ứng.

2. **Bảng Cân đối kế toán (Mẫu B01-DN):**
   - Phân loại chuẩn xác Tài sản ngắn hạn/dài hạn, Nợ phải trả và Vốn chủ sở hữu theo Thông tư 200.
   - Tự động xác định và phân bổ lợi nhuận chưa phân phối năm nay vào Vốn chủ sở hữu.
   - **Cam kết cân đối tuyệt đối 100% (Tổng Tài sản = Tổng Nguồn vốn)** tại mọi mốc thời gian báo cáo.
   - So sánh biến động số dư qua các quý và tỷ lệ tăng trưởng lũy kế cả năm.

3. **Báo cáo Kết quả hoạt động kinh doanh (Mẫu B02-DN):**
   - Doanh thu thuần, Giá vốn, Lợi nhuận gộp, Doanh thu tài chính, Chi phí tài chính/quản lý, Lợi nhuận trước thuế và sau thuế.
   - Chi tiết từng quý (Q1, Q2, Q3, Q4) và cột tổng cộng Lũy kế cả năm (Full Year).
   - Tỷ lệ tăng trưởng kết quả kinh doanh kỳ cuối so với kỳ trước.

4. **Báo cáo Lưu chuyển tiền tệ (Mẫu B03-DN - Phương pháp gián tiếp):**
   - Tự động bóc tách lưu chuyển tiền từ Hoạt động kinh doanh (CFO), Hoạt động đầu tư (CFI) và Hoạt động tài chính (CFF).
   - Đối chiếu số dư Tiền và tương đương tiền cuối kỳ khớp 100% với Bảng cân đối kế toán.

5. **Dashboard & Phân tích KPI tài chính quản trị:**
   - Biên lợi nhuận gộp (Gross Margin), Biên lợi nhuận ròng (Net Margin).
   - Tỷ số thanh toán hiện hành (Current Ratio), Tỷ số thanh toán tiền mặt tức thời (Cash Ratio), Tỷ số nợ/Tổng tài sản (D/A).
   - Tóm tắt và đánh giá xu hướng tài chính doanh nghiệp.

6. **Bảng Cân đối tài khoản tổng hợp:**
   - Đối chiếu toàn bộ số dư đầu kỳ, phát sinh Nợ/Có và số dư cuối kỳ của các tài khoản cấp 1 và cấp 2 qua tất cả các quý.

---

## 🛠 4. Hướng dẫn chạy trên máy tính cá nhân (Python CLI)

### Cài đặt thư viện
Yêu cầu Python 3.9+
```bash
pip install -r requirements.txt
```

### Các lệnh chạy công cụ

#### Cách 1: Chỉ định thư mục chứa các file Bảng cân đối tài khoản (Khuyên dùng)
Chỉ cần thả toàn bộ file Excel các quý (ví dụ: Q1, Q2, Q3, Q4 năm 2025 hoặc 2026) vào thư mục `data/`:
```bash
python3 lap_bao_cao_tai_chinh.py --dir data --output output/Bo_Bao_Cao_Tai_Chinh.xlsx
```

#### Cách 2: Truyền danh sách file cụ thể
```bash
python3 lap_bao_cao_tai_chinh.py \
    --files data/Bang_can_doi_tai_khoanQ12025.xlsx \
            data/Bang_can_doi_tai_khoanQ22025.xlsx \
            data/Bang_can_doi_tai_khoanQ32025.xlsx \
            data/Bang_can_doi_tai_khoanQ42025.xlsx \
    --output output/Bo_Bao_Cao_Tai_Chinh_2025.xlsx
```

#### Cách 3: Lệnh 2 quý (Tương thích ngược)
```bash
python3 lap_bao_cao_tai_chinh.py \
    --q1 data/Bang_can_doi_tai_khoanQ12026.xlsx \
    --q2 data/Bang_can_doi_tai_khoanQ22026.xlsx \
    --output output/Bo_Bao_Cao_Tai_Chinh_Q1_Q2_2026.xlsx
```

---

## 📁 Cấu trúc dự án

```
financial-statement-tool/
├── Bo_Bao_Cao_Tai_Chinh_TT200.xlsm      # File Excel Macro Native (Bấm nút chọn thư mục là chạy)
├── Bao_Cao_Tai_Chinh_Colab.ipynb         # Google Colab Notebook tương tác 1-click & biểu đồ
├── vba/                                  # Thư mục mã nguồn VBA
│   ├── BaoCaoTaiChinh_TT200.bas          # Module VBA mã nguồn chính
│   └── HD_SU_DUNG_VBA.md                 # Hướng dẫn chi tiết cách Import và sử dụng
├── data/                                 # Chứa các file Excel Bảng cân đối tài khoản đầu vào
│   ├── Bang_can_doi_tai_khoanQ12026.xlsx
│   └── Bang_can_doi_tai_khoanQ22026.xlsx
├── output/                               # Chứa file Excel báo cáo hoàn chỉnh xuất ra
│   └── Bo_Bao_Cao_Tai_Chinh_Q1_Q2_2026.xlsx
├── lap_bao_cao_tai_chinh.py             # Script xử lý đa kỳ và lập báo cáo (Python)
├── requirements.txt                      # Danh mục thư viện phụ thuộc
├── .gitignore
└── README.md
```
