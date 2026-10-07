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

## 🏢 5. Luồng Báo cáo Công ty con gửi Công ty mẹ (Phục vụ Hợp nhất BCTC)

Luồng độc lập hoàn toàn chuyên phục vụ quy đổi Bảng cân đối tài khoản doanh nghiệp (Thông tư 200) sang hệ thống tài khoản và mẫu biểu của **Tổ chức Tín dụng (TCTD - NHNN)** để công ty con gửi lên Ngân hàng mẹ phục vụ lập Báo cáo Tài chính hợp nhất:

### Tính năng:
- **Tự động ánh xạ (Mapping):** Chuyển đổi chính xác 100% từ các tài khoản Thông tư 200 sang 84 tài khoản cấp 5 theo hệ thống kế toán TCTD (Sheet `so lieu`).
- **Xử lý bút toán cấn trừ (Intercompany Elimination):** Hỗ trợ cấn trừ tài khoản trung gian ngoại tệ (TK 33886.02 / 13885.02), chuyển phải thu về tiền gửi (TK 112 / 13881), cấn trừ đối tác Western Union (WU), WASH...
- **Bộ Báo cáo chuẩn gửi Công ty Mẹ:**
  - `B02` (Mẫu B02a/TCTD): Báo cáo tình hình tài chính giữa niên độ.
  - `B03` (Mẫu B03a/TCTD): Báo cáo kết quả hoạt động kinh doanh.
  - `B04` (Mẫu B04a/TCTD): Báo cáo lưu chuyển tiền tệ (Phương pháp trực tiếp).
  - `B05` (Mẫu B05a/TCTD): Thuyết minh Báo cáo Tài chính giữa niên độ.
  - `Thue`: Bảng tình hình thực hiện nghĩa vụ đối với Nhà nước.
- **Cam kết cân đối:** Đảm bảo Tổng Dư Nợ = Tổng Dư Có, Tổng Phát sinh Nợ = Tổng Phát sinh Có (Chênh lệch = 0 VNĐ).

### Lệnh thực thi:
```bash
# Chạy cho Quý 1:
python3 lap_bao_cao_cong_ty_con.py -i data/Bang_can_doi_tai_khoanQ12026.xlsx -o output/Bao_Cao_Hop_Nhat_Q1_2026.xlsx

# Chạy cho Quý 2:
python3 lap_bao_cao_cong_ty_con.py -i data/Bang_can_doi_tai_khoanQ22026.xlsx -o output/Bao_Cao_Hop_Nhat_Q2_2026.xlsx

# Tùy biến số tiền bút toán cấn trừ thủ công (nếu cần):
python3 lap_bao_cao_cong_ty_con.py \
    -i data/Bang_can_doi_tai_khoanQ22026.xlsx \
    --can-tru-d7 10460694846 \
    --can-tru-d25 18358019545 \
    -o output/Bao_Cao_Hop_Nhat_Q2_2026.xlsx
```

---

## 📁 Cấu trúc dự án

```
financial-statement-tool/
├── Bo_Bao_Cao_Tai_Chinh_TT200.xlsm      # [Luồng 1] File Excel Macro Native TT200
├── Bao_Cao_Tai_Chinh_Colab.ipynb         # [Luồng 1] Google Colab Notebook TT200
├── lap_bao_cao_tai_chinh.py             # [Luồng 1] Script Python BCTC TT200
├── lap_bao_cao_cong_ty_con.py           # [Luồng 2] Script Python BCTC gửi Công ty mẹ (Hợp nhất TCTD)
├── templates/
│   └── Mau_BCTC_Hop_Nhat_TCTD.xlsx      # [Luồng 2] Template chuẩn BCTC B02-B05 gửi Công ty mẹ
├── data/                                 # Chứa file Bảng cân đối tài khoản đầu vào (Q1, Q2...)
│   ├── Bang_can_doi_tai_khoanQ12026.xlsx
│   └── Bang_can_doi_tai_khoanQ22026.xlsx
├── output/                               # Thư mục xuất báo cáo hoàn chỉnh
│   ├── Bo_Bao_Cao_Tai_Chinh_Q1_Q2_2026.xlsx
│   ├── Bao_Cao_Hop_Nhat_Q1_2026.xlsx
│   └── Bao_Cao_Hop_Nhat_Q2_2026.xlsx
├── vba/                                  # Thư mục mã nguồn VBA
│   ├── BaoCaoTaiChinh_TT200.bas          # [Luồng 1] Module VBA BCTC TT200
│   ├── HD_SU_DUNG_VBA.md                 # [Luồng 1] Hướng dẫn sử dụng VBA TT200
│   ├── BaoCaoHopNhat_TCTD.bas            # [Luồng 2] Module VBA BCTC Hợp nhất TCTD (Native Excel)
│   └── HD_SU_DUNG_VBA_HOP_NHAT.md        # [Luồng 2] Hướng dẫn sử dụng VBA Hợp nhất TCTD
├── requirements.txt
├── .gitignore
└── README.md
```
