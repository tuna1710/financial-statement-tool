# Financial Statement Tool (Thông tư 200/2014/TT-BTC)

Công cụ Python tự động hóa phân tích Bảng cân đối tài khoản (Trial Balance) và lập Bộ Báo cáo Tài chính hoàn chỉnh theo chuẩn mực Kế toán Doanh nghiệp Việt Nam (Thông tư 200/2014/TT-BTC).

## 🚀 Tính năng chính

1. **Bảng Cân đối kế toán (Mẫu B01-DN):**
   - Tự động phân loại tài sản ngắn hạn/dài hạn, nợ phải trả và vốn chủ sở hữu.
   - Tự động kết chuyển và đối soát lợi nhuận trong kỳ đảm bảo cân đối 100% (Tổng Tài sản = Tổng Nguồn vốn).
   - So sánh biến động giữa các quý (Đầu năm, Quý 1, Quý 2) và tính tỷ lệ tăng trưởng (%).

2. **Báo cáo Kết quả hoạt động kinh doanh (Mẫu B02-DN):**
   - Đầy đủ các chỉ tiêu: Doanh thu thuần, Giá vốn, Lợi nhuận gộp, Doanh thu tài chính, Chi phí tài chính/quản lý, Lợi nhuận trước thuế và sau thuế.
   - Phân tích chi tiết số liệu từng quý và lũy kế 6 tháng.

3. **Báo cáo Lưu chuyển tiền tệ (Mẫu B03-DN - Phương pháp gián tiếp):**
   - Tính toán lưu chuyển tiền thuần từ hoạt động kinh doanh, đầu tư và tài chính.
   - Đối chiếu số dư tiền cuối kỳ khớp 100% với Bảng cân đối kế toán.

4. **Dashboard & Phân tích KPI tài chính:**
   - Biên lợi nhuận gộp (Gross Margin), Biên lợi nhuận ròng (Net Margin).
   - Tỷ số thanh toán hiện hành (Current Ratio), Tỷ số thanh toán tức thời (Cash Ratio), Tỷ số nợ/Tổng tài sản.
   - Nhận xét và đánh giá chuyên môn tài chính quản trị.

5. **Bảng Cân đối tài khoản tổng hợp:**
   - Tổng hợp số dư, phát sinh các tài khoản cấp 1 và cấp 2 để đối chiếu kiểm toán.

---

## 🛠 Cài đặt & Sử dụng

### 1. Cài đặt môi trường
Yêu cầu Python 3.9+
```bash
pip install -r requirements.txt
```

### 2. Chạy lệnh lập báo cáo
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
├── data/                                 # Dữ liệu Bảng cân đối tài khoản đầu vào
│   ├── Bang_can_doi_tai_khoanQ12026.xlsx
│   └── Bang_can_doi_tai_khoanQ22026.xlsx
├── output/                               # File báo cáo kết xuất
│   └── Bo_Bao_Cao_Tai_Chinh_Q1_Q2_2026.xlsx
├── lap_bao_cao_tai_chinh.py             # Script xử lý và tự động hóa chính
├── requirements.txt                      # Thư viện phụ thuộc
├── .gitignore
└── README.md
```
