# Hướng dẫn sử dụng VBA Macro (Dành cho Kế toán dùng trực tiếp trong Excel)

Nếu doanh nghiệp của bạn ưu tiên sử dụng **Microsoft Excel thuần túy** mà không muốn cài đặt Python hay mở trình duyệt Google Colab, bạn có thể sử dụng trực tiếp module VBA `BaoCaoTaiChinh_TT200.bas`.

---

## 🚀 Các bước cài đặt & sử dụng trong 1 phút

### Bước 1: Mở cửa sổ soạn thảo VBA
1. Mở bất kỳ file Excel nào (hoặc file Bảng cân đối tài khoản của bạn).
2. Nhấn tổ hợp phím **`Alt + F11`** (trên máy Mac là `Fn + Option + F11`) để mở cửa sổ Microsoft Visual Basic for Applications.

### Bước 2: Import mã nguồn VBA
1. Trên thanh menu, chọn **File** ➔ **Import File...** (hoặc bấm `Ctrl + M`).
2. Điều hướng đến file **`BaoCaoTaiChinh_TT200.bas`** và nhấn **Open**.
3. Bạn sẽ thấy module `BaoCaoTaiChinh_TT200` xuất hiện trong mục **Modules** ở cây thư mục bên trái.

### Bước 3: Tạo nút bấm tiện lợi (1-Click Run)
1. Quay lại màn hình Excel chính (`Alt + F11`).
2. Vào thẻ **Insert** ➔ **Illustrations** ➔ **Shapes** ➔ Chọn vẽ một hình chữ nhật bo tròn.
3. Ghi chữ lên hình: **"⚡ LẬP BÁO CÁO TÀI CHÍNH TT200"**, chọn màu xanh Navy chuyên nghiệp.
4. Chuột phải vào hình chữ nhật vừa vẽ ➔ Chọn **Assign Macro...**.
5. Chọn macro **`Chay_Lap_Bao_Cao_Tai_Chinh`** ➔ Bấm **OK**.

### Bước 4: Chạy báo cáo
1. Bấm vào nút vừa tạo.
2. Excel sẽ hỏi:
   - Chọn **YES** nếu bạn muốn mở một file Excel Bảng cân đối tài khoản khác từ máy tính.
   - Chọn **NO** nếu bạn muốn lấy dữ liệu ngay trên Sheet đang mở.
3. Trong vòng 1 giây, Macro sẽ:
   - Tự lọc bỏ các dòng lặp tiêu đề trang in của phần mềm kế toán.
   - Lập Bảng Cân đối kế toán (Mẫu B01-DN) chuẩn TT200 với chênh lệch **0 VNĐ**.
   - Lập Báo cáo Kết quả kinh doanh (Mẫu B02-DN).
   - Kẻ khung viền, định dạng số tiền `#,##0` và tô màu giao diện chuẩn chỉnh.

---

## 💡 Lưu ý khi lưu file
Để lưu lại macro trong file Excel cho các lần dùng sau, khi bấm **Save As**, bạn nhớ chọn định dạng **Excel Macro-Enabled Workbook (*.xlsm)**.
