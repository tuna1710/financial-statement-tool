# Hướng dẫn sử dụng VBA Macro Luồng 2 (100% Native Excel)
## Lập Báo Cáo Tài Chính Hợp Nhất Chuẩn TCTD

Module VBA **`vba/BaoCaoHopNhat_TCTD.bas`** được thiết kế để kế toán sử dụng **100% trực tiếp trong Microsoft Excel** mà **không cần cài đặt Python hay bất kỳ phần mềm nào khác**.

---

### 🚀 Cài đặt 1 lần duy nhất (trong 1 phút)

#### Bước 1: Mở file mẫu báo cáo
* Mở file **`templates/Mau_BCTC_Hop_Nhat_TCTD.xlsx`** bằng Microsoft Excel.

#### Bước 2: Import module VBA
1. Nhấn tổ hợp phím **`Alt + F11`** (trên máy Mac là `Option + F11` hoặc `Fn + Option + F11`) để mở cửa sổ soạn thảo VBA.
2. Trên thanh menu, chọn **File** ➔ **Import File...** (hoặc bấm `Ctrl + M`).
3. Điều hướng và chọn file **`vba/BaoCaoHopNhat_TCTD.bas`** ➔ bấm **Open**.
4. Bạn sẽ thấy module **`BaoCaoHopNhat_TCTD`** xuất hiện trong thư mục **Modules** ở cây thư mục bên trái.

#### Bước 3: Gán Macro vào nút bấm trên trang chủ
1. Nhấn **`Alt + F11`** để quay lại màn hình Excel chính tại sheet **`TRANG_CHU`**.
2. **Gán nút Bước 1:**
   - Vào thẻ **Insert** ➔ **Illustrations** ➔ **Shapes** ➔ Vẽ một hình chữ nhật bo tròn đè lên vùng ô C6:D6.
   - Ghi chữ: **"📁 BẤM ĐỂ CHỌN FILE CĐTK"** (chọn màu xanh Navy).
   - Chuột phải vào nút vừa vẽ ➔ Chọn **Assign Macro...** ➔ Chọn macro **`ChonFileBangCanDoi`** ➔ Bấm **OK**.
3. **Gán nút Bước 2:**
   - Vẽ tiếp một nút bo tròn đè lên ô C18:F18.
   - Ghi chữ: **"⚡ BẤM ĐỂ LẬP BÁO CÁO HỢP NHẤT"** (chọn màu xanh dương).
   - Chuột phải vào nút vừa vẽ ➔ Chọn **Assign Macro...** ➔ Chọn macro **`CapNhatVaLapBaoCaoHopNhat`** ➔ Bấm **OK**.

#### Bước 4: Lưu file định dạng Macro
* Nhấn **F12** (hoặc vào **File** ➔ **Save As**).
* Tại mục *Save as type*, chọn **`Excel Macro-Enabled Workbook (*.xlsm)`**.
* Đặt tên file là: **`Bo_Bao_Cao_Hop_Nhat_TCTD.xlsm`** và bấm **Save**.

---

### 💻 Hướng dẫn sử dụng định kỳ hàng quý (Q1, Q2, Q3, Q4...)

Từ nay về sau, mỗi khi có kỳ báo cáo mới, bạn chỉ cần mở file **`Bo_Bao_Cao_Hop_Nhat_TCTD.xlsm`**:

1. **Bấm Nút Bước 1 ("📁 BẤM ĐỂ CHỌN FILE CĐTK"):**
   - Hộp thoại chọn file sẽ hiện ra, bạn chỉ cần chọn file Excel Bảng cân đối tài khoản Thông tư 200 (ví dụ `Bang_can_doi_tai_khoanQ22026.xlsx`).
   - Hệ thống sẽ tự nhận diện: *QUÝ 2 NĂM 2026*.
   - Hệ thống tự trích xuất số dư tài khoản `13885.02` và `13881` để điền số tiền gợi ý cấn trừ vào bảng.

2. **Kiểm tra / Sửa số tiền cấn trừ (nếu cần):**
   - Nếu bạn đồng ý với số tự động gợi ý ➔ Để trống cột màu vàng.
   - Nếu bạn muốn chỉnh sửa số tiền cấn trừ theo ý mình ➔ Nhập số tiền vào cột màu vàng *Sửa thủ công (nếu có)*.

3. **Bấm Nút Bước 2 ("⚡ BẤM ĐỂ LẬP BÁO CÁO HỢP NHẤT"):**
   - Trong vòng 1 giây, Macro sẽ:
     * Quy đổi chính xác 100% sang 84 tài khoản TCTD trên sheet **`so lieu`**.
     * Áp dụng các bút toán cấn trừ ngoại tệ và chuyển phải thu.
     * Cập nhật tiêu đề Quý, Năm, Ngày báo cáo trên tất cả các sheet **`B02`**, **`B03`**, **`B04`**, **`B05`**, **`Thue`**.
     * Kiểm tra tính cân đối và cập nhật tích xanh **`CÂN ĐỐI (0 đ)`** lên Dashboard trang chủ.

4. **Xem và gửi báo cáo:**
   - Bạn chỉ cần bấm sang các sheet `B02`, `B03`, `B04`, `B05` là đã có bộ báo cáo tài chính hoàn chỉnh gửi công ty mẹ!
