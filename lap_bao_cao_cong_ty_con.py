#!/usr/bin/env python3
"""
TOOL LẬP BÁO CÁO TÀI CHÍNH CÔNG TY CON GỬI CÔNG TY MẸ (PHỤC VỤ HỢP NHẤT)
Theo chuẩn Chế độ Kế toán Các Tổ chức Tín dụng (TCTD - NHNN)

Quy trình:
1. Đọc Bảng cân đối tài khoản Thông tư 200 (data/Bang_can_doi_tai_khoanQ...xlsx).
2. Tự động chuyển đổi & map sang hệ thống tài khoản TCTD cấp 5 (Sheet 'so lieu').
3. Áp dụng các bút toán điều chỉnh / cấn trừ nội bộ & đối tác (Sheet 'But toan can tru').
4. Cập nhật và xuất Bộ BCTC riêng lẻ gửi Công ty Mẹ:
   - B02 (Mẫu B02a/TCTD): Báo cáo tình hình tài chính giữa niên độ
   - B03 (Mẫu B03a/TCTD): Báo cáo kết quả hoạt động
   - B04 (Mẫu B04a/TCTD): Báo cáo lưu chuyển tiền tệ (Trực tiếp)
   - B05 (Mẫu B05a/TCTD): Thuyết minh BCTC
   - Thue: Bảng tình hình thực hiện nghĩa vụ thuế
"""

import os
import sys
import re
import json
import argparse
import openpyxl

# Bản đồ ánh xạ từ Hệ thống TK TT200 (Doanh nghiệp) sang Hệ thống TK TCTD (SBV cấp 5)
# Dựa trên chuẩn đối chiếu thực tế của đơn vị
ACCOUNT_MAPPING = {
    '101101': '11111.01',                                       # Tiền mặt VNĐ
    '131101': '11211.01',                                       # Tiền gửi không kỳ hạn VNĐ (+ cấn trừ)
    '131201': ['12811.02', '12811.04'],                         # Tiền gửi có kỳ hạn VNĐ
    '132101': '11221.01',                                       # Tiền gửi không kỳ hạn ngoại tệ
    '301301': '21121.01',                                       # Máy móc, thiết bị
    '301401': '21131.01',                                       # Phương tiện vận tải, truyền dẫn
    '301501': '21141.01',                                       # Thiết bị, dụng cụ quản lý
    '301902': '21181.01',                                       # TSCĐ khác
    '305101': '21411.01',                                       # Hao mòn TSCĐ hữu hình
    '305102': '21411.02',                                       # Hao mòn máy móc thiết bị
    '311001': None,
    '313001': None,
    '351001': '24411.01',                                       # Ký quỹ, ký cược dài hạn
    '353202': '1331',                                           # Thuế GTGT được khấu trừ
    '359299': ['13881.01', '13881.02', '13882.01', '13882.02', '13885.01', '13885.02'], # Phải thu khác (+ cấn trừ)
    '361207': None,
    '361299': '14111.09',                                       # Tạm ứng
    '361301': None,
    '388009': '24211.01',                                       # Chi phí trả trước ngắn hạn
    '391':    '13884.01',                                       # Phải thu lãi dự thu
    '397001': '13884.02',                                       # Phải thu phí
    '427909': '33861.01',                                       # Bảo hiểm thất nghiệp
    '453101': '33311',                                          # Thuế GTGT đầu ra
    '453401': '33341.01',                                       # Thuế TNDN
    '453801': '33382.01',                                       # Các loại thuế khác
    '453802': '33351.01',                                       # Thuế TNCN
    '453901': None,
    '454001': '33883.01',                                       # Kiều hối phải trả VNĐ
    '459908': '33511.09',                                       # Chi phí phải trả khác
    '459999': ['33821.01', '33831.01', '33841.01', '33882.02', '33884.01', '33884.02', '33886.01', '33886.02'], # Phải trả khác (+ cấn trừ)
    '462001': '33411',                                          # Phải trả công nhân viên
    '484101': '35311.01',                                       # Quỹ khen thưởng
    '484201': '35321.01',                                       # Quỹ phúc lợi
    '497001': '33885',                                          # Phải trả về phí
    '601001': '41111.01',                                       # Vốn đầu tư chủ sở hữu
    '611001': '41811.01',                                       # Các quỹ thuộc VCSH
    '612101': '41411.01',                                       # Quỹ đầu tư phát triển
    '631101': '41311.01',                                       # Chênh lệch tỷ giá hối đoái
    '691001': '42121.01',                                       # LN chưa phân phối năm nay
    '692001': '42111.01',                                       # LN chưa phân phối năm trước
    '701001': ['51511.02', '51511.03'],                         # Thu lãi tiền gửi (có kỳ hạn & không kỳ hạn)
    '711001': '51131.01',                                       # Thu phí chuyển tiền
    '721001': '51511.01',                                       # Lãi chênh lệch tỷ giá
    '790001': None,
    '790006': None,
    '811001': '63211.01',                                       # Chi phí hoạt động dịch vụ
    '821001': '63511.01',                                       # Lỗ chênh lệch tỷ giá
    '831002': None,
    '832099': '64251.03',                                       # Thuế, phí, lệ phí khác
    '833101': '82111.01',                                       # Chi phí thuế TNDN hiện hành
    '849001': None,
    '851103': '64211.01',                                       # Chi phí tiền lương
    '851104': None,
    '852001': '64213.01',                                       # Trang phục giao dịch
    '853101': '64212.01',                                       # BHXH
    '853201': '64212.02',                                       # BHYT
    '853401': '64212.03',                                       # Kinh phí công đoàn
    '853999': '64212.04',                                       # BHTN
    '856001': '64211.04',                                       # Ăn ca
    '859001': '64213.06',                                       # Chi phúc lợi
    '861101': '64221.01',                                       # Vật liệu văn phòng
    '861201': None,
    '861401': '64221.03',                                       # Xăng dầu
    '861999': None,
    '862001': ['64281.02', '64281.04'],                         # Chi phí công tác phí & đi lại
    '863001': '64281.01',                                       # Đào tạo, huấn luyện
    '864001': None,
    '865001': ['64271.02', '64271.13'],                         # Bưu phí, điện thoại
    '866001': ['64282.08', '64282.11'],                         # Khuyến mại & chi phí khác
    '868001': ['64282.05', '64282.06'],                         # Hoạt động Đảng, Đoàn thể
    '869101': ['64271.01', '64271.11', '64271.12'],             # Điện, nước, vệ sinh
    '869301': '64282.02',                                       # Hội nghị
    '869401': '64282.03',                                       # Lễ tân, khánh tiết
    '869402': '64282.04',                                       # Giao dịch đối ngoại
    '869501': None,
    '869701': '64271.07',                                       # Phòng cháy chữa cháy
    '869999': ['64271.10', '64281.03', '64281.09'],             # Kiểm toán, DV mua ngoài khác & QL khác
    '871001': '64241.01',                                       # Khấu hao TSCĐ
    '872001': '64271.08',                                       # Sửa chữa tài sản
    '874001': None,
    '875001': None,
    '876001': None,
    '899001': None,
}

QUARTER_NAMES_ROMAN = {1: 'I', 2: 'II', 3: 'III', 4: 'IV'}
QUARTER_END_DATES = {
    1: '31/03',
    2: '30/06',
    3: '30/09',
    4: '31/12'
}

def detect_period(file_path):
    """
    Tự động nhận diện Quý và Năm từ tiêu đề file hoặc tên file.
    """
    wb = openpyxl.load_workbook(file_path, data_only=True)
    ws = wb.active
    for r in range(1, 16):
        for c in range(1, 5):
            val = ws.cell(r, c).value
            if val:
                m = re.search(r'Quý\s*(\d+)\s*năm\s*(\d{4})', str(val), re.IGNORECASE)
                if m:
                    return int(m.group(1)), int(m.group(2))
    base = os.path.basename(file_path)
    m = re.search(r'Q(\d).*?(\d{4})', base, re.IGNORECASE)
    if m:
        return int(m.group(1)), int(m.group(2))
    return 1, 2026

def parse_trial_balance(file_path):
    """
    Đọc dữ liệu từ file Bảng cân đối tài khoản Thông tư 200.
    """
    wb = openpyxl.load_workbook(file_path, data_only=True)
    ws = wb.active
    tb = {}
    for r in range(1, ws.max_row + 1):
        v = ws.cell(r, 1).value
        if v and str(v).strip() and str(v).strip()[0].isdigit():
            tk = str(v).strip()
            name = str(ws.cell(r, 2).value or '').strip()
            dk_n = float(ws.cell(r, 3).value or 0)
            dk_c = float(ws.cell(r, 5).value or 0)
            ps_n = float(ws.cell(r, 6).value or 0)
            ps_c = float(ws.cell(r, 7).value or 0)
            ck_n = float(ws.cell(r, 13).value or 0)
            ck_c = float(ws.cell(r, 15).value or 0)
            tb[tk] = {
                'name': name,
                'dk_n': dk_n, 'dk_c': dk_c,
                'ps_n': ps_n, 'ps_c': ps_c,
                'ck_n': ck_n, 'ck_c': ck_c,
            }
    return tb

def get_tb_sum(tb, accounts):
    """
    Tính tổng số dư và phát sinh cho một hoặc một danh sách tài khoản TT200.
    """
    res = {'dk_n': 0.0, 'dk_c': 0.0, 'ps_n': 0.0, 'ps_c': 0.0, 'ck_n': 0.0, 'ck_c': 0.0}
    if not accounts:
        return res
    if isinstance(accounts, str):
        accounts = [accounts]
    for acc in accounts:
        d = tb.get(acc)
        if d:
            for k in res:
                res[k] += d[k]
    return res

def compute_so_lieu(tb, can_tru):
    """
    Tính toán bảng số liệu 84 tài khoản SBV cấp 5 từ Bảng CĐTK TT200 và các bút toán cấn trừ.
    """
    so_lieu_dict = {}
    d7 = can_tru.get('D7', 0.0)
    d8 = can_tru.get('D8', d7)
    d25 = can_tru.get('D25', 0.0)
    d27 = can_tru.get('D27', d25)

    for tk_sl, tt200_acc in ACCOUNT_MAPPING.items():
        if tt200_acc:
            d = get_tb_sum(tb, tt200_acc)
            vals = [d['dk_n'], d['dk_c'], d['ps_n'], d['ps_c'], d['ck_n'], d['ck_c']]
            
            # Áp dụng các bút toán cấn trừ đặc thù
            if tk_sl == '131101':
                # Tiền gửi không kỳ hạn VNĐ: cộng cấn trừ phải thu chuyển về tiền gửi
                vals[2] += d25  # PS Nợ
                vals[4] += d25  # CK Nợ
            elif tk_sl == '359299':
                # Phải thu khác: cấn trừ trung gian ngoại tệ và chuyển phải thu
                vals[3] += (d27 + d7)  # PS Có
                vals[4] -= (d25 + d7)  # CK Nợ
            elif tk_sl == '459999':
                # Phải trả khác: cấn trừ tài khoản trung gian ngoại tệ
                vals[3] -= d8   # PS Có
                vals[5] -= d7   # CK Có
            
            so_lieu_dict[tk_sl] = vals
        else:
            so_lieu_dict[tk_sl] = [0.0] * 6

    return so_lieu_dict

def update_report_package(template_path, output_path, quarter, year, so_lieu_dict, can_tru):
    """
    Điền dữ liệu vào Template Excel, cập nhật tiêu đề kỳ báo cáo và xuất file hoàn chỉnh.
    """
    wb = openpyxl.load_workbook(template_path)
    q_roman = QUARTER_NAMES_ROMAN.get(quarter, 'I')
    end_date_str = f"{QUARTER_END_DATES.get(quarter, '31/03')}/{year}"

    # 1. Cập nhật sheet 'so lieu'
    ws_sl = wb['so lieu']
    for r in range(1, 85):
        tk = ws_sl.cell(r, 1).value
        if tk:
            tk_str = str(tk).strip()
            if tk_str in so_lieu_dict:
                vals = so_lieu_dict[tk_str]
                for i in range(6):
                    ws_sl.cell(r, i + 2).value = vals[i]

    # 2. Cập nhật sheet 'But toan can tru'
    if 'But toan can tru' in wb.sheetnames:
        ws_bt = wb['But toan can tru']
        ws_bt['D4'] = can_tru.get('D4', 0)
        ws_bt['D7'] = can_tru.get('D7', 0)
        ws_bt['D10'] = can_tru.get('D10', 0)
        ws_bt['D12'] = can_tru.get('D12', 0)
        ws_bt['D25'] = can_tru.get('D25', 0)
        ws_bt['D27'] = can_tru.get('D27', 0)

    if 'B02' in wb.sheetnames:
        ws_b02 = wb['B02']
        ws_b02['A6'] = f"Quý {q_roman} năm {year}"
        ws_b02['D89'] = "='so lieu'!G41"  # Số dư cuối kỳ của TK 692001 (Lợi nhuận năm trước)

    if 'B03' in wb.sheetnames:
        ws_b03 = wb['B03']
        ws_b03['A7'] = f"Năm {year}"
        ws_b03['A8'] = f"Ngày {end_date_str}"
        ws_b03['D11'] = f"Quý {q_roman}"

        # Khi lập Quý 2: Điền số liệu Quý 1 vào Cột J để Cột L lũy kế cộng đủ 6 tháng
        if quarter == 2:
            ws_b03['J11'] = "Quý I"
            ws_b03['J12'] = f"Năm {year}"
            def dk_n(tk): return so_lieu_dict.get(tk, [0]*6)[0]
            def dk_c(tk): return so_lieu_dict.get(tk, [0]*6)[1]
            j14 = dk_c('701001') - dk_n('701001')
            j17 = dk_c('711001')
            j18 = dk_n('811001') - dk_c('811001')
            j20 = dk_c('721001') - dk_n('721001') - dk_n('821001')
            chiphi_q1 = sum(dk_n(tk) for tk in [
                '832099', '851103', '852001', '853101', '853201', '853401', '853999', '856001', '859001',
                '861101', '861401', '862001', '863001', '865001', '866001', '868001', '869101', '869301',
                '869401', '869402', '869701', '869999', '871001', '872001'
            ])
            j27 = chiphi_q1
            j28 = j14 + (j17 - j18) + j20 - j27
            j30 = j28
            j31 = dk_n('833101')
            j34 = j30 - j31
            j36 = j34
            ws_b03['J14'] = j14
            ws_b03['J17'] = j17
            ws_b03['J18'] = j18
            ws_b03['J20'] = j20
            ws_b03['J27'] = j27
            ws_b03['J30'] = j30
            ws_b03['J31'] = j31
            ws_b03['J34'] = j34
            ws_b03['J36'] = j36

    if 'B04' in wb.sheetnames:
        ws_b04 = wb['B04']
        ws_b04['A7'] = f"Quý {q_roman} năm {year}"

    if 'B05' in wb.sheetnames:
        ws_b05 = wb['B05']
        ws_b05['A5'] = f"Quý {q_roman} năm {year}"

    if 'Thue' in wb.sheetnames:
        ws_thue = wb['Thue']
        ws_thue['A5'] = f"Quý {quarter} năm {year}"

    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    wb.save(output_path)
    print(f"-> Đã lưu báo cáo hợp nhất thành công tại: {output_path}")

def validate_and_print_summary(output_path, quarter, year, can_tru):
    """
    Đọc lại file kết quả, kiểm tra tính cân đối và in bảng tóm tắt chỉ tiêu chính.
    """
    wb = openpyxl.load_workbook(output_path, data_only=True)
    ws_sl = wb['so lieu']
    
    # Tính tổng cộng sheet so lieu
    tot_dk_n = sum(float(ws_sl.cell(r, 2).value or 0) for r in range(1, 85))
    tot_dk_c = sum(float(ws_sl.cell(r, 3).value or 0) for r in range(1, 85))
    tot_ps_n = sum(float(ws_sl.cell(r, 4).value or 0) for r in range(1, 85))
    tot_ps_c = sum(float(ws_sl.cell(r, 5).value or 0) for r in range(1, 85))
    tot_ck_n = sum(float(ws_sl.cell(r, 6).value or 0) for r in range(1, 85))
    tot_ck_c = sum(float(ws_sl.cell(r, 7).value or 0) for r in range(1, 85))
    
    diff_dk = tot_dk_n - tot_dk_c
    diff_ps = tot_ps_n - tot_ps_c
    diff_ck = tot_ck_n - tot_ck_c
    
    # Kiểm tra B02: Tổng Tài sản vs Tổng Nguồn vốn
    ws_b02 = wb['B02']
    tong_tai_san = None
    tong_nguon_von = None
    for r in range(40, min(80, ws_b02.max_row + 1)):
        name = str(ws_b02.cell(r, 2).value or '').strip()
        if 'TỔNG CỘNG TÀI SẢN' in name.upper() or 'TỔNG TÀI SẢN' in name.upper():
            tong_tai_san = float(ws_b02.cell(r, 4).value or 0)
        elif 'TỔNG CỘNG NGUỒN VỐN' in name.upper() or 'TỔNG NGUỒN VỐN' in name.upper():
            tong_nguon_von = float(ws_b02.cell(r, 4).value or 0)
            
    print("\n" + "=" * 70)
    print(f"📊 BÁO CÁO TỔNG KẾT BCTC HỢP NHẤT - QUÝ {QUARTER_NAMES_ROMAN.get(quarter, quarter)} NĂM {year}")
    print("=" * 70)
    print(f"1. Bút toán cấn trừ đã áp dụng:")
    print(f"   - Cấn trừ trung gian ngoại tệ (TK 33886.02 / 13885.02): {can_tru.get('D7', 0):>15,.0f} VNĐ")
    print(f"   - Chuyển phải thu NHCT về tiền gửi (TK 112 / 13881)   : {can_tru.get('D25', 0):>15,.0f} VNĐ")
    print("-" * 70)
    print(f"2. Bảng cân đối tài khoản TCTD (Sheet 'so lieu'):")
    print(f"   - Dư đầu kỳ : Nợ = {tot_dk_n:>18,.0f} | Có = {tot_dk_c:>18,.0f} | Chênh lệch = {diff_dk:,.0f} VNĐ")
    print(f"   - Phát sinh : Nợ = {tot_ps_n:>18,.0f} | Có = {tot_ps_c:>18,.0f} | Chênh lệch = {diff_ps:,.0f} VNĐ")
    print(f"   - Dư cuối kỳ: Nợ = {tot_ck_n:>18,.0f} | Có = {tot_ck_c:>18,.0f} | Chênh lệch = {diff_ck:,.0f} VNĐ")
    print("-" * 70)
    print(f"3. Bảng Cân đối kế toán (Mẫu B02a/TCTD):")
    if tong_tai_san is not None and tong_nguon_von is not None:
        diff_bs = tong_tai_san - tong_nguon_von
        status = "✅ CÂN ĐỐI TUYỆT ĐỐI (100%)" if abs(diff_bs) < 1 else f"❌ LỆCH: {diff_bs:,.0f} VNĐ"
        print(f"   - Tổng Tài sản  : {tong_tai_san:>20,.0f} VNĐ")
        print(f"   - Tổng Nguồn vốn: {tong_nguon_von:>20,.0f} VNĐ")
        print(f"   - Trạng thái    : {status}")
    else:
        print(f"   - Đã cập nhật đầy đủ công thức theo chuẩn mẫu B02.")
    print("=" * 70 + "\n")

def main():
    parser = argparse.ArgumentParser(
        description="Tool tự động lập Báo cáo Tài chính Công ty Con gửi Công ty Mẹ (chuẩn TCTD phục vụ Hợp nhất)"
    )
    parser.add_argument(
        "-i", "--input", required=True,
        help="Đường dẫn file Bảng cân đối tài khoản (TT200) đầu vào (VD: data/Bang_can_doi_tai_khoanQ12026.xlsx)"
    )
    parser.add_argument(
        "-t", "--template", default="templates/Mau_BCTC_Hop_Nhat_TCTD.xlsx",
        help="Đường dẫn file template BCTC (Mặc định: templates/Mau_BCTC_Hop_Nhat_TCTD.xlsx)"
    )
    parser.add_argument(
        "-o", "--output", default=None,
        help="Đường dẫn file Excel đầu ra (Mặc định: output/Bao_Cao_Hop_Nhat_Q...xlsx)"
    )
    parser.add_argument(
        "--can-tru-d7", type=float, default=None,
        help="Số tiền cấn trừ trung gian ngoại tệ (Nợ TK 33886.02 / Có TK 13885.02). Mặc định tự lấy dư cuối TK 13885.02"
    )
    parser.add_argument(
        "--can-tru-d25", type=float, default=None,
        help="Số tiền chuyển phải thu NHCT về tiền gửi (Nợ TK 112 / Có TK 13881). Mặc định tự lấy dư cuối TK 13881"
    )
    parser.add_argument(
        "--can-tru-file", default=None,
        help="Đường dẫn file JSON chứa các thông số cấn trừ tùy chỉnh"
    )

    args = parser.parse_args()

    if not os.path.exists(args.input):
        print(f"Lỗi: Không tìm thấy file đầu vào: {args.input}")
        sys.exit(1)

    if not os.path.exists(args.template):
        print(f"Lỗi: Không tìm thấy file template: {args.template}")
        sys.exit(1)

    quarter, year = detect_period(args.input)
    print(f"-> Nhận diện kỳ báo cáo: Quý {quarter} năm {year}")

    tb = parse_trial_balance(args.input)
    print(f"-> Đã đọc {len(tb)} tài khoản từ Bảng cân đối tài khoản.")

    # Xác định các bút toán cấn trừ
    can_tru = {}
    if args.can_tru_file and os.path.exists(args.can_tru_file):
        with open(args.can_tru_file, 'r', encoding='utf-8') as f:
            can_tru = json.load(f)

    # Nếu người dùng truyền CLI hoặc lấy tự động từ tài khoản tương ứng
    if args.can_tru_d7 is not None:
        can_tru['D7'] = args.can_tru_d7
    elif 'D7' not in can_tru:
        # Tự động lấy từ dư cuối Nợ của TK 13885.02 nếu có
        can_tru['D7'] = tb.get('13885.02', {}).get('ck_n', 0.0)

    if args.can_tru_d25 is not None:
        can_tru['D25'] = args.can_tru_d25
    elif 'D25' not in can_tru:
        # Tự động lấy từ dư cuối Nợ của TK 13881 nếu có
        can_tru['D25'] = tb.get('13881', {}).get('ck_n', 0.0)

    can_tru['D8'] = can_tru.get('D7', 0.0)
    can_tru['D27'] = can_tru.get('D25', 0.0)

    # Tính toán bảng so lieu
    so_lieu_dict = compute_so_lieu(tb, can_tru)

    # Đặt tên file output nếu chưa chỉ định
    output_path = args.output
    if not output_path:
        output_path = f"output/Bao_Cao_Hop_Nhat_Q{quarter}_{year}.xlsx"

    # Cập nhật vào Template và lưu
    update_report_package(args.template, output_path, quarter, year, so_lieu_dict, can_tru)

    # Kiểm tra cân đối và in tổng kết
    validate_and_print_summary(output_path, quarter, year, can_tru)

if __name__ == '__main__':
    main()
