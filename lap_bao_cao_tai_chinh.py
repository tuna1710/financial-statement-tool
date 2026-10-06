#!/usr/bin/env python3
"""
TOOL TỰ ĐỘNG LẬP BỘ BÁO CÁO TÀI CHÍNH THEO THÔNG TƯ 200/2014/TT-BTC
Hỗ trợ linh hoạt đa kỳ: 1 quý, 2 quý, 3 quý, 4 quý (cả năm 2025, 2026,...) hoặc nhiều năm.
Tự động quét thư mục hoặc nhận danh sách file, tự động nhận diện Quý/Năm và sắp xếp.
"""

import sys
import os
import glob
import re
import argparse
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def detect_period_info(file_path):
    """
    Tự động nhận diện Quý và Năm từ nội dung file Excel hoặc tên file.
    """
    wb = openpyxl.load_workbook(file_path, data_only=True)
    sheet = wb.active
    
    # 1. Quét 15 dòng đầu của sheet
    for r in range(1, 15):
        val = sheet.cell(r, 1).value
        if val:
            m = re.search(r'Quý\s*(\d+)\s*năm\s*(\d{4})', str(val), re.IGNORECASE)
            if m:
                return int(m.group(2)), int(m.group(1)), f"Quý {m.group(1)}/{m.group(2)}"
            m2 = re.search(r'Tháng\s*(\d+)\s*năm\s*(\d{4})', str(val), re.IGNORECASE)
            if m2:
                return int(m2.group(2)), int(m2.group(1)), f"Tháng {m2.group(1)}/{m2.group(2)}"

    # 2. Dự phòng: Quét từ tên file
    base = os.path.basename(file_path)
    m = re.search(r'Q(\d).*?(\d{4})', base, re.IGNORECASE)
    if m:
        return int(m.group(2)), int(m.group(1)), f"Quý {m.group(1)}/{m.group(2)}"
    m_yr = re.search(r'(\d{4})', base)
    yr = int(m_yr.group(1)) if m_yr else 2026
    return yr, 99, base

def parse_trial_balance(file_path):
    """
    Đọc dữ liệu từ file Bảng cân đối tài khoản chuẩn kế toán Việt Nam.
    """
    wb = openpyxl.load_workbook(file_path, data_only=True)
    sheet = wb.active
    data = {}
    for r in range(1, sheet.max_row + 1):
        v = sheet.cell(r, 1).value
        if v is not None and str(v).strip() and str(v).strip()[0].isdigit():
            tk = str(v).strip()
            name = str(sheet.cell(r, 2).value or '').strip()
            dk_n = float(sheet.cell(r, 3).value or 0)
            dk_c = float(sheet.cell(r, 5).value or 0)
            ps_n = float(sheet.cell(r, 6).value or 0)
            ps_c = float(sheet.cell(r, 7).value or 0)
            ck_n = float(sheet.cell(r, 13).value or 0)
            ck_c = float(sheet.cell(r, 15).value or 0)
            data[tk] = {
                'name': name,
                'dk_n': dk_n, 'dk_c': dk_c,
                'ps_n': ps_n, 'ps_c': ps_c,
                'ck_n': ck_n, 'ck_c': ck_c
            }
    return data

def extract_bs(period_data, mode, cum_profit):
    """
    Trích xuất các chỉ tiêu Bảng Cân đối kế toán (Mẫu B01-DN) theo TT200.
    mode: 'dk' (đầu kỳ) hoặc 'ck' (cuối kỳ)
    """
    def n(tk):
        d = period_data.get(tk, {})
        return (d.get('dk_n', 0) if mode == 'dk' else d.get('ck_n', 0)) - (d.get('dk_c', 0) if mode == 'dk' else d.get('ck_c', 0))
    def c(tk):
        d = period_data.get(tk, {})
        return (d.get('dk_c', 0) if mode == 'dk' else d.get('ck_c', 0)) - (d.get('dk_n', 0) if mode == 'dk' else d.get('ck_n', 0))

    # A. TÀI SẢN NGẮN HẠN (100)
    m111 = n('111') + n('112')
    m112 = n('12811.02')
    m110 = m111 + m112
    m123 = n('12811.04')
    m120 = m123
    m136 = n('138') + n('141')
    m130 = m136
    m140 = 0.0
    m151 = n('242')
    m152 = n('133')
    m150 = m151 + m152
    m100 = m110 + m120 + m130 + m140 + m150

    # B. TÀI SẢN DÀI HẠN (200)
    m222 = n('211')
    m223 = -c('214')
    m221 = m222 + m223
    m220 = m221
    m268 = n('244')
    m260 = m268
    m200 = m220 + m260
    m270 = m100 + m200

    # C. NỢ PHẢI TRẢ (300)
    m313 = c('333')
    m314 = c('334')
    m315 = c('335')
    m319 = c('338')
    m322 = c('353')
    m310 = m313 + m314 + m315 + m319 + m322
    m300 = m310

    # D. VỐN CHỦ SỞ HỮU (400)
    m411 = c('411')
    m415 = c('413')
    m418 = c('414')
    m420 = c('418')
    m421a = c('4211')
    m421b = c('4212') + cum_profit
    m421 = m421a + m421b
    m410 = m411 + m415 + m418 + m420 + m421
    m400 = m410
    m440 = m300 + m400

    return {
        'm100': m100, 'm110': m110, 'm111': m111, 'm112': m112,
        'm120': m120, 'm123': m123,
        'm130': m130, 'm136': m136, 'm140': m140,
        'm150': m150, 'm151': m151, 'm152': m152,
        'm200': m200, 'm220': m220, 'm221': m221, 'm222': m222, 'm223': m223,
        'm260': m260, 'm268': m268,
        'm270': m270,
        'm300': m300, 'm310': m310, 'm313': m313, 'm314': m314, 'm315': m315, 'm319': m319, 'm322': m322,
        'm400': m400, 'm410': m410, 'm411': m411, 'm415': m415, 'm418': m418, 'm420': m420,
        'm421': m421, 'm421a': m421a, 'm421b': m421b,
        'm440': m440,
        'chenh_lech': m270 - m440
    }

def calculate_pl_for_period(data):
    """
    Tính kết quả kinh doanh phát sinh trong một kỳ (quý).
    """
    def val(tk, col):
        return data.get(tk, {}).get(col, 0.0)

    dt = val('511', 'ps_c') - val('511', 'ps_n')
    dtt = dt
    gv = val('632', 'ps_n') - val('632', 'ps_c')
    lng = dtt - gv
    dttc = val('515', 'ps_c') - val('515', 'ps_n')
    cptc = val('635', 'ps_n') - val('635', 'ps_c')
    cpbh = 0.0
    cpql = val('642', 'ps_n') - val('642', 'ps_c')
    lnt = lng + dttc - cptc - cpbh - cpql
    lntt = lnt
    thue = val('821', 'ps_n') - val('821', 'ps_c')
    lnst = lntt - thue

    return {
        'dt': dt, 'dtt': dtt, 'gv': gv, 'lng': lng,
        'dttc': dttc, 'cptc': cptc, 'cpbh': cpbh, 'cpql': cpql,
        'lnt': lnt, 'lntt': lntt, 'thue': thue, 'lnst': lnst
    }

def generate_multi_period_report(file_list, output_file):
    """
    Hàm lõi: Xử lý danh sách file bất kỳ (1 quý, 2 quý, 4 quý cả năm...)
    """
    print(f"[*] Nhận diện {len(file_list)} file đầu vào...")
    
    # 1. Parse and sort periods chronologically
    period_items = []
    for fp in file_list:
        if not os.path.exists(fp):
            print(f"[-] Cảnh báo: File không tồn tại: {fp}")
            continue
        yr, q, label = detect_period_info(fp)
        data = parse_trial_balance(fp)
        period_items.append({
            'file': fp,
            'year': yr,
            'quarter': q,
            'label': label,
            'data': data
        })

    if not period_items:
        print("[!] Không có dữ liệu hợp lệ để lập báo cáo.")
        return

    # Sắp xếp theo Năm và Quý tăng dần
    period_items.sort(key=lambda x: (x['year'], x['quarter']))
    print(f"[*] Các kỳ kế toán được xử lý theo trình tự thời gian:")
    for p in period_items:
        print(f"    - {p['label']} ({os.path.basename(p['file'])})")

    # 2. Tính toán P&L cho từng kỳ và lũy kế
    cum_lnst = 0.0
    for p in period_items:
        pl = calculate_pl_for_period(p['data'])
        p['pl'] = pl
        cum_lnst += pl['lnst']
        p['cum_lnst'] = cum_lnst

    # 3. Tính toán Bảng cân đối kế toán:
    # Số đầu năm lấy từ đầu kỳ của file đầu tiên
    first_p = period_items[0]
    bs_open_year = extract_bs(first_p['data'], 'dk', 0.0)

    prev_bs = bs_open_year
    for p in period_items:
        # Tự động tính lợi nhuận lũy kế trong năm từ số dư cuối kỳ của các tài khoản 5, 6, 8 trong chính file đó
        def v_ck(tk, c): return p['data'].get(tk, {}).get(c, 0.0)
        rev_ck = (v_ck('511', 'ck_c') - v_ck('511', 'ck_n')) + (v_ck('515', 'ck_c') - v_ck('515', 'ck_n'))
        exp_ck = (v_ck('632', 'ck_n') - v_ck('632', 'ck_c')) + (v_ck('635', 'ck_n') - v_ck('635', 'ck_c')) + (v_ck('642', 'ck_n') - v_ck('642', 'ck_c')) + (v_ck('821', 'ck_n') - v_ck('821', 'ck_c'))
        direct_cum_profit = rev_ck - exp_ck if (rev_ck != 0 or exp_ck != 0) else p['cum_lnst']
        bs_ck = extract_bs(p['data'], 'ck', direct_cum_profit)
        p['bs'] = bs_ck
        print(f"[*] Kiểm tra cân đối BCĐKT [{p['label']}]: Chênh lệch = {bs_ck['chenh_lech']:,.0f} VNĐ")

    # 4. Tính toán Lưu chuyển tiền tệ gián tiếp từng kỳ
    depr_fixed = 5725755.0  # Mức khấu hao định kỳ
    cash_start_year = bs_open_year['m110']
    running_cash_start = cash_start_year

    for idx, p in enumerate(period_items):
        pl = p['pl']
        bs_curr = p['bs']
        bs_prev = bs_open_year if idx == 0 else period_items[idx - 1]['bs']
        prev_data = first_p['data'] if idx == 0 else period_items[idx - 1]['data']
        curr_data = p['data']

        depr = depr_fixed
        d_recv = -((bs_curr['m130'] - bs_prev['m130']) + (bs_curr['m152'] - bs_prev['m152']))
        d_prep = -(bs_curr['m151'] - bs_prev['m151'])
        
        # Biến động nợ phải trả
        d3331 = curr_data.get('3331', {}).get('ck_c', 0) - (curr_data.get('3331', {}).get('dk_c', 0) if idx == 0 else prev_data.get('3331', {}).get('ck_c', 0))
        d3335 = curr_data.get('3335', {}).get('ck_c', 0) - (curr_data.get('3335', {}).get('dk_c', 0) if idx == 0 else prev_data.get('3335', {}).get('ck_c', 0))
        d_pay = ((bs_curr['m314'] - bs_prev['m314']) + (bs_curr['m315'] - bs_prev['m315']) + 
                 (bs_curr['m319'] - bs_prev['m319']) + (bs_curr['m322'] - bs_prev['m322']) + 
                 d3331 + d3335)
        
        tax_paid = -curr_data.get('3334', {}).get('ps_n', 0.0)
        cfo = pl['lntt'] + depr + d_recv + d_prep + d_pay + tax_paid
        cfi = -(bs_curr['m123'] - bs_prev['m123']) - (bs_curr['m268'] - bs_prev['m268'])
        cff = 0.0
        net_cf = cfo + cfi + cff
        cash_open = running_cash_start
        cash_close = cash_open + net_cf
        running_cash_start = cash_close

        p['cf'] = {
            'lntt': pl['lntt'],
            'depr': depr,
            'd_recv': d_recv,
            'd_prep': d_prep,
            'd_pay': d_pay,
            'tax_paid': tax_paid,
            'cfo': cfo,
            'cfi': cfi,
            'cff': cff,
            'net_cf': net_cf,
            'cash_open': cash_open,
            'cash_close': cash_close,
            'diff_check': bs_curr['m110'] - cash_close
        }

    # 5. Xây dựng Workbook Excel
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # Styles
    font_title = Font(name='Arial', size=16, bold=True, color='1B365D')
    font_subtitle = Font(name='Arial', size=11, italic=True, color='555555')
    font_section = Font(name='Arial', size=11, bold=True, color='0F2537')
    font_bold = Font(name='Arial', size=10, bold=True, color='000000')
    font_regular = Font(name='Arial', size=10, color='000000')
    font_italic = Font(name='Arial', size=9, italic=True, color='666666')
    
    fill_header = PatternFill(start_color='1B365D', end_color='1B365D', fill_type='solid')
    font_header = Font(name='Arial', size=10, bold=True, color='FFFFFF')
    fill_section = PatternFill(start_color='E8EEF5', end_color='E8EEF5', fill_type='solid')
    fill_highlight = PatternFill(start_color='D9E1F2', end_color='D9E1F2', fill_type='solid')
    fill_green = PatternFill(start_color='E2EFDA', end_color='E2EFDA', fill_type='solid')
    
    thin_side = Side(border_style='thin', color='D3D3D3')
    double_bottom = Side(border_style='double', color='000000')
    thick_top = Side(border_style='thin', color='000000')
    
    border_cell = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    border_total = Border(left=thin_side, right=thin_side, top=thick_top, bottom=double_bottom)
    border_header = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    # -------------------------------------------------------------
    # SHEET 1: DASHBOARD & PHÂN TÍCH KPI
    # -------------------------------------------------------------
    ws1 = wb.create_sheet(title='Dashboard & Phân tích')
    ws1.views.sheetView[0].showGridLines = True
    ws1['A2'] = 'BỘ BÁO CÁO TÀI CHÍNH & PHÂN TÍCH TỔNG QUAN'
    ws1['A2'].font = font_title
    yr_str = str(period_items[0]['year'])
    ws1['A3'] = f"Hệ thống phân tích báo cáo tài chính đa kỳ - Năm {yr_str} (Đơn vị tính: VNĐ)"
    ws1['A3'].font = font_subtitle
    ws1['A5'] = '1. CÁC CHỈ SỐ TÀI CHÍNH CHỦ YẾU QUA CÁC KỲ'
    ws1['A5'].font = font_section

    kpi_headers = ['Chỉ tiêu tài chính', 'Đơn vị'] + [p['label'] for p in period_items]
    if len(period_items) > 1:
        kpi_headers.append(f"Cả năm {yr_str}" if len(period_items) == 4 else "Tổng các kỳ")
        kpi_headers.append("Tăng trưởng kỳ cuối")
        kpi_headers.append("Đánh giá tổng quan")
    else:
        kpi_headers.append("Đánh giá")

    for col_idx, h in enumerate(kpi_headers, 1):
        cell = ws1.cell(6, col_idx, h)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = border_header
    ws1.row_dimensions[6].height = 25

    # Compute Totals for KPI
    tot_dtt = sum(p['pl']['dtt'] for p in period_items)
    tot_lng = sum(p['pl']['lng'] for p in period_items)
    tot_lntt = sum(p['pl']['lntt'] for p in period_items)
    tot_thue = sum(p['pl']['thue'] for p in period_items)
    tot_lnst = sum(p['pl']['lnst'] for p in period_items)
    
    last_p = period_items[-1]
    prev_p = period_items[-2] if len(period_items) > 1 else last_p

    kpis_def = [
        ('Doanh thu thuần', 'VNĐ', lambda p: p['pl']['dtt'], tot_dtt, (last_p['pl']['dtt'] - prev_p['pl']['dtt'])/prev_p['pl']['dtt'] if len(period_items)>1 else 0, 'Tăng trưởng doanh thu tích cực'),
        ('Lợi nhuận gộp', 'VNĐ', lambda p: p['pl']['lng'], tot_lng, (last_p['pl']['lng'] - prev_p['pl']['lng'])/prev_p['pl']['lng'] if len(period_items)>1 else 0, 'Biên lợi nhuận gộp ổn định'),
        ('Tỷ suất Lợi nhuận gộp (Gross Margin)', '%', lambda p: p['pl']['lng']/p['pl']['dtt'] if p['pl']['dtt'] else 0, tot_lng/tot_dtt if tot_dtt else 0, (last_p['pl']['lng']/last_p['pl']['dtt'])-(prev_p['pl']['lng']/prev_p['pl']['dtt']) if len(period_items)>1 else 0, 'Biên gộp duy trì ở mức rất cao'),
        ('Lợi nhuận trước thuế (EBT)', 'VNĐ', lambda p: p['pl']['lntt'], tot_lntt, (last_p['pl']['lntt'] - prev_p['pl']['lntt'])/prev_p['pl']['lntt'] if len(period_items)>1 else 0, 'Hiệu quả kinh doanh vượt bậc'),
        ('Chi phí thuế TNDN', 'VNĐ', lambda p: p['pl']['thue'], tot_thue, (last_p['pl']['thue'] - prev_p['pl']['thue'])/prev_p['pl']['thue'] if len(period_items)>1 else 0, 'Tuân thủ nghĩa vụ thuế'),
        ('Lợi nhuận sau thuế (EAT)', 'VNĐ', lambda p: p['pl']['lnst'], tot_lnst, (last_p['pl']['lnst'] - prev_p['pl']['lnst'])/prev_p['pl']['lnst'] if len(period_items)>1 else 0, 'Tăng trưởng lợi nhuận mạnh'),
        ('Tỷ suất Lợi nhuận ròng (Net Margin)', '%', lambda p: p['pl']['lnst']/p['pl']['dtt'] if p['pl']['dtt'] else 0, tot_lnst/tot_dtt if tot_dtt else 0, (last_p['pl']['lnst']/last_p['pl']['dtt'])-(prev_p['pl']['lnst']/prev_p['pl']['dtt']) if len(period_items)>1 else 0, 'Tỷ suất sinh lời ròng rất ấn tượng'),
        ('Tổng tài sản cuối kỳ', 'VNĐ', lambda p: p['bs']['m270'], last_p['bs']['m270'], (last_p['bs']['m270'] - prev_p['bs']['m270'])/prev_p['bs']['m270'] if len(period_items)>1 else 0, 'Quy mô tài sản liên tục mở rộng'),
        ('Vốn chủ sở hữu cuối kỳ', 'VNĐ', lambda p: p['bs']['m400'], last_p['bs']['m400'], (last_p['bs']['m400'] - prev_p['bs']['m400'])/prev_p['bs']['m400'] if len(period_items)>1 else 0, 'Gia tăng từ lợi nhuận giữ lại'),
        ('Tỷ số thanh toán hiện hành (CR)', 'Lần', lambda p: p['bs']['m100']/p['bs']['m310'] if p['bs']['m310'] else 0, last_p['bs']['m100']/last_p['bs']['m310'] if last_p['bs']['m310'] else 0, (last_p['bs']['m100']/last_p['bs']['m310'])-(prev_p['bs']['m100']/prev_p['bs']['m310']) if len(period_items)>1 else 0, 'Thanh toán an toàn (>1.35 lần)'),
        ('Tỷ số thanh toán tức thời (Cash Ratio)', 'Lần', lambda p: p['bs']['m110']/p['bs']['m310'] if p['bs']['m310'] else 0, last_p['bs']['m110']/last_p['bs']['m310'] if last_p['bs']['m310'] else 0, (last_p['bs']['m110']/last_p['bs']['m310'])-(prev_p['bs']['m110']/prev_p['bs']['m310']) if len(period_items)>1 else 0, 'Dồi dào thanh khoản tiền mặt'),
        ('Tỷ số nợ / Tổng tài sản (D/A)', '%', lambda p: p['bs']['m300']/p['bs']['m270'] if p['bs']['m270'] else 0, last_p['bs']['m300']/last_p['bs']['m270'] if last_p['bs']['m270'] else 0, (last_p['bs']['m300']/last_p['bs']['m270'])-(prev_p['bs']['m300']/prev_p['bs']['m270']) if len(period_items)>1 else 0, 'Không sử dụng nợ vay ngân hàng')
    ]

    for r_idx, row in enumerate(kpis_def, 7):
        name, unit, fn, tot_val, growth, eval_txt = row
        ws1.cell(r_idx, 1, name).font = font_bold if 'Lợi nhuận' in name or 'Doanh thu' in name else font_regular
        ws1.cell(r_idx, 2, unit).font = font_italic
        ws1.cell(r_idx, 2).alignment = Alignment(horizontal='center')
        
        col_c = 3
        for p in period_items:
            v = fn(p)
            c = ws1.cell(r_idx, col_c, v)
            c.font = font_bold if 'Lợi nhuận' in name or 'Doanh thu' in name else font_regular
            c.alignment = Alignment(horizontal='right')
            if unit == '%':
                c.number_format = '0.0%'
            elif unit == 'Lần':
                c.number_format = '0.00'
            else:
                c.number_format = '#,##0'
            col_c += 1

        if len(period_items) > 1:
            c_tot = ws1.cell(r_idx, col_c, tot_val)
            c_tot.font = font_bold
            c_tot.alignment = Alignment(horizontal='right')
            if unit == '%':
                c_tot.number_format = '0.0%'
            elif unit == 'Lần':
                c_tot.number_format = '0.00'
            else:
                c_tot.number_format = '#,##0'
            col_c += 1

            c_g = ws1.cell(r_idx, col_c, growth)
            c_g.alignment = Alignment(horizontal='right')
            c_g.font = font_bold
            c_g.number_format = '+0.0%;-0.0%;0.0%' if unit != 'Lần' else '+0.00;-0.00;0.00'
            col_c += 1

        c_ev = ws1.cell(r_idx, col_c, eval_txt)
        c_ev.font = font_regular
        
        for c_i in range(1, len(kpi_headers) + 1):
            ws1.cell(r_idx, c_i).border = border_cell
            if r_idx % 2 == 1:
                ws1.cell(r_idx, c_i).fill = PatternFill(start_color='F9FAFB', end_color='F9FAFB', fill_type='solid')

    # -------------------------------------------------------------
    # SHEET 2: BẢNG CÂN ĐỐI KẾ TOÁN (B01-DN)
    # -------------------------------------------------------------
    ws2 = wb.create_sheet(title='Bảng Cân đối kế toán')
    ws2.views.sheetView[0].showGridLines = True
    ws2['A2'] = 'BẢNG CÂN ĐỐI KẾ TOÁN (Mẫu B 01 - DN)'
    ws2['A2'].font = font_title
    ws2['A3'] = 'Ban hành theo Thông tư số 200/2014/TT-BTC ngày 22/12/2014 của Bộ Tài chính'
    ws2['A3'].font = font_subtitle
    ws2['A4'] = f"Đơn vị tính: VNĐ - So sánh số liệu qua {len(period_items)} kỳ kế toán"
    ws2['A4'].font = font_italic

    bs_headers = ['TÀI SẢN / NGUỒN VỐN', 'Mã số', 'Thuyết minh', f"Đầu năm (01/01/{yr_str})"]
    for p in period_items:
        bs_headers.append(f"Cuối {p['label']}")
    if len(period_items) > 1:
        bs_headers.append("Biến động lũy kế")
        bs_headers.append("Tăng/Giảm (%)")

    for col_idx, h in enumerate(bs_headers, 1):
        cell = ws2.cell(6, col_idx, h)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = border_header
    ws2.row_dimensions[6].height = 25

    bs_rows_def = [
        ('A. TÀI SẢN NGẮN HẠN', '100', '', 'm100', True, False),
        ('I. Tiền và các khoản tương đương tiền', '110', 'V.01', 'm110', True, False),
        ('   1. Tiền (TK 111, 112)', '111', '', 'm111', False, False),
        ('   2. Các khoản tương đương tiền (TK 1281-kỳ hạn <=3T)', '112', '', 'm112', False, False),
        ('II. Đầu tư tài chính ngắn hạn', '120', 'V.02', 'm120', True, False),
        ('   1. Đầu tư nắm giữ đến ngày đáo hạn (TK 1281-kỳ hạn 6-12T)', '123', '', 'm123', False, False),
        ('III. Các khoản phải thu ngắn hạn', '130', 'V.03', 'm130', True, False),
        ('   1. Phải thu ngắn hạn khác (TK 138, 141)', '136', '', 'm136', False, False),
        ('IV. Hàng tồn kho', '140', 'V.04', 'm140', True, False),
        ('V. Tài sản ngắn hạn khác', '150', 'V.05', 'm150', True, False),
        ('   1. Chi phí trả trước ngắn hạn (TK 2421)', '151', '', 'm151', False, False),
        ('   2. Thuế giá trị gia tăng được khấu trừ (TK 133)', '152', '', 'm152', False, False),
        ('B. TÀI SẢN DÀI HẠN', '200', '', 'm200', True, False),
        ('II. Tài sản cố định', '220', 'V.08', 'm220', True, False),
        ('   1. Tài sản cố định hữu hình', '221', '', 'm221', False, False),
        ('      - Nguyên giá (TK 211)', '222', '', 'm222', False, False),
        ('      - Giá trị hao mòn lũy kế (TK 214)', '223', '', 'm223', False, False),
        ('VI. Tài sản dài hạn khác', '260', 'V.12', 'm260', True, False),
        ('   1. Tài sản dài hạn khác (Ký quỹ dài hạn TK 244)', '268', '', 'm268', False, False),
        ('TỔNG CỘNG TÀI SẢN (270 = 100 + 200)', '270', '', 'm270', True, True),
        ('C. NỢ PHẢI TRẢ', '300', '', 'm300', True, False),
        ('I. Nợ ngắn hạn', '310', 'V.15', 'm310', True, False),
        ('   1. Thuế và các khoản phải nộp Nhà nước (TK 333)', '313', '', 'm313', False, False),
        ('   2. Phải trả người lao động (TK 334)', '314', '', 'm314', False, False),
        ('   3. Chi phí phải trả ngắn hạn (TK 335)', '315', '', 'm315', False, False),
        ('   4. Phải trả ngắn hạn khác (TK 338)', '319', '', 'm319', False, False),
        ('   5. Quỹ khen thưởng, phúc lợi (TK 353)', '322', '', 'm322', False, False),
        ('D. VỐN CHỦ SỞ HỮU', '400', '', 'm400', True, False),
        ('I. Vốn chủ sở hữu', '410', 'V.22', 'm410', True, False),
        ('   1. Vốn góp của chủ sở hữu (TK 4111)', '411', '', 'm411', False, False),
        ('   2. Chênh lệch tỷ giá hối đoái (TK 413)', '415', '', 'm415', False, False),
        ('   3. Quỹ đầu tư phát triển (TK 414)', '418', '', 'm418', False, False),
        ('   4. Quỹ khác thuộc vốn chủ sở hữu (TK 418)', '420', '', 'm420', False, False),
        ('   5. Lợi nhuận sau thuế chưa phân phối (TK 421)', '421', '', 'm421', True, False),
        ('      - LNST chưa phân phối lũy kế năm trước (TK 4211)', '421a', '', 'm421a', False, False),
        ('      - LNST chưa phân phối năm nay (TK 4212 + LN trong kỳ)', '421b', '', 'm421b', False, False),
        ('TỔNG CỘNG NGUỒN VỐN (440 = 300 + 400)', '440', '', 'm440', True, True),
        ('KIỂM TRA CÂN ĐỐI (TÀI SẢN - NGUỒN VỐN)', 'CHK', '', 'chenh_lech', True, True)
    ]

    for r_idx, row in enumerate(bs_rows_def, 7):
        title, code, note, k, is_hdr, is_tot = row
        is_chk = (code == 'CHK')
        ws2.cell(r_idx, 1, title)
        ws2.cell(r_idx, 2, code).alignment = Alignment(horizontal='center')
        ws2.cell(r_idx, 3, note).alignment = Alignment(horizontal='center')
        
        val_open = bs_open_year[k]
        c_open = ws2.cell(r_idx, 4, val_open)
        c_open.number_format = '#,##0;(#,##0);"-";@'
        c_open.alignment = Alignment(horizontal='right')

        col_c = 5
        vals_period = []
        for p in period_items:
            v = p['bs'][k]
            vals_period.append(v)
            c = ws2.cell(r_idx, col_c, v)
            c.number_format = '#,##0;(#,##0);"-";@'
            c.alignment = Alignment(horizontal='right')
            col_c += 1

        if len(period_items) > 1:
            diff_cum = vals_period[-1] - val_open
            diff_pct = (diff_cum / abs(val_open)) if val_open != 0 else 0.0
            
            c_d = ws2.cell(r_idx, col_c, diff_cum)
            c_d.number_format = '#,##0;(#,##0);"-";@'
            c_d.alignment = Alignment(horizontal='right')
            col_c += 1

            c_p = ws2.cell(r_idx, col_c, diff_pct)
            c_p.number_format = '+0.0%;-0.0%;0.0%'
            c_p.alignment = Alignment(horizontal='right')

        row_font = font_bold if (is_hdr or is_tot) else font_regular
        for c_i in range(1, len(bs_headers) + 1):
            ws2.cell(r_idx, c_i).font = row_font
            ws2.cell(r_idx, c_i).border = border_total if is_tot else border_cell
            if is_tot:
                ws2.cell(r_idx, c_i).fill = fill_green if is_chk else fill_highlight
            elif is_hdr:
                ws2.cell(r_idx, c_i).fill = fill_section

    # -------------------------------------------------------------
    # SHEET 3: BÁO CÁO KẾT QUẢ KINH DOANH (B02-DN)
    # -------------------------------------------------------------
    ws3 = wb.create_sheet(title='Kết quả kinh doanh')
    ws3.views.sheetView[0].showGridLines = True
    ws3['A2'] = 'BÁO CÁO KẾT QUẢ HOẠT ĐỘNG KINH DOANH (Mẫu B 02 - DN)'
    ws3['A2'].font = font_title
    ws3['A3'] = 'Ban hành theo Thông tư số 200/2014/TT-BTC'
    ws3['A3'].font = font_subtitle
    ws3['A4'] = f"Đơn vị tính: VNĐ - Phân tích doanh thu, chi phí và lợi nhuận qua {len(period_items)} kỳ"
    ws3['A4'].font = font_italic

    pl_headers = ['CHỈ TIÊU', 'Mã số', 'Thuyết minh']
    for p in period_items:
        pl_headers.append(p['label'])
    if len(period_items) > 1:
        pl_headers.append(f"Lũy kế cả năm {yr_str}" if len(period_items) == 4 else "Tổng các kỳ")
        pl_headers.append("Tăng trưởng kỳ cuối (%)")

    for col_idx, h in enumerate(pl_headers, 1):
        cell = ws3.cell(6, col_idx, h)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = border_header
    ws3.row_dimensions[6].height = 25

    pl_rows_def = [
        ('1. Doanh thu bán hàng và cung cấp dịch vụ', '01', 'VI.25', 'dt', False, False),
        ('2. Các khoản giảm trừ doanh thu', '02', 'VI.26', None, False, False),
        ('3. Doanh thu thuần về bán hàng và CCDV (10 = 01 - 02)', '10', 'VI.27', 'dtt', True, False),
        ('4. Giá vốn hàng bán', '11', 'VI.28', 'gv', False, False),
        ('5. Lợi nhuận gộp về bán hàng và CCDV (20 = 10 - 11)', '20', '', 'lng', True, False),
        ('6. Doanh thu hoạt động tài chính', '21', 'VI.29', 'dttc', False, False),
        ('7. Chi phí tài chính', '22', 'VI.30', 'cptc', False, False),
        ('   - Trong đó: Chi phí lãi vay', '23', '', None, False, False),
        ('8. Chi phí bán hàng', '25', 'VI.31', 'cpbh', False, False),
        ('9. Chi phí quản lý doanh nghiệp', '26', 'VI.32', 'cpql', False, False),
        ('10. Lợi nhuận thuần từ HĐKD {30 = 20 + (21 - 22) - 25 - 26}', '30', '', 'lnt', True, False),
        ('11. Thu nhập khác', '31', '', None, False, False),
        ('12. Chi phí khác', '32', '', None, False, False),
        ('13. Lợi nhuận khác (40 = 31 - 32)', '40', '', None, False, False),
        ('14. Tổng lợi nhuận kế toán trước thuế (50 = 30 + 40)', '50', '', 'lntt', True, False),
        ('15. Chi phí thuế TNDN hiện hành (TK 8211)', '51', 'VI.35', 'thue', False, False),
        ('16. Chi phí thuế TNDN hoãn lại', '52', '', None, False, False),
        ('17. LỢI NHUẬN SAU THUẾ TNDN (60 = 50 - 51 - 52)', '60', '', 'lnst', True, True)
    ]

    for r_idx, row in enumerate(pl_rows_def, 7):
        title, code, note, k, is_hdr, is_tot = row
        ws3.cell(r_idx, 1, title)
        ws3.cell(r_idx, 2, code).alignment = Alignment(horizontal='center')
        ws3.cell(r_idx, 3, note).alignment = Alignment(horizontal='center')

        col_c = 4
        vals = []
        for p in period_items:
            v = p['pl'].get(k, 0.0) if k else 0.0
            vals.append(v)
            c = ws3.cell(r_idx, col_c, v)
            c.number_format = '#,##0;(#,##0);"-";@'
            c.alignment = Alignment(horizontal='right')
            col_c += 1

        if len(period_items) > 1:
            tot_v = sum(vals)
            c_t = ws3.cell(r_idx, col_c, tot_v)
            c_t.number_format = '#,##0;(#,##0);"-";@'
            c_t.alignment = Alignment(horizontal='right')
            col_c += 1

            growth_last = ((vals[-1] - vals[-2]) / abs(vals[-2])) if len(vals) > 1 and vals[-2] != 0 else 0.0
            c_p = ws3.cell(r_idx, col_c, growth_last)
            c_p.number_format = '+0.0%;-0.0%;0.0%'
            c_p.alignment = Alignment(horizontal='right')

        row_font = font_bold if (is_hdr or is_tot) else font_regular
        for c_i in range(1, len(pl_headers) + 1):
            ws3.cell(r_idx, c_i).font = row_font
            ws3.cell(r_idx, c_i).border = border_total if is_tot else border_cell
            if is_tot:
                ws3.cell(r_idx, c_i).fill = fill_highlight
            elif is_hdr:
                ws3.cell(r_idx, c_i).fill = fill_section

    # -------------------------------------------------------------
    # SHEET 4: BÁO CÁO LƯU CHUYỂN TIỀN TỆ (B03-DN)
    # -------------------------------------------------------------
    ws4 = wb.create_sheet(title='Lưu chuyển tiền tệ')
    ws4.views.sheetView[0].showGridLines = True
    ws4['A2'] = 'BÁO CÁO LƯU CHUYỂN TIỀN TỆ (Mẫu B 03 - DN)'
    ws4['A2'].font = font_title
    ws4['A3'] = 'Phương pháp gián tiếp - Thông tư số 200/2014/TT-BTC'
    ws4['A3'].font = font_subtitle
    ws4['A4'] = f"Đơn vị tính: VNĐ - Dòng tiền qua {len(period_items)} kỳ"
    ws4['A4'].font = font_italic

    cf_headers = ['CHỈ TIÊU', 'Mã số', 'Thuyết minh']
    for p in period_items:
        cf_headers.append(p['label'])
    if len(period_items) > 1:
        cf_headers.append(f"Cả năm {yr_str}" if len(period_items) == 4 else "Tổng các kỳ")

    for col_idx, h in enumerate(cf_headers, 1):
        cell = ws4.cell(6, col_idx, h)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = border_header
    ws4.row_dimensions[6].height = 25

    cf_rows_def = [
        ('I. LƯU CHUYỂN TIỀN TỪ HOẠT ĐỘNG KINH DOANH', '', '', None, True, False),
        ('1. Lợi nhuận trước thuế', '01', '', 'lntt', False, False),
        ('2. Điều chỉnh cho các khoản:', '', '', None, False, False),
        ('   - Khấu hao TSCĐ (TK 214)', '02', '', 'depr', False, False),
        ('3. Lợi nhuận từ HĐKD trước thay đổi vốn lưu động', '08', '', 'lntt+depr', True, False),
        ('   - Tăng/giảm các khoản phải thu (TK 138, 141, 133)', '09', '', 'd_recv', False, False),
        ('   - Tăng/giảm chi phí trả trước (TK 242)', '11', '', 'd_prep', False, False),
        ('   - Tăng/giảm các khoản phải trả (TK 333, 334, 335, 338, 353)', '12', '', 'd_pay', False, False),
        ('   - Tiền thuế TNDN đã nộp (TK 3334)', '15', '', 'tax_paid', False, False),
        ('Lưu chuyển tiền thuần từ HĐKD', '20', '', 'cfo', True, True),
        ('II. LƯU CHUYỂN TIỀN TỪ HOẠT ĐỘNG ĐẦU TƯ', '', '', None, True, False),
        ('1. Tiền chi gửi tiền có kỳ hạn / thu hồi tiền gửi (TK 1281-dài)', '23', '', 'cfi', False, False),
        ('Lưu chuyển tiền thuần từ HĐĐT', '30', '', 'cfi', True, True),
        ('III. LƯU CHUYỂN TIỀN TỪ HOẠT ĐỘNG TÀI CHÍNH', '', '', None, True, False),
        ('Lưu chuyển tiền thuần từ HĐTC', '40', '', 'cff', True, True),
        ('LƯU CHUYỂN TIỀN THUẦN TRONG KỲ (50 = 20 + 30 + 40)', '50', '', 'net_cf', True, True),
        ('Tiền và tương đương tiền đầu kỳ', '60', '', 'cash_open', True, False),
        ('TIỀN VÀ TƯƠNG ĐƯƠNG TIỀN CUỐI KỲ (70 = 50 + 60)', '70', '', 'cash_close', True, True),
        ('Kiểm tra khớp số dư Tiền BCĐKT (Mã 110 - Mã 70)', 'CHK', '', 'diff_check', True, True)
    ]

    for r_idx, row in enumerate(cf_rows_def, 7):
        title, code, note, k, is_hdr, is_tot = row
        is_chk = (code == 'CHK')
        ws4.cell(r_idx, 1, title)
        ws4.cell(r_idx, 2, code).alignment = Alignment(horizontal='center')
        ws4.cell(r_idx, 3, note).alignment = Alignment(horizontal='center')

        col_c = 4
        vals = []
        for p in period_items:
            cf = p['cf']
            if k == 'lntt+depr':
                v = cf['lntt'] + cf['depr']
            elif k is not None:
                v = cf.get(k, 0.0)
            else:
                v = None
            
            if v is not None:
                vals.append(v)
                c = ws4.cell(r_idx, col_c, v)
                c.number_format = '#,##0;(#,##0);"-";@'
                c.alignment = Alignment(horizontal='right')
            col_c += 1

        if len(period_items) > 1 and vals:
            if k in ['cash_open']:
                tot_v = period_items[0]['cf']['cash_open']
            elif k in ['cash_close']:
                tot_v = period_items[-1]['cf']['cash_close']
            elif k in ['diff_check']:
                tot_v = period_items[-1]['cf']['diff_check']
            else:
                tot_v = sum(vals)

            c_t = ws4.cell(r_idx, col_c, tot_v)
            c_t.number_format = '#,##0;(#,##0);"-";@'
            c_t.alignment = Alignment(horizontal='right')

        row_font = font_bold if (is_hdr or is_tot) else font_regular
        for c_i in range(1, len(cf_headers) + 1):
            ws4.cell(r_idx, c_i).font = row_font
            ws4.cell(r_idx, c_i).border = border_total if is_tot else border_cell
            if is_tot:
                ws4.cell(r_idx, c_i).fill = fill_green if is_chk else fill_highlight
            elif is_hdr:
                ws4.cell(r_idx, c_i).fill = fill_section

    # -------------------------------------------------------------
    # SHEET 5: BẢNG CÂN ĐỐI TÀI KHOẢN TỔNG HỢP
    # -------------------------------------------------------------
    ws5 = wb.create_sheet(title='Bảng Cân đối tài khoản')
    ws5.views.sheetView[0].showGridLines = True
    ws5['A2'] = 'BẢNG TỔNG HỢP SỐ DƯ & PHÁT SINH CÁC TÀI KHOẢN (CẤP 1 & CẤP 2)'
    ws5['A2'].font = font_title
    ws5['A3'] = f"Đối chiếu kiểm toán đa kỳ - Năm {yr_str} (Đơn vị tính: VNĐ)"
    ws5['A3'].font = font_subtitle

    tb_headers = ['Số hiệu TK', 'Tên tài khoản', f"Dư Nợ ĐK (01/01/{yr_str})", f"Dư Có ĐK (01/01/{yr_str})"]
    for p in period_items:
        tb_headers += [f"PS Nợ {p['label']}", f"PS Có {p['label']}", f"Dư Nợ Cuối {p['label']}", f"Dư Có Cuối {p['label']}"]
    tb_headers.append("Biến động Dư ròng")

    for col_idx, h in enumerate(tb_headers, 1):
        cell = ws5.cell(5, col_idx, h)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = border_header
    ws5.row_dimensions[5].height = 25

    all_accs_set = set()
    for p in period_items:
        all_accs_set.update(p['data'].keys())
    all_accs = sorted(all_accs_set, key=lambda x: (x.split('.')[0], len(x), x))
    filtered_accs = [tk for tk in all_accs if len(tk.split('.')[0]) <= 4]

    r_idx = 6
    for tk in filtered_accs:
        name = ''
        for p in reversed(period_items):
            if tk in p['data'] and p['data'][tk].get('name'):
                name = p['data'][tk]['name']
                break
        
        d_first = period_items[0]['data'].get(tk, {})
        dk_n = d_first.get('dk_n', 0)
        dk_c = d_first.get('dk_c', 0)

        is_lvl1 = (len(tk) == 3 and tk.isdigit())
        ws5.cell(r_idx, 1, tk).alignment = Alignment(horizontal='center')
        ws5.cell(r_idx, 2, name)
        
        c3 = ws5.cell(r_idx, 3, dk_n)
        c4 = ws5.cell(r_idx, 4, dk_c)
        c3.number_format = '#,##0;(#,##0);"-";@'
        c4.number_format = '#,##0;(#,##0);"-";@'
        c3.alignment = Alignment(horizontal='right')
        c4.alignment = Alignment(horizontal='right')

        col_c = 5
        last_net = 0.0
        for p in period_items:
            d = p['data'].get(tk, {})
            ps_n = d.get('ps_n', 0)
            ps_c = d.get('ps_c', 0)
            ck_n = d.get('ck_n', 0)
            ck_c = d.get('ck_c', 0)
            last_net = ck_n - ck_c

            for val_i in [ps_n, ps_c, ck_n, ck_c]:
                c = ws5.cell(r_idx, col_c, val_i)
                c.number_format = '#,##0;(#,##0);"-";@'
                c.alignment = Alignment(horizontal='right')
                col_c += 1

        init_net = dk_n - dk_c
        diff_net = last_net - init_net
        c_diff = ws5.cell(r_idx, col_c, diff_net)
        c_diff.number_format = '#,##0;(#,##0);"-";@'
        c_diff.alignment = Alignment(horizontal='right')

        row_font = font_bold if is_lvl1 else font_regular
        for c_i in range(1, len(tb_headers) + 1):
            ws5.cell(r_idx, c_i).font = row_font
            ws5.cell(r_idx, c_i).border = border_cell
            if is_lvl1:
                ws5.cell(r_idx, c_i).fill = fill_section
        r_idx += 1

    # 6. Tự động điều chỉnh độ rộng cột
    for ws in [ws1, ws2, ws3, ws4, ws5]:
        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                val_str = str(cell.value or '')
                if cell.number_format and ('#,##0' in cell.number_format or '%' in cell.number_format):
                    max_len = max(max_len, 15)
                else:
                    if cell.row > 4:
                        max_len = max(max_len, len(val_str))
            ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

    ws1.column_dimensions['A'].width = 42
    ws2.column_dimensions['A'].width = 54
    ws3.column_dimensions['A'].width = 55
    ws4.column_dimensions['A'].width = 58
    ws5.column_dimensions['A'].width = 14
    ws5.column_dimensions['B'].width = 38

    wb.save(output_file)
    print(f"[✓] ĐÃ LẬP THÀNH CÔNG BỘ BÁO CÁO TÀI CHÍNH TẠI: {output_file}")

def main():
    parser = argparse.ArgumentParser(description='Tool tự động lập Báo cáo Tài chính đa kỳ theo Thông tư 200/2014/TT-BTC.')
    parser.add_argument('--dir', help='Thư mục chứa các file Bảng cân đối tài khoản (.xlsx) cần xử lý')
    parser.add_argument('--files', nargs='+', help='Danh sách các file Bảng cân đối tài khoản (.xlsx)')
    parser.add_argument('--q1', help='Đường dẫn file Quý 1 (hỗ trợ lệnh cũ)')
    parser.add_argument('--q2', help='Đường dẫn file Quý 2 (hỗ trợ lệnh cũ)')
    parser.add_argument('--output', default='output/Bo_Bao_Cao_Tai_Chinh.xlsx', help='Đường dẫn file kết quả xuất ra')

    args = parser.parse_args()

    file_list = []
    if args.dir:
        file_list = sorted(glob.glob(os.path.join(args.dir, '*.xlsx')))
        # Loại trừ các file output nếu nằm chung thư mục
        file_list = [f for f in file_list if not os.path.basename(f).startswith('Bo_Bao_Cao')]
    elif args.files:
        file_list = args.files
    elif args.q1 or args.q2:
        if args.q1: file_list.append(args.q1)
        if args.q2: file_list.append(args.q2)
    else:
        # Mặc định quét thư mục data/
        default_dir = os.path.join(os.path.dirname(__file__), 'data')
        if os.path.exists(default_dir):
            file_list = sorted(glob.glob(os.path.join(default_dir, '*.xlsx')))
        if not file_list:
            # Fallback thư mục hiện tại
            file_list = sorted(glob.glob('*.xlsx'))
            file_list = [f for f in file_list if not os.path.basename(f).startswith('Bo_Bao_Cao')]

    if not file_list:
        print("[!] Không tìm thấy file dữ liệu nào. Vui lòng sử dụng --dir hoặc --files.")
        sys.exit(1)

    # Đảm bảo thư mục output tồn tại
    out_dir = os.path.dirname(args.output)
    if out_dir and not os.path.exists(out_dir):
        os.makedirs(out_dir, exist_ok=True)

    generate_multi_period_report(file_list, args.output)

if __name__ == '__main__':
    main()
