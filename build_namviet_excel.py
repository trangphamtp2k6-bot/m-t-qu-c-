import json
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, PieChart, Reference
import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

with open('namviet_100_clean.json', 'r', encoding='utf-8') as f:
    records = json.load(f)

wb = openpyxl.Workbook()

# Colors & Styles
NAVY = "1B365D"
SLATE = "234975"
LIGHT_BG = "F4F7FB"
ZEBRA = "F9FBFC"
RED_BG = "FDECEC"
RED_TXT = "B42318"
AMBER_BG = "FEF3C7"
AMBER_TXT = "B45309"
GREEN_BG = "ECFDF3"
GREEN_TXT = "027A48"
BORDER_COLOR = "D0D7DE"

thin_side = Side(style='thin', color=BORDER_COLOR)
med_navy = Side(style='medium', color=NAVY)
border_all = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

font_title = Font(name="Calibri", size=16, bold=True, color=NAVY)
font_sub = Font(name="Calibri", size=11, italic=True, color="475467")
font_sec = Font(name="Calibri", size=12, bold=True, color=NAVY)
font_hdr = Font(name="Calibri", size=10.5, bold=True, color="FFFFFF")
font_body = Font(name="Calibri", size=10.5, color="1D2939")
font_bold = Font(name="Calibri", size=10.5, bold=True, color="1D2939")

fill_hdr = PatternFill("solid", fgColor=NAVY)
fill_subhdr = PatternFill("solid", fgColor=SLATE)
fill_card = PatternFill("solid", fgColor=LIGHT_BG)
fill_zebra = PatternFill("solid", fgColor=ZEBRA)
fill_red = PatternFill("solid", fgColor=RED_BG)
fill_amber = PatternFill("solid", fgColor=AMBER_BG)
fill_green = PatternFill("solid", fgColor=GREEN_BG)

# Helper Audit Flag function
def get_audit_flag(r):
    code = r['Mã Hồ Sơ']
    budget = r['Ngân Sách (VNĐ)']
    status = r['Trạng Thái Xử Lý']
    if code in ['L044', 'L034']:
        return f"⚠️ VIP ≥500tr BỊ ĐÓNG TỰ ĐỘNG ({budget/1e6:.0f}tr) - Cần TPKD/KAM thẩm định"
    if code in ['L055', 'L084', 'L090', 'L100']:
        return f"⚠️ LỖI PHÂN VÙNG/NGÀNH ({r['Tỉnh / Thành Phố']}) - Thiếu quy tắc Miền Trung"
    if code == 'L013':
        return "🔍 LỊCH SỬ KHIẾU NẠI CSKH - Đẩy TPKD xử lý đúng Bước 1"
    if status == 'CẦN TPKD DUYỆT':
        return "⚖️ ĐA LĨNH VỰC - Cần TPKD phân công"
    if status == 'CHỜ BỔ SUNG THÔNG TIN':
        return "📩 THIẾU THÔNG TIN - Đang gửi khảo sát tự động"
    if status == 'ĐÃ GỘP HỒ SƠ':
        return "🔗 ĐÃ GỘP MASTER - Chống trùng lặp thành công"
    if status == 'TỪ CHỐI / ĐÓNG LEAD':
        return "🛑 NGOÀI PHẠM VI SMB - Đã từ chối"
    return "✅ HỢP LỆ - Theo dõi SLA 4h / 24h"

# ==================== SHEET 2: DATA 100 LEADS AUDIT ====================
ws_data = wb.active
ws_data.title = "2_Data_100_Leads_Audit"
ws_data.views.sheetView[0].showGridLines = True

headers = [
    "Mã Hồ Sơ", "Người Đại Diện", "Tên Doanh Nghiệp Chuẩn Hóa", "MST / Domain",
    "Số Điện Thoại", "Email", "Tỉnh / Thành Phố", "Nhóm Ngành", "Gói Sản Phẩm",
    "Ngân Sách (VNĐ)", "Mức Ưu Tiên", "NVKD Phụ Trách", "Trạng Thái Xử Lý",
    "Hành Động Đề Xuất", "Cảnh Báo Kiểm Toán BI", "SLA Mục Tiêu"
]

ws_data.append(headers)
for col_idx in range(1, len(headers) + 1):
    cell = ws_data.cell(1, col_idx)
    cell.font = font_hdr
    cell.fill = fill_hdr
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = border_all
ws_data.row_dimensions[1].height = 28

for idx, r in enumerate(records, start=2):
    flag = get_audit_flag(r)
    sla = "Theo dõi 4h / 24h" if r['Trạng Thái Xử Lý'] == 'ĐÃ PHÂN CÔNG' else ("Khẩn cấp TPKD" if "TPKD" in flag or "VIP" in flag else "Không áp dụng")
    row_vals = [
        r['Mã Hồ Sơ'], r['Người Đại Diện'], r['Tên Doanh Nghiệp Chuẩn Hóa'], r['Mã Số Thuế / Tên Miền'],
        r['Số Điện Thoại'], r['Địa Chỉ Email'], r['Tỉnh / Thành Phố'], r['Nhóm Ngành Chuẩn Hóa'],
        r['Gói Sản Phẩm'], r['Ngân Sách (VNĐ)'], r['Mức Ưu Tiên'], r['NVKD Phụ Trách'],
        r['Trạng Thái Xử Lý'], r['Hành Động Đề Xuất'], flag, sla
    ]
    ws_data.append(row_vals)
    for c_idx in range(1, len(row_vals) + 1):
        c = ws_data.cell(idx, c_idx)
        c.font = font_body
        c.border = border_all
        c.alignment = Alignment(vertical="center", wrap_text=(c_idx in [14, 15]))
        if c_idx == 10:
            c.number_format = "#,##0"
            c.alignment = Alignment(horizontal="right", vertical="center")
        elif c_idx in [1, 11, 13, 16]:
            c.alignment = Alignment(horizontal="center", vertical="center")
        
        if "⚠️" in flag:
            c.fill = fill_red if "VIP" in flag else fill_amber
            if c_idx == 15:
                c.font = Font(name="Calibri", size=10, bold=True, color=RED_TXT if "VIP" in flag else AMBER_TXT)
        elif idx % 2 == 1:
            c.fill = fill_zebra

col_widths_data = [11, 18, 34, 15, 15, 24, 16, 24, 18, 16, 16, 24, 22, 42, 38, 18]
for i, w in enumerate(col_widths_data, 1):
    ws_data.column_dimensions[get_column_letter(i)].width = w
ws_data.freeze_panes = "B2"
ws_data.auto_filter.ref = f"A1:P101"

# ==================== SHEET 1: EXECUTIVE PIPELINE DASHBOARD ====================
ws_dash = wb.create_sheet(title="1_Pipeline_Dashboard", index=0)
ws_dash.views.sheetView[0].showGridLines = False

ws_dash["B2"] = "NAM VIỆT CRM – PIPELINE DASHBOARD & KIỂM TOÁN ĐIỀU PHỐI AI (100 LEADS MẪU)"
ws_dash["B2"].font = font_title
ws_dash["B3"] = "Giám sát luồng 4 bước AI Agent, Tải trọng 8 NVKD (Scale dự báo 700 Leads/tuần), SLA & Phát hiện rò rỉ cơ hội"
ws_dash["B3"].font = font_sub

# KPI Summary Block (Row 5-7)
kpis = [
    ("TỔNG LEAD ĐẦU VÀO", "=COUNTA('2_Data_100_Leads_Audit'!A2:A101)", "Tổng ngân sách: 6,765 tỷ ₫", "B"),
    ("BÀN GIAO 8 NVKD", "=COUNTIF('2_Data_100_Leads_Audit'!M2:M101,\"ĐÃ PHÂN CÔNG\")", "=SUMIF('2_Data_100_Leads_Audit'!M2:M101,\"ĐÃ PHÂN CÔNG\",'2_Data_100_Leads_Audit'!J2:J101)", "D"),
    ("HỒ SƠ TRÙNG ĐÃ GỘP", "=COUNTIF('2_Data_100_Leads_Audit'!M2:M101,\"ĐÃ GỘP HỒ SƠ\")", "Chống gọi đụng độ (6 Master)", "F"),
    ("TỪ CHỐI NGOÀI PHẠM VI", "=COUNTIF('2_Data_100_Leads_Audit'!M2:M101,\"TỪ CHỐI / ĐÓNG LEAD\")", "⚠️ Cảnh báo: 2 VIP ≥500tr (1,4 tỷ)", "H"),
    ("TPKD DUYỆT / XÁC MINH", "=COUNTIF('2_Data_100_Leads_Audit'!M2:M101,\"CẦN TPKD DUYỆT\")+COUNTIF('2_Data_100_Leads_Audit'!M2:M101,\"CHỜ BỔ SUNG THÔNG TIN\")", "4 Cần duyệt + 2 Thiếu thông tin", "J")
]

for title, formula, sub, col_start in kpis:
    col_idx = openpyxl.utils.column_index_from_string(col_start)
    col_end = get_column_letter(col_idx + 1)
    ws_dash.merge_cells(f"{col_start}5:{col_end}5")
    ws_dash.merge_cells(f"{col_start}6:{col_end}6")
    ws_dash.merge_cells(f"{col_start}7:{col_end}7")
    
    c5 = ws_dash[f"{col_start}5"]
    c5.value = title
    c5.font = Font(name="Calibri", size=9.5, bold=True, color="475467")
    c5.fill = fill_card
    c5.alignment = Alignment(horizontal="center", vertical="center")
    
    c6 = ws_dash[f"{col_start}6"]
    c6.value = formula
    c6.font = Font(name="Calibri", size=18, bold=True, color=NAVY)
    c6.fill = fill_card
    c6.alignment = Alignment(horizontal="center", vertical="center")
    
    c7 = ws_dash[f"{col_start}7"]
    c7.value = sub
    if str(sub).startswith("="):
        c7.number_format = '#,##0" ₫"'
        c7.font = Font(name="Calibri", size=10, bold=True, color=GREEN_TXT)
    elif "⚠️" in str(sub):
        c7.font = Font(name="Calibri", size=9.5, bold=True, color=RED_TXT)
    else:
        c7.font = Font(name="Calibri", size=9.5, color="475467")
    c7.fill = fill_card
    c7.alignment = Alignment(horizontal="center", vertical="center")
    
    for r_i in [5, 6, 7]:
        for c_i in [col_idx, col_idx + 1]:
            ws_dash.cell(r_i, c_i).border = border_all

ws_dash.row_dimensions[5].height = 20
ws_dash.row_dimensions[6].height = 30
ws_dash.row_dimensions[7].height = 20

# Section 1: Workload Table for 8 NVKD (Dynamic Formulas)
ws_dash["B9"] = "1. MA TRẬN TẢI TRỌNG & GIÁ TRỊ PIPELINE 8 NVKD (TÍNH TỰ ĐỘNG BẰNG CÔNG THỨC)"
ws_dash["B9"].font = font_sec

nv_headers = [
    "Mã & Tên NVKD", "Phạm Vi & Ngành Phụ Trách", "1 - Cao", "2 - Trung Bình", "3 - Thấp",
    "Tổng Lead Tuần Mẫu", "Dự Báo 700 Lead/Tuần", "Cảnh Báo Trần 30 Leads", "Tổng Ngân Sách (VNĐ)", "Trung Bình / Lead"
]
for i, h in enumerate(nv_headers, start=2):
    c = ws_dash.cell(10, i, h)
    c.font = font_hdr
    c.fill = fill_hdr
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    c.border = border_all
ws_dash.row_dimensions[10].height = 28

nvkd_list = [
    ("NV01 - Nguyễn Văn An", "Bán lẻ (PROD_RETAIL) - Miền Bắc"),
    ("NV02 - Trần Thị Bình", "Bán lẻ (PROD_RETAIL) - Miền Nam"),
    ("NV03 - Lê Hoàng Cường", "F&B (PROD_FNB) - Miền Bắc & Trung"),
    ("NV04 - Phạm Minh Dũng", "F&B (PROD_FNB) - Miền Nam"),
    ("NV05 - Vũ Thị Hoa", "Dịch vụ & Y tế (PROD_SERVICE) - Toàn quốc"),
    ("NV06 - Đỗ Tuấn Kiệt", "Sản xuất (PROD_MFG) - Toàn quốc"),
    ("NV07 - Bùi Thu Lan", "Đa ngành & Mới - Miền Bắc"),
    ("NV08 - Ngô Quang Nam", "Đa ngành & Mới - Miền Nam")
]

for idx, (nv_name, scope) in enumerate(nvkd_list, start=11):
    ws_dash.cell(idx, 2, nv_name).font = font_bold
    ws_dash.cell(idx, 3, scope).font = font_body
    ws_dash.cell(idx, 4, f'=COUNTIFS(\'2_Data_100_Leads_Audit\'!$L$2:$L$101, B{idx}, \'2_Data_100_Leads_Audit\'!$K$2:$K$101, "1 - CAO")')
    ws_dash.cell(idx, 5, f'=COUNTIFS(\'2_Data_100_Leads_Audit\'!$L$2:$L$101, B{idx}, \'2_Data_100_Leads_Audit\'!$K$2:$K$101, "2 - TRUNG BÌNH")')
    ws_dash.cell(idx, 6, f'=COUNTIFS(\'2_Data_100_Leads_Audit\'!$L$2:$L$101, B{idx}, \'2_Data_100_Leads_Audit\'!$K$2:$K$101, "3 - THẤP")')
    ws_dash.cell(idx, 7, f'=SUM(D{idx}:F{idx})')
    ws_dash.cell(idx, 8, f'=G{idx}*7')
    ws_dash.cell(idx, 9, f'=IF(H{idx}>30, "🔴 QUÁ TẢI (>30) - Kích hoạt Overflow", "✅ Trong ngưỡng SLA")')
    ws_dash.cell(idx, 10, f'=SUMIFS(\'2_Data_100_Leads_Audit\'!$J$2:$J$101, \'2_Data_100_Leads_Audit\'!$L$2:$L$101, B{idx})')
    ws_dash.cell(idx, 11, f'=IF(G{idx}>0, J{idx}/G{idx}, 0)')
    
    for col_i in range(2, 12):
        cell = ws_dash.cell(idx, col_i)
        cell.border = border_all
        if col_i in [4, 5, 6, 7, 8]:
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.font = font_bold if col_i in [7, 8] else font_body
        elif col_i in [10, 11]:
            cell.number_format = "#,##0"
            cell.alignment = Alignment(horizontal="right", vertical="center")
            cell.font = font_bold if col_i == 10 else font_body
        elif col_i == 9:
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.font = Font(name="Calibri", size=9.5, bold=True, color=RED_TXT if idx in [11, 12, 13, 14, 15, 16, 17] else GREEN_TXT)
        if idx % 2 == 0:
            cell.fill = fill_zebra

# Total Row (Row 19)
ws_dash.cell(19, 2, "TỔNG CỘNG TEAM SALES").font = font_hdr
ws_dash.cell(19, 3, "8 Nhân sự chuyên trách").font = font_hdr
for col_i, col_let in [(4, 'D'), (5, 'E'), (6, 'F'), (7, 'G'), (8, 'H'), (10, 'J')]:
    c = ws_dash.cell(19, col_i, f"=SUM({col_let}11:{col_let}18)")
    c.font = font_hdr
    if col_i == 10:
        c.number_format = "#,##0"
        c.alignment = Alignment(horizontal="right", vertical="center")
    else:
        c.alignment = Alignment(horizontal="center", vertical="center")
ws_dash.cell(19, 9, "7/8 NVKD quá tải khi lên 700 Lead").font = font_hdr
ws_dash.cell(19, 9).alignment = Alignment(horizontal="center", vertical="center")
ws_dash.cell(19, 11, "=J19/G19").font = font_hdr
ws_dash.cell(19, 11).number_format = "#,##0"
ws_dash.cell(19, 11).alignment = Alignment(horizontal="right", vertical="center")

for col_i in range(2, 12):
    ws_dash.cell(19, col_i).fill = fill_subhdr
    ws_dash.cell(19, col_i).border = border_all

# Add Stacked Bar Chart for 8 NVKD
chart1 = BarChart()
chart1.type = "col"
chart1.grouping = "stacked"
chart1.overlap = 100
chart1.title = "Phân bổ tải trọng Leads theo Mức độ Ưu tiên của 8 NVKD"
chart1.y_axis.title = "Số lượng Leads"
data_ref = Reference(ws_dash, min_col=4, min_row=10, max_col=6, max_row=18)
cats_ref = Reference(ws_dash, min_col=2, min_row=11, max_row=18)
chart1.add_data(data_ref, titles_from_data=True)
chart1.set_categories(cats_ref)
chart1.width = 18
chart1.height = 10
ws_dash.add_chart(chart1, "B21")

# Add Package Budget Summary & Audit Findings next to Chart (Columns H to K, Row 21+)
ws_dash["H21"] = "2. CƠ CẤU NGÂN SÁCH THEO GÓI SẢN PHẨM (80 LEADS)"
ws_dash["H21"].font = font_sec

pkg_hdrs = ["Gói Sản Phẩm", "Số Lead", "Tổng Ngân Sách (VNĐ)", "Tỷ Trọng %"]
for i, h in enumerate(pkg_hdrs, start=8):
    c = ws_dash.cell(22, i, h)
    c.font = font_hdr
    c.fill = fill_hdr
    c.alignment = Alignment(horizontal="center", vertical="center")
    c.border = border_all

pkgs = ["PROD_RETAIL", "PROD_MFG", "PROD_FNB", "PROD_SERVICE"]
for idx, pk in enumerate(pkgs, start=23):
    ws_dash.cell(idx, 8, pk).font = font_bold
    ws_dash.cell(idx, 9, f'=COUNTIFS(\'2_Data_100_Leads_Audit\'!$I$2:$I$101, H{idx}, \'2_Data_100_Leads_Audit\'!$M$2:$M$101, "ĐÃ PHÂN CÔNG")')
    ws_dash.cell(idx, 10, f'=SUMIFS(\'2_Data_100_Leads_Audit\'!$J$2:$J$101, \'2_Data_100_Leads_Audit\'!$I$2:$I$101, H{idx}, \'2_Data_100_Leads_Audit\'!$M$2:$M$101, "ĐÃ PHÂN CÔNG")')
    ws_dash.cell(idx, 11, f'=J{idx}/$J$27')
    for ci in range(8, 12):
        cell = ws_dash.cell(idx, ci)
        cell.border = border_all
        if ci == 9: cell.alignment = Alignment(horizontal="center")
        elif ci == 10:
            cell.number_format = "#,##0"
            cell.alignment = Alignment(horizontal="right")
        elif ci == 11:
            cell.number_format = "0.0%"
            cell.alignment = Alignment(horizontal="right")

ws_dash.cell(27, 8, "TỔNG PIPELINE").font = font_hdr
ws_dash.cell(27, 9, "=SUM(I23:I26)").font = font_hdr
ws_dash.cell(27, 9).alignment = Alignment(horizontal="center")
ws_dash.cell(27, 10, "=SUM(J23:J26)").font = font_hdr
ws_dash.cell(27, 10).number_format = "#,##0"
ws_dash.cell(27, 11, "=SUM(K23:K26)").font = font_hdr
ws_dash.cell(27, 11).number_format = "0.0%"
for ci in range(8, 12):
    ws_dash.cell(27, ci).fill = fill_subhdr
    ws_dash.cell(27, ci).border = border_all

# Audit Insights Box in Dashboard (Row 29-35, H to K)
ws_dash["H29"] = "3. CẢNH BÁO KIỂM TOÁN DỮ LIỆU & LỖI QUY TẮC AI"
ws_dash["H29"].font = Font(name="Calibri", size=11, bold=True, color=RED_TXT)
audit_notes = [
    "1. Lọt Luồng VIP (>500tr): L044 (900tr) & L034 (500tr) bị đóng tự động do ngoài SMB. Cần TPKD thẩm định Enterprise tránh rơi 1,4 tỷ.",
    "2. Lệch Tải Trọng: NV01 (17 leads), NV06 (15 leads, 1,005 tỷ - 12 Cao) quá tải nặng khi lên 700 leads/tuần (~105-119 leads), trong khi NV08 chỉ có 3 leads.",
    "3. Sai Vùng Miền Trung: L055 (Thanh Hóa), L090 (Khánh Hòa), L100 (Quảng Ngãi) giao cho NV07 Miền Bắc; L084 (Huế) giao cho NV08 Miền Nam.",
    "4. Thiếu Trường Tracking SLA: Cần bổ sung cột Timestamp_Assigned & First_Call_Time để đếm ngược 4h cảnh báo / 24h thu hồi theo Bước 4."
]
for i, note in enumerate(audit_notes, start=30):
    ws_dash.merge_cells(f"H{i}:K{i}")
    c = ws_dash[f"H{i}"]
    c.value = note
    c.font = Font(name="Calibri", size=9.5, color="1D2939")
    c.fill = fill_amber
    c.alignment = Alignment(wrap_text=True, vertical="center")
    ws_dash.row_dimensions[i].height = 28
    for ci in range(8, 12):
        ws_dash.cell(i, ci).border = border_all

col_widths_dash = [3, 24, 34, 10, 13, 10, 16, 18, 28, 20, 16]
for i, w in enumerate(col_widths_dash, 1):
    ws_dash.column_dimensions[get_column_letter(i)].width = w

# ==================== SHEET 3: DAILY WORKLIST 8 NVKD ====================
ws_work = wb.create_sheet(title="3_Daily_Worklist_NVKD")
assigned_recs = [r for r in records if r['Trạng Thái Xử Lý'] == 'ĐÃ PHÂN CÔNG']
# Sort by NVKD, then Priority (1-CAO, 2-TB, 3-THAP), then Budget desc
assigned_recs.sort(key=lambda x: (x['NVKD Phụ Trách'], x['Mức Ưu Tiên'], -x['Ngân Sách (VNĐ)']))

work_headers = [
    "NVKD Phụ Trách", "Thứ Tự Ưu Tiên", "Mã Hồ Sơ", "Tên Doanh Nghiệp", "Người Đại Diện",
    "Số Điện Thoại", "Email", "Tỉnh/TP", "Gói Sản Phẩm", "Ngân Sách (VNĐ)", "Hành Động Đề Xuất", "SLA Tối Đa"
]
ws_work.append(work_headers)
for ci in range(1, len(work_headers)+1):
    c = ws_work.cell(1, ci)
    c.font = font_hdr
    c.fill = fill_hdr
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    c.border = border_all
ws_work.row_dimensions[1].height = 28

for idx, r in enumerate(assigned_recs, start=2):
    row_v = [
        r['NVKD Phụ Trách'], r['Mức Ưu Tiên'], r['Mã Hồ Sơ'], r['Tên Doanh Nghiệp Chuẩn Hóa'],
        r['Người Đại Diện'], r['Số Điện Thoại'], r['Địa Chỉ Email'], r['Tỉnh / Thành Phố'],
        r['Gói Sản Phẩm'], r['Ngân Sách (VNĐ)'], r['Hành Động Đề Xuất'], "4h Nhắc / 24h Thu hồi"
    ]
    ws_work.append(row_v)
    for ci in range(1, len(row_v)+1):
        c = ws_work.cell(idx, ci)
        c.font = font_body
        c.border = border_all
        if ci == 10:
            c.number_format = "#,##0"
            c.alignment = Alignment(horizontal="right", vertical="center")
        elif ci in [2, 3, 9, 12]:
            c.alignment = Alignment(horizontal="center", vertical="center")
        if r['Mức Ưu Tiên'] == '1 - CAO' and ci == 2:
            c.fill = fill_red
            c.font = Font(name="Calibri", size=10, bold=True, color=RED_TXT)
        elif idx % 2 == 1:
            c.fill = fill_zebra

w_widths = [24, 16, 11, 34, 18, 15, 24, 16, 15, 16, 44, 20]
for i, w in enumerate(w_widths, 1):
    ws_work.column_dimensions[get_column_letter(i)].width = w
ws_work.freeze_panes = "B2"
ws_work.auto_filter.ref = f"A1:L{len(assigned_recs)+1}"

# ==================== SHEET 4: TPKD CAN DUYET & EXCEPTIONS ====================
ws_exc = wb.create_sheet(title="4_TPKD_Duyet_Va_Loi")
exc_recs = [r for r in records if r['Trạng Thái Xử Lý'] != 'ĐÃ PHÂN CÔNG' or r['Mã Hồ Sơ'] in ['L055', 'L084', 'L090', 'L100']]
exc_headers = [
    "Phân Nhóm Xử Lý", "Mã Hồ Sơ", "Tên Doanh Nghiệp", "Người Đại Diện", "Tỉnh/TP",
    "Ngân Sách (VNĐ)", "Trạng Thái Hiện Tại", "Cảnh Báo BI & Lý Do", "Khuyến Nghị TPKD Hành Động"
]
ws_exc.append(exc_headers)
for ci in range(1, len(exc_headers)+1):
    c = ws_exc.cell(1, ci)
    c.font = font_hdr
    c.fill = fill_hdr
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    c.border = border_all
ws_exc.row_dimensions[1].height = 28

for idx, r in enumerate(exc_recs, start=2):
    code = r['Mã Hồ Sơ']
    if code in ['L044', 'L034']:
        grp = "1. KHẨN CẤP - VIP ≥500TR BỊ ĐÓNG"
        rec_act = "Mở lại hồ sơ ngay, chỉ định Key Account Manager thẩm định gói Enterprise."
    elif code == 'L013':
        grp = "2. CẦN TPKD DUYỆT - LỊCH SỬ CSKH"
        rec_act = "TPKD liên hệ trực tiếp xoa dịu khiếu nại cũ trước khi giao Sales."
    elif r['Trạng Thái Xử Lý'] == 'CẦN TPKD DUYỆT':
        grp = "3. CẦN TPKD DUYỆT - ĐA LĨNH VỰC"
        rec_act = r['Hành Động Đề Xuất']
    elif code in ['L055', 'L084', 'L090', 'L100']:
        grp = "4. LỖI PHÂN VÙNG MIỀN TRUNG"
        rec_act = "Bổ sung quy tắc Miền Trung vào Ma trận (Giao NV03/NV07/NV08 theo vùng gần nhất)."
    elif r['Trạng Thái Xử Lý'] == 'CHỜ BỔ SUNG THÔNG TIN':
        grp = "5. THIẾU THÔNG TIN"
        rec_act = "Chờ khách điền form khảo sát tự động."
    elif r['Trạng Thái Xử Lý'] == 'ĐÃ GỘP HỒ SƠ':
        grp = "6. ĐÃ GỘP MASTER"
        rec_act = r['Hành Động Đề Xuất']
    else:
        grp = "7. ĐÓNG HỒ SƠ HỢP LỆ"
        rec_act = r['Hành Động Đề Xuất']

    row_v = [
        grp, code, r['Tên Doanh Nghiệp Chuẩn Hóa'], r['Người Đại Diện'], r['Tỉnh / Thành Phố'],
        r['Ngân Sách (VNĐ)'], r['Trạng Thái Xử Lý'], get_audit_flag(r), rec_act
    ]
    ws_exc.append(row_v)
    for ci in range(1, len(row_v)+1):
        c = ws_exc.cell(idx, ci)
        c.font = font_body
        c.border = border_all
        if ci == 6:
            c.number_format = "#,##0"
            c.alignment = Alignment(horizontal="right", vertical="center")
        if "KHẨN CẤP" in grp:
            c.fill = fill_red
        elif "LỖI PHÂN VÙNG" in grp or "LỊCH SỬ" in grp:
            c.fill = fill_amber

e_widths = [32, 11, 32, 18, 16, 16, 22, 38, 44]
for i, w in enumerate(e_widths, 1):
    ws_exc.column_dimensions[get_column_letter(i)].width = w
ws_exc.freeze_panes = "B2"
ws_exc.auto_filter.ref = f"A1:I{len(exc_recs)+1}"

out_ws = r"c:\Users\HP\Downloads\MÔN AI AGENT\NamViet_BI_Pipeline_Dashboard.xlsx"
out_dl = r"C:\Users\HP\Downloads\NamViet_BI_Pipeline_Dashboard.xlsx"
wb.save(out_ws)
wb.save(out_dl)
print("Saved Excel Dashboard successfully!")

