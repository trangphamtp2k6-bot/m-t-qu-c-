import os
import sys
import docx
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

sys.stdout.reconfigure(encoding='utf-8')

IMG_DIR = r'C:\Users\HP\Downloads\extracted_hasaki_images'
OUT_LOCAL = r'c:\Users\HP\Downloads\MÔN AI AGENT\Bao_Cao_BI_Hasaki_Rut_Gon.docx'
OUT_DL = r'C:\Users\HP\Downloads\Bao_Cao_BI_Hasaki_Rut_Gon.docx'

doc = docx.Document()

# ==================== PAGE SETUP ====================
section = doc.sections[0]
section.page_width = Cm(21.0)
section.page_height = Cm(29.7)
section.top_margin = Cm(2.0)
section.bottom_margin = Cm(2.0)
section.left_margin = Cm(3.0)
section.right_margin = Cm(2.0)

# Colors
COLOR_PRIMARY = RGBColor(27, 54, 93)     # Deep Navy #1B365D
COLOR_SECONDARY = RGBColor(35, 73, 117)  # Slate Blue #234975
COLOR_TEXT = RGBColor(30, 30, 30)
COLOR_MUTED = RGBColor(100, 100, 100)
HEX_PRIMARY = "1B365D"
HEX_LIGHT_BG = "F4F7FB"
HEX_HEADER_BG = "1B365D"
HEX_ZEBRA = "F9FBFC"

def set_cell_background(cell, hex_color):
    shading_elm = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
    cell._tc.get_or_add_tcPr().append(shading_elm)

def set_cell_margins(cell, top=80, bottom=80, left=120, right=120):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(f'''
        <w:tcMar {nsdecls("w")}>
            <w:top w:w="{top}" w:type="dxa"/>
            <w:bottom w:w="{bottom}" w:type="dxa"/>
            <w:left w:w="{left}" w:type="dxa"/>
            <w:right w:w="{right}" w:type="dxa"/>
        </w:tcMar>
    ''')
    tcPr.append(tcMar)

def set_table_borders(table, color="B0C0D0", sz="4", val="single"):
    tblPr = table._tbl.tblPr
    borders = parse_xml(f'''
        <w:tblBorders {nsdecls("w")}>
            <w:top w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>
            <w:bottom w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>
            <w:insideH w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>
            <w:insideV w:val="none"/>
            <w:left w:val="none"/>
            <w:right w:val="none"/>
        </w:tblBorders>
    ''')
    tblPr.append(borders)

def format_run(run, font_name="Times New Roman", size_pt=12.5, bold=False, italic=False, color=COLOR_TEXT):
    run.font.name = font_name
    run.font.size = Pt(size_pt)
    run.bold = bold
    run.italic = italic
    run.font.color.rgb = color

def add_h1(text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.keep_with_next = True
    r = p.add_run(text)
    format_run(r, size_pt=14, bold=True, color=COLOR_PRIMARY)
    return p

def add_h2(text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.keep_with_next = True
    r = p.add_run(text)
    format_run(r, size_pt=12.5, bold=True, color=COLOR_SECONDARY)
    return p

def add_p(text, bold_prefix=None):
    p = doc.add_paragraph()
    p.paragraph_format.line_spacing = 1.3
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.first_line_indent = Cm(0.8)
    if bold_prefix:
        r_pre = p.add_run(bold_prefix)
        format_run(r_pre, size_pt=12, bold=True, color=COLOR_TEXT)
    r = p.add_run(text)
    format_run(r, size_pt=12, bold=False, color=COLOR_TEXT)
    return p

def add_b(text, bold_prefix=None):
    p = doc.add_paragraph(style='List Bullet')
    p.paragraph_format.line_spacing = 1.2
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    if bold_prefix:
        r_pre = p.add_run(bold_prefix)
        format_run(r_pre, size_pt=12, bold=True, color=COLOR_TEXT)
    r = p.add_run(text)
    format_run(r, size_pt=12, bold=False, color=COLOR_TEXT)
    return p

def add_box(title, items):
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    cell = table.cell(0, 0)
    cell.width = Cm(16.0)
    set_cell_background(cell, HEX_LIGHT_BG)
    set_cell_margins(cell, top=100, bottom=100, left=150, right=150)
    
    tcPr = cell._tc.get_or_add_tcPr()
    borders = parse_xml(f'''
        <w:tcBorders {nsdecls("w")}>
            <w:left w:val="single" w:sz="18" w:space="0" w:color="{HEX_PRIMARY}"/>
            <w:top w:val="none"/><w:right w:val="none"/><w:bottom w:val="none"/>
        </w:tcBorders>
    ''')
    tcPr.append(borders)
    
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(2)
    r_title = p.add_run(title)
    format_run(r_title, size_pt=11.5, bold=True, color=COLOR_PRIMARY)
    
    for it in items:
        p2 = cell.add_paragraph()
        p2.paragraph_format.space_before = Pt(1)
        p2.paragraph_format.space_after = Pt(1)
        p2.paragraph_format.line_spacing = 1.15
        p2.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        if isinstance(it, tuple):
            r_b = p2.add_run(f"• {it[0]}: ")
            format_run(r_b, size_pt=11, bold=True, color=COLOR_TEXT)
            r_t = p2.add_run(it[1])
            format_run(r_t, size_pt=11, bold=False, color=COLOR_TEXT)
        else:
            r_t = p2.add_run(f"• {it}")
            format_run(r_t, size_pt=11, bold=False, color=COLOR_TEXT)
    
    p_sp = doc.add_paragraph()
    p_sp.paragraph_format.space_after = Pt(2)

def add_fig(img_name, title, width=Cm(14.0)):
    img_path = os.path.join(IMG_DIR, img_name)
    if not os.path.exists(img_path):
        return
    p_img = doc.add_paragraph()
    p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_img.paragraph_format.space_before = Pt(4)
    p_img.paragraph_format.space_after = Pt(2)
    p_img.paragraph_format.keep_with_next = True
    r = p_img.add_run()
    r.add_picture(img_path, width=width)
    
    p_cap = doc.add_paragraph()
    p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_cap.paragraph_format.space_before = Pt(0)
    p_cap.paragraph_format.space_after = Pt(6)
    r_cap = p_cap.add_run(title)
    format_run(r_cap, size_pt=10.5, bold=True, italic=True, color=COLOR_TEXT)

def add_tbl(title, headers, rows, widths, aligns):
    p_cap = doc.add_paragraph()
    p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_cap.paragraph_format.space_before = Pt(6)
    p_cap.paragraph_format.space_after = Pt(2)
    p_cap.paragraph_format.keep_with_next = True
    r_cap = p_cap.add_run(title)
    format_run(r_cap, size_pt=11, bold=True, color=COLOR_PRIMARY)
    
    table = doc.add_table(rows=len(rows)+1, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    set_table_borders(table)
    
    # Header
    for i, h in enumerate(headers):
        c = table.rows[0].cells[i]
        c.width = widths[i]
        set_cell_background(c, HEX_HEADER_BG)
        set_cell_margins(c, top=80, bottom=80, left=80, right=80)
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(h)
        format_run(r, size_pt=10.5, bold=True, color=RGBColor(255, 255, 255))
        
    for r_i, row in enumerate(rows):
        bg = HEX_ZEBRA if r_i % 2 == 1 else "FFFFFF"
        for c_i, val in enumerate(row):
            c = table.rows[r_i+1].cells[c_i]
            c.width = widths[c_i]
            if bg != "FFFFFF":
                set_cell_background(c, bg)
            set_cell_margins(c, top=60, bottom=60, left=80, right=80)
            p = c.paragraphs[0]
            p.alignment = aligns[c_i]
            r = p.add_run(str(val))
            format_run(r, size_pt=10, bold=False, color=COLOR_TEXT)
            
    p_sp = doc.add_paragraph()
    p_sp.paragraph_format.space_after = Pt(4)

# ==================== 1. COVER PAGE ====================
tbl_cov = doc.add_table(rows=1, cols=1)
tbl_cov.alignment = WD_TABLE_ALIGNMENT.CENTER
tbl_cov.autofit = False
c_cov = tbl_cov.cell(0, 0)
c_cov.width = Cm(16.0)
set_cell_margins(c_cov, top=140, bottom=140, left=180, right=180)
tcPr = c_cov._tc.get_or_add_tcPr()
borders = parse_xml(f'''
    <w:tcBorders {nsdecls("w")}>
        <w:top w:val="double" w:sz="18" w:space="0" w:color="{HEX_PRIMARY}"/>
        <w:left w:val="double" w:sz="18" w:space="0" w:color="{HEX_PRIMARY}"/>
        <w:bottom w:val="double" w:sz="18" w:space="0" w:color="{HEX_PRIMARY}"/>
        <w:right w:val="double" w:sz="18" w:space="0" w:color="{HEX_PRIMARY}"/>
    </w:tcBorders>
''')
tcPr.append(borders)

p0 = c_cov.paragraphs[0]
p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
r0 = p0.add_run("HỌC VIỆN CÔNG NGHỆ BƯU CHÍNH VIỄN THÔNG\nKHOA QUẢN TRỊ KINH DOANH SỐ\n")
format_run(r0, size_pt=12, bold=True, color=COLOR_PRIMARY)

# Logo
p_logo = c_cov.add_paragraph()
p_logo.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_logo.paragraph_format.space_before = Pt(4)
p_logo.paragraph_format.space_after = Pt(10)
r_l1 = p_logo.add_run()
img1 = os.path.join(IMG_DIR, 'image1.png')
img2 = os.path.join(IMG_DIR, 'image2.png')
if os.path.exists(img1):
    r_l1.add_picture(img1, width=Cm(2.2))
r_sp = p_logo.add_run("     ")
r_l2 = p_logo.add_run()
if os.path.exists(img2):
    r_l2.add_picture(img2, width=Cm(3.2))

p_t1 = c_cov.add_paragraph()
p_t1.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_t1.paragraph_format.space_before = Pt(8)
p_t1.paragraph_format.space_after = Pt(4)
r_t1 = p_t1.add_run("BÁO CÁO PHÂN TÍCH BUSINESS INTELLIGENCE (BI301)")
format_run(r_t1, size_pt=13, bold=True, color=COLOR_PRIMARY)

p_t2 = c_cov.add_paragraph()
p_t2.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_t2.paragraph_format.space_before = Pt(4)
p_t2.paragraph_format.space_after = Pt(12)
r_t2 = p_t2.add_run("TỐI ƯU HÓA CHIẾN LƯỢC KHUYẾN MÃI, PHÂN TÍCH KÊNH BÁN HÀNG VÀ QUẢN TRỊ LỢI NHUẬN DANH MỤC SKINCARE – CHUỖI HASAKI BEAUTY & CLINIC")
format_run(r_t2, size_pt=14, bold=True, color=COLOR_SECONDARY)

p_info = c_cov.add_paragraph()
p_info.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_info.paragraph_format.space_before = Pt(10)
p_info.paragraph_format.space_after = Pt(14)
r_info = p_info.add_run("Giảng viên hướng dẫn: ThS. Hoàng Thị Kim Oanh\nSinh viên thực hiện: Phạm Thuỳ Trang  |  MSV: B25DTQT066\nĐơn vị đào tạo: PTIT x Rikkei Education  |  Năm 2026")
format_run(r_info, size_pt=11.5, bold=False, color=COLOR_TEXT)

doc.add_page_break()

# ==================== BƯỚC 1: XÁC ĐỊNH YÊU CẦU KINH DOANH ====================
add_h1("1. XÁC ĐỊNH YÊU CẦU KINH DOANH & CÂY CHỈ SỐ KPIs")
add_p("Hasaki Beauty & Clinic là chuỗi bán lẻ mỹ phẩm hàng đầu với hơn 140 chi nhánh và hệ thống bán hàng đa kênh (Cửa hàng, Website/App, Sàn TMĐT Shopee/TikTok/Lazada). Với cam kết 100% hàng chính hãng, giá cạnh tranh và giao nhanh 2h, Hasaki đối mặt với 3 thách thức lớn:", bold_prefix="Bối cảnh & Thách thức: ")
add_b("Khách hàng Gen Z nhạy cảm giá; việc lạm dụng giảm giá sâu (≥20%) ăn mòn lợi nhuận ròng.", bold_prefix="Cuộc chiến giá & Chiết khấu: ")
add_b("Sàn TMĐT mang lại lưu lượng nhưng phát sinh phí hoa hồng 8-15%; kênh cửa hàng/App giữ biên nhưng tốn chi phí thu hút khách.", bold_prefix="Xung đột đa kênh: ")
add_b("Mỹ phẩm, nước hoa vỏ thủy tinh dễ vỡ; giao trễ hoặc vỡ hỏng đẩy tỷ lệ hoàn trả tăng cao.", bold_prefix="Rủi ro Logistics chặng cuối: ")

add_p("Mục tiêu của dự án là phân tích 15.000 đơn hàng thực tế (2024–2025) để trả lời câu hỏi cốt lõi: Làm thế nào để tăng trưởng doanh số mà không đánh đổi biên lợi nhuận, chiết khấu bao nhiêu là tối ưu, và điểm nghẽn vận hành nào đang gây mất trải nghiệm khách hàng?", bold_prefix="Mục tiêu phân tích: ")

add_tbl("Bảng 1. Khung chỉ số hiệu suất chính (KPI Framework)",
    ["Nhóm KPI", "Tên chỉ số", "Cách tính toán", "Mục tiêu quản trị"],
    [
        ["Tài chính", "Doanh thu thuần (Net Sales)", "SUM(Sales) - SUM(Discount)", "Đo lường dòng tiền thực nhận"],
        ["Tài chính", "Lợi nhuận gộp & Biên gộp %", "(Profit / Net Sales) * 100", "Đo lường hiệu quả sinh lời (Mục tiêu ≥ 30%)"],
        ["Đơn hàng", "Giá trị đơn trung bình (AOV)", "Net Sales / Total Orders", "Theo dõi chất lượng giỏ hàng"],
        ["Vận hành", "Tỷ lệ đơn giao trễ (> 3 ngày)", "Đơn > 3 ngày / Tổng đơn giao từ xa", "Kiểm soát cam kết SLA (Mục tiêu < 5%)"],
        ["Khách hàng", "Điểm hài lòng CSAT & Hoàn trả %", "CSAT TB & Đơn hoàn / Tổng đơn", "Đo lường trải nghiệm (CSAT ≥ 4.0; Hoàn < 4%)"]
    ],
    [Cm(2.5), Cm(4.5), Cm(4.5), Cm(4.5)],
    [WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT]
)

# ==================== BƯỚC 2: THU THẬP & TÍCH HỢP DỮ LIỆU ETL ====================
add_h1("2. THU THẬP VÀ TÍCH HỢP DỮ LIỆU (DATA COLLECTION & ETL)")
add_p("Kiến trúc luồng dữ liệu trích xuất từ 4 nguồn chính: Hệ thống POS tại 140+ cửa hàng (bán lẻ tại quầy); Hệ thống CRM/ERP (hồ sơ khách hàng, phân hạng Cá nhân/VIP/B2B); Open APIs từ sàn TMĐT (Shopee, TikTok, Lazada); và Web/App log. Toàn bộ dữ liệu được nạp vào Staging Area, tiến hành làm sạch, chuyển đổi kiểu dữ liệu và tải vào Data Warehouse theo mô hình Lược đồ hình sao (Star Schema).", bold_prefix="Quy trình ETL: ")

# ==================== BƯỚC 3: LÀM SẠCH & CHUẨN BỊ DỮ LIỆU ====================
add_h1("3. LÀM SẠCH VÀ CHUẨN BỊ DỮ LIỆU (DATA CLEANING)")
add_p("Kiểm toán chất lượng dữ liệu trên toàn bộ 15.000 đơn hàng ghi nhận 100% tính duy nhất của mã Order_ID, không có bản ghi trùng lặp. Các công thức số học (Sales = Quantity * Unit_Price; Net_Sales = Sales - Discount_Amount; Profit = Net_Sales - Cost) khớp chính xác tuyệt đối.", bold_prefix="Kiểm toán chất lượng: ")
add_b("1.785 đơn không có đánh giá được gắn nhãn 'No Rating' thay vì điền số ngẫu nhiên để tránh làm lệch phân phối CSAT.", bold_prefix="Xử lý khuyết thiếu CSAT: ")
add_b("Order_Date chuyển về định dạng ISO (YYYY-MM-DD); Chiết khấu chuyển từ Text (%) sang Decimal (0.00 – 0.30); Đơn mua tại cửa hàng gán Shipping_Days = 0.", bold_prefix="Chuẩn hóa định dạng: ")
add_b("Tạo trường Discount_Bin phân nhóm chiết khấu ('0%', '1-10%', '11-20%', '>20%') phục vụ phân tích nhóm.", bold_prefix="Tạo trường dẫn xuất: ")

add_tbl("Bảng 2. Từ điển dữ liệu rút gọn (Các thuộc tính nghiệp vụ cốt lõi)",
    ["Thuộc tính", "Kiểu dữ liệu", "Mô tả nghiệp vụ", "Thuộc tính", "Kiểu dữ liệu", "Mô tả nghiệp vụ"],
    [
        ["Order_ID", "Varchar", "Mã giao dịch đơn hàng", "Category", "Varchar", "4 danh mục chính (Skincare, Nước hoa...)"],
        ["Order_Date", "Date", "Ngày đặt hàng thành công", "Unit_Price", "VND", "Đơn giá niêm yết bán lẻ"],
        ["Customer_ID", "Varchar", "Mã định danh khách hàng", "Discount", "Decimal", "Tỷ lệ chiết khấu áp dụng (0 - 30%)"],
        ["Segment", "Varchar", "Cá nhân, VIP, B2B", "Net_Sales", "VND", "Doanh thu thuần sau chiết khấu"],
        ["Channel", "Varchar", "Cửa hàng, Web/App, TMĐT", "Profit", "VND", "Lợi nhuận gộp đơn hàng"],
        ["Region", "Varchar", "Miền Bắc, Trung, Nam", "Shipping_Days", "Int", "Số ngày giao (0: quầy; 1-5+: online)"],
        ["Payment_Method", "Varchar", "VNPay, Thẻ, MoMo, COD", "Returned", "Varchar", "Trạng thái hoàn trả ('Có'/'Không')"]
    ],
    [Cm(2.6), Cm(1.8), Cm(3.6), Cm(2.6), Cm(1.8), Cm(3.6)],
    [WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT]
)

# ==================== BƯỚC 4: MÔ HÌNH HÓA DỮ LIỆU ====================
add_h1("4. MÔ HÌNH HÓA DỮ LIỆU ĐA CHIỀU (DATA MODELING)")
add_p("Hệ thống áp dụng mô hình Lược đồ hình sao (Star Schema) Kimball với bảng sự kiện trung tâm Fact_Sales kết nối 1-Nhiều với 4 bảng chiều: Dim_Date (thời gian), Dim_Product (sản phẩm), Dim_Customer (khách hàng), Dim_OrderContext (kênh, khu vực, thanh toán). Mô hình này tối ưu tốc độ truy vấn, hỗ trợ cắt lát đa chiều và lập trình các phép đo DAX phức tạp.", bold_prefix="Kiến trúc Star Schema: ")

add_fig('4ad2157b308ef30a8adf0921e2c764ccac083625.png', "Hình 1. Lược đồ hình sao (Star Schema): Fact_Sales ở trung tâm và 4 bảng chiều", width=Cm(14.5))

add_p("Các công thức DAX nghiệp vụ cốt lõi:", bold_prefix="Độ đo DAX chính: ")
add_b("Gross Margin % = DIVIDE(SUM('Fact_Sales'[Profit_VND]), SUM('Fact_Sales'[Net_Sales_VND]), 0)", bold_prefix="Biên gộp: ")
add_b("Return Rate % = DIVIDE(CALCULATE(COUNTROWS('Fact_Sales'), 'Fact_Sales'[Returned]=\"Có\"), COUNTROWS('Fact_Sales'), 0)", bold_prefix="Tỷ lệ hoàn: ")
add_b("Delay Rate % = DIVIDE(CALCULATE(COUNTROWS('Fact_Sales'), 'Fact_Sales'[Shipping_Days] > 3), COUNTROWS('Fact_Sales'), 0)", bold_prefix="Giao trễ: ")

# ==================== BƯỚC 5: TRỰC QUAN HÓA DASHBOARD ====================
add_h1("5. TRỰC QUAN HÓA DỮ LIỆU VÀ DASHBOARDS")
add_p("Dashboard Tổng quan Điều hành được thiết kế tương tác với 4 thẻ KPI tổng hợp và 5 biểu đồ cắt lớp theo kênh, tháng, khu vực, danh mục, phân khúc, tích hợp bộ lọc Slicer và Timeline:", bold_prefix="Giao diện Dashboard: ")

add_fig('2d3900aad2a5ea3e797d162f3bf54021759482ad.png', "Hình 2. Dashboard Tổng quan Điều hành: Doanh thu thuần và Lợi nhuận gộp đa chiều", width=Cm(15.2))

add_tbl("Bảng 3. Bảng chỉ số tổng hợp toàn bộ dữ liệu 2024 – 2025",
    ["Chỉ số tổng hợp", "Năm 2024", "Năm 2025", "Tăng trưởng / Thay đổi", "Nhận định"],
    [
        ["Doanh thu thuần (Net Sales)", "10,51 tỷ ₫", "13,81 tỷ ₫", "+31,4%", "Quy mô mở rộng mạnh mẽ"],
        ["Lợi nhuận gộp (Gross Profit)", "3,21 tỷ ₫", "4,23 tỷ ₫", "+31,8%", "Tăng tương ứng doanh thu"],
        ["Biên lợi nhuận gộp", "30,57%", "30,63%", "+0,06%", "Giữ ổn định quanh 30,6%"],
        ["Số lượng đơn hàng", "6.286 đơn", "8.714 đơn", "+38,6%", "Tăng nhanh hơn doanh thu"],
        ["Giá trị đơn TB (AOV)", "1,67 tr ₫", "1,58 tr ₫", "-5,4%", "Chất lượng đơn hàng giảm nhẹ"],
        ["Tỷ lệ hoàn trả chung", "3,85%", "3,92%", "+0,07%", "Tổng 583 đơn hoàn (mất 250 tr lãi)"]
    ],
    [Cm(4.5), Cm(2.5), Cm(2.5), Cm(3.0), Cm(3.5)],
    [WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.LEFT]
)

# ==================== BƯỚC 6: ĐÁNH GIÁ INSIGHT & ĐỀ XUẤT ====================
add_h1("6. ĐÁNH GIÁ, TRÍCH XUẤT INSIGHT VÀ ĐỀ XUẤT HÀNH ĐỘNG")

# 6.1 Doanh thu & Danh mục
add_h2("6.1. Hiệu suất Doanh thu và Cơ cấu Lợi nhuận")
add_fig('705af6b635f8c83181a758913c32d32c743a8076.png', "Hình 3. Doanh thu thuần và Lợi nhuận gộp theo tháng (Đỉnh tháng 3/2025 đạt 2,09 tỷ ₫)", width=Cm(14.0))

add_p("Doanh thu đạt đỉnh vào tháng 03/2025 (2,09 tỷ ₫, 1.262 đơn do dịp 8/3), sau đó giảm liên tục đến tháng 12/2025 (chỉ còn 59 đơn – cần đối soát lỗi dữ liệu ERP cuối kỳ). Nước hoa cao cấp chỉ chiếm 13,7% số đơn nhưng tạo 46,1% doanh thu và 49,1% lợi nhuận gộp (biên gộp 32,6%, AOV 5,47 tr ₫). Ngược lại, Skincare chiếm 40,8% số đơn nhưng biên gộp thấp nhất chuỗi (28,2%). Nhóm khách VIP nhận chiết khấu bình quân 11,5% nên có biên gộp thấp nhất (29,05%), trong khi khách B2B có biên gộp cao nhất (31,7%) và AOV gấp đôi bình quân (3,13 tr ₫).", bold_prefix="Phát hiện then chốt: ")

# 6.2 Khuyến mãi & Độ nhạy giá
add_h2("6.2. Khuyến mãi và Độ nhạy Giá: 'Bẫy chiết khấu sâu'")
add_fig('b204238171ea5cc0ddebd79ebc852d5d58106a2f.png', "Hình 4. Tỷ trọng số đơn và Biên lợi nhuận gộp theo từng mức chiết khấu", width=Cm(14.0))

add_tbl("Bảng 4. Hiệu quả tài chính theo các mức chiết khấu",
    ["Mức giảm", "Số đơn", "% Đơn", "% Lãi gộp", "Biên gộp", "AOV", "Hoàn trả %", "Đánh giá nghiệp vụ"],
    [
        ["0%", "1.740", "11,6%", "14,3%", "37,4%", "1,96 tr ₫", "3,1%", "Biên cao nhất, khách mua nguyên giá"],
        ["5%", "3.840", "25,6%", "28,4%", "33,8%", "1,78 tr ₫", "3,3%", "Vùng 'Điểm ngọt' (Sweet Spot) tạo ra"],
        ["10%", "3.825", "25,5%", "24,8%", "30,2%", "1,65 tr ₫", "3,5%", "hơn 53% lợi nhuận toàn hệ thống"],
        ["15%", "3.195", "21,3%", "24,3%", "26,5%", "1,52 tr ₫", "4,0%", "Ngưỡng an toàn tối đa cho phép"],
        ["20%", "1.305", "8,7%", "5,4%", "21,9%", "1,38 tr ₫", "5,1%", "Vùng nguy hiểm: Chiết khấu ≥20%"],
        ["25%", "675", "4,5%", "1,9%", "16,5%", "1,25 tr ₫", "6,8%", "chiếm 16% đơn nhưng chỉ đem 8,2% lãi;"],
        ["30%", "420", "2,8%", "0,9%", "10,2%", "1,17 tr ₫", "7,9%", "AOV giảm 40%, tỷ lệ hoàn vọt lên 7,9%"]
    ],
    [Cm(1.8), Cm(1.6), Cm(1.6), Cm(1.8), Cm(1.8), Cm(1.8), Cm(1.8), Cm(4.0)],
    [WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.LEFT]
)

add_fig('d3c28026e85ea244c9be68f2607b5d381b8e3c70.png', "Hình 5. Độ co giãn biên lợi nhuận của Skincare sụp đổ từ 35% xuống 7,3% khi giảm 30%", width=Cm(14.0))

# 6.3 Đa kênh & Vận hành
add_h2("6.3. Hiệu suất Đa kênh, Ngưỡng gãy SLA và Tổn thất Hoàn trả")
add_fig('9894f6daeca7f0d5a6eea2d9ecd362e734dcd177.png', "Hình 6. So sánh Doanh thu thuần và Lợi nhuận gộp theo Kênh bán hàng", width=Cm(14.0))

add_p("Cửa hàng vật lý là trụ cột (45,6% doanh thu, 45,8% lợi nhuận, CSAT 4,21, hoàn trả 1,9%). Sàn TMĐT đóng góp 19,1% doanh thu nhưng tỷ lệ hoàn trả lên tới 7,2% (gấp gần 4 lần cửa hàng); biên gộp sàn trên sổ sách (30,4%) là 'ảo' vì chưa trừ chi phí hoa hồng sàn 8-15%.", bold_prefix="Xung đột kênh: ")

add_fig('23d01866298b2b456c68b413e9bd8c3d6d872e4e.png', "Hình 7. Điểm hài lòng CSAT và Tỷ lệ hoàn trả theo Thời gian giao hàng (Ngưỡng gãy 3 ngày)", width=Cm(14.0))

add_p("Dữ liệu chỉ ra ngưỡng gãy trải nghiệm chính xác tại mốc 3 ngày. Giao 1-3 ngày giữ CSAT 4,07-4,09; sang ngày thứ 4 CSAT sụp đổ xuống 2,70 (43,5% bị chấm 1-2 sao). Nhóm giao trễ có 2.060 đơn (13,7%), tập trung ở Website/App và TMĐT. Tổng thiệt hại hoàn trả 583 đơn làm bốc hơi 0,25 tỷ ₫ lợi nhuận gộp và đọng 0,58 tỷ ₫ giá vốn hàng bán (Nước hoa chiếm 34,6% giá vốn hoàn).", bold_prefix="Ngưỡng gãy Logistics: ")

# 6.4 Thanh toán
add_h2("6.4. Phương thức Thanh toán: Bác bỏ định kiến về COD")
add_fig('6bd2fa01c79b3856095f9ddcac34036d73529499.png', "Hình 8. Cơ cấu Phương thức thanh toán (81,6% thanh toán số) và Tỷ lệ hoàn trả", width=Cm(14.0))
add_p("Thanh toán số đạt 81,6% (VNPay 25,3%, Thẻ 21,9%, MoMo 20,5%), COD chỉ 18,4%. Trên kênh giao hàng online từ xa, tỷ lệ hoàn trả của COD là 5,4% ngang ngửa thanh toán trước (5,4%–5,6%). Nguyên nhân hoàn trả nằm ở chất lượng bao bì và thời gian giao hàng, không phụ thuộc vào hình thức trả tiền.", bold_prefix="Bản chất thanh toán: ")

# 6.5 Nguyên nhân gốc rễ & Giải pháp
add_h2("6.5. Nguyên nhân Gốc rễ và Gói 7 Giải pháp Chiến lược (Lộ trình 90 ngày)")
add_box("4 NGUYÊN NHÂN GỐC RỄ (ROOT CAUSES)", [
    ("N1 - Chính sách giá cào bằng", "Áp cùng mức giảm cho mọi ngành hàng, khiến Skincare giảm 30% thì biên rơi về 7,3% (bán lỗ); VIP giảm 11,5% có biên thấp hơn khách thường."),
    ("N2 - Chiết khấu sâu không ràng buộc AOV", "Giảm giá sâu thu hút khách săn deal nhỏ, AOV giảm 40% (1,96 xuống 1,17 tr ₫), hoàn trả tăng lên 7,9%."),
    ("N3 - Thiếu chuẩn đóng gói chống sốc & Cảnh báo SLA 3 ngày", "Hàng thủy tinh dễ vỡ chặng cuối gây hoàn 583 tr vốn; đơn trễ >3 ngày kéo CSAT sụt dốc."),
    ("N4 - Thiếu ghi nhận chi phí ẩn sàn TMĐT", "Chưa phản ánh hoa hồng 8-15% trên DWH dẫn đến hiểu sai lợi nhuận thực tế giữa các kênh.")
])

add_tbl("Bảng 5. Gói 7 giải pháp chiến lược và Lộ trình triển khai 90 ngày",
    ["Giải pháp", "Hành động thực thi cốt lõi", "Phòng ban", "Lộ trình", "Kỳ vọng tài chính"],
    [
        ["GP1. Trần chiết khấu 15%", "Khóa mã giảm ≥20% cho Skincare; dồn đơn về dải Sweet Spot 5-10%", "Commercial", "Ngày 1-30", "+333 triệu ₫ lãi gộp"],
        ["GP2. Đổi ưu đãi VIP phi giá", "Thay voucher tiền bằng dịch vụ soi da/spa miễn phí tại Hasaki Clinic", "Marketing", "Ngày 31-60", "Biên VIP tăng lên ≥31,5%"],
        ["GP3. Cảnh báo sớm SLA 48h", "Bật cờ cảnh báo đơn quá 48h; mở kho đệm trung chuyển tại Đà Nẵng", "Operations", "Ngày 1-30", "Đơn trễ giảm xuống <5%"],
        ["GP4. Chuẩn hóa đóng gói 3 lớp", "Dùng màng bọc khí 3 lớp, hộp cứng dán nhãn 'Hàng dễ vỡ' cho Nước hoa", "Logistics", "Ngày 31-60", "Tiết kiệm +120 triệu ₫ vốn"],
        ["GP5. Kéo khách từ sàn sang App", "Kẹp voucher độc quyền trong gói hàng TMĐT để dẫn khách tải Hasaki App", "E-Commerce", "Ngày 31-60", "15% khách sàn sang App"],
        ["GP6. Bán Combo quà tặng", "Thiết kế Combo Skincare + Nước hoa để tăng giá trị giỏ hàng thay vì giảm giá", "Sales", "Ngày 61-90", "AOV tăng lên ≥1,85 tr ₫"],
        ["GP7. Đo lường chi phí ẩn", "Kết nối API hoa hồng sàn vào Data Warehouse; chạy thử nghiệm A/B giá", "BI & IT", "Ngày 61-90", "Chuẩn hóa 100% P&L ròng"]
    ],
    [Cm(3.0), Cm(5.0), Cm(2.2), Cm(2.0), Cm(3.8)],
    [WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT]
)

# Save
doc.save(OUT_LOCAL)
doc.save(OUT_DL)
print(f"SUCCESS! Concise report saved at:\n  - {OUT_LOCAL}\n  - {OUT_DL}")
print(f"File size: {os.path.getsize(OUT_LOCAL):,} bytes")
