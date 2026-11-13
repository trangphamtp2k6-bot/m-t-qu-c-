import os
import sys
import docx
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

from docx_helpers import (
    COLOR_PRIMARY, COLOR_SECONDARY, COLOR_TEXT, COLOR_MUTED,
    HEX_PRIMARY, HEX_LIGHT_BG, HEX_HEADER_BG, HEX_ZEBRA, HEX_BORDER,
    set_cell_background, set_cell_margins, set_table_borders,
    format_run, add_heading_1, add_heading_2, add_heading_3,
    add_body, add_bullet, add_callout, add_figure, add_table
)

sys.stdout.reconfigure(encoding='utf-8')

IMG_DIR = r'C:\Users\HP\Downloads\extracted_hasaki_images'
OUTPUT_DOCX = r'C:\Users\HP\Downloads\Bao_Cao_Chuyen_Gia_BI_Hasaki_Khoa_Hoc_Chuan.docx'
OUTPUT_DOCX_LOCAL = r'c:\Users\HP\Downloads\MÔN AI AGENT\Bao_Cao_Chuyen_Gia_BI_Hasaki_Khoa_Hoc_Chuan.docx'

print("Starting document creation...")
doc = docx.Document()

# ==================== PAGE SETUP ====================
section = doc.sections[0]
section.page_width = Cm(21.0)
section.page_height = Cm(29.7)
section.top_margin = Cm(2.0)
section.bottom_margin = Cm(2.0)
section.left_margin = Cm(3.0)
section.right_margin = Cm(2.0)

# Configure default Normal style
style_normal = doc.styles['Normal']
style_normal.font.name = 'Times New Roman'
style_normal.font.size = Pt(13)
style_normal.font.color.rgb = COLOR_TEXT

# Helper for cover border
def add_cover_page(doc):
    # Cover page container table with border
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    cell = table.cell(0, 0)
    cell.width = Cm(16.0)
    set_cell_margins(cell, top=200, bottom=200, left=240, right=240)
    
    # Outer double border
    tcPr = cell._tc.get_or_add_tcPr()
    borders = parse_xml(f'''
        <w:tcBorders {nsdecls("w")}>
            <w:top w:val="double" w:sz="18" w:space="0" w:color="{HEX_PRIMARY}"/>
            <w:left w:val="double" w:sz="18" w:space="0" w:color="{HEX_PRIMARY}"/>
            <w:bottom w:val="double" w:sz="18" w:space="0" w:color="{HEX_PRIMARY}"/>
            <w:right w:val="double" w:sz="18" w:space="0" w:color="{HEX_PRIMARY}"/>
        </w:tcBorders>
    ''')
    tcPr.append(borders)
    
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(2)
    r = p.add_run("BỘ THÔNG TIN VÀ TRUYỀN THÔNG\nHỌC VIỆN CÔNG NGHỆ BƯU CHÍNH VIỄN THÔNG")
    format_run(r, size_pt=12.5, bold=True, color=COLOR_PRIMARY)
    
    p_sub = cell.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sub.paragraph_format.space_before = Pt(2)
    p_sub.paragraph_format.space_after = Pt(14)
    r_sub = p_sub.add_run("KHOA QUẢN TRỊ KINH DOANH - VIỆN ĐÀO TẠO SỐ")
    format_run(r_sub, size_pt=11.5, bold=True, color=COLOR_SECONDARY)
    
    # Add logos
    p_logo = cell.add_paragraph()
    p_logo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_logo.paragraph_format.space_before = Pt(4)
    p_logo.paragraph_format.space_after = Pt(16)
    r_l1 = p_logo.add_run()
    img1 = os.path.join(IMG_DIR, 'image1.png')
    img2 = os.path.join(IMG_DIR, 'image2.png')
    if os.path.exists(img1):
        r_l1.add_picture(img1, width=Cm(2.5))
    r_sp = p_logo.add_run("       ")
    r_l2 = p_logo.add_run()
    if os.path.exists(img2):
        r_l2.add_picture(img2, width=Cm(3.8))
        
    p_rep = cell.add_paragraph()
    p_rep.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_rep.paragraph_format.space_before = Pt(14)
    p_rep.paragraph_format.space_after = Pt(6)
    r_rep = p_rep.add_run("BÁO CÁO DỰ ÁN KHOA HỌC CHUYÊN SÂU\nMÔN HỌC: BUSINESS INTELLIGENCE (BI301)")
    format_run(r_rep, size_pt=13.5, bold=True, color=COLOR_PRIMARY)
    
    # Divider line
    p_div = cell.add_paragraph()
    p_div.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_div.paragraph_format.space_before = Pt(2)
    p_div.paragraph_format.space_after = Pt(10)
    r_div = p_div.add_run("—" * 28)
    format_run(r_div, size_pt=12, bold=True, color=COLOR_SECONDARY)
    
    p_title = cell.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_before = Pt(8)
    p_title.paragraph_format.space_after = Pt(8)
    r_title = p_title.add_run("ĐỀ TÀI NGHIÊN CỨU & ỨNG DỤNG:\nXÂY DỰNG HỆ THỐNG BUSINESS INTELLIGENCE TỐI ƯU HÓA CHIẾN LƯỢC KHUYẾN MÃI, KÊNH PHÂN PHỐI VÀ QUẢN TRỊ LỢI NHUẬN DANH MỤC SKINCARE TẠI CHUỖI BÁN LẺ HASAKI BEAUTY & CLINIC")
    format_run(r_title, size_pt=15.5, bold=True, color=COLOR_PRIMARY)
    
    p_div2 = cell.add_paragraph()
    p_div2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_div2.paragraph_format.space_before = Pt(4)
    p_div2.paragraph_format.space_after = Pt(24)
    r_div2 = p_div2.add_run("—" * 28)
    format_run(r_div2, size_pt=12, bold=True, color=COLOR_SECONDARY)
    
    # Meta info table inside cover
    tbl_meta = cell.add_table(rows=4, cols=2)
    tbl_meta.alignment = WD_TABLE_ALIGNMENT.CENTER
    col_w = [Cm(5.5), Cm(8.5)]
    meta_data = [
        ("Giảng viên hướng dẫn:", "ThS. Hoàng Thị Kim Oanh"),
        ("Sinh viên thực hiện:", "Phạm Thuỳ Trang"),
        ("Mã sinh viên:", "B25DTQT066"),
        ("Chuyên ngành:", "Quản trị Kinh doanh Số (Digital Business)")
    ]
    for idx, (lbl, val) in enumerate(meta_data):
        c0, c1 = tbl_meta.rows[idx].cells
        c0.width, c1.width = col_w
        set_cell_margins(c0, top=40, bottom=40, left=80, right=80)
        set_cell_margins(c1, top=40, bottom=40, left=80, right=80)
        p0 = c0.paragraphs[0]
        p0.alignment = WD_ALIGN_PARAGRAPH.LEFT
        r0 = p0.add_run(lbl)
        format_run(r0, size_pt=12, bold=True, color=COLOR_TEXT)
        p1 = c1.paragraphs[0]
        p1.alignment = WD_ALIGN_PARAGRAPH.LEFT
        r1 = p1.add_run(val)
        format_run(r1, size_pt=12, bold=False, color=COLOR_TEXT)
        
    p_loc = cell.add_paragraph()
    p_loc.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_loc.paragraph_format.space_before = Pt(36)
    p_loc.paragraph_format.space_after = Pt(0)
    r_loc = p_loc.add_run("HÀ NỘI – NĂM 2026")
    format_run(r_loc, size_pt=12, bold=True, color=COLOR_PRIMARY)
    
    doc.add_page_break()

add_cover_page(doc)
print("Cover page created successfully!")

