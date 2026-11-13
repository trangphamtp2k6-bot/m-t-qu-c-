import os
import sys
import docx
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

sys.stdout.reconfigure(encoding='utf-8')

# Colors
COLOR_PRIMARY = RGBColor(27, 54, 93)     # Deep Navy #1B365D
COLOR_SECONDARY = RGBColor(35, 73, 117)  # Slate Blue #234975
COLOR_TEXT = RGBColor(30, 30, 30)        # Dark Charcoal
COLOR_MUTED = RGBColor(100, 100, 100)    # Gray
HEX_PRIMARY = "1B365D"
HEX_LIGHT_BG = "F4F7FB"
HEX_HEADER_BG = "1B365D"
HEX_ZEBRA = "F9FBFC"
HEX_BORDER = "D0D7DE"

def set_cell_background(cell, hex_color):
    shading_elm = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
    cell._tc.get_or_add_tcPr().append(shading_elm)

def set_cell_margins(cell, top=120, bottom=120, left=150, right=150):
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

def set_table_borders(table, color="D0D7DE", sz="4", val="single"):
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

def format_run(run, font_name="Times New Roman", size_pt=13, bold=False, italic=False, color=COLOR_TEXT):
    run.font.name = font_name
    run.font.size = Pt(size_pt)
    run.bold = bold
    run.italic = italic
    run.font.color.rgb = color

def add_heading_1(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(14)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.keep_with_next = True
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run(text)
    format_run(run, size_pt=15, bold=True, color=COLOR_PRIMARY)
    return p

def add_heading_2(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.keep_with_next = True
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run(text)
    format_run(run, size_pt=13.5, bold=True, color=COLOR_SECONDARY)
    return p

def add_heading_3(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.keep_with_next = True
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run(text)
    format_run(run, size_pt=13, bold=True, italic=True, color=COLOR_TEXT)
    return p

def add_body(doc, text, bold_prefix=None, indent=True):
    p = doc.add_paragraph()
    p.paragraph_format.line_spacing = 1.4
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    if indent:
        p.paragraph_format.first_line_indent = Cm(1.0)
    if bold_prefix:
        r_pre = p.add_run(bold_prefix)
        format_run(r_pre, size_pt=13, bold=True, color=COLOR_TEXT)
    r_body = p.add_run(text)
    format_run(r_body, size_pt=13, bold=False, color=COLOR_TEXT)
    return p

def add_bullet(doc, text, bold_prefix=None):
    p = doc.add_paragraph(style='List Bullet')
    p.paragraph_format.line_spacing = 1.3
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    if bold_prefix:
        r_pre = p.add_run(bold_prefix)
        format_run(r_pre, size_pt=13, bold=True, color=COLOR_TEXT)
    r_body = p.add_run(text)
    format_run(r_body, size_pt=13, bold=False, color=COLOR_TEXT)
    return p

def add_callout(doc, title, items):
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    cell = table.cell(0, 0)
    cell.width = Cm(16.0)
    set_cell_background(cell, HEX_LIGHT_BG)
    set_cell_margins(cell, top=140, bottom=140, left=200, right=200)
    
    # Left border accent
    tcPr = cell._tc.get_or_add_tcPr()
    borders = parse_xml(f'''
        <w:tcBorders {nsdecls("w")}>
            <w:left w:val="single" w:sz="24" w:space="0" w:color="{HEX_PRIMARY}"/>
            <w:top w:val="none"/>
            <w:right w:val="none"/>
            <w:bottom w:val="none"/>
        </w:tcBorders>
    ''')
    tcPr.append(borders)
    
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(4)
    r_title = p.add_run(title)
    format_run(r_title, size_pt=12.5, bold=True, color=COLOR_PRIMARY)
    
    for it in items:
        p2 = cell.add_paragraph()
        p2.paragraph_format.space_before = Pt(2)
        p2.paragraph_format.space_after = Pt(2)
        p2.paragraph_format.line_spacing = 1.25
        p2.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        if isinstance(it, tuple):
            r_b = p2.add_run(f"• {it[0]}: ")
            format_run(r_b, size_pt=11.5, bold=True, color=COLOR_TEXT)
            r_t = p2.add_run(it[1])
            format_run(r_t, size_pt=11.5, bold=False, color=COLOR_TEXT)
        else:
            r_t = p2.add_run(f"• {it}")
            format_run(r_t, size_pt=11.5, bold=False, color=COLOR_TEXT)
    
    p_spacer = doc.add_paragraph()
    p_spacer.paragraph_format.space_after = Pt(4)

def add_figure(doc, img_path, fig_title, source_note, width=Cm(15.0)):
    if not os.path.exists(img_path):
        print(f"Warning: Image not found: {img_path}")
        return
    p_img = doc.add_paragraph()
    p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_img.paragraph_format.space_before = Pt(8)
    p_img.paragraph_format.space_after = Pt(4)
    p_img.paragraph_format.keep_with_next = True
    run_img = p_img.add_run()
    run_img.add_picture(img_path, width=width)
    
    # Figure caption (Below image, centered, bold italic)
    p_cap = doc.add_paragraph()
    p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_cap.paragraph_format.space_before = Pt(2)
    p_cap.paragraph_format.space_after = Pt(2)
    p_cap.paragraph_format.keep_with_next = True
    r_cap = p_cap.add_run(fig_title)
    format_run(r_cap, size_pt=11, bold=True, italic=True, color=COLOR_TEXT)
    
    # Source note
    p_src = doc.add_paragraph()
    p_src.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_src.paragraph_format.space_before = Pt(0)
    p_src.paragraph_format.space_after = Pt(8)
    r_src = p_src.add_run(f"({source_note})")
    format_run(r_src, size_pt=10, bold=False, italic=True, color=COLOR_MUTED)

def add_table(doc, table_title, headers, rows_data, col_widths, alignments, source_note=None):
    # Table caption (Above table, centered, bold)
    p_cap = doc.add_paragraph()
    p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_cap.paragraph_format.space_before = Pt(10)
    p_cap.paragraph_format.space_after = Pt(4)
    p_cap.paragraph_format.keep_with_next = True
    r_cap = p_cap.add_run(table_title)
    format_run(r_cap, size_pt=11.5, bold=True, color=COLOR_PRIMARY)
    
    table = doc.add_table(rows=len(rows_data) + 1, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    set_table_borders(table, color="B0C0D0", sz="6", val="single")
    
    # Header row
    hdr_cells = table.rows[0].cells
    for i, h_text in enumerate(headers):
        hdr_cells[i].width = col_widths[i]
        set_cell_background(hdr_cells[i], HEX_HEADER_BG)
        set_cell_margins(hdr_cells[i], top=120, bottom=120, left=120, right=120)
        p = hdr_cells[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(2)
        r = p.add_run(h_text)
        format_run(r, size_pt=11, bold=True, color=RGBColor(255, 255, 255))
        
    # Repeat header row on new pages
    trPr = table.rows[0]._tr.get_or_add_trPr()
    trPr.append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))
    
    # Body rows
    for r_idx, row in enumerate(rows_data):
        row_cells = table.rows[r_idx + 1].cells
        bg_color = HEX_ZEBRA if r_idx % 2 == 1 else "FFFFFF"
        for c_idx, val in enumerate(row):
            row_cells[c_idx].width = col_widths[c_idx]
            if bg_color != "FFFFFF":
                set_cell_background(row_cells[c_idx], bg_color)
            set_cell_margins(row_cells[c_idx], top=80, bottom=80, left=120, right=120)
            p = row_cells[c_idx].paragraphs[0]
            p.alignment = alignments[c_idx]
            p.paragraph_format.space_before = Pt(1)
            p.paragraph_format.space_after = Pt(1)
            p.paragraph_format.line_spacing = 1.15
            r = p.add_run(str(val))
            format_run(r, size_pt=10.5, bold=False, color=COLOR_TEXT)
            
    # Source note below table
    if source_note:
        p_src = doc.add_paragraph()
        p_src.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p_src.paragraph_format.space_before = Pt(2)
        p_src.paragraph_format.space_after = Pt(8)
        r_src = p_src.add_run(f"Nguồn: {source_note}")
        format_run(r_src, size_pt=10, bold=False, italic=True, color=COLOR_MUTED)

print("Helper functions ready!")

