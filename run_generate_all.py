import os
import sys
import docx
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

from docx_helpers import (
    COLOR_PRIMARY, COLOR_SECONDARY, COLOR_TEXT, COLOR_MUTED,
    HEX_PRIMARY, HEX_LIGHT_BG, HEX_HEADER_BG, HEX_ZEBRA, HEX_BORDER,
    set_cell_background, set_cell_margins, set_table_borders,
    format_run
)

from generate_full_report import build_cover, doc, section, style_normal
from section_frontmatter import add_frontmatter
from section_intro_ch1 import add_intro_and_chapter1
from section_ch2_ch3 import add_chapter2_and_3
from section_ch4_ch5 import add_chapter4_and_5
from section_ch6_insights import add_chapter6
from section_ch7_end import add_chapter7_and_conclusion

sys.stdout.reconfigure(encoding='utf-8')

print("Starting full assembly of Business Intelligence Master Document...")

# Assemble sections
print("Adding Frontmatter...")
add_frontmatter(doc)

print("Adding Intro and Chapter 1...")
add_intro_and_chapter1(doc)

print("Adding Chapter 2 & Chapter 3...")
add_chapter2_and_3(doc)

print("Adding Chapter 4 & Chapter 5...")
add_chapter4_and_5(doc)

print("Adding Chapter 6 (Data Analysis & Deep Insights)...")
add_chapter6(doc)

print("Adding Chapter 7, Conclusion, References & Appendix...")
add_chapter7_and_conclusion(doc)

# Save document
out_local = r'c:\Users\HP\Downloads\MÔN AI AGENT\Bao_Cao_Chuyen_Gia_BI_Hasaki_Khoa_Hoc_Chuan.docx'
out_downloads = r'C:\Users\HP\Downloads\Bao_Cao_Chuyen_Gia_BI_Hasaki_Khoa_Hoc_Chuan.docx'

print("Saving local file...")
doc.save(out_local)

print("Saving Downloads file...")
doc.save(out_downloads)

print(f"SUCCESS! Master Report saved at:\n  - {out_local}\n  - {out_downloads}")
print(f"File size: {os.path.getsize(out_local):,} bytes")

