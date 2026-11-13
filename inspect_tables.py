import docx
import sys

sys.stdout.reconfigure(encoding='utf-8')
doc = docx.Document(r'C:\Users\HP\Downloads\Bao_cao_BI301_Hasaki_Pham_Thuy_Trang_B25DTQT066.docx')

print(f'Total tables: {len(doc.tables)}')
for i, t in enumerate(doc.tables):
    rows = len(t.rows)
    cols = len(t.columns)
    header = [cell.text.replace('\n', ' ').strip() for cell in t.rows[0].cells]
    print(f'Table {i+1} ({rows}x{cols}): {header[:5]}')
