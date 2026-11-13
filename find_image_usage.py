import docx
import sys

sys.stdout.reconfigure(encoding='utf-8')
doc = docx.Document(r'C:\Users\HP\Downloads\Bao_cao_BI301_Hasaki_Pham_Thuy_Trang_B25DTQT066.docx')

rels = doc.part.rels
for r_id, rel in rels.items():
    if 'image' in rel.target_ref:
        print(f'{r_id} -> {rel.target_ref}')

# check if any paragraph or table references them
xml_str = doc._element.xml
for r_id in rels:
    if r_id in xml_str and 'image' in rels[r_id].target_ref:
        count = xml_str.count(r_id)
        print(f'{r_id} ({rels[r_id].target_ref}) used {count} times')
