import docx
import re
import sys
import os

sys.stdout.reconfigure(encoding='utf-8')
doc = docx.Document(r'C:\Users\HP\Downloads\Bao_cao_BI301_Hasaki_Pham_Thuy_Trang_B25DTQT066.docx')
rels = doc.part.rels

for i, p in enumerate(doc.paragraphs):
    if 'blip' in p._element.xml:
        embeds = re.findall(r'r:embed="([^"]+)"', p._element.xml)
        for emb in embeds:
            target = rels[emb].target_ref
            caption = ''
            for offset in [1, 2, -1, -2]:
                idx = i + offset
                if 0 <= idx < len(doc.paragraphs):
                    t = doc.paragraphs[idx].text.strip()
                    if t.startswith('Hình') or 'Hình' in t:
                        caption = t
                        break
            print(f'Paragraph {i} | Image: {target} | Caption: {caption}')
