import zipfile
import xml.etree.ElementTree as ET
import sys
import os

sys.stdout.reconfigure(encoding='utf-8')
pptx_path = r'C:\Users\HP\Downloads\Hasaki BI301 – Báo cáo phân tích.pptx'

with zipfile.ZipFile(pptx_path, 'r') as z:
    for name in z.namelist():
        if name.startswith('ppt/slides/slide') and name.endswith('.xml'):
            content = z.read(name).decode('utf-8')
            # extract text
            texts = re.findall(r'<a:t>([^<]+)</a:t>', content) if 're' in globals() else []
            import re
            texts = re.findall(r'<a:t>([^<]+)</a:t>', content)
            print(f'=== {name} ===')
            print(' '.join(texts[:10]))
