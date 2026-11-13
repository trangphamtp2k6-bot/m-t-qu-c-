import docx
import zipfile
import os

docx_path = r'C:\Users\HP\Downloads\Bao_cao_BI301_Hasaki_Pham_Thuy_Trang_B25DTQT066.docx'
with zipfile.ZipFile(docx_path, 'r') as z:
    for f in z.namelist():
        if f.startswith('word/media/'):
            print(f, z.getinfo(f).file_size)
