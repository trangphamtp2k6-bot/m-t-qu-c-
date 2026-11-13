from PIL import Image
import os
import sys

sys.stdout.reconfigure(encoding='utf-8')
dir_path = r'C:\Users\HP\Downloads\extracted_hasaki_images'

for f in sorted(os.listdir(dir_path)):
    if f.endswith('.png'):
        p = os.path.join(dir_path, f)
        im = Image.open(p)
        print(f'{f}: size={im.size}, mode={im.mode}, file_size={os.path.getsize(p)} bytes')
