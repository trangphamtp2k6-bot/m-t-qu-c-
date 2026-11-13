import pandas as pd
import json
import sys

sys.stdout.reconfigure(encoding='utf-8')

with open('namviet_leads_raw.csv', 'r', encoding='utf-8') as f:
    lines = f.readlines()

start_idx = 0
for i, line in enumerate(lines):
    if line.startswith('Mã Hồ Sơ,'):
        start_idx = i
        break

df = pd.read_csv('namviet_leads_raw.csv', skiprows=start_idx, nrows=100)
df['Ngân Sách (VNĐ)'] = pd.to_numeric(df['Ngân Sách (VNĐ)'], errors='coerce').fillna(0).astype(int)
df = df.fillna('')

# Save clean JSON for HTML builder
records = df.to_dict(orient='records')
with open('namviet_100_clean.json', 'w', encoding='utf-8') as f:
    json.dump(records, f, ensure_ascii=False, indent=2)

print("Saved 100 clean JSON records.")
print("Total budget all 100:", df['Ngân Sách (VNĐ)'].sum())
print("Total budget 80 assigned:", df[df['Trạng Thái Xử Lý'] == 'ĐÃ PHÂN CÔNG']['Ngân Sách (VNĐ)'].sum())
print("Count by Priority in 80 assigned:")
print(df[df['Trạng Thái Xử Lý'] == 'ĐÃ PHÂN CÔNG']['Mức Ưu Tiên'].value_counts())

