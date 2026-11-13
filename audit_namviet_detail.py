import pandas as pd
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
df['Ngân Sách (VNĐ)'] = pd.to_numeric(df['Ngân Sách (VNĐ)'], errors='coerce').fillna(0)

assigned = df[df['Trạng Thái Xử Lý'] == 'ĐÃ PHÂN CÔNG'].copy()

print("=== TẢI TRỌNG VÀ NGÂN SÁCH THEO 8 NVKD ===")
summary = assigned.groupby('NVKD Phụ Trách').agg(
    Total_Leads=('Mã Hồ Sơ', 'count'),
    High=('Mức Ưu Tiên', lambda x: (x == '1 - CAO').sum()),
    Med=('Mức Ưu Tiên', lambda x: (x == '2 - TRUNG BÌNH').sum()),
    Low=('Mức Ưu Tiên', lambda x: (x == '3 - THẤP').sum()),
    Total_Budget=('Ngân Sách (VNĐ)', 'sum'),
    Avg_Budget=('Ngân Sách (VNĐ)', 'mean')
).reset_index()
print(summary.to_string())

print("\n=== TỈ LỆ PHÂN BỔ NGÀNH HÀNG & GÓI SẢN PHẨM TRONG 80 LEADS ===")
print(assigned.groupby('Gói Sản Phẩm')['Ngân Sách (VNĐ)'].agg(['count', 'sum', 'mean']))

print("\n=== KIỂM TRA LỖI ĐỊNH TUYẾN (ROUTING ANOMALIES) ===")
# Check if any NVKD got leads outside their specified scope:
# NV01: PROD_RETAIL - Miền Bắc
# NV02: PROD_RETAIL - Miền Nam
# NV03: PROD_FNB - Miền Bắc & Miền Trung
# NV04: PROD_FNB - Miền Nam
# NV05: PROD_SERVICE - Toàn quốc
# NV06: PROD_MFG - Toàn quốc
# NV07: Đa ngành & Mới - Miền Bắc
# NV08: Đa ngành & Mới - Miền Nam

for idx, row in assigned.iterrows():
    nv = row['NVKD Phụ Trách']
    pkg = row['Gói Sản Phẩm']
    city = row['Tỉnh / Thành Phố']
    ind = row['Nhóm Ngành Chuẩn Hóa']
    # Check mismatch
    if 'NV01' in nv and pkg != 'PROD_RETAIL':
        print(f"Mismatch Pkg NV01: {row['Mã Hồ Sơ']} - {pkg}")
    if 'NV02' in nv and pkg != 'PROD_RETAIL':
        print(f"Mismatch Pkg NV02: {row['Mã Hồ Sơ']} - {pkg}")
    if 'NV07' in nv and ind != 'Doanh nghiệp Đa ngành / Khác':
        print(f"Note NV07 (Đa ngành Miền Bắc) received: {row['Mã Hồ Sơ']} | City: {city} | Ind: {ind} | Pkg: {pkg}")
    if 'NV08' in nv and ind != 'Doanh nghiệp Đa ngành / Khác':
        print(f"Note NV08 (Đa ngành Miền Nam) received: {row['Mã Hồ Sơ']} | City: {city} | Ind: {ind} | Pkg: {pkg}")
    if ('NV07' in nv or 'NV08' in nv) and city in ['Khánh Hòa', 'Quảng Ngãi', 'Thừa Thiên Huế']:
        print(f"Region Check {nv}: {row['Mã Hồ Sơ']} is in Central VN ({city}) | Ind: {ind} | Pkg: {pkg}")
    if 'NV05' in nv and city in ['Hà Nội', 'TP. Hồ Chí Minh', 'Bắc Ninh'] and ind == 'Doanh nghiệp Đa ngành / Khác':
        print(f"NV05 anomaly: {row['Mã Hồ Sơ']}")
    if 'NV08' in nv and city == 'Thừa Thiên Huế':
        print(f"Region Anomaly NV08 (Miền Nam) got Huế: {row['Mã Hồ Sơ']}")

