import pandas as pd
import sys

sys.stdout.reconfigure(encoding='utf-8')

# Read lines starting from the 100 leads header
with open('namviet_leads_raw.csv', 'r', encoding='utf-8') as f:
    lines = f.readlines()

# Find start of 100 leads table
start_idx = 0
for i, line in enumerate(lines):
    if line.startswith('Mã Hồ Sơ,'):
        start_idx = i
        break

df = pd.read_csv('namviet_leads_raw.csv', skiprows=start_idx)
print("Total rows loaded:", len(df))
print("\n--- Trạng Thái Xử Lý ---")
print(df['Trạng Thái Xử Lý'].value_counts())

print("\n--- Phân bổ NVKD ---")
print(df['NVKD Phụ Trách'].value_counts())

print("\n--- Kiểm tra Ngân sách > 500 triệu (VIP Rule > 500M) ---")
df['Ngân Sách (VNĐ)'] = pd.to_numeric(df['Ngân Sách (VNĐ)'], errors='coerce').fillna(0)
vip_leads = df[df['Ngân Sách (VNĐ)'] > 500000000]
print(vip_leads[['Mã Hồ Sơ', 'Tên Doanh Nghiệp Chuẩn Hóa', 'Ngân Sách (VNĐ)', 'Gói Sản Phẩm', 'Trạng Thái Xử Lý']])

print("\n--- Kiểm tra Ngân sách = 500 triệu ---")
print(df[df['Ngân Sách (VNĐ)'] == 500000000][['Mã Hồ Sơ', 'Tên Doanh Nghiệp Chuẩn Hóa', 'Ngân Sách (VNĐ)', 'Gói Sản Phẩm', 'Trạng Thái Xử Lý']])

print("\n--- Kiểm tra tổng ngân sách theo trạng thái ---")
print(df.groupby('Trạng Thái Xử Lý')['Ngân Sách (VNĐ)'].agg(['count', 'sum', 'mean']))

print("\n--- Kiểm tra các hồ sơ có sự cố lịch sử / trùng / ngoại lệ ---")
exceptions = df[df['Trạng Thái Xử Lý'] != 'ĐÃ PHÂN CÔNG']
print(exceptions[['Mã Hồ Sơ', 'Tên Doanh Nghiệp Chuẩn Hóa', 'Ngân Sách (VNĐ)', 'Trạng Thái Xử Lý', 'Hành Động Đề Xuất']].to_string())

print("\n--- Kiểm tra phân bổ khu vực & NVKD có bất hợp lý? ---")
assigned = df[df['Trạng Thái Xử Lý'] == 'ĐÃ PHÂN CÔNG']
print("\nChi tiết NV07 & NV08 (Đa ngành):")
print(assigned[assigned['NVKD Phụ Trách'].str.contains('NV07|NV08')][['Mã Hồ Sơ', 'Tỉnh / Thành Phố', 'Nhóm Ngành Chuẩn Hóa', 'Gói Sản Phẩm', 'NVKD Phụ Trách']])

