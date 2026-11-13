import matplotlib.pyplot as plt
import numpy as np
import json
import sys

sys.stdout.reconfigure(encoding='utf-8')
plt.rcParams['font.sans-serif'] = 'Segoe UI'
plt.rcParams['axes.unicode_minus'] = False

# Load clean data
with open('namviet_100_clean.json', 'r', encoding='utf-8') as f:
    records = json.load(f)

# Chart 1: Funnel Breakdown (Counts & Budgets)
fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(11, 4.5))

labels = ['Bàn giao 8 NVKD\n(Hợp lệ)', 'Từ chối\nNgoài phạm vi', 'Đã gộp\nTrùng lặp', 'Cần TPKD\nDuyệt / Bổ sung']
counts = [80, 7, 7, 6]
budgets_ty = [3.033, 3.100, 0.335, 0.297]
colors = ['#1B365D', '#DC2626', '#64748B', '#D97706']

bars1 = ax1.bar(labels, counts, color=colors, width=0.55)
ax1.set_title('Phân loại 100 Hồ sơ Leads Đầu vào (Số lượng)', fontsize=11, fontweight='bold', color='#1B365D')
ax1.set_ylabel('Số lượng hồ sơ (Leads)', fontsize=10)
ax1.set_ylim(0, 95)
ax1.grid(axis='y', linestyle='--', alpha=0.4)
for bar in bars1:
    yval = bar.get_height()
    ax1.text(bar.get_x() + bar.get_width()/2, yval + 2, f'{int(yval)} ({int(yval)}%)', ha='center', va='bottom', fontsize=9.5, fontweight='bold')

bars2 = ax2.bar(labels, budgets_ty, color=colors, width=0.55)
ax2.set_title('Giá trị Ngân sách theo Nhóm (Tỷ VNĐ)', fontsize=11, fontweight='bold', color='#1B365D')
ax2.set_ylabel('Tổng ngân sách (Tỷ VNĐ)', fontsize=10)
ax2.set_ylim(0, 3.8)
ax2.grid(axis='y', linestyle='--', alpha=0.4)
for i, bar in enumerate(bars2):
    yval = bar.get_height()
    lbl = f'{yval:.2f} tỷ'
    if i == 1:
        lbl += '\n(Lọt 1.4 tỷ VIP!)'
    ax2.text(bar.get_x() + bar.get_width()/2, yval + 0.08, lbl, ha='center', va='bottom', fontsize=9, fontweight='bold', color='#DC2626' if i==1 else '#1E293B')

plt.tight_layout()
plt.savefig('chart_namviet_funnel.png', dpi=200)
plt.close()

# Chart 2: Workload Forecast 700 leads/week vs 30-lead Cap
nv_names = ['NV01\nAn', 'NV06\nKiệt', 'NV03\nCường', 'NV07\nLan', 'NV05\nHoa', 'NV04\nDũng', 'NV02\nBình', 'NV08\nNam']
sample_100 = [17, 15, 14, 10, 8, 7, 6, 3]
scale_700 = [x * 7 for x in sample_100]

fig, ax = plt.subplots(figsize=(10, 4.8))
x = np.arange(len(nv_names))
bar_colors = ['#DC2626' if v > 60 else ('#F59E0B' if v > 30 else '#059669') for v in scale_700]

bars = ax.bar(x, scale_700, color=bar_colors, width=0.55, label='Dự báo tải trọng (700 leads/tuần)')
ax.axhline(30, color='#1B365D', linestyle='--', linewidth=2, label='Ngưỡng trần Backlog cho phép (30 leads)')

ax.set_title('Mất Cân Bằng Tải Trọng & Dự Báo Vỡ Trần SLA Khi Lên Quy Mô 700 Leads/Tuần', fontsize=12, fontweight='bold', color='#1B365D')
ax.set_ylabel('Số lượng Leads mỗi nhân viên/tuần', fontsize=10)
ax.set_xticks(x)
ax.set_xticklabels(nv_names, fontsize=9.5, fontweight='bold')
ax.set_ylim(0, 140)
ax.grid(axis='y', linestyle='--', alpha=0.4)
ax.legend(loc='upper right', fontsize=9.5)

for i, bar in enumerate(bars):
    yval = bar.get_height()
    ax.text(bar.get_x() + bar.get_width()/2, yval + 3, f'{yval} leads\n(Mẫu: {sample_100[i]})', ha='center', va='bottom', fontsize=8.5, fontweight='bold')

plt.tight_layout()
plt.savefig('chart_namviet_workload.png', dpi=200)
plt.close()

print("Charts generated successfully!")

