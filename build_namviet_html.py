import json
import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

with open('namviet_100_clean.json', 'r', encoding='utf-8') as f:
    records = json.load(f)

data_json_str = json.dumps(records, ensure_ascii=False)

html_template = f"""<!DOCTYPE html>
<html lang="vi">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Nam Việt CRM – BI Pipeline Dashboard & AI Agent Control Center</title>
<style>
  :root {{
    --navy: #1B365D;
    --navy-dark: #0F2342;
    --blue: #2563EB;
    --bg: #F1F5F9;
    --card: #FFFFFF;
    --border: #E2E8F0;
    --text: #1E293B;
    --muted: #64748B;
    --red: #DC2626;
    --red-bg: #FEF2F2;
    --amber: #D97706;
    --amber-bg: #FFFBEB;
    --green: #059669;
    --green-bg: #ECFDF5;
  }}
  * {{ box-sizing: border-box; margin: 0; padding: 0; font-family: 'Segoe UI', system-ui, -apple-system, sans-serif; }}
  body {{ background: var(--bg); color: var(--text); line-height: 1.5; }}
  header {{ background: linear-gradient(135deg, var(--navy-dark), var(--navy)); color: #fff; padding: 18px 28px; display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 12px; box-shadow: 0 2px 10px rgba(0,0,0,0.12); }}
  .brand h1 {{ font-size: 20px; font-weight: 700; letter-spacing: 0.3px; }}
  .brand p {{ font-size: 13px; color: #CBD5E1; margin-top: 2px; }}
  .tabs {{ display: flex; gap: 8px; background: rgba(255,255,255,0.12); padding: 4px; border-radius: 10px; }}
  .tab-btn {{ border: none; background: transparent; color: #E2E8F0; padding: 8px 16px; border-radius: 8px; cursor: pointer; font-size: 13.5px; font-weight: 600; transition: all 0.2s; }}
  .tab-btn.active {{ background: #fff; color: var(--navy); box-shadow: 0 2px 6px rgba(0,0,0,0.15); }}
  .container {{ max-width: 1400px; margin: 20px auto; padding: 0 20px; }}
  .tab-panel {{ display: none; }}
  .tab-panel.active {{ display: block; }}

  /* KPI Grid */
  .kpi-grid {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(220px, 1fr)); gap: 16px; margin-bottom: 20px; }}
  .kpi-card {{ background: var(--card); border-radius: 12px; padding: 18px; border: 1px solid var(--border); box-shadow: 0 1px 3px rgba(0,0,0,0.04); border-top: 4px solid var(--navy); }}
  .kpi-card.warn {{ border-top-color: var(--red); }}
  .kpi-card.ok {{ border-top-color: var(--green); }}
  .kpi-title {{ font-size: 12px; font-weight: 700; color: var(--muted); text-transform: uppercase; letter-spacing: 0.5px; }}
  .kpi-val {{ font-size: 28px; font-weight: 800; color: var(--navy); margin: 6px 0 4px; }}
  .kpi-sub {{ font-size: 12.5px; color: var(--muted); }}
  .kpi-sub.danger {{ color: var(--red); font-weight: 600; }}
  .kpi-sub.success {{ color: var(--green); font-weight: 600; }}

  /* Layout Grids */
  .grid-2 {{ display: grid; grid-template-columns: 1.4fr 1fr; gap: 20px; margin-bottom: 20px; }}
  @media(max-width: 960px) {{ .grid-2 {{ grid-template-columns: 1fr; }} }}
  .card {{ background: var(--card); border-radius: 12px; padding: 20px; border: 1px solid var(--border); box-shadow: 0 1px 3px rgba(0,0,0,0.04); margin-bottom: 20px; }}
  .card-title {{ font-size: 15px; font-weight: 700; color: var(--navy); margin-bottom: 14px; display: flex; justify-content: space-between; align-items: center; }}
  .badge {{ display: inline-block; padding: 3px 9px; border-radius: 999px; font-size: 11.5px; font-weight: 700; }}
  .b-red {{ background: var(--red-bg); color: var(--red); }}
  .b-amber {{ background: var(--amber-bg); color: var(--amber); }}
  .b-green {{ background: var(--green-bg); color: var(--green); }}
  .b-blue {{ background: #EFF6FF; color: var(--blue); }}
  .b-gray {{ background: #F1F5F9; color: var(--muted); }}

  /* Bar chart CSS */
  .bar-row {{ margin-bottom: 12px; }}
  .bar-meta {{ display: flex; justify-content: space-between; font-size: 12.5px; margin-bottom: 4px; font-weight: 600; }}
  .bar-track {{ height: 22px; background: #F1F5F9; border-radius: 6px; overflow: hidden; display: flex; position: relative; }}
  .seg-high {{ background: #DC2626; height: 100%; display: flex; align-items: center; justify-content: center; color: #fff; font-size: 11px; font-weight: 700; }}
  .seg-med {{ background: #F59E0B; height: 100%; display: flex; align-items: center; justify-content: center; color: #fff; font-size: 11px; font-weight: 700; }}
  .seg-low {{ background: #64748B; height: 100%; display: flex; align-items: center; justify-content: center; color: #fff; font-size: 11px; font-weight: 700; }}
  .legend {{ display: flex; gap: 16px; font-size: 12px; margin-bottom: 12px; }}
  .legend span {{ display: flex; align-items: center; gap: 5px; }}
  .dot {{ width: 10px; height: 10px; border-radius: 2px; display: inline-block; }}

  /* Table */
  .tbl-wrap {{ overflow-x: auto; }}
  table {{ width: 100%; border-collapse: collapse; font-size: 13px; }}
  th {{ background: #F8FAFC; color: var(--navy); font-weight: 700; text-align: left; padding: 10px 12px; border-bottom: 2px solid var(--border); white-space: nowrap; }}
  td {{ padding: 10px 12px; border-bottom: 1px solid var(--border); vertical-align: middle; }}
  tr:hover td {{ background: #F8FAFC; }}
  .num {{ text-align: right; font-variant-numeric: tabular-nums; font-weight: 600; }}

  /* Controls */
  .controls {{ display: flex; gap: 12px; flex-wrap: wrap; align-items: center; margin-bottom: 16px; background: #F8FAFC; padding: 12px 16px; border-radius: 10px; border: 1px solid var(--border); }}
  select, input {{ padding: 8px 12px; border: 1px solid #CBD5E1; border-radius: 8px; font-size: 13.5px; background: #fff; color: var(--text); }}
  .btn-action {{ padding: 6px 12px; border: none; border-radius: 6px; font-size: 12px; font-weight: 600; cursor: pointer; background: var(--blue); color: #fff; }}
  .btn-action.done {{ background: var(--green); }}
  .alert-box {{ padding: 12px 14px; border-radius: 8px; margin-bottom: 10px; font-size: 13px; border-left: 4px solid; }}
  .alert-red {{ background: var(--red-bg); border-color: var(--red); }}
  .alert-amber {{ background: var(--amber-bg); border-color: var(--amber); }}

  /* AI Pipeline Flow */
  .flow-grid {{ display: grid; grid-template-columns: repeat(4, 1fr); gap: 14px; margin-bottom: 20px; }}
  @media(max-width: 960px) {{ .flow-grid {{ grid-template-columns: 1fr 1fr; }} }}
  .step-box {{ background: #fff; border: 1px solid var(--border); border-radius: 10px; padding: 16px; position: relative; border-top: 4px solid var(--blue); }}
  .step-num {{ font-size: 11px; font-weight: 800; color: var(--blue); text-transform: uppercase; }}
  .step-title {{ font-size: 14.5px; font-weight: 700; color: var(--navy); margin: 4px 0 8px; }}
  .step-stat {{ font-size: 12.5px; color: var(--text); background: #F8FAFC; padding: 8px; border-radius: 6px; margin-top: 8px; }}
</style>
</head>
<body>

<header>
  <div class="brand">
    <h1>NAM VIỆT CRM – BI PIPELINE & AI LEAD ROUTING DASHBOARD</h1>
    <p>Hệ thống Kiểm toán Điều phối 100 Leads Tuần Mẫu (Dự báo quy mô thực tế: 700 Leads/Tuần) • 8 NVKD + TPKD</p>
  </div>
  <div class="tabs">
    <button class="tab-btn active" onclick="switchTab('tab-tpkd', this)">📊 Pipeline Dashboard (TPKD)</button>
    <button class="tab-btn" onclick="switchTab('tab-nvkd', this)">📋 Việc Hàng Ngày (8 NVKD)</button>
    <button class="tab-btn" onclick="switchTab('tab-ai', this)">🤖 Giám Sát 4 Bước AI & Schema</button>
  </div>
</header>

<div class="container">

  <!-- ==================== TAB 1: TPKD DASHBOARD ==================== -->
  <div id="tab-tpkd" class="tab-panel active">
    <div class="kpi-grid">
      <div class="kpi-card">
        <div class="kpi-title">Tổng Lead Đầu Vào</div>
        <div class="kpi-val">100 <span style="font-size:14px;color:#64748B;font-weight:500;">(700/tuần)</span></div>
        <div class="kpi-sub">Tổng ngân sách danh nghĩa: <b>6,765 tỷ ₫</b></div>
      </div>
      <div class="kpi-card ok">
        <div class="kpi-title">Bàn Giao 8 NVKD</div>
        <div class="kpi-val">80 <span style="font-size:14px;color:#059669;font-weight:600;">(80%)</span></div>
        <div class="kpi-sub success">Pipeline khả thi: 3,033 tỷ ₫ (TB 37,9 tr/lead)</div>
      </div>
      <div class="kpi-card">
        <div class="kpi-title">Hồ Sơ Trùng Đã Gộp</div>
        <div class="kpi-val">7 <span style="font-size:14px;color:#64748B;font-weight:500;">(6 Master)</span></div>
        <div class="kpi-sub">Chống gọi đụng độ (Gộp 335 tr ₫ trùng)</div>
      </div>
      <div class="kpi-card warn">
        <div class="kpi-title">Từ Chối Ngoài Phạm Vi</div>
        <div class="kpi-val">7 <span style="font-size:14px;color:#DC2626;font-weight:600;">(3,100 tỷ ₫)</span></div>
        <div class="kpi-sub danger">⚠️ Lọt 2 hồ sơ VIP ≥500tr (1,4 tỷ ₫)</div>
      </div>
      <div class="kpi-card">
        <div class="kpi-title">TPKD Duyệt / Bổ Sung</div>
        <div class="kpi-val">6 <span style="font-size:14px;color:#D97706;font-weight:600;">(297 tr ₫)</span></div>
        <div class="kpi-sub">4 Cần duyệt (1 khiếu nại CSKH) + 2 Thiếu tin</div>
      </div>
    </div>

    <div class="grid-2">
      <!-- Workload Chart -->
      <div class="card">
        <div class="card-title">
          <span>1. Phân Bổ Tải Trọng 8 NVKD & Cảnh Báo Quá Tải (Scale 700 Leads/Tuần)</span>
          <span class="badge b-red">Trần Backlog: 30 Leads</span>
        </div>
        <div class="legend">
          <span><i class="dot" style="background:#DC2626"></i> 1 - Cao (26)</span>
          <span><i class="dot" style="background:#F59E0B"></i> 2 - Trung bình (34)</span>
          <span><i class="dot" style="background:#64748B"></i> 3 - Thấp (20)</span>
        </div>
        <div id="workload-chart"></div>
      </div>

      <!-- BI Audit Warnings -->
      <div class="card">
        <div class="card-title">
          <span>2. Phát Hiện Kiểm Toán BI & Lỗi Cần Xử Lý Ngay</span>
          <span class="badge b-red">5 Điểm Nóng</span>
        </div>
        <div class="alert-box alert-red">
          <b>🚨 Vi phạm Luồng VIP (&gt;500tr):</b> Hồ sơ <b>L044 (Đại Nam Resort - 900 tr)</b> và <b>L034 (FPT Partner - 500 tr)</b> bị AI tự động đóng do "Ngoài SMB". Cần kích hoạt Push Notification lên TPKD/KAM ngay để không mất <b>1,4 tỷ ₫</b>.
        </div>
        <div class="alert-box alert-amber">
          <b>⚖️ Mất cân bằng tải trọng cực đoan:</b> <b>NV06</b> gánh 15 leads (<b>1,005 tỷ ₫</b>, 12 ca Cao) và <b>NV01</b> gánh 17 leads (686 tr), trong khi <b>NV08</b> chỉ có 3 leads (88 tr). Khi lên 700 leads/tuần, NV01 (~119 leads) và NV06 (~105 leads) sẽ vỡ trận SLA nếu không kích hoạt luồng Overflow.
        </div>
        <div class="alert-box alert-amber">
          <b>🗺️ Lỗi Ma trận Vùng Miền Trung (4 hồ sơ):</b> <b>L055</b> (Thanh Hóa), <b>L090</b> (Khánh Hòa), <b>L100</b> (Quảng Ngãi) thuộc nhóm Bán lẻ nhưng giao nhầm cho NV07 (Đa ngành Miền Bắc); <b>L084</b> (Huế) giao cho NV08 (Miền Nam).
        </div>
        <div class="alert-box alert-red">
          <b>⏱️ Thiếu trường Tracking SLA (Bước 4):</b> Bảng CRM hiện tại chưa có cột <code>Timestamp_Giao_Viec</code> và <code>First_Call_Time</code> để kích hoạt cảnh báo 4h và thu hồi quá hạn 24h.
        </div>
      </div>
    </div>

    <!-- Package & SLA Violation Simulation -->
    <div class="grid-2">
      <div class="card">
        <div class="card-title">3. Bảng Xếp Hạng Vi Phạm & Giám Sát SLA (Bước 4 - Mô phỏng Thời gian thực)</div>
        <div class="tbl-wrap">
          <table>
            <thead>
              <tr>
                <th>Nhân Viên</th>
                <th class="num">Lead Giao</th>
                <th class="num">Scale 700/Tuần</th>
                <th class="num">Ngân Sách (VNĐ)</th>
                <th>Trạng Thái Tải (Bước 3)</th>
                <th>Rủi Ro SLA 24h</th>
              </tr>
            </thead>
            <tbody id="sla-rank-body"></tbody>
          </table>
        </div>
      </div>

      <div class="card">
        <div class="card-title">4. Danh Sách Hồ Sơ Ngoại Lệ / VIP Cần TPKD Thẩm Định (12 Hồ Sơ)</div>
        <div class="tbl-wrap" style="max-height:360px;overflow-y:auto;">
          <table>
            <thead>
              <tr>
                <th>Mã</th>
                <th>Doanh Nghiệp</th>
                <th class="num">Ngân Sách</th>
                <th>Cảnh Báo BI</th>
              </tr>
            </thead>
            <tbody id="tpkd-action-body"></tbody>
          </table>
        </div>
      </div>
    </div>
  </div>

  <!-- ==================== TAB 2: NVKD DAILY WORKLIST ==================== -->
  <div id="tab-nvkd" class="tab-panel">
    <div class="card">
      <div class="card-title">
        <span>Danh Sách Công Việc Gọn Gàng Mỗi Sáng (100% Đúng Chuyên Môn • Đã Xếp Thứ Tự Ưu Tiên)</span>
        <span class="badge b-blue" id="nvkd-count-badge">80 Leads Đang Giao</span>
      </div>
      <div class="controls">
        <label><b>Chọn Nhân Viên Kinh Doanh:</b></label>
        <select id="filter-nvkd" onchange="renderWorklist()">
          <option value="ALL">-- Tất cả 8 NVKD (80 Leads) --</option>
          <option value="NV01 - Nguyễn Văn An">NV01 - Nguyễn Văn An (Bán lẻ Miền Bắc - 17)</option>
          <option value="NV02 - Trần Thị Bình">NV02 - Trần Thị Bình (Bán lẻ Miền Nam - 6)</option>
          <option value="NV03 - Lê Hoàng Cường">NV03 - Lê Hoàng Cường (F&B Bắc & Trung - 14)</option>
          <option value="NV04 - Phạm Minh Dũng">NV04 - Phạm Minh Dũng (F&B Miền Nam - 7)</option>
          <option value="NV05 - Vũ Thị Hoa">NV05 - Vũ Thị Hoa (Dịch vụ Toàn quốc - 8)</option>
          <option value="NV06 - Đỗ Tuấn Kiệt">NV06 - Đỗ Tuấn Kiệt (Sản xuất Toàn quốc - 15)</option>
          <option value="NV07 - Bùi Thu Lan">NV07 - Bùi Thu Lan (Đa ngành Miền Bắc - 10)</option>
          <option value="NV08 - Ngô Quang Nam">NV08 - Ngô Quang Nam (Đa ngành Miền Nam - 3)</option>
        </select>
        <select id="filter-prio" onchange="renderWorklist()">
          <option value="ALL">-- Mọi mức ưu tiên --</option>
          <option value="1 - CAO">🔥 1 - CAO</option>
          <option value="2 - TRUNG BÌNH">⚡ 2 - TRUNG BÌNH</option>
          <option value="3 - THẤP">❄️ 3 - THẤP</option>
        </select>
        <input type="text" id="search-box" placeholder="Tìm tên công ty, SĐT..." oninput="renderWorklist()">
      </div>
      <div class="tbl-wrap">
        <table>
          <thead>
            <tr>
              <th>Ưu Tiên</th>
              <th>Mã</th>
              <th>NVKD Phụ Trách</th>
              <th>Doanh Nghiệp & Liên Hệ</th>
              <th>Khu Vực & Gói</th>
              <th class="num">Ngân Sách</th>
              <th>Hành Động Đề Xuất</th>
              <th>Cập Nhật SLA (4h/24h)</th>
            </tr>
          </thead>
          <tbody id="worklist-body"></tbody>
        </table>
      </div>
    </div>
  </div>

  <!-- ==================== TAB 3: AI 4 STEPS & SCHEMA ==================== -->
  <div id="tab-ai" class="tab-panel">
    <div class="flow-grid">
      <div class="step-box">
        <div class="step-num">Bước 1 • History Check</div>
        <div class="step-title">Quét Blacklist & Lịch Sử CSKH</div>
        <p style="font-size:12.5px;color:#475467;">Truy vấn SĐT/Email với CRM cũ để chặn khách từ chối hoặc đẩy lại khách cũ cho NVKD đang chăm sóc.</p>
        <div class="step-stat">
          • <b>7 hồ sơ trùng</b> gộp về 6 Master<br>
          • <b>1 hồ sơ khiếu nại cũ (L013)</b> bị chặn đẩy lên TPKD duyệt<br>
          • <b>0 hồ sơ Blacklist</b> trong tuần mẫu
        </div>
      </div>
      <div class="step-box">
        <div class="step-num">Bước 2 • Smart Routing</div>
        <div class="step-title">Định Tuyến Ma Trận & Luồng VIP</div>
        <p style="font-size:12.5px;color:#475467;">Khớp Khu vực + Ngành nghề với 8 NVKD. Tách luồng VIP (&gt;500tr), Đa ngành và Thiếu thông tin.</p>
        <div class="step-stat">
          • <b>80 hồ sơ</b> định tuyến vào 8 NVKD<br>
          • <b>3 hồ sơ đa ngành</b> & <b>2 thiếu tin</b> tách riêng<br>
          • <span style="color:#DC2626"><b>Lỗi:</b> L044 (900tr) & L034 (500tr) bị đóng thay vì báo VIP!</span>
        </div>
      </div>
      <div class="step-box">
        <div class="step-num">Bước 3 • Workload Balance</div>
        <div class="step-title">Kiểm Soát Tải Trọng (Trần 30)</div>
        <p style="font-size:12.5px;color:#475467;">Đếm Backlog chưa gọi của NVKD. Nếu vượt 30 Leads, kích hoạt luồng tràn (Overflow).</p>
        <div class="step-stat">
          • Tuần mẫu 100: Max NV01 (17 leads)<br>
          • <b>Khi lên 700/tuần:</b> 7/8 NVKD vượt trần 30! NV01 (~119), NV06 (~105)<br>
          • Cần mở Overflow sang NV08 (chỉ 21 leads)
        </div>
      </div>
      <div class="step-box">
        <div class="step-num">Bước 4 • SLA Timer</div>
        <div class="step-title">Đếm Ngược 4h & Thu Hồi 24h</div>
        <p style="font-size:12.5px;color:#475467;">Gắn đồng hồ ngầm khi giao việc. Quá 4h nhắc Zalo/Lark; quá 24h thu hồi về kho và báo đỏ TPKD.</p>
        <div class="step-stat">
          • <b>Khuyến nghị BI:</b> Bổ sung ngay 3 trường vào Lark Base / CRM:<br>
          1. <code>Assigned_At</code> (Datetime)<br>
          2. <code>First_Contact_At</code> (Datetime)<br>
          3. <code>SLA_Status</code> (Formula)
        </div>
      </div>
    </div>

    <div class="card">
      <div class="card-title">Đề Xuất Chuẩn Hóa Quy Tắc Phân Quyền & Sửa Lỗi Ma Trận Cho Công Ty Nam Việt</div>
      <div class="tbl-wrap">
        <table>
          <thead>
            <tr>
              <th>Hạng Mục Lỗi</th>
              <th>Hiện Trạng Trong Dữ Liệu 100 Leads</th>
              <th>Giải Pháp Chuẩn Hóa Của BI</th>
            </tr>
          </thead>
          <tbody>
            <tr>
              <td><b>1. Bẫy "Ngoài phạm vi" với Khách VIP</b></td>
              <td>L044 (900 triệu) và L034 (500 triệu) bị từ chối tự động vì vượt gói SMB.</td>
              <td>Đảo thứ tự điều kiện AI: <b>Kiểm tra Ngân sách &ge; 500 triệu TRƯỚC khi check phạm vi SMB</b>. Chuyển thẳng TPKD/KAM thẩm định.</td>
            </tr>
            <tr>
              <td><b>2. Khoảng trống Miền Trung</b></td>
              <td>L055 (Thanh Hóa), L090 (Khánh Hòa), L100 (Quảng Ngãi) giao cho NV07 Bắc; L084 (Huế) giao cho NV08 Nam.</td>
              <td>Quy hoạch lại ranh giới: Bắc Trung Bộ (Thanh Hóa, Nghệ An, Hà Tĩnh) &rarr; NV Miền Bắc; Đà Nẵng trở vào &rarr; NV Miền Nam hoặc lập nhóm Miền Trung.</td>
            </tr>
            <tr>
              <td><b>3. Lệch tải trọng nặng ở PROD_MFG &amp; RETAIL</b></td>
              <td>NV06 ôm 1,005 tỷ (12 ca Cao) và NV01 ôm 17 ca; trong khi NV08 chỉ nhận 3 ca TB (88 triệu).</td>
              <td>Kích hoạt thuật toán <b>Workload Overflow</b> ở Bước 3: Tự động san sẻ lead mức 2-TB/3-Thấp từ NV01/NV06 sang NV07/NV08 khi Backlog &gt; 30.</td>
            </tr>
            <tr>
              <td><b>4. Thiếu log dữ liệu SLA</b></td>
              <td>Chưa có cột lưu giờ giao và giờ gọi đầu tiên trên bảng CRM.</td>
              <td>Thêm <code>Assigned_At</code>, <code>First_Call_At</code>, <code>Overdue_Count</code> để vẽ bảng xếp hạng vi phạm SLA thực tế trên Lark Base.</td>
            </tr>
          </tbody>
        </table>
      </div>
    </div>
  </div>

</div>

<script>
const rawData = {data_json_str};
const slaCalled = new Set();

function switchTab(tabId, btn) {{
  document.querySelectorAll('.tab-panel').forEach(p => p.classList.remove('active'));
  document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
  document.getElementById(tabId).classList.add('active');
  btn.classList.add('active');
}}

function fmtMoney(val) {{
  if (!val) return '0 ₫';
  if (val >= 1e9) return (val/1e9).toFixed(2).replace('.', ',') + ' tỷ ₫';
  if (val >= 1e6) return (val/1e6).toFixed(0) + ' tr ₫';
  return val.toLocaleString('vi-VN') + ' ₫';
}}

function renderTPKD() {{
  const assigned = rawData.filter(r => r['Trạng Thái Xử Lý'] === 'ĐÃ PHÂN CÔNG');
  const nvMap = {{}};
  assigned.forEach(r => {{
    const k = r['NVKD Phụ Trách'];
    if (!nvMap[k]) nvMap[k] = {{ name: k, high: 0, med: 0, low: 0, total: 0, budget: 0 }};
    nvMap[k].total++;
    nvMap[k].budget += r['Ngân Sách (VNĐ)'];
    if (r['Mức Ưu Tiên'] === '1 - CAO') nvMap[k].high++;
    else if (r['Mức Ưu Tiên'] === '2 - TRUNG BÌNH') nvMap[k].med++;
    else nvMap[k].low++;
  }});

  const nvs = Object.values(nvMap).sort((a,b) => b.total - a.total);
  const maxTotal = Math.max(...nvs.map(x => x.total));

  // Render Stacked Bars
  let chartHtml = '';
  nvs.forEach(nv => {{
    const wHigh = (nv.high / 20) * 100;
    const wMed = (nv.med / 20) * 100;
    const wLow = (nv.low / 20) * 100;
    const scale700 = nv.total * 7;
    const warnTxt = scale700 > 30 ? `<span style="color:#DC2626">Dự báo 700/tuần: ${{scale700}} (Quá tải &gt;30)</span>` : `<span style="color:#059669">Dự báo 700/tuần: ${{scale700}} (An toàn)</span>`;
    chartHtml += `
      <div class="bar-row">
        <div class="bar-meta">
          <span>${{nv.name}} — <b>${{nv.total}} leads</b> (${{fmtMoney(nv.budget)}})</span>
          <span>${{warnTxt}}</span>
        </div>
        <div class="bar-track">
          ${{nv.high ? `<div class="seg-high" style="width:${{wHigh}}%">${{nv.high}}</div>` : ''}}
          ${{nv.med ? `<div class="seg-med" style="width:${{wMed}}%">${{nv.med}}</div>` : ''}}
          ${{nv.low ? `<div class="seg-low" style="width:${{wLow}}%">${{nv.low}}</div>` : ''}}
        </div>
      </div>
    `;
  }});
  document.getElementById('workload-chart').innerHTML = chartHtml;

  // Render SLA & Workload Table
  let slaHtml = '';
  nvs.forEach(nv => {{
    const scale = nv.total * 7;
    const badge = scale > 60 ? '<span class="badge b-red">🔴 Quá tải nghiêm trọng</span>' : (scale > 30 ? '<span class="badge b-amber">⚠️ Vượt trần 30</span>' : '<span class="badge b-green">✅ Còn dư tải</span>');
    const risk = scale > 60 ? 'Nguy cơ vi phạm SLA 24h rất cao' : (scale > 30 ? 'Cần Overflow hoặc nhắc 4h' : 'Đảm bảo SLA < 4h');
    slaHtml += `
      <tr>
        <td><b>${{nv.name}}</b></td>
        <td class="num">${{nv.total}}</td>
        <td class="num">${{scale}}</td>
        <td class="num">${{fmtMoney(nv.budget)}}</td>
        <td>${{badge}}</td>
        <td style="font-size:12px;">${{risk}}</td>
      </tr>
    `;
  }});
  document.getElementById('sla-rank-body').innerHTML = slaHtml;

  // Render TPKD Exceptions Table
  const exc = rawData.filter(r => r['Trạng Thái Xử Lý'] !== 'ĐÃ PHÂN CÔNG' || ['L055','L084','L090','L100'].includes(r['Mã Hồ Sơ']));
  let excHtml = '';
  exc.forEach(r => {{
    const c = r['Mã Hồ Sơ'];
    let tag = `<span class="badge b-gray">${{r['Trạng Thái Xử Lý']}}</span>`;
    if (c === 'L044' || c === 'L034') tag = '<span class="badge b-red">🚨 VIP ≥500tr bị đóng!</span>';
    else if (c === 'L013') tag = '<span class="badge b-red">🔍 Lịch sử khiếu nại CSKH</span>';
    else if (['L055','L084','L090','L100'].includes(c)) tag = `<span class="badge b-amber">🗺️ Sai vùng (${{r['Tỉnh / Thành Phố']}})</span>`;
    else if (r['Trạng Thái Xử Lý'] === 'CẦN TPKD DUYỆT') tag = '<span class="badge b-amber">⚖️ Đa ngành cần duyệt</span>';
    else if (r['Trạng Thái Xử Lý'] === 'ĐÃ GỘP HỒ SƠ') return; // Skip merged to keep urgent focus

    excHtml += `
      <tr>
        <td><b>${{c}}</b></td>
        <td>${{r['Tên Doanh Nghiệp Chuẩn Hóa']}}<br><small style="color:#64748B">${{r['Người Đại Diện']}}</small></td>
        <td class="num">${{fmtMoney(r['Ngân Sách (VNĐ)'])}}</td>
        <td>${{tag}}</td>
      </tr>
    `;
  }});
  document.getElementById('tpkd-action-body').innerHTML = excHtml;
}}

function renderWorklist() {{
  const nvFilter = document.getElementById('filter-nvkd').value;
  const prioFilter = document.getElementById('filter-prio').value;
  const q = document.getElementById('search-box').value.toLowerCase();

  let list = rawData.filter(r => r['Trạng Thái Xử Lý'] === 'ĐÃ PHÂN CÔNG');
  if (nvFilter !== 'ALL') list = list.filter(r => r['NVKD Phụ Trách'] === nvFilter);
  if (prioFilter !== 'ALL') list = list.filter(r => r['Mức Ưu Tiên'] === prioFilter);
  if (q) list = list.filter(r =>
    r['Tên Doanh Nghiệp Chuẩn Hóa'].toLowerCase().includes(q) ||
    r['Người Đại Diện'].toLowerCase().includes(q) ||
    r['Số Điện Thoại'].toLowerCase().includes(q) ||
    r['Mã Hồ Sơ'].toLowerCase().includes(q)
  );

  // Sort by Priority (1 -> 2 -> 3) then Budget Desc
  list.sort((a,b) => a['Mức Ưu Tiên'].localeCompare(b['Mức Ưu Tiên']) || b['Ngân Sách (VNĐ)'] - a['Ngân Sách (VNĐ)']);

  document.getElementById('nvkd-count-badge').textContent = `${{list.length}} Leads Cần Tiếp Cận`;

  let html = '';
  list.forEach(r => {{
    const p = r['Mức Ưu Tiên'];
    const pBadge = p === '1 - CAO' ? '<span class="badge b-red">🔥 1 - CAO</span>' : (p === '2 - TRUNG BÌNH' ? '<span class="badge b-amber">⚡ 2 - TB</span>' : '<span class="badge b-gray">❄️ 3 - THẤP</span>');
    const isDone = slaCalled.has(r['Mã Hồ Sơ']);
    const btnHtml = isDone
      ? `<button class="btn-action done" onclick="toggleSLA('${{r['Mã Hồ Sơ']}}')">✅ Đã gọi (Đóng SLA)</button>`
      : `<button class="btn-action" onclick="toggleSLA('${{r['Mã Hồ Sơ']}}')">⏱️ Đang đếm 4h | Gọi ngay</button>`;

    html += `
      <tr>
        <td>${{pBadge}}</td>
        <td><b>${{r['Mã Hồ Sơ']}}</b></td>
        <td style="font-size:12px;">${{r['NVKD Phụ Trách']}}</td>
        <td><b>${{r['Tên Doanh Nghiệp Chuẩn Hóa']}}</b><br><small>${{r['Người Đại Diện']}} • 📞 ${{r['Số Điện Thoại']}} • ✉️ ${{r['Địa Chỉ Email']}}</small></td>
        <td>${{r['Tỉnh / Thành Phố']}}<br><span class="badge b-blue">${{r['Gói Sản Phẩm']}}</span></td>
        <td class="num">${{fmtMoney(r['Ngân Sách (VNĐ)'])}}</td>
        <td style="font-size:12px;max-width:240px;">${{r['Hành Động Đề Xuất']}}</td>
        <td>${{btnHtml}}</td>
      </tr>
    `;
  }});
  document.getElementById('worklist-body').innerHTML = html;
}}

function toggleSLA(id) {{
  if (slaCalled.has(id)) slaCalled.delete(id);
  else slaCalled.add(id);
  renderWorklist();
}}

renderTPKD();
renderWorklist();
</script>
</body>
</html>
"""

out_ws = r"c:\Users\HP\Downloads\MÔN AI AGENT\NamViet_Pipeline_Dashboard.html"
out_dl = r"C:\Users\HP\Downloads\NamViet_Pipeline_Dashboard.html"
with open(out_ws, "w", encoding="utf-8") as f:
    f.write(html_template)
with open(out_dl, "w", encoding="utf-8") as f:
    f.write(html_template)

print("HTML Dashboard built successfully!")

