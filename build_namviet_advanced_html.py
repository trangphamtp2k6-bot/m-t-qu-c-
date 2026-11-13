# -*- coding: utf-8 -*-
"""
Script xây dựng Web Dashboard có:
- Bước 1: Kiểm Tra Tính Hợp Lệ Nghiêm Ngặt (Strict Input Validation)
- Bước 2: AI Agent Làm Sạch & Điều Phối 4 Bước
- Bước 3: XÁC NHẬN CỦA KHÁCH HÀNG / QUẢN LÝ TRƯỚC KHI XUẤT BÁO CÁO (Human-in-the-Loop Confirmation)
  * Bảng tổng hợp tóm tắt kết quả tính toán
  * Yêu cầu người dùng tích chọn xác nhận phê duyệt
  * Chỉ khi xác nhận mới mở khóa và cho phép xuất các file báo cáo (.xlsx, .html, .docx, PDF)
"""

import json
import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

with open('namviet_100_clean.json', 'r', encoding='utf-8') as f:
    records = json.load(f)

data_json_str = json.dumps(records, ensure_ascii=False)

html_code = f"""<!DOCTYPE html>
<html lang="vi">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Nam Việt CRM – Trung Tâm Điều Phối Leads & Xuất Báo Cáo AI</title>
<!-- Local SheetJS with CDN fallback for Excel parsing & export -->
<script src="xlsx.full.min.js"></script>
<script>
  if (typeof XLSX === 'undefined') {{
    document.write('<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"><\\/script>');
  }}
</script>
<style>
  :root {{
    --navy: #1B365D;
    --navy-dark: #0F2342;
    --blue: #2563EB;
    --blue-hover: #1D4ED8;
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
    --purple: #7C3AED;
    --purple-bg: #F5F3FF;
  }}
  * {{ box-sizing: border-box; margin: 0; padding: 0; font-family: 'Segoe UI', system-ui, -apple-system, sans-serif; }}
  body {{ background: var(--bg); color: var(--text); line-height: 1.5; }}
  
  header {{
    background: linear-gradient(135deg, var(--navy-dark), var(--navy));
    color: #fff;
    padding: 16px 28px;
    display: flex;
    justify-content: space-between;
    align-items: center;
    flex-wrap: wrap;
    gap: 14px;
    box-shadow: 0 4px 12px rgba(0,0,0,0.12);
  }}
  .brand h1 {{ font-size: 20px; font-weight: 700; letter-spacing: 0.3px; display: flex; align-items: center; gap: 8px; }}
  .brand p {{ font-size: 13px; color: #CBD5E1; margin-top: 3px; }}
  
  .header-actions {{ display: flex; gap: 8px; flex-wrap: wrap; align-items: center; }}
  .btn-hdr {{
    border: none;
    padding: 8px 14px;
    border-radius: 8px;
    font-size: 13px;
    font-weight: 600;
    cursor: pointer;
    display: inline-flex;
    align-items: center;
    gap: 6px;
    transition: all 0.2s;
    text-decoration: none;
  }}
  .btn-upload-toggle {{ background: var(--blue); color: #fff; }}
  .btn-upload-toggle:hover {{ background: var(--blue-hover); transform: translateY(-1px); }}
  .btn-template {{ background: #0891B2; color: #fff; }}
  .btn-template:hover {{ background: #0e7490; transform: translateY(-1px); }}
  .btn-confirm-trigger {{ background: #7C3AED; color: #fff; }}
  .btn-confirm-trigger:hover {{ background: #6D28D9; transform: translateY(-1px); }}
  .btn-export-excel {{ background: #107C41; color: #fff; }}
  .btn-export-excel:hover {{ background: #0b5c30; transform: translateY(-1px); }}
  .btn-export-html {{ background: #D97706; color: #fff; }}
  .btn-export-html:hover {{ background: #b45309; transform: translateY(-1px); }}
  .btn-print {{ background: #475467; color: #fff; }}
  .btn-print:hover {{ background: #344054; transform: translateY(-1px); }}

  .container {{ max-width: 1440px; margin: 20px auto; padding: 0 20px; }}

  /* Approval Notification Strip */
  #approval-strip {{
    background: #FFFBEB;
    border: 1px solid #FCD34D;
    border-left: 6px solid #D97706;
    border-radius: 10px;
    padding: 12px 18px;
    margin-bottom: 20px;
    display: flex;
    justify-content: space-between;
    align-items: center;
    flex-wrap: wrap;
    gap: 12px;
  }}
  #approval-strip.confirmed {{
    background: #ECFDF5;
    border-color: #6EE7B7;
    border-left-color: #059669;
  }}
  .strip-title {{ font-size: 13.5px; font-weight: 600; color: #92400E; display: flex; align-items: center; gap: 8px; }}
  #approval-strip.confirmed .strip-title {{ color: #065F46; }}

  /* Collapsible Upload Hub */
  #upload-hub {{
    background: var(--card);
    border: 2px dashed #93C5FD;
    border-radius: 14px;
    padding: 22px;
    margin-bottom: 24px;
    box-shadow: 0 4px 14px rgba(37,99,235,0.06);
    transition: all 0.3s;
  }}
  .hub-header {{ display: flex; justify-content: space-between; align-items: center; margin-bottom: 16px; flex-wrap: wrap; gap: 10px; }}
  .hub-header h2 {{ font-size: 16px; font-weight: 700; color: var(--navy); display: flex; align-items: center; gap: 8px; }}
  .hub-tabs {{ display: flex; gap: 8px; margin-bottom: 16px; flex-wrap: wrap; }}
  .hub-tab-btn {{
    padding: 7px 14px;
    border: 1px solid var(--border);
    background: #F8FAFC;
    border-radius: 8px;
    font-size: 13px;
    font-weight: 600;
    cursor: pointer;
    color: var(--muted);
  }}
  .hub-tab-btn.active {{ background: #EFF6FF; border-color: var(--blue); color: var(--blue); }}
  
  .drop-zone {{
    border: 2px dashed #CBD5E1;
    border-radius: 10px;
    padding: 28px;
    text-align: center;
    background: #F8FAFC;
    cursor: pointer;
    transition: all 0.2s;
  }}
  .drop-zone:hover, .drop-zone.dragover {{ border-color: var(--blue); background: #EFF6FF; }}
  .drop-zone p {{ margin-top: 8px; font-size: 13px; color: var(--muted); }}
  .drop-zone b {{ color: var(--navy); font-size: 14px; }}
  
  .gs-box {{ display: flex; gap: 10px; flex-wrap: wrap; }}
  .gs-box input {{ flex: 1; min-width: 300px; padding: 10px 14px; border: 1px solid #CBD5E1; border-radius: 8px; font-size: 13.5px; }}
  .btn-primary {{
    background: var(--blue);
    color: #fff;
    border: none;
    padding: 10px 20px;
    border-radius: 8px;
    font-size: 13.5px;
    font-weight: 600;
    cursor: pointer;
    transition: all 0.2s;
  }}
  .btn-primary:hover {{ background: var(--blue-hover); }}
  
  .status-banner {{
    margin-top: 14px;
    padding: 12px 16px;
    border-radius: 8px;
    font-size: 13.5px;
    display: none;
    align-items: center;
    justify-content: space-between;
  }}
  .status-success {{ background: var(--green-bg); color: var(--green); border: 1px solid #A7F3D0; }}
  .status-error {{ background: var(--red-bg); color: var(--red); border: 1px solid #FECACA; }}

  /* Modals */
  .modal-overlay {{
    display: none;
    position: fixed;
    top: 0; left: 0; width: 100%; height: 100%;
    background: rgba(15, 23, 42, 0.65);
    z-index: 9999;
    align-items: center;
    justify-content: center;
    padding: 20px;
    backdrop-filter: blur(2px);
  }}
  .modal-content {{
    background: #fff;
    border-radius: 14px;
    max-width: 650px;
    width: 100%;
    padding: 26px;
    box-shadow: 0 20px 25px -5px rgba(0,0,0,0.25);
    animation: modalPop 0.25s ease-out;
  }}
  @keyframes modalPop {{
    0% {{ transform: scale(0.92); opacity: 0; }}
    100% {{ transform: scale(1); opacity: 1; }}
  }}
  .modal-title {{
    font-size: 18px;
    font-weight: 700;
    display: flex;
    align-items: center;
    gap: 8px;
    margin-bottom: 14px;
  }}
  .modal-body {{
    font-size: 13.5px;
    color: #334155;
    line-height: 1.6;
    margin-bottom: 20px;
  }}
  .modal-actions {{
    display: flex;
    justify-content: flex-end;
    gap: 10px;
    flex-wrap: wrap;
  }}

  /* Confirmation Review Box inside Modal */
  .review-grid {{
    display: grid;
    grid-template-columns: repeat(2, 1fr);
    gap: 12px;
    margin: 14px 0;
  }}
  .review-card {{
    background: #F8FAFC;
    border: 1px solid #E2E8F0;
    border-radius: 8px;
    padding: 12px;
  }}
  .review-card-title {{ font-size: 11.5px; font-weight: 700; color: var(--muted); text-transform: uppercase; }}
  .review-card-val {{ font-size: 20px; font-weight: 800; color: var(--navy); margin-top: 4px; }}
  
  .confirm-checkbox-box {{
    background: #F0FDF4;
    border: 1px solid #BBF7D0;
    padding: 12px 14px;
    border-radius: 8px;
    margin-top: 14px;
    display: flex;
    align-items: flex-start;
    gap: 10px;
    cursor: pointer;
  }}
  .confirm-checkbox-box input {{ margin-top: 3px; cursor: pointer; width: 16px; height: 16px; }}
  .confirm-checkbox-box label {{ font-size: 13px; font-weight: 600; color: #166534; cursor: pointer; }}

  /* Main Navigation Tabs */
  .nav-tabs {{
    display: flex;
    gap: 8px;
    background: var(--card);
    padding: 6px;
    border-radius: 12px;
    border: 1px solid var(--border);
    margin-bottom: 20px;
  }}
  .tab-btn {{
    flex: 1;
    border: none;
    background: transparent;
    color: var(--muted);
    padding: 10px 16px;
    border-radius: 8px;
    cursor: pointer;
    font-size: 14px;
    font-weight: 600;
    transition: all 0.2s;
    text-align: center;
  }}
  .tab-btn.active {{ background: var(--navy); color: #fff; box-shadow: 0 2px 6px rgba(27,54,93,0.2); }}

  .tab-panel {{ display: none; }}
  .tab-panel.active {{ display: block; }}

  /* KPI Grid */
  .kpi-grid {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(220px, 1fr)); gap: 16px; margin-bottom: 20px; }}
  .kpi-card {{ background: var(--card); border-radius: 12px; padding: 18px; border: 1px solid var(--border); box-shadow: 0 1px 3px rgba(0,0,0,0.04); border-top: 4px solid var(--navy); }}
  .kpi-card.warn {{ border-top-color: var(--red); }}
  .kpi-card.ok {{ border-top-color: var(--green); }}
  .kpi-title {{ font-size: 11.5px; font-weight: 700; color: var(--muted); text-transform: uppercase; letter-spacing: 0.5px; }}
  .kpi-val {{ font-size: 26px; font-weight: 800; color: var(--navy); margin: 6px 0 4px; }}
  .kpi-sub {{ font-size: 12px; color: var(--muted); }}
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
  .b-purple {{ background: var(--purple-bg); color: var(--purple); }}
  .b-gray {{ background: #F1F5F9; color: var(--muted); }}

  /* Workload Bar Chart */
  .bar-row {{ margin-bottom: 12px; }}
  .bar-meta {{ display: flex; justify-content: space-between; font-size: 12.5px; margin-bottom: 4px; font-weight: 600; }}
  .bar-track {{ height: 22px; background: #F1F5F9; border-radius: 6px; overflow: hidden; display: flex; position: relative; }}
  .seg-high {{ background: #DC2626; height: 100%; display: flex; align-items: center; justify-content: center; color: #fff; font-size: 11px; font-weight: 700; }}
  .seg-med {{ background: #F59E0B; height: 100%; display: flex; align-items: center; justify-content: center; color: #fff; font-size: 11px; font-weight: 700; }}
  .seg-low {{ background: #64748B; height: 100%; display: flex; align-items: center; justify-content: center; color: #fff; font-size: 11px; font-weight: 700; }}
  .legend {{ display: flex; gap: 16px; font-size: 12px; margin-bottom: 12px; }}
  .legend span {{ display: flex; align-items: center; gap: 5px; }}
  .dot {{ width: 10px; height: 10px; border-radius: 2px; display: inline-block; }}

  /* Tables */
  .tbl-wrap {{ overflow-x: auto; max-height: 600px; }}
  table {{ width: 100%; border-collapse: collapse; font-size: 13px; }}
  th {{ background: #F8FAFC; color: var(--navy); font-weight: 700; text-align: left; padding: 10px 12px; border-bottom: 2px solid var(--border); white-space: nowrap; position: sticky; top: 0; z-index: 1; }}
  td {{ padding: 10px 12px; border-bottom: 1px solid var(--border); vertical-align: middle; }}
  tr:hover td {{ background: #F8FAFC; }}
  .num {{ text-align: right; font-variant-numeric: tabular-nums; font-weight: 600; }}

  /* Worklist Controls */
  .controls {{ display: flex; gap: 12px; flex-wrap: wrap; align-items: center; margin-bottom: 16px; background: #F8FAFC; padding: 12px 16px; border-radius: 10px; border: 1px solid var(--border); }}
  select, input {{ padding: 8px 12px; border: 1px solid #CBD5E1; border-radius: 8px; font-size: 13px; background: #fff; color: var(--text); }}
  .btn-action {{ padding: 6px 12px; border: none; border-radius: 6px; font-size: 12px; font-weight: 600; cursor: pointer; background: var(--blue); color: #fff; transition: all 0.2s; text-decoration: none; display: inline-flex; align-items: center; gap: 4px; }}
  .btn-action.done {{ background: var(--green); }}

  /* Print Styles */
  @media print {{
    header, #upload-hub, #approval-strip, .nav-tabs, .controls, .btn-action, .modal-overlay {{ display: none !important; }}
    .tab-panel {{ display: block !important; page-break-after: always; }}
    body {{ background: #fff; color: #000; }}
    .card {{ box-shadow: none; border: 1px solid #ccc; }}
  }}
</style>
</head>
<body>

<header>
  <div class="brand">
    <h1>🚀 Nam Việt CRM – AI Lead Automation Hub</h1>
    <p>Nhập liệu nghiêm ngặt ➔ AI điều phối ➔ Khách hàng xác nhận ➔ Xuất báo cáo</p>
  </div>
  <div class="header-actions">
    <button class="btn-hdr btn-upload-toggle" onclick="toggleUploadHub()">📥 Tải Lên / Đổi Dữ Liệu</button>
    <a href="Mau_Nhap_Lieu_Nam_Viet_CRM.xlsx" download class="btn-hdr btn-template">📋 File Mẫu (.xlsx)</a>
    <button class="btn-hdr btn-confirm-trigger" onclick="openConfirmModal()">🛡️ Phê Duyệt Xuất Báo Cáo</button>
    <button class="btn-hdr btn-export-excel" onclick="attemptExport('excel')">📊 Xuất File Excel (.xlsx)</button>
    <button class="btn-hdr btn-export-html" onclick="attemptExport('html')">🌐 Xuất Báo Cáo (.html)</button>
    <button class="btn-hdr btn-print" onclick="attemptExport('print')">🖨️ In / Xuất PDF</button>
  </div>
</header>

<div class="container">

  <!-- THANH TRẠNG THÁI XÁC NHẬN CỦA KHÁCH HÀNG (HUMAN-IN-THE-LOOP STRIP) -->
  <div id="approval-strip">
    <div class="strip-title" id="strip-title-text">
      <span>⚠️ <b>TRẠNG THÁI:</b> Dữ liệu đã làm sạch & tính toán xong nhưng <b>CHƯA ĐƯỢC XÁC NHẬN</b>. Vui lòng phê duyệt trước khi xuất báo cáo.</span>
    </div>
    <div style="display:flex; gap:8px;">
      <button class="btn-action" style="background:#7C3AED; padding:8px 16px;" onclick="openConfirmModal()">🔍 Xem Tóm Tắt & Phê Duyệt Ngay</button>
    </div>
  </div>

  <!-- KHU VỰC NHẬP LIỆU ĐẦU VÀO ĐA ĐỊNH DẠNG -->
  <div id="upload-hub" style="display:none;">
    <div class="hub-header">
      <h2>📥 BƯỚC 1: NHẬP LIỆU ĐẦU VÀO (KIỂM TRA ĐỊNH DẠNG NGHIÊM NGẶT)</h2>
      <div style="display:flex; gap:8px;">
        <a href="Mau_Nhap_Lieu_Nam_Viet_CRM.xlsx" download class="btn-action" style="background:#0891B2;">📥 Tải File Mẫu (.xlsx)</a>
        <button class="btn-action" style="background:#64748B;" onclick="resetToDefaultData()">↺ Khôi phục 100 Leads Mẫu</button>
      </div>
    </div>
    
    <div class="hub-tabs">
      <button class="hub-tab-btn active" onclick="switchHubTab('tab-file', this)">📁 Tải File (Excel / CSV / JSON)</button>
      <button class="hub-tab-btn" onclick="switchHubTab('tab-gsheet', this)">🌐 Link Google Sheets</button>
      <button class="hub-tab-btn" onclick="switchHubTab('tab-paste', this)">📋 Dán Dữ Liệu Bảng Tính</button>
    </div>

    <!-- Panel 1: File Upload -->
    <div id="tab-file" class="hub-subpanel">
      <div class="drop-zone" id="drop-zone" onclick="document.getElementById('file-input').click()">
        <input type="file" id="file-input" style="display:none;" accept=".xlsx, .xls, .csv, .json" onchange="handleFileSelect(event)">
        <div style="font-size:32px;">📄</div>
        <b>Kéo thả file vào đây hoặc nhấp để chọn file</b>
        <p>Hệ thống chỉ chấp nhận đúng định dạng bảng tính: <b>.xlsx, .xls, .csv, .json</b></p>
      </div>
    </div>

    <!-- Panel 2: Google Sheets Link -->
    <div id="tab-gsheet" class="hub-subpanel" style="display:none;">
      <p style="font-size:13px; color:var(--muted); margin-bottom:10px;">
        💡 Link Google Sheets phải có dạng: <code>https://docs.google.com/spreadsheets/d/{{SHEET_ID}}/edit</code> và đã mở quyền xem công khai (<i>"Bất kỳ ai có đường liên kết đều có thể xem"</i>).
      </p>
      <div class="gs-box">
        <input type="text" id="gsheet-url" placeholder="https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/edit#gid=0">
        <button class="btn-primary" onclick="fetchGoogleSheet()">🚀 Kiểm Tra & Nạp Link</button>
      </div>
    </div>

    <!-- Panel 3: Paste Raw Text -->
    <div id="tab-paste" class="hub-subpanel" style="display:none;">
      <textarea id="paste-input" rows="5" style="width:100%; padding:10px; border-radius:8px; border:1px solid #CBD5E1; font-size:12px;" placeholder="Copy các cột từ bảng tính (phải gồm Tên DN, SĐT, Tỉnh, Gói/Ngành, Ngân sách) rồi dán vào đây..."></textarea>
      <div style="margin-top:10px; text-align:right;">
        <button class="btn-primary" onclick="parsePastedData()">🚀 Kiểm Tra & Xử Lý Dữ Liệu</button>
      </div>
    </div>

    <!-- Status Alert Banner -->
    <div id="status-banner" class="status-banner">
      <span id="status-msg"></span>
      <span id="status-details" style="font-weight:600;"></span>
    </div>
  </div>

  <!-- NAVIGATION TABS -->
  <div class="nav-tabs">
    <button class="tab-btn active" onclick="switchTab('tab-tpkd', this)">📊 1. Tổng Quan TPKD & Pipeline</button>
    <button class="tab-btn" onclick="switchTab('tab-worklist', this)">📋 2. Bàn Làm Việc 8 NVKD</button>
    <button class="tab-btn" onclick="switchTab('tab-audit', this)">🔍 3. Kiểm Toán 4 Bước AI</button>
    <button class="tab-btn" onclick="switchTab('tab-sla', this)">⏱️ 4. Giám Sát SLA & Đồng Hồ Đếm Ngược</button>
  </div>

  <!-- TAB 1: TỔNG QUAN TPKD -->
  <div id="tab-tpkd" class="tab-panel active">
    <div class="kpi-grid">
      <div class="kpi-card">
        <div class="kpi-title">Tổng Leads Tiếp Nhận</div>
        <div class="kpi-val" id="kpi-total">100</div>
        <div class="kpi-sub" id="kpi-total-sub">Đầu vào toàn luồng</div>
      </div>
      <div class="kpi-card ok">
        <div class="kpi-title">Bàn Giao Cho 8 NVKD</div>
        <div class="kpi-val" id="kpi-assigned">80</div>
        <div class="kpi-sub success" id="kpi-pipeline">Pipeline: 3,03 tỷ ₫</div>
      </div>
      <div class="kpi-card">
        <div class="kpi-title">Hồ Sơ Trùng Đã Gộp</div>
        <div class="kpi-val" id="kpi-merged">7</div>
        <div class="kpi-sub">Chống đụng độ (6 Master)</div>
      </div>
      <div class="kpi-card warn">
        <div class="kpi-title">Từ Chối / Blacklist</div>
        <div class="kpi-val" id="kpi-rejected">7</div>
        <div class="kpi-sub danger" id="kpi-vip-sub">Đã gửi phản hồi</div>
      </div>
      <div class="kpi-card warn">
        <div class="kpi-title">TPKD Duyệt / Bổ Sung</div>
        <div class="kpi-val" id="kpi-tpkd">6</div>
        <div class="kpi-sub danger">Bao gồm Deal VIP & CSKH</div>
      </div>
    </div>

    <div class="grid-2">
      <!-- Cột trái: Biểu đồ tải trọng 8 NVKD -->
      <div class="card">
        <div class="card-title">
          <span>📊 Phân Bổ Tải Trọng & Mức Độ Ưu Tiên 8 NVKD</span>
          <span class="badge b-blue" id="assigned-badge">80 Leads Hợp Lệ</span>
        </div>
        <div class="legend">
          <span><span class="dot" style="background:#DC2626;"></span> 1 - Cao</span>
          <span><span class="dot" style="background:#F59E0B;"></span> 2 - Trung Bình</span>
          <span><span class="dot" style="background:#64748B;"></span> 3 - Thấp</span>
        </div>
        <div id="workload-chart"></div>
      </div>

      <!-- Cột phải: Xếp hạng nguy cơ SLA & Ca khẩn -->
      <div>
        <div class="card">
          <div class="card-title">
            <span>⚠️ Cảnh Báo Quá Tải & Nguy Cơ Vi Phạm SLA</span>
            <span class="badge b-red">Ngưỡng Trần 30 Leads</span>
          </div>
          <div class="tbl-wrap">
            <table>
              <thead>
                <tr>
                  <th>NVKD</th>
                  <th class="num">Tuần Mẫu</th>
                  <th class="num">Dự Báo 700</th>
                  <th>Trạng Thái</th>
                </tr>
              </thead>
              <tbody id="sla-rank-body"></tbody>
            </table>
          </div>
        </div>

        <div class="card">
          <div class="card-title">
            <span>🚨 Hồ Sơ Cần TPKD Can Thiệp Ngay</span>
            <span class="badge b-amber" id="urgent-badge">Xem xét</span>
          </div>
          <div class="tbl-wrap" style="max-height:220px;">
            <table>
              <thead>
                <tr>
                  <th>Mã</th>
                  <th>Doanh Nghiệp</th>
                  <th class="num">Ngân Sách</th>
                  <th>Lý Do</th>
                </tr>
              </thead>
              <tbody id="tpkd-action-body"></tbody>
            </table>
          </div>
        </div>
      </div>
    </div>
  </div>

  <!-- TAB 2: BÀN LÀM VIỆC 8 NVKD -->
  <div id="tab-worklist" class="tab-panel">
    <div class="card">
      <div class="card-title">
        <span>📋 Bàn Làm Việc Hằng Ngày Cho Từng NVKD (Daily Worklist)</span>
        <span class="badge b-green" id="nvkd-count-badge">80 Leads</span>
      </div>

      <div class="controls">
        <div>
          <label style="font-size:12px; font-weight:600; color:var(--muted); display:block; margin-bottom:4px;">Lọc theo Nhân Viên:</label>
          <select id="filter-nvkd" onchange="renderWorklist()">
            <option value="ALL">-- Tất cả 8 NVKD --</option>
            <option value="NV01 - Nguyễn Văn An">NV01 - Nguyễn Văn An (Bán lẻ - Bắc)</option>
            <option value="NV02 - Trần Thị Bình">NV02 - Trần Thị Bình (Bán lẻ - Nam)</option>
            <option value="NV03 - Lê Hoàng Cường">NV03 - Lê Hoàng Cường (F&B - Bắc & Trung)</option>
            <option value="NV04 - Phạm Minh Dũng">NV04 - Phạm Minh Dũng (F&B - Nam)</option>
            <option value="NV05 - Vũ Thị Hoa">NV05 - Vũ Thị Hoa (Dịch vụ - Toàn quốc)</option>
            <option value="NV06 - Đỗ Tuấn Kiệt">NV06 - Đỗ Tuấn Kiệt (Sản xuất - Toàn quốc)</option>
            <option value="NV07 - Bùi Thu Lan">NV07 - Bùi Thu Lan (Đa ngành - Bắc)</option>
            <option value="NV08 - Ngô Quang Nam">NV08 - Ngô Quang Nam (Đa ngành - Nam)</option>
          </select>
        </div>

        <div>
          <label style="font-size:12px; font-weight:600; color:var(--muted); display:block; margin-bottom:4px;">Mức Ưu Tiên:</label>
          <select id="filter-prio" onchange="renderWorklist()">
            <option value="ALL">-- Tất cả độ ưu tiên --</option>
            <option value="1 - CAO">🔥 1 - CAO (Ưu tiên gọi trong 4h)</option>
            <option value="2 - TRUNG BÌNH">⚡ 2 - TRUNG BÌNH</option>
            <option value="3 - THẤP">❄️ 3 - THẤP</option>
          </select>
        </div>

        <div style="flex:1;">
          <label style="font-size:12px; font-weight:600; color:var(--muted); display:block; margin-bottom:4px;">Tìm kiếm nhanh:</label>
          <input type="text" id="search-box" placeholder="Tìm theo tên DN, người liên hệ, SĐT, mã..." oninput="renderWorklist()" style="width:100%;">
        </div>
      </div>

      <div class="tbl-wrap">
        <table>
          <thead>
            <tr>
              <th>Ưu Tiên</th>
              <th>Mã</th>
              <th>NVKD</th>
              <th>Doanh Nghiệp & Người Đại Diện</th>
              <th>Tỉnh / Gói SP</th>
              <th class="num">Ngân Sách</th>
              <th>Hành Động Đề Xuất</th>
              <th>Hành Động Gọi</th>
            </tr>
          </thead>
          <tbody id="worklist-body"></tbody>
        </table>
      </div>
    </div>
  </div>

  <!-- TAB 3: KIỂM TOÁN 4 BƯỚC AI -->
  <div id="tab-audit" class="tab-panel">
    <div class="card">
      <div class="card-title">
        <span>🔍 Nhật Ký Kiểm Toán Toàn Bộ Bản Ghi (Audit Log 4 Bước)</span>
        <button class="btn-action" style="background:#107C41;" onclick="attemptExport('excel')">📥 Tải Excel Kiểm Toán</button>
      </div>

      <div class="tbl-wrap">
        <table>
          <thead>
            <tr>
              <th>Mã</th>
              <th>Doanh Nghiệp</th>
              <th>MST / SĐT</th>
              <th>Tỉnh / Ngành</th>
              <th>Gói SP</th>
              <th class="num">Ngân Sách</th>
              <th>Phân Công Cho</th>
              <th>Trạng Thái</th>
              <th>Cảnh Báo Kiểm Toán AI</th>
            </tr>
          </thead>
          <tbody id="audit-table-body"></tbody>
        </table>
      </div>
    </div>
  </div>

  <!-- TAB 4: GIÁM SÁT SLA -->
  <div id="tab-sla" class="tab-panel">
    <div class="card">
      <div class="card-title">
        <span>⏱️ Giám Sát Đồng Hồ SLA & Mô Phỏng Cảnh Báo Vi Phạm</span>
        <span class="badge b-purple">Chu Kỳ 4h Nhắc Nhở / 24h Thu Hồi</span>
      </div>
      <p style="font-size:13px; color:var(--muted); margin-bottom:16px;">
        Hệ thống tự động kích hoạt đồng hồ kể từ thời điểm Lead được bàn giao vào tài khoản NVKD. Nếu sau 4h NVKD chưa liên hệ, bot gửi tin nhắn nhắc nhở. Sau 24h không liên hệ, quyền chăm sóc tự động bị thu hồi về kho chung.
      </p>

      <div class="tbl-wrap">
        <table>
          <thead>
            <tr>
              <th>Mã Lead</th>
              <th>Doanh Nghiệp</th>
              <th>NVKD Phụ Trách</th>
              <th>Mức Ưu Tiên</th>
              <th>Thời Gian Trôi Qua (Mô phỏng)</th>
              <th>Trạng Thái Đồng Hồ SLA</th>
              <th>Hành Động Khắc Phục</th>
            </tr>
          </thead>
          <tbody id="sla-simulation-body"></tbody>
        </table>
      </div>
    </div>
  </div>

</div>

<!-- MODAL 1: BÁO LỖI ĐẦU VÀO NGHIÊM NGẶT -->
<div id="error-modal" class="modal-overlay">
  <div class="modal-content" style="border-top:6px solid var(--red);">
    <div class="modal-title" style="color:var(--red);">
      <span>❌</span>
      <span id="modal-err-title">Lỗi Định Dạng Dữ Liệu Đầu Vào!</span>
    </div>
    <div class="modal-body" id="modal-err-body">
      <!-- Error details injected here -->
    </div>
    <div class="modal-actions">
      <a href="Mau_Nhap_Lieu_Nam_Viet_CRM.xlsx" download class="btn-action" style="background:#0891B2;">📥 Tải File Mẫu Chuẩn (.xlsx)</a>
      <button class="btn-primary" style="background:#64748B;" onclick="closeModal('error-modal')">Đã Hiểu & Nhập Lại</button>
    </div>
  </div>
</div>

<!-- MODAL 2: XÁC NHẬN PHÊ DUYỆT CỦA KHÁCH HÀNG / QUẢN LÝ (CONFIRMATION MODAL) -->
<div id="confirm-modal" class="modal-overlay">
  <div class="modal-content" style="border-top:6px solid #7C3AED; max-width:680px;">
    <div class="modal-title" style="color:#7C3AED;">
      <span>🛡️</span>
      <span>Xác Nhận & Phê Duyệt Kết Quả Điều Phối Leads</span>
    </div>
    <div class="modal-body">
      <p style="font-size:13.5px; color:#475467; margin-bottom:12px;">
        Dữ liệu đã được AI hoàn thành các công đoạn làm sạch, gộp trùng, lọc VIP và phân bổ tải trọng cho 8 NVKD. Vui lòng đối soát bảng tóm tắt bên dưới trước khi ký duyệt xuất file báo cáo chính thức.
      </p>

      <div class="review-grid">
        <div class="review-card">
          <div class="review-card-title">Tổng Leads Đã Xử Lý</div>
          <div class="review-card-val" id="rev-total">100</div>
          <small style="color:var(--muted)">Đầu vào toàn luồng</small>
        </div>
        <div class="review-card" style="border-color:#A7F3D0; background:#F0FDF4;">
          <div class="review-card-title" style="color:#059669">Bàn Giao Cho 8 NVKD</div>
          <div class="review-card-val" id="rev-assigned" style="color:#059669">80 Leads</div>
          <small id="rev-pipeline" style="color:#059669; font-weight:600">3,03 tỷ ₫</small>
        </div>
        <div class="review-card">
          <div class="review-card-title">Hồ Sơ Trùng Lặp Đã Gộp</div>
          <div class="review-card-val" id="rev-merged">7</div>
          <small style="color:var(--muted)">Tiết kiệm thời gian, chống đụng độ</small>
        </div>
        <div class="review-card" style="border-color:#FECACA; background:#FEF2F2;">
          <div class="review-card-title" style="color:#DC2626">Cần TPKD Can Thiệp / VIP</div>
          <div class="review-card-val" id="rev-tpkd" style="color:#DC2626">6 Leads</div>
          <small style="color:#DC2626; font-weight:600">Gồm đơn VIP ≥ 500tr & CSKH</small>
        </div>
      </div>

      <div style="background:#F8FAFC; border:1px solid #E2E8F0; padding:12px; border-radius:8px; margin-top:10px;">
        <label style="font-size:12px; font-weight:700; color:var(--navy); display:block; margin-bottom:4px;">
          Họ Tên Người Xác Nhận / Quản Lý:
        </label>
        <input type="text" id="signer-name" placeholder="Ví dụ: Nguyễn Văn Trưởng Phòng / Khách Hàng Đại Diện" value="Trưởng Phòng Kinh Doanh Nam Việt" style="width:100%; padding:8px 10px; font-size:13px; border:1px solid #CBD5E1; border-radius:6px;">
      </div>

      <div class="confirm-checkbox-box" onclick="toggleConfirmCheckbox()">
        <input type="checkbox" id="confirm-chk">
        <label for="confirm-chk">
          Tôi đã đối soát và <b>XÁC NHẬN</b> kết quả làm sạch, phân bổ tải trọng và đồng ý xuất các file báo cáo chính thức.
        </label>
      </div>
    </div>
    <div class="modal-actions">
      <button class="btn-action" style="background:#64748B;" onclick="closeModal('confirm-modal')">Hủy / Xem Lại Dữ Liệu</button>
      <button class="btn-primary" style="background:#059669;" onclick="submitCustomerConfirmation()">✅ Ký Duyệt & Cho Phép Xuất Báo Cáo</button>
    </div>
  </div>
</div>

<script>
// ================= DATA & INITIALIZATION =================
let currentData = {data_json_str};
const defaultData = JSON.parse(JSON.stringify(currentData));
const slaCalled = new Set();
let isCustomerConfirmed = true; // Mặc định dữ liệu mẫu đã xác nhận; khi tải file mới sẽ reset về false
let confirmedBy = "Trưởng Phòng Kinh Doanh Nam Việt";
let confirmedAt = new Date().toLocaleString('vi-VN');
let pendingExportAction = null; // Lưu loại báo cáo đang chờ xuất

// Ma trận 8 NVKD chuẩn
const SALES_MATRIX = [
  {{ id: "NV01", name: "NV01 - Nguyễn Văn An", pkg: "PROD_RETAIL", provinces: ["Hà Nội", "Hải Phòng", "Quảng Ninh", "Bắc Ninh", "Hải Dương", "Hưng Yên", "Nam Định", "Thái Bình", "Vĩnh Phúc", "Phú Thọ"] }},
  {{ id: "NV02", name: "NV02 - Trần Thị Bình", pkg: "PROD_RETAIL", provinces: ["TP. Hồ Chí Minh", "Bình Dương", "Đồng Nai", "Long An", "Cần Thơ", "Tiền Giang", "Bến Tre", "Tây Ninh"] }},
  {{ id: "NV03", name: "NV03 - Lê Hoàng Cường", pkg: "PROD_FNB", provinces: ["Hà Nội", "Đà Nẵng", "Quảng Nam", "Thừa Thiên Huế", "Khánh Hòa", "Hải Phòng", "Quảng Ninh"] }},
  {{ id: "NV04", name: "NV04 - Phạm Minh Dũng", pkg: "PROD_FNB", provinces: ["TP. Hồ Chí Minh", "Bình Dương", "Đồng Nai", "Bà Rịa - Vũng Tàu", "Cần Thơ", "Kiên Giang"] }},
  {{ id: "NV05", name: "NV05 - Vũ Thị Hoa", pkg: "PROD_SERVICE", provinces: ["Toàn quốc"] }},
  {{ id: "NV06", name: "NV06 - Đỗ Tuấn Kiệt", pkg: "PROD_MFG", provinces: ["Toàn quốc"] }},
  {{ id: "NV07", name: "NV07 - Bùi Thu Lan", pkg: "ĐA_NGANH", provinces: ["Hà Nội", "Hải Phòng", "Bắc Ninh", "Thái Nguyên", "Vĩnh Phúc", "Thanh Hóa", "Nghệ An", "Hà Tĩnh"] }},
  {{ id: "NV08", name: "NV08 - Ngô Quang Nam", pkg: "ĐA_NGANH", provinces: ["TP. Hồ Chí Minh", "Bình Dương", "Đồng Nai", "Long An", "Bình Thuận", "Ninh Thuận", "Lâm Đồng", "Đắk Lắk", "Cần Thơ"] }}
];

const BLACKLIST = {{
  "0909999000": {{ type: "Blacklist", note: "Do Not Contact" }},
  "0988000111": {{ type: "Khách cũ", owner: "NV03 - Lê Hoàng Cường" }},
  "0918000222": {{ type: "Khiếu nại CSKH", note: "Từng khiếu nại chất lượng" }}
}};

const REQUIRED_FIELDS_SCHEMA = [
  {{ key: "doanh_nghiep", label: "Tên Doanh Nghiệp", aliases: ["tên doanh nghiệp", "doanh nghiệp", "công ty", "company", "khách hàng"] }},
  {{ key: "sdt", label: "Số Điện Thoại", aliases: ["số điện thoại", "điện thoại", "sđt", "phone", "hotline", "di động"] }},
  {{ key: "tinh_thanh", label: "Tỉnh / Thành Phố", aliases: ["tỉnh", "thành phố", "khu vực", "city", "province"] }},
  {{ key: "san_pham_nganh", label: "Gói Sản Phẩm hoặc Nhóm Ngành", aliases: ["gói sản phẩm", "sản phẩm", "gói sp", "package", "nhóm ngành", "ngành nghề", "ngành", "lĩnh vực", "industry"] }},
  {{ key: "ngan_sach", label: "Ngân Sách (VNĐ)", aliases: ["ngân sách", "budget", "doanh thu", "chi phí"] }}
];

function toggleUploadHub() {{
  const hub = document.getElementById('upload-hub');
  hub.style.display = hub.style.display === 'none' ? 'block' : 'none';
}}

function switchHubTab(tabId, btn) {{
  document.querySelectorAll('.hub-subpanel').forEach(p => p.style.display = 'none');
  document.querySelectorAll('.hub-tab-btn').forEach(b => b.classList.remove('active'));
  document.getElementById(tabId).style.display = 'block';
  btn.classList.add('active');
}}

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
  return Number(val).toLocaleString('vi-VN') + ' ₫';
}}

function showErrorModal(title, messageHtml) {{
  document.getElementById('modal-err-title').textContent = title;
  document.getElementById('modal-err-body').innerHTML = messageHtml;
  document.getElementById('error-modal').style.display = 'flex';
}}

function closeModal(modalId) {{
  document.getElementById(modalId).style.display = 'none';
}}

// ================= HUMAN CONFIRMATION WORKFLOW =================
function updateApprovalStrip() {{
  const strip = document.getElementById('approval-strip');
  const txt = document.getElementById('strip-title-text');

  if (isCustomerConfirmed) {{
    strip.className = 'confirmed';
    txt.innerHTML = `<span>✅ <b>ĐÃ ĐƯỢC XÁC NHẬN:</b> Báo cáo đã được ký duyệt bởi <b>${{confirmedBy}}</b> lúc <i>${{confirmedAt}}</i>. Sẵn sàng xuất tất cả định dạng!</span>`;
  }} else {{
    strip.className = '';
    txt.innerHTML = `<span>⚠️ <b>TRẠNG THÁI:</b> Dữ liệu mới đã làm sạch & tính toán xong nhưng <b>CHƯA ĐƯỢC XÁC NHẬN</b>. Vui lòng phê duyệt trước khi xuất báo cáo.</span>`;
  }}
}}

function openConfirmModal(actionAfterConfirm = null) {{
  pendingExportAction = actionAfterConfirm;
  
  const total = currentData.length;
  const assigned = currentData.filter(r => r['Trạng Thái Xử Lý'] === 'ĐÃ PHÂN CÔNG');
  const merged = currentData.filter(r => r['Trạng Thái Xử Lý'] === 'ĐÃ GỘP HỒ SƠ');
  const tpkd = currentData.filter(r => r['Trạng Thái Xử Lý'] === 'CẦN TPKD DUYỆT' || r['Trạng Thái Xử Lý'] === 'CHỜ BỔ SUNG THÔNG TIN');
  const totalBudget = assigned.reduce((s, r) => s + (r['Ngân Sách (VNĐ)'] || 0), 0);

  document.getElementById('rev-total').textContent = total.toLocaleString() + ' Leads';
  document.getElementById('rev-assigned').textContent = assigned.length.toLocaleString() + ' Leads';
  document.getElementById('rev-pipeline').textContent = 'Pipeline: ' + fmtMoney(totalBudget);
  document.getElementById('rev-merged').textContent = merged.length.toLocaleString() + ' Hồ sơ';
  document.getElementById('rev-tpkd').textContent = tpkd.length.toLocaleString() + ' Leads';
  document.getElementById('confirm-chk').checked = isCustomerConfirmed;

  document.getElementById('confirm-modal').style.display = 'flex';
}}

function toggleConfirmCheckbox() {{
  const chk = document.getElementById('confirm-chk');
  chk.checked = !chk.checked;
}}

function submitCustomerConfirmation() {{
  const chk = document.getElementById('confirm-chk');
  if (!chk.checked) {{
    alert('Vui lòng tích chọn vào ô: "Tôi đã đối soát và XÁC NHẬN kết quả..." để hoàn tất phê duyệt!');
    return;
  }}

  const signer = document.getElementById('signer-name').value.trim() || 'Người Đại Diện Khách Hàng';
  confirmedBy = signer;
  confirmedAt = new Date().toLocaleString('vi-VN');
  isCustomerConfirmed = true;

  closeModal('confirm-modal');
  updateApprovalStrip();
  showStatus(`🎉 <b>Xác nhận thành công!</b> Phê duyệt bởi: ${{confirmedBy}} lúc ${{confirmedAt}}. Hệ thống đã mở khóa toàn bộ quyền xuất báo cáo.`, 'success');

  // Nếu có hành động xuất file đang chờ, thực thi ngay
  if (pendingExportAction === 'excel') exportToExcel();
  else if (pendingExportAction === 'html') exportToHTML();
  else if (pendingExportAction === 'print') window.print();
  pendingExportAction = null;
}}

function attemptExport(type) {{
  if (!isCustomerConfirmed) {{
    openConfirmModal(type);
    return;
  }}
  if (type === 'excel') exportToExcel();
  else if (type === 'html') exportToHTML();
  else if (type === 'print') window.print();
}}

// ================= DRAG & DROP AND FILE IMPORT WITH STRICT VALIDATION =================
const dropZone = document.getElementById('drop-zone');
['dragenter', 'dragover'].forEach(name => {{
  dropZone.addEventListener(name, (e) => {{ e.preventDefault(); dropZone.classList.add('dragover'); }}, false);
}});
['dragleave', 'drop'].forEach(name => {{
  dropZone.addEventListener(name, (e) => {{ e.preventDefault(); dropZone.classList.remove('dragover'); }}, false);
}});
dropZone.addEventListener('drop', (e) => {{
  const files = e.dataTransfer.files;
  if (files.length) processSelectedFile(files[0]);
}});

function handleFileSelect(e) {{
  if (e.target.files.length) {{
    processSelectedFile(e.target.files[0]);
    e.target.value = '';
  }}
}}

function processSelectedFile(file) {{
  const validExtensions = ['.xlsx', '.xls', '.csv', '.json'];
  const ext = file.name.substring(file.name.lastIndexOf('.')).toLowerCase();

  if (!validExtensions.includes(ext)) {{
    showErrorModal(
      '❌ Định Dạng File Không Hợp Lệ!',
      `Tập tin <b>"${{file.name}}"</b> có định dạng <b>"${{ext}}"</b> không được hỗ trợ.<br><br>` +
      `<b>Quy định định dạng:</b> Hệ thống chỉ chấp nhận các file bảng tính chuẩn:<br>` +
      `• Microsoft Excel: <code>.xlsx</code>, <code>.xls</code><br>` +
      `• Tệp phân tách dấu phẩy: <code>.csv</code><br>` +
      `• Tệp dữ liệu cấu trúc: <code>.json</code><br><br>` +
      `👉 <b>Yêu cầu:</b> Vui lòng xuất bảng tính của bạn ra file <code>.xlsx</code> hoặc <code>.csv</code> và thử tải lại.`
    );
    return;
  }}

  showStatus(`Đang đọc file: ${{file.name}}...`, 'info');
  const reader = new FileReader();

  if (ext === '.json') {{
    reader.onload = (e) => {{
      try {{
        const json = JSON.parse(e.target.result);
        validateAndRunAIEngine(json, file.name);
      }} catch(err) {{
        showErrorModal('❌ File JSON Bị Lỗi Cú Pháp!', `Nội dung file JSON bị lỗi: <i>${{err.message}}</i>. Vui lòng kiểm tra lại cấu trúc file.`);
      }}
    }};
    reader.readAsText(file);
  }} else {{
    reader.onload = (e) => {{
      try {{
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, {{ type: 'array' }});
        
        let targetSheetName = workbook.SheetNames[0];
        for (let s of workbook.SheetNames) {{
          if (s.toLowerCase().includes('data') || s.toLowerCase().includes('lead') || s.toLowerCase().includes('hồ sơ') || s.toLowerCase().includes('audit')) {{
            targetSheetName = s;
            break;
          }}
        }}

        const worksheet = workbook.Sheets[targetSheetName];
        const json = XLSX.utils.sheet_to_json(worksheet);
        validateAndRunAIEngine(json, file.name);
      }} catch(err) {{
        showErrorModal('❌ Không Thể Đọc Bảng Tính!', `Lỗi phân tích file Excel/CSV: <i>${{err.message}}</i>. Vui lòng đảm bảo file không bị hỏng hoặc khóa mật khẩu.`);
      }}
    }};
    reader.readAsArrayBuffer(file);
  }}
}}

// ================= GOOGLE SHEETS IMPORT =================
function fetchGoogleSheet() {{
  let url = document.getElementById('gsheet-url').value.trim();

  if (!url) {{
    showErrorModal('❌ Link Google Sheets Bị Trống!', 'Vui lòng dán đường link liên kết của trang tính Google Sheets vào ô nhập liệu.');
    return;
  }}

  if (!url.includes('docs.google.com/spreadsheets/d/')) {{
    showErrorModal(
      '❌ Sai Định Dạng Link Google Sheets!',
      `Đường link bạn nhập: <br><code>${{url}}</code><br><br>` +
      `<b>Cú pháp hợp lệ của Google Sheets:</b><br>` +
      `<code>https://docs.google.com/spreadsheets/d/{{SHEET_ID}}/edit#gid=0</code><br><br>` +
      `👉 <b>Yêu cầu:</b> Vui lòng sao chép lại toàn bộ link trên thanh địa chỉ và dán lại vào đây.`
    );
    return;
  }}

  let exportUrl = url;
  const idMatch = url.match(/\/d\/([a-zA-Z0-9-_]+)/);
  const gidMatch = url.match(/[#&]gid=([0-9]+)/);
  if (idMatch) {{
    const sheetId = idMatch[1];
    const gid = gidMatch ? gidMatch[1] : '0';
    exportUrl = `https://docs.google.com/spreadsheets/d/${{sheetId}}/export?format=csv&gid=${{gid}}`;
  }}

  showStatus('Đang kết nối tải dữ liệu từ Google Sheets...', 'info');

  fetch(exportUrl)
    .then(r => {{
      if (r.status === 403 || r.status === 401) {{
        throw new Error('Google Sheets đang ở chế độ RIÊNG TƯ (Private). Quyền truy cập bị từ chối.');
      }}
      if (r.status === 404) {{
        throw new Error('Không tìm thấy bảng tính Google Sheets (Mã ID sheet không tồn tại).');
      }}
      if (!r.ok) throw new Error(`Mã lỗi mạng HTTP: ${{r.status}}`);
      return r.text();
    }})
    .then(csvText => {{
      if (!csvText || !csvText.trim()) throw new Error('Trang tính Google Sheets không chứa bất kỳ dữ liệu nào (File rỗng).');
      const workbook = XLSX.read(csvText, {{ type: 'string' }});
      const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
      const json = XLSX.utils.sheet_to_json(firstSheet);
      validateAndRunAIEngine(json, 'Google Sheets');
    }})
    .catch(err => {{
      showErrorModal(
        '❌ Lỗi Kết Nối Google Sheets!',
        `Chi tiết lỗi: <b>${{err.message}}</b><br><br>` +
        `<b>Cách khắc phục:</b><br>` +
        `1. Mở trang tính Google Sheets của bạn.<br>` +
        `2. Nhấn nút <b>Chia sẻ (Share)</b> ➔ Đổi quyền thành: <b>"Bất kỳ ai có đường liên kết đều có thể xem"</b>.<br>` +
        `3. Hoặc tải file về: <i>Tệp ➔ Tải xuống ➔ Microsoft Excel (.xlsx)</i> rồi kéo thả trực tiếp vào đây.`
      );
    }});
}}

// ================= PASTE RAW TEXT =================
function parsePastedData() {{
  const text = document.getElementById('paste-input').value.trim();
  if (!text) {{
    showErrorModal('❌ Dữ Liệu Dán Bị Trống!', 'Vui lòng copy dữ liệu từ Excel / Google Sheets rồi dán vào khung văn bản.');
    return;
  }}

  try {{
    const workbook = XLSX.read(text, {{ type: 'string' }});
    const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
    const json = XLSX.utils.sheet_to_json(firstSheet);
    validateAndRunAIEngine(json, 'Dữ liệu dán trực tiếp');
  }} catch(err) {{
    showErrorModal('❌ Lỗi Phân Tích Cú Pháp!', `Không thể nhận diện bảng dữ liệu vừa dán: <i>${{err.message}}</i>.`);
  }}
}}

// ================= SCHEMA VALIDATION =================
function validateAndRunAIEngine(rawList, sourceName) {{
  if (!rawList || !Array.isArray(rawList) || rawList.length === 0) {{
    showErrorModal(
      '❌ Bảng Tính Không Có Dữ Liệu!',
      `Tập tin từ <b>${{sourceName}}</b> không chứa bất kỳ dòng dữ liệu khách hàng nào.<br><br>` +
      `👉 <b>Yêu cầu:</b> Vui lòng kiểm tra lại file và nạp lại bảng tính có dữ liệu.`
    );
    return;
  }}

  const firstRow = rawList[0];
  const detectedKeys = Object.keys(firstRow).map(k => String(k).trim().toLowerCase());
  const missingFields = [];

  REQUIRED_FIELDS_SCHEMA.forEach(schema => {{
    const isFound = schema.aliases.some(alias => detectedKeys.some(dk => dk.includes(alias)));
    if (!isFound) missingFields.push(schema);
  }});

  if (missingFields.length > 0) {{
    const missingHtml = missingFields.map(m => 
      `<li><b>${{m.label}}</b> (Chấp nhận các từ khóa: <i>${{m.aliases.slice(0, 3).join(', ')}}...</i>)</li>`
    ).join('');

    showErrorModal(
      '❌ Dữ Liệu Thiếu Các Cột Bắt Buộc!',
      `Tập tin từ <b>${{sourceName}}</b> chưa đạt tiêu chuẩn để AI có thể điều phối.<br><br>` +
      `<b>Bảng tính của bạn đang THIẾU các trường bắt buộc sau:</b><br>` +
      `<ul>${{missingHtml}}</ul><br>` +
      `👉 <b>Yêu cầu:</b> Vui lòng bổ sung các cột trên (hoặc tải file mẫu chuẩn bên dưới) và tải lên lại.`
    );
    return;
  }}

  // Dữ liệu hợp lệ ➔ Chạy AI Engine
  runAIEngine(rawList, sourceName);
}}

// ================= THUẬT TOÁN ĐIỀU PHỐI AI 4 BƯỚC =================
function normalizeKey(k) {{
  const s = String(k).toLowerCase();
  if (s.includes('mã') || s.includes('id')) return 'Mã Hồ Sơ';
  if (s.includes('đại diện') || s.includes('liên hệ') || s.includes('họ tên')) return 'Người Đại Diện';
  if (s.includes('doanh nghiệp') || s.includes('công ty') || s.includes('company')) return 'Tên Doanh Nghiệp Chuẩn Hóa';
  if (s.includes('thuế') || s.includes('mst') || s.includes('domain') || s.includes('miền')) return 'Mã Số Thuế / Tên Miền';
  if (s.includes('thoại') || s.includes('sđt') || s.includes('phone')) return 'Số Điện Thoại';
  if (s.includes('mail')) return 'Địa Chỉ Email';
  if (s.includes('tỉnh') || s.includes('thành phố') || s.includes('khu vực')) return 'Tỉnh / Thành Phố';
  if (s.includes('ngành')) return 'Nhóm Ngành Chuẩn Hóa';
  if (s.includes('gói') || s.includes('sản phẩm')) return 'Gói Sản Phẩm';
  if (s.includes('ngân sách') || s.includes('budget') || s.includes('doanh thu')) return 'Ngân Sách (VNĐ)';
  if (s.includes('ưu tiên') || s.includes('priority')) return 'Mức Ưu Tiên';
  return k;
}}

function runAIEngine(rawList, sourceName) {{
  const seenMst = {{}};
  const seenPhone = {{}};
  const backlog = {{}};
  SALES_MATRIX.forEach(n => backlog[n.name] = 0);

  const processed = [];

  rawList.forEach((raw, idx) => {{
    const r = {{}};
    Object.keys(raw).forEach(k => r[normalizeKey(k)] = raw[k]);

    const code = String(r['Mã Hồ Sơ'] || `L${{String(idx+1).padStart(3, '0')}}`).trim();
    const rep = String(r['Người Đại Diện'] || 'Chưa rõ').trim();
    const comp = String(r['Tên Doanh Nghiệp Chuẩn Hóa'] || 'Doanh nghiệp SMB').trim();
    const mst = String(r['Mã Số Thuế / Tên Miền'] || '').trim();
    const phone = String(r['Số Điện Thoại'] || '').replace(/[\s\.\-\(\)]/g, '');
    const email = String(r['Địa Chỉ Email'] || '').trim().toLowerCase();
    const city = String(r['Tỉnh / Thành Phố'] || 'Chưa xác định').trim();
    const ind = String(r['Nhóm Ngành Chuẩn Hóa'] || 'Đa ngành').trim();
    let pkg = String(r['Gói Sản Phẩm'] || 'CHƯA XÁC ĐỊNH').trim();
    
    let budget = 0;
    if (r['Ngân Sách (VNĐ)']) {{
      const bStr = String(r['Ngân Sách (VNĐ)']).replace(/[^0-9]/g, '');
      budget = Number(bStr) || 0;
    }}

    let prio = budget >= 50000000 ? '1 - CAO' : (budget >= 20000000 ? '2 - TRUNG BÌNH' : '3 - THẤP');
    let status = 'ĐÃ PHÂN CÔNG';
    let assigned = '';
    let action = '';
    let flag = '✅ HỢP LỆ - Theo dõi SLA 4h / 24h';
    let slaTarget = 'Theo dõi 4h / 24h';

    // BƯỚC 1: Blacklist & Lịch sử & Trùng lặp
    if (BLACKLIST[phone] || BLACKLIST[email]) {{
      const hit = BLACKLIST[phone] || BLACKLIST[email];
      if (hit.type === 'Blacklist') {{
        status = 'TỪ CHỐI / ĐÓNG LEAD';
        prio = '0 - ĐÃ ĐÓNG';
        assigned = 'HỆ THỐNG ĐÓNG';
        action = 'Do Not Contact - Khách từ chối tiếp cận.';
        flag = '🛑 BLACKLIST - Đã đóng tự động';
        slaTarget = 'Không áp dụng';
      }} else if (hit.type === 'Khiếu nại CSKH') {{
        status = 'CẦN TPKD DUYỆT';
        prio = '1 - CAO (CSKH)';
        assigned = 'CHỜ TPKD CHỈ ĐỊNH';
        action = 'Lịch sử khiếu nại CSKH. Trình TPKD chỉ định nhân sự xử lý.';
        flag = '🔍 LỊCH SỬ KHIẾU NẠI CSKH - Đẩy TPKD đúng Bước 1';
        slaTarget = 'Khẩn cấp TPKD';
      }} else if (hit.type === 'Khách cũ') {{
        assigned = hit.owner || 'NV03 - Lê Hoàng Cường';
        action = 'Khách cũ gia hạn/upsell. Giao thẳng NVKD từng chăm sóc.';
        flag = `🤝 KHÁCH CŨ UPSELL - Giao thẳng ${{assigned}}`;
        backlog[assigned] = (backlog[assigned] || 0) + 1;
      }}
    }}

    // Trùng lặp
    if (status !== 'TỪ CHỐI / ĐÓNG LEAD' && status !== 'CẦN TPKD DUYỆT') {{
      let masterId = null;
      if (mst && mst !== 'CHƯA XÁC ĐỊNH' && mst !== '0') masterId = seenMst[mst];
      if (!masterId && phone && phone.length >= 9) masterId = seenPhone[phone];

      if (masterId) {{
        status = 'ĐÃ GỘP HỒ SƠ';
        prio = '0 - ĐÃ GỘP';
        assigned = `Theo Master ${{masterId}}`;
        action = `Gộp vào hồ sơ ${{masterId}}. Lưu ${{rep}} làm liên hệ phụ.`;
        flag = `🔗 ĐÃ GỘP MASTER (${{masterId}}) - Chống gọi đụng độ`;
        slaTarget = 'Không áp dụng';
      }} else {{
        if (mst && mst !== 'CHƯA XÁC ĐỊNH' && mst !== '0') seenMst[mst] = code;
        if (phone && phone.length >= 9) seenPhone[phone] = code;
      }}
    }}

    // BƯỚC 2: Định Tuyến Thông Minh & Luồng VIP
    if (status === 'ĐÃ PHÂN CÔNG' && !assigned) {{
      if (budget >= 500000000) {{
        status = 'CẦN TPKD DUYỆT';
        prio = 'TPKD DUYỆT (VIP)';
        assigned = 'CHỜ TPKD CHỈ ĐỊNH';
        action = `Deal VIP ${{fmtMoney(budget)}} (≥500tr). Báo đỏ TPKD chỉ định KAM.`;
        flag = `⚠️ VIP ≥500TR (${{fmtMoney(budget)}}) - Cần TPKD/KAM thẩm định`;
        slaTarget = 'Khẩn cấp TPKD';
      }} else if (pkg === 'NGOÀI PHẠM VI' || pkg === 'KHÔNG PHÙ HỢP') {{
        status = 'TỪ CHỐI / ĐÓNG LEAD';
        prio = '3 - THẤP';
        assigned = 'CHƯA PHÂN BỔ';
        action = 'Ngoài phạm vi gói phần mềm SMB. Gửi email từ chối lịch sự.';
        flag = '🛑 NGOÀI PHẠM VI SMB - Đã đóng tự động';
        slaTarget = 'Không áp dụng';
      }} else if (pkg === 'CHƯA XÁC ĐỊNH' || comp === 'Doanh nghiệp SMB') {{
        status = 'CHỜ BỔ SUNG THÔNG TIN';
        prio = '3 - THẤP';
        assigned = 'HỆ THỐNG FORM TỰ ĐỘNG';
        action = 'Chưa rõ nhu cầu. Tự động gửi Form khảo sát thông tin.';
        flag = '📩 THIẾU THÔNG TIN - Đang gửi khảo sát tự động';
        slaTarget = 'Chờ khách phản hồi';
      }} else {{
        let matched = SALES_MATRIX.find(n => n.pkg === pkg && (n.provinces.includes('Toàn quốc') || n.provinces.some(p => city.includes(p))));
        if (!matched) {{
          const isSouth = ['Hồ Chí Minh', 'Bình Dương', 'Đồng Nai', 'Cần Thơ', 'Long An'].some(p => city.includes(p));
          matched = isSouth ? SALES_MATRIX[7] : SALES_MATRIX[6];
          flag = `⚖️ ĐA LĨNH VỰC/VÙNG (${{city}}) - NVKD Đa ngành phụ trách`;
        }}

        // BƯỚC 3: Cân Bằng Tải Trọng
        if (backlog[matched.name] >= 30 && prio !== '1 - CAO') {{
          const alt = SALES_MATRIX.filter(n => n.name !== matched.name && backlog[n.name] < 30).sort((a,b) => backlog[a.name] - backlog[b.name])[0];
          if (alt) {{
            flag = `🔀 KÍCH HOẠT OVERFLOW - ${{matched.id}} quá tải 30 leads, chuyển ${{alt.id}}`;
            matched = alt;
          }}
        }}

        assigned = matched.name;
        action = `Tư vấn giải pháp ${{pkg}}. NVKD liên hệ trong 24 giờ.`;
        backlog[assigned] = (backlog[assigned] || 0) + 1;
      }}
    }}

    processed.push({{
      'Mã Hồ Sơ': code,
      'Người Đại Diện': rep,
      'Tên Doanh Nghiệp Chuẩn Hóa': comp,
      'Mã Số Thuế / Tên Miền': mst || 'CHƯA XÁC ĐỊNH',
      'Số Điện Thoại': phone || 'CHƯA CÓ',
      'Địa Chỉ Email': email || 'CHƯA CÓ',
      'Tỉnh / Thành Phố': city,
      'Nhóm Ngành Chuẩn Hóa': ind,
      'Gói Sản Phẩm': pkg,
      'Ngân Sách (VNĐ)': budget,
      'Mức Ưu Tiên': prio,
      'NVKD Phụ Trách': assigned,
      'Trạng Thái Xử Lý': status,
      'Hành Động Đề Xuất': action,
      'Cảnh Báo Kiểm Toán BI': flag,
      'SLA Mục Tiêu': slaTarget
    }});
  }});

  currentData = processed;
  slaCalled.clear();
  
  // RESET TRẠNG THÁI XÁC NHẬN KHI CÓ DỮ LIỆU MỚI TÍNH TOÁN XONG
  isCustomerConfirmed = false;
  updateApprovalStrip();
  renderAllViews();
  
  showStatus(`✅ Đã nạp và tính toán ${{processed.length}} leads từ: ${{sourceName}}! Vui lòng <b>XÁC NHẬN</b> trước khi xuất báo cáo.`, 'info');
  
  // Tự động mở Modal Xác Nhận Phê Duyệt cho người dùng đối soát
  setTimeout(() => {{
    openConfirmModal();
  }}, 400);
}}

function resetToDefaultData() {{
  currentData = JSON.parse(JSON.stringify(defaultData));
  slaCalled.clear();
  isCustomerConfirmed = true;
  confirmedBy = "Trưởng Phòng Kinh Doanh Nam Việt";
  confirmedAt = new Date().toLocaleString('vi-VN');
  updateApprovalStrip();
  renderAllViews();
  showStatus('Đã khôi phục dữ liệu mẫu 100 Leads chuẩn.', 'success');
}}

function showStatus(msg, type) {{
  const b = document.getElementById('status-banner');
  b.style.display = 'flex';
  b.className = 'status-banner ' + (type === 'success' ? 'status-success' : (type === 'error' ? 'status-error' : 'status-success'));
  document.getElementById('status-msg').innerHTML = msg;
}}

// ================= RENDER FUNCTIONS =================
function renderAllViews() {{
  renderTPKD();
  renderWorklist();
  renderAuditTable();
  renderSLASimulation();
}}

function renderTPKD() {{
  const total = currentData.length;
  const assigned = currentData.filter(r => r['Trạng Thái Xử Lý'] === 'ĐÃ PHÂN CÔNG');
  const merged = currentData.filter(r => r['Trạng Thái Xử Lý'] === 'ĐÃ GỘP HỒ SƠ');
  const rejected = currentData.filter(r => r['Trạng Thái Xử Lý'] === 'TỪ CHỐI / ĐÓNG LEAD');
  const tpkd = currentData.filter(r => r['Trạng Thái Xử Lý'] === 'CẦN TPKD DUYỆT' || r['Trạng Thái Xử Lý'] === 'CHỜ BỔ SUNG THÔNG TIN');
  const totalBudget = assigned.reduce((s, r) => s + (r['Ngân Sách (VNĐ)'] || 0), 0);

  document.getElementById('kpi-total').textContent = total.toLocaleString();
  document.getElementById('kpi-assigned').textContent = assigned.length.toLocaleString();
  document.getElementById('kpi-pipeline').textContent = 'Pipeline: ' + fmtMoney(totalBudget);
  document.getElementById('kpi-merged').textContent = merged.length.toLocaleString();
  document.getElementById('kpi-rejected').textContent = rejected.length.toLocaleString();
  document.getElementById('kpi-tpkd').textContent = tpkd.length.toLocaleString();
  document.getElementById('assigned-badge').textContent = `${{assigned.length}} Leads Hợp Lệ`;

  // Workload Chart
  const nvMap = {{}};
  SALES_MATRIX.forEach(n => nvMap[n.name] = {{ name: n.name, high: 0, med: 0, low: 0, total: 0, budget: 0 }});
  assigned.forEach(r => {{
    const k = r['NVKD Phụ Trách'];
    if (nvMap[k]) {{
      nvMap[k].total++;
      nvMap[k].budget += r['Ngân Sách (VNĐ)'] || 0;
      if (r['Mức Ưu Tiên'] === '1 - CAO') nvMap[k].high++;
      else if (r['Mức Ưu Tiên'] === '2 - TRUNG BÌNH') nvMap[k].med++;
      else nvMap[k].low++;
    }}
  }});

  const nvs = Object.values(nvMap).sort((a,b) => b.total - a.total);
  const maxTotal = Math.max(...nvs.map(x => x.total), 1);

  let chartHtml = '';
  nvs.forEach(nv => {{
    const wHigh = (nv.high / Math.max(maxTotal, 1)) * 100;
    const wMed = (nv.med / Math.max(maxTotal, 1)) * 100;
    const wLow = (nv.low / Math.max(maxTotal, 1)) * 100;
    const scale700 = Math.round(nv.total * (700 / Math.max(total, 1)));
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

  // SLA Table
  let slaHtml = '';
  nvs.forEach(nv => {{
    const scale = Math.round(nv.total * (700 / Math.max(total, 1)));
    const badge = scale > 40 ? '<span class="badge b-red">🔴 Quá tải</span>' : (scale > 30 ? '<span class="badge b-amber">⚠️ Vượt trần 30</span>' : '<span class="badge b-green">✅ An toàn</span>');
    slaHtml += `
      <tr>
        <td><b>${{nv.name.split(' - ')[0]}}</b></td>
        <td class="num">${{nv.total}}</td>
        <td class="num">${{scale}}</td>
        <td>${{badge}}</td>
      </tr>
    `;
  }});
  document.getElementById('sla-rank-body').innerHTML = slaHtml;

  // Urgent TPKD Action Table
  const urgent = currentData.filter(r => r['Trạng Thái Xử Lý'] !== 'ĐÃ PHÂN CÔNG' && r['Trạng Thái Xử Lý'] !== 'ĐÃ GỘP HỒ SƠ');
  document.getElementById('urgent-badge').textContent = `${{urgent.length}} Ca Cần Xử Lý`;
  let excHtml = '';
  urgent.slice(0, 15).forEach(r => {{
    let tag = `<span class="badge b-amber">${{r['Trạng Thái Xử Lý']}}</span>`;
    if (r['Cảnh Báo Kiểm Toán BI'].includes('VIP')) tag = '<span class="badge b-red">🚨 VIP ≥500tr</span>';
    else if (r['Cảnh Báo Kiểm Toán BI'].includes('CSKH')) tag = '<span class="badge b-red">🔍 CSKH khiếu nại</span>';

    excHtml += `
      <tr>
        <td><b>${{r['Mã Hồ Sơ']}}</b></td>
        <td>${{r['Tên Doanh Nghiệp Chuẩn Hóa']}}</td>
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

  let list = currentData.filter(r => r['Trạng Thái Xử Lý'] === 'ĐÃ PHÂN CÔNG');
  if (nvFilter !== 'ALL') list = list.filter(r => r['NVKD Phụ Trách'] === nvFilter);
  if (prioFilter !== 'ALL') list = list.filter(r => r['Mức Ưu Tiên'] === prioFilter);
  if (q) list = list.filter(r =>
    (r['Tên Doanh Nghiệp Chuẩn Hóa'] || '').toLowerCase().includes(q) ||
    (r['Người Đại Diện'] || '').toLowerCase().includes(q) ||
    (r['Số Điện Thoại'] || '').includes(q) ||
    (r['Mã Hồ Sơ'] || '').toLowerCase().includes(q)
  );

  list.sort((a,b) => String(a['Mức Ưu Tiên']).localeCompare(String(b['Mức Ưu Tiên'])) || (b['Ngân Sách (VNĐ)'] - a['Ngân Sách (VNĐ)']));

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
  document.getElementById('worklist-body').innerHTML = html || '<tr><td colspan="8" style="text-align:center; padding:30px; color:var(--muted);">Không tìm thấy hồ sơ phù hợp.</td></tr>';
}}

function toggleSLA(id) {{
  if (slaCalled.has(id)) slaCalled.delete(id);
  else slaCalled.add(id);
  renderWorklist();
  renderSLASimulation();
}}

function renderAuditTable() {{
  let html = '';
  currentData.forEach(r => {{
    html += `
      <tr>
        <td><b>${{r['Mã Hồ Sơ']}}</b></td>
        <td>${{r['Tên Doanh Nghiệp Chuẩn Hóa']}}<br><small style="color:var(--muted);">${{r['Người Đại Diện']}}</small></td>
        <td>${{r['Số Điện Thoại']}}<br><small>${{r['Mã Số Thuế / Tên Miền']}}</small></td>
        <td>${{r['Tỉnh / Thành Phố']}}<br><small>${{r['Nhóm Ngành Chuẩn Hóa']}}</small></td>
        <td><span class="badge b-blue">${{r['Gói Sản Phẩm']}}</span></td>
        <td class="num">${{fmtMoney(r['Ngân Sách (VNĐ)'])}}</td>
        <td>${{r['NVKD Phụ Trách']}}</td>
        <td><span class="badge b-gray">${{r['Trạng Thái Xử Lý']}}</span></td>
        <td style="font-size:12px;">${{r['Cảnh Báo Kiểm Toán BI']}}</td>
      </tr>
    `;
  }});
  document.getElementById('audit-table-body').innerHTML = html;
}}

function renderSLASimulation() {{
  const assigned = currentData.filter(r => r['Trạng Thái Xử Lý'] === 'ĐÃ PHÂN CÔNG').slice(0, 15);
  let html = '';
  assigned.forEach((r, idx) => {{
    const isDone = slaCalled.has(r['Mã Hồ Sơ']);
    let elapsed = ((idx * 2.3) % 28).toFixed(1);
    let statusBadge = '<span class="badge b-green">⏱️ Trong hạn 4h</span>';
    let actionTxt = 'Đang trong thời gian cho phép gọi điện.';

    if (isDone) {{
      statusBadge = '<span class="badge b-blue">✅ Đã liên hệ thành công</span>';
      actionTxt = 'Đã tắt đồng hồ SLA.';
    }} else if (elapsed > 24) {{
      statusBadge = '<span class="badge b-red">🚨 Quá hạn 24h - THU HỒI</span>';
      actionTxt = 'Tự động chuyển về kho chung, cảnh báo đỏ TPKD.';
    }} else if (elapsed > 4) {{
      statusBadge = '<span class="badge b-amber">⏰ Quá hạn 4h - ĐÃ NHẮC</span>';
      actionTxt = 'Đã bắn tin nhắn Zalo/Telegram giục gọi ngay.';
    }}

    html += `
      <tr>
        <td><b>${{r['Mã Hồ Sơ']}}</b></td>
        <td>${{r['Tên Doanh Nghiệp Chuẩn Hóa']}}</td>
        <td>${{r['NVKD Phụ Trách']}}</td>
        <td>${{r['Mức Ưu Tiên']}}</td>
        <td class="num">${{isDone ? 'Hoàn thành' : elapsed + ' giờ'}}</td>
        <td>${{statusBadge}}</td>
        <td style="font-size:12px;">${{actionTxt}}</td>
      </tr>
    `;
  }});
  document.getElementById('sla-simulation-body').innerHTML = html;
}}

// ================= EXPORT FUNCTIONS =================
function exportToExcel() {{
  if (typeof XLSX === 'undefined') return alert('Thư viện XLSX chưa sẵn sàng. Vui lòng thử lại sau 2 giây.');
  
  const wb = XLSX.utils.book_new();

  // Sheet 1: Dashboard
  const dashData = [
    ["BÁO CÁO PIPELINE & PHÂN BỔ LEADS - CÔNG TY NAM VIỆT"],
    ["Thời gian xuất:", new Date().toLocaleString()],
    ["Người phê duyệt / Xác nhận:", confirmedBy],
    ["Thời điểm phê duyệt:", confirmedAt],
    [],
    ["TỔNG SỐ LEADS ĐẦU VÀO", currentData.length],
    ["BÀN GIAO CHO 8 NVKD", currentData.filter(r => r['Trạng Thái Xử Lý'] === 'ĐÃ PHÂN CÔNG').length],
    ["HỒ SƠ TRÙNG ĐÃ GỘP", currentData.filter(r => r['Trạng Thái Xử Lý'] === 'ĐÃ GỘP HỒ SƠ').length],
    ["TỪ CHỐI / BLACKLIST", currentData.filter(r => r['Trạng Thái Xử Lý'] === 'TỪ CHỐI / ĐÓNG LEAD').length],
    ["CẦN TPKD DUYỆT", currentData.filter(r => r['Trạng Thái Xử Lý'] === 'CẦN TPKD DUYỆT').length]
  ];
  const ws1 = XLSX.utils.aoa_to_sheet(dashData);
  XLSX.utils.book_append_sheet(wb, ws1, "1_Dashboard");

  // Sheet 2: Data Audit
  const ws2 = XLSX.utils.json_to_sheet(currentData);
  XLSX.utils.book_append_sheet(wb, ws2, "2_Data_Audit");

  // Sheet 3: Phan Cong 8 NVKD
  const assigned = currentData.filter(r => r['Trạng Thái Xử Lý'] === 'ĐÃ PHÂN CÔNG');
  const ws3 = XLSX.utils.json_to_sheet(assigned);
  XLSX.utils.book_append_sheet(wb, ws3, "3_Phan_Cong_8_NVKD");

  // Sheet 4: TPKD Xu Ly Gap
  const urgent = currentData.filter(r => r['Trạng Thái Xử Lý'] !== 'ĐÃ PHÂN CÔNG' && r['Trạng Thái Xử Lý'] !== 'ĐÃ GỘP HỒ SƠ');
  const ws4 = XLSX.utils.json_to_sheet(urgent);
  XLSX.utils.book_append_sheet(wb, ws4, "4_TPKD_Xu_Ly_Gap");

  const fileName = `NamViet_CRM_Export_${{new Date().toISOString().slice(0,10)}}.xlsx`;
  XLSX.writeFile(wb, fileName);
}}

function exportToHTML() {{
  const fullHtml = document.documentElement.outerHTML;
  const blob = new Blob([fullHtml], {{ type: 'text/html;charset=utf-8' }});
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `NamViet_Pipeline_Report_${{new Date().toISOString().slice(0,10)}}.html`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
}}

// Run on load
updateApprovalStrip();
renderAllViews();
</script>
</body>
</html>
"""

with open("c:\\Users\\HP\\Downloads\\MÔN AI AGENT\\NamViet_Pipeline_Dashboard.html", "w", encoding="utf-8") as f:
    f.write(html_code)

with open("c:\\Users\\HP\\Downloads\\MÔN AI AGENT\\index.html", "w", encoding="utf-8") as f:
    f.write(html_code)

print("HTML Dashboard with Customer Confirmation Step built successfully!")
