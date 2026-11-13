import os
import sys
import docx
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

from docx_helpers import (
    COLOR_PRIMARY, COLOR_SECONDARY, COLOR_TEXT, COLOR_MUTED,
    HEX_PRIMARY, HEX_LIGHT_BG, HEX_HEADER_BG, HEX_ZEBRA,
    set_cell_background, set_cell_margins, set_table_borders,
    format_run, add_heading_1, add_heading_2, add_heading_3,
    add_body, add_bullet, add_callout, add_figure, add_table
)

sys.stdout.reconfigure(encoding='utf-8')

OUT_WS = r'c:\Users\HP\Downloads\MÔN AI AGENT\Bao_Cao_Phan_Tich_Va_Giai_Phap_Nam_Viet_CRM.docx'
OUT_DL = r'C:\Users\HP\Downloads\Bao_Cao_Phan_Tich_Va_Giai_Phap_Nam_Viet_CRM.docx'

doc = docx.Document()

# Page setup (Academic / Business standard)
section = doc.sections[0]
section.page_width = Cm(21.0)
section.page_height = Cm(29.7)
section.top_margin = Cm(2.0)
section.bottom_margin = Cm(2.0)
section.left_margin = Cm(3.0)
section.right_margin = Cm(2.0)

style_normal = doc.styles['Normal']
style_normal.font.name = 'Times New Roman'
style_normal.font.size = Pt(12.5)
style_normal.font.color.rgb = COLOR_TEXT

# ==================== COVER / HEADER BLOCK ====================
tbl_hdr = doc.add_table(rows=1, cols=1)
tbl_hdr.alignment = WD_TABLE_ALIGNMENT.CENTER
tbl_hdr.autofit = False
cell = tbl_hdr.cell(0, 0)
cell.width = Cm(16.0)
set_cell_background(cell, HEX_LIGHT_BG)
set_cell_margins(cell, top=160, bottom=160, left=200, right=200)
tcPr = cell._tc.get_or_add_tcPr()
tcPr.append(parse_xml(f'''
    <w:tcBorders {nsdecls("w")}>
        <w:top w:val="single" w:sz="18" w:space="0" w:color="{HEX_PRIMARY}"/>
        <w:bottom w:val="single" w:sz="18" w:space="0" w:color="{HEX_PRIMARY}"/>
        <w:left w:val="none"/><w:right w:val="none"/>
    </w:tcBorders>
'''))

p0 = cell.paragraphs[0]
p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
p0.paragraph_format.space_after = Pt(4)
r0 = p0.add_run("CÔNG TY CỔ PHẦN PHẦN MỀM NAM VIỆT • BÁO CÁO CHẨN ĐOÁN & GIẢI PHÁP BI")
format_run(r0, size_pt=11.5, bold=True, color=COLOR_SECONDARY)

p1 = cell.add_paragraph()
p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
p1.paragraph_format.space_after = Pt(6)
r1 = p1.add_run("PHÂN TÍCH GỐC RỄ SỰ CỐ ĐIỀU PHỐI LEADS (CON NGƯỜI VS. HỆ THỐNG),\nĐÁNH GIÁ TÁC ĐỘNG VÀ HƯỚNG DẪN TRIỂN KHAI GIẢI PHÁP TRÊN LARK BASE")
format_run(r1, size_pt=14.5, bold=True, color=COLOR_PRIMARY)

p2 = cell.add_paragraph()
p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
p2.paragraph_format.space_after = Pt(0)
r2 = p2.add_run("Người thực hiện: Chuyên viên Business Intelligence (BI)  |  Quy mô: 700 Leads/tuần (Kiểm toán mẫu 100 Leads)  |  Đội ngũ: 8 NVKD & 1 TPKD")
format_run(r2, size_pt=11, italic=True, color=COLOR_TEXT)

doc.add_paragraph().paragraph_format.space_after = Pt(4)

# ==================== PHẦN 1: PHÂN TÍCH VẤN ĐỀ & CHẨN ĐOÁN GỐC RỄ ====================
add_heading_1(doc, "PHẦN 1: PHÂN TÍCH KỸ BỐI CẢNH VÀ CHẨN ĐOÁN GỐC RỄ VẤN ĐỀ")

add_heading_2(doc, "1.1. Bức tranh Thực tế từ Dữ liệu Kiểm toán (100 Leads Mẫu)")
add_body(doc, "Công ty Nam Việt tiếp nhận khoảng 700 hồ sơ khách hàng tiềm năng (Leads) mỗi tuần từ hội thảo, website và các chiến dịch tiếp thị cho phần mềm quản trị doanh nghiệp vừa và nhỏ (SMB), phân bổ cho 8 Nhân viên Kinh doanh (NVKD) theo ngành hàng và vùng địa lý. Qua kiểm toán trực tiếp 100 hồ sơ chuẩn hóa mẫu (tổng ngân sách danh nghĩa 6,765 tỷ VNĐ), cấu trúc phễu phân bổ được thể hiện như sau:", indent=True)

add_figure(doc, r'c:\Users\HP\Downloads\MÔN AI AGENT\chart_namviet_funnel.png', "Hình 1.1: Phân loại 100 hồ sơ Leads đầu vào và Giá trị Ngân sách theo nhóm trạng thái", "Nguồn: Kết xuất kiểm toán dữ liệu 100 Leads Nam Việt", width=Cm(15.5))

add_bullet(doc, "80 hồ sơ (80%) đủ điều kiện được bàn giao cho 8 NVKD (tổng Pipeline đạt 3,033 tỷ VNĐ; trung bình 37,91 triệu/Lead).", bold_prefix="Bàn giao hợp lệ: ")
add_bullet(doc, "7 hồ sơ (7%) bị trùng lặp cùng một doanh nghiệp (như An Phước, Golden Coffee, Seoul Spa...) đã được gộp về 6 hồ sơ Master (tránh lãng phí 335 triệu VNĐ dữ liệu trùng).", bold_prefix="Hồ sơ trùng đã gộp: ")
add_bullet(doc, "7 hồ sơ (7%) bị đóng vì đánh giá ngoài phạm vi SMB, nhưng tổng ngân sách nhóm này lên tới 3,100 tỷ VNĐ — đặc biệt có 2 hồ sơ VIP ≥ 500 triệu bị đóng oan.", bold_prefix="Từ chối ngoài phạm vi: ")
add_bullet(doc, "6 hồ sơ (6%) chờ xử lý gồm 4 ca cần TPKD duyệt (trong đó L013 có lịch sử khiếu nại CSKH cũ) và 2 ca thiếu thông tin doanh nghiệp/nhu cầu.", bold_prefix="Cần TPKD duyệt / Bổ sung: ")

add_heading_2(doc, "1.2. Chẩn đoán Gốc rễ: Lỗi nằm ở Con người hay do Máy móc / Hệ thống?")
add_body(doc, "Dưới góc độ kiểm toán BI, tình trạng hỗn loạn (gọi trùng khách, bỏ sót khách phù hợp, chạy theo đơn to sai sản phẩm, trưởng phòng mù thông tin) là sự cộng hưởng giữa 70% Lỗi từ Con người & Quy trình quản lý và 30% Lỗi từ Thiết lập Máy móc / Công cụ:", indent=True)

diag_headers = ["Phân nhóm nguyên nhân", "Biểu hiện cụ thể ghi nhận trên dữ liệu", "Bản chất gốc rễ của vấn đề"]
diag_widths = [Cm(3.5), Cm(6.5), Cm(6.0)]
diag_align = [WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT]
diag_data = [
    ["1. Lỗi CON NGƯỜI (Marketing & Sales)\n[Chiếm ~45%]",
     "• Cùng một công ty (An Phước, Golden Coffee) nhưng các bộ phận đăng ký khác nhau bị đẩy thô xuống Sales.\n• NVKD tự đọc danh sách, thấy ngân sách lớn (300–900tr) là ưu tiên gọi trước dù sai gói SMB; bỏ qua các hồ sơ thiếu thông tin (L008, L071).",
     "Thiếu kỷ luật chuẩn hóa đầu vào và tâm lý 'hái quả dễ / tự gắp hồ sơ' của NVKD khi không có rào chắn quyền xem dữ liệu."],
    ["2. Lỗi QUY TRÌNH QUẢN LÝ (TPKD)\n[Chiếm ~25%]",
     "• Ma trận phân quyền bỏ quên khu vực Miền Trung cho mảng Bán lẻ & Đa ngành: L055 (Thanh Hóa), L090 (Khánh Hòa), L100 (Quảng Ngãi) giao cho NV07 Miền Bắc; L084 (Huế) giao cho NV08 Miền Nam.\n• Chưa quy định ngưỡng trần 30 leads/người và SLA phản hồi.",
     "Trưởng phòng chưa thiết kế Ma trận tỉnh thành khép kín 63 tỉnh và chưa thiết lập cơ chế định mức tải trọng, thời hạn cam kết."],
    ["3. Lỗi MÁY MÓC / LUẬT HỆ THỐNG\n[Chiếm ~30%]",
     "• Lỗi thứ tự câu lệnh AI: Máy chạy luật 'Từ chối ngoài phạm vi SMB' trước luật 'Chặn VIP ≥ 500tr', làm đóng tự động L044 (900tr) và L034 (500tr).\n• Bảng CRM hiện tại thiếu hoàn toàn cột lưu Giờ giao việc và Giờ gọi đầu.",
     "Công cụ phân loại cấu hình sai thứ tự ưu tiên (Priority Logic) và thiếu các trường Timestamp để máy móc có thể đếm ngược SLA 4h/24h."]
]
add_table(doc, "Bảng 1.1: Phân tách lỗi Con người vs. Quy trình vs. Máy móc tại Công ty Nam Việt", diag_headers, diag_data, diag_widths, diag_align, "Phân tích gốc rễ BI Nam Việt")

add_heading_2(doc, "1.3. Ảnh hưởng và Thiệt hại Cụ thể nếu không xử lý triệt để")
add_body(doc, "Những lỗ hổng trên không chỉ gây khó chịu trong vận hành mà đang tạo ra 4 thiệt hại định lượng trực tiếp cho doanh thu và nhân sự của công ty:", indent=True)

add_figure(doc, r'c:\Users\HP\Downloads\MÔN AI AGENT\chart_namviet_workload.png', "Hình 1.2: Mất cân bằng tải trọng 8 NVKD và Dự báo vỡ trần 30 Leads khi lên quy mô 700 Leads/tuần", "Nguồn: Mô hình dự phóng tải trọng BI Nam Việt", width=Cm(15.2))

add_bullet(doc, "Chỉ riêng trong 100 mẫu đã để lọt 2 khách hàng ngân sách lớn là Đại Nam Resort (L044 - 900 triệu) và FPT Partner (L034 - 500 triệu) do đóng tự động nhầm. Quy đổi lên quy mô 700 leads/tuần, công ty có nguy cơ đánh rơi tới ~9,8 tỷ VNĐ giá trị cơ hội dự án lớn mỗi tuần.", bold_prefix="1. Đánh rơi cơ hội doanh thu tỷ đồng (VIP Leakage): ")
add_bullet(doc, "7% hồ sơ trùng + 7% hồ sơ ngoài phạm vi = 14% thời gian của đội ngũ 8 NVKD bị đốt cháy vào những cuộc gọi vô nghĩa (tương đương gần 100 cuộc gọi lãng phí mỗi tuần). Nghiêm trọng hơn, việc 2 nhân viên gọi đụng một công ty làm giảm uy tín chuyên nghiệp của Nam Việt.", bold_prefix="2. Đốt 14% công suất đội Sales & Mất uy tín thương hiệu: ")
add_bullet(doc, "Dữ liệu cho thấy sự chênh lệch cực đoan: NV06 ôm tới 1,005 tỷ VNĐ (15 leads, 12 ca Cao), NV01 ôm 17 leads, NV03 ôm 14 leads; trong khi NV08 chỉ có 3 leads (88 triệu). Khi lên 700 leads/tuần, nhóm NV01/NV06/NV03 phải gánh 98–119 leads/tuần (vượt gấp 3–4 lần ngưỡng chịu đựng 30 leads), dẫn tới quá tải (Burnout), ngâm khách quá 24h và rơi rụng khách sang đối thủ.", bold_prefix="3. Vỡ trận SLA & Lệch tải nghiêm trọng giữa 8 NVKD: ")
add_bullet(doc, "Vì thiếu đồng hồ đo thời gian thực và phân loại trạng thái, Trưởng phòng không biết nhân sự nào đang ôm lead quá hạn 24h hay hồ sơ nào bị bỏ quên để kịp thời điều phối lại.", bold_prefix="4. Trưởng phòng rơi vào 'Điểm mù quản trị': ")

doc.add_page_break()

# ==================== PHẦN 2: LỰA CHỌN GIẢI PHÁP THÂN QUEN & DỄ THỰC HIỆN ====================
add_heading_1(doc, "PHẦN 2: LỰA CHỌN GIẢI PHÁP DỄ THỰC HIỆN VÀ THÂN QUEN NHẤT")

add_body(doc, "Để giải quyết triệt để vấn đề cho doanh nghiệp SME có 8 NVKD mà không cần đội ngũ IT lập trình phức tạp hay mua phần mềm CRM đắt đỏ, tôi lựa chọn giải pháp: Xây dựng Hệ thống Điều phối & Giám sát trên LARK BASE kết hợp Công thức dạng Excel và Automation kéo-thả có sẵn.", indent=True)

comp_headers = ["Tiêu chí đánh giá", "Cách cũ (File Excel rời + Zalo)", "Phần mềm CRM lớn (Salesforce/Tự code)", "GIẢI PHÁP CHỌN: LARK BASE (Chuẩn hóa)"]
comp_widths = [Cm(3.2), Cm(4.0), Cm(4.2), Cm(4.6)]
comp_align = [WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT]
comp_data = [
    ["Độ dễ hiểu & Thân quen", "Dễ dùng nhưng rời rạc, ai cũng xem được dẫn đến gọi trùng.", "Giao diện phức tạp, khó hiểu, nhân viên ngại nhập liệu.", "Giao diện 90% giống hệt Excel/Google Sheets, người mới dùng 15 phút là quen."],
    ["Khả năng tự động hóa", "Không có. Phải lọc tay, nhắc miệng qua Zalo.", "Mạnh nhưng tốn hàng trăm triệu thuê IT viết code cài đặt.", "Có sẵn nút Automation kéo-thả & Tin nhắn Bot Lark nhắc việc trực tiếp."],
    ["Kiểm soát tải trọng & SLA", "Không đo được giờ chờ, không biết ai quá tải.", "Đo được nhưng mất nhiều tháng triển khai.", "Dùng hàm IF, COUNTIF quen thuộc để đếm ngược 4h/24h và chặn quá tải 30 leads."],
    ["Chi phí & Thời gian làm", "Miễn phí nhưng mất 14% nhân lực gọi rác.", "Rất đắt, thời gian triển khai 3–6 tháng.", "Miễn phí/Chi phí cực thấp. Thiết lập hoàn chỉnh trong vòng 60 phút."]
]
add_table(doc, "Bảng 2.1: So sánh và lý do lựa chọn nền tảng Lark Base làm giải pháp trọng tâm", comp_headers, comp_data, comp_widths, comp_align, "Đánh giá lựa chọn công cụ BI Nam Việt")

add_callout(doc, "TẠI SAO GIẢI PHÁP NÀY PHÙ HỢP VÀ DỄ THỰC HIỆN VỚI BẠN NHẤT?", [
    ("Tư duy bảng tính quen thuộc", "Toàn bộ 3 bảng dữ liệu (Leads, Nhân sự, Lịch sử) trên Lark Base hoạt động y hệt 3 Sheet trong Excel. Công thức tính giờ chờ SLA hay đếm Backlog dùng cú pháp IF/NOW/COUNTIF giống hệt Excel."),
    ("Tất cả trong một (All-in-One)", "Không cần nối sang Slack hay Zalo phức tạp. Lark Base có sẵn Lark Messenger đi kèm để AI/Automation tự động nhắn tin nhắc nhở NVKD khi quá 4 giờ và báo động đỏ TPKD khi quá 24 giờ."),
    ("Giải quyết cả Con người lẫn Máy móc", "Chức năng Phân quyền nâng cao khóa không cho NVKD nhìn thấy khách của đồng nghiệp (trị tận gốc lỗi con người tự gắp hồ sơ), kết hợp luật tự động sửa lỗi máy móc (chặn VIP ≥ 500tr & chia Miền Trung).")
])

doc.add_page_break()

# ==================== PHẦN 3: HƯỚNG DẪN THỰC HIỆN CHI TIẾT TỪNG BƯỚC ====================
add_heading_1(doc, "PHẦN 3: HƯỚNG DẪN CHI TIẾT TỪNG BƯỚC TRIỂN KHAI TRÊN LARK BASE")

add_body(doc, "Dưới đây là quy trình cầm tay chỉ việc gồm 4 bước tuần tự, bạn chỉ cần làm đúng theo từng mục là hệ thống tự động chạy trơn tru:", indent=True)

# BƯỚC 1
add_heading_2(doc, "Bước 1: Tạo 3 Bảng (3 Tabs) Dữ liệu Chuẩn hóa trong 1 File Lark Base")
add_body(doc, "Tạo một file Lark Base mới tên 'NamViet_CRM_DieuPhoi' gồm 3 bảng (giống 3 sheet Excel) với cấu trúc cột như sau:", indent=True)

b1_headers = ["Tên Bảng (Table)", "Các cột cần tạo thêm (ngoài 14 cột có sẵn)", "Kiểu dữ liệu & Công thức cài đặt"]
b1_widths = [Cm(3.5), Cm(5.5), Cm(7.0)]
b1_align = [WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT]
b1_data = [
    ["Bảng 1:\nDanh_Sach_Leads\n(Chứa dữ liệu khách)",
     "• Key_Nhan_Dien_Trung\n• Gio_Giao_Viec (Assigned_At)\n• Gio_Goi_Dau (First_Call_At)\n• So_Gio_Cho\n• Trang_Thai_SLA\n• NVKD_Bi_Thu_Hoi",
     "• Formula: Ghép MST/Domain hoặc SĐT để phát hiện trùng.\n• Date & Time: Tự điền khi có NVKD.\n• Date & Time: Tự điền khi NVKD đổi trạng thái sang 'Đã gọi'.\n• Formula: IF(ISBLANK(Gio_Goi_Dau), (NOW()-Gio_Giao_Viec)*24, (Gio_Goi_Dau-Gio_Giao_Viec)*24)\n• Formula: Phân nhãn '✅ Đạt', '⏰ Cảnh báo 4h', '🚨 Quá hạn 24h'.\n• User/Text: Lưu tên người làm trễ 24h để xếp hạng vi phạm."],
    ["Bảng 2:\nMa_Tran_8_NVKD\n(Quản lý định mức)",
     "• Ma_NV & Ten_NVKD\n• Nganh_Phu_Trach & Gói SP\n• Tinh_Thanh_Phu_Trach\n• Backlog_Chua_Goi\n• Trang_Thai_Tai\n• So_Lan_Tre_SLA_24h",
     "• Text / User: Danh sách NV01 đến NV08.\n• Multi-select: Gắn đúng gói PROD_RETAIL, FNB, SERVICE, MFG, Đa ngành.\n• Multi-select: Ghi rõ 63 tỉnh (Chia rạch ròi tỉnh Miền Trung vào Bắc hoặc Nam).\n• Rollup (Count): Đếm số dòng bên Bảng 1 có Trạng thái = 'Chưa gọi'.\n• Formula: IF(Backlog_Chua_Goi >= 30, '🔴 QUÁ TẢI', '🟢 CÒN TẢI').\n• Rollup (Count): Đếm số hồ sơ có tên ở cột NVKD_Bi_Thu_Hoi."],
    ["Bảng 3:\nBlacklist_Va_CSKH\n(Ràng buộc lịch sử)",
     "• So_Dien_Thoai / Email\n• Phan_Loai_Lich_Su\n• NVKD_Dang_Cham_Soc\n• Ghi_Chu_Su_Co",
     "• Text / Phone / Email: Khóa tìm kiếm khi khách mới đổ về.\n• Single Select: 'Từ chối/Blacklist' | 'Khách cũ Upsell' | 'Khiếu nại CSKH cũ'.\n• User: Ghi tên NV cũ (VD: NV03) nếu là khách đang có hợp đồng.\n• Text: VD như ca L013 từng khiếu nại dịch vụ."]
]
add_table(doc, "Bảng 3.1: Cấu trúc chi tiết 3 bảng dữ liệu trên Lark Base", b1_headers, b1_data, b1_widths, b1_align, "Thiết kế Schema Lark Base Nam Việt")

# BƯỚC 2
add_heading_2(doc, "Bước 2: Cài đặt 4 Luật Tự động hóa theo Đúng Thứ tự Logic")
add_body(doc, "Vào mục Automation (Tự động hóa) ở góc trên bên phải màn hình Lark Base (hoặc cấu hình trong AI Agent nếu dùng API) và cài đặt tuần tự 4 luồng xử lý sau khi có khách hàng mới vào Bảng 1:", indent=True)

add_bullet(doc, "Khi có Lead mới (L001..L100), hệ thống tra cứu cột SĐT/Email sang Bảng 3 (Blacklist_Va_CSKH) và cột Key_Nhan_Dien_Trung ở Bảng 1. Nếu khớp Blacklist -> Đóng hồ sơ ('Do Not Contact'). Nếu trùng MST/Domain với hồ sơ có sẵn -> Tự động chuyển trạng thái 'ĐÃ GỘP HỒ SƠ' vào hồ sơ Master. Nếu có cờ 'Khiếu nại CSKH' (như L013) -> Chuyển 'CẦN TPKD DUYỆT' và báo TPKD.", bold_prefix="Luật 1 – Lọc Trùng, Blacklist & Lịch sử CSKH: ")
add_bullet(doc, "Với khách mới hợp lệ, hệ thống PHẢI kiểm tra điều kiện Ngân sách ≥ 500 triệu VNĐ TRƯỚC tiên. Nếu ≥ 500 triệu (như L044 - 900tr, L034 - 500tr) -> Gắn nhãn 'VIP - CẦN TPKD DUYỆT', bắn tin nhắn Push Notification cho Trưởng phòng chỉ định Key Account Manager (tuyệt đối không tự đóng). Với hồ sơ < 500 triệu, nếu ngoài phạm vi SMB mới chuyển 'TỪ CHỐI'; nếu hợp lệ thì dò Bảng 2 (Ma_Tran_8_NVKD) khớp đúng Gói SP + Tỉnh thành (đã bao gồm Miền Trung) để tìm người phụ trách. Nếu thiếu thông tin (L008, L071) -> Tự gửi email/tin nhắn khảo sát bổ sung hoặc chia xoay vòng.", bold_prefix="Luật 2 – Đảo Luật VIP ≥ 500tr & Định tuyến Chuẩn Vùng/Ngành: ")
add_bullet(doc, "Trước khi ghi tên NVKD vào hồ sơ, hệ thống kiểm tra cột Trang_Thai_Tai của người đó tại Bảng 2. Nếu Backlog_Chua_Goi >= 30 ('🔴 QUÁ TẢI'), hệ thống kích hoạt Luồng tràn (Overflow): Chỉ giữ lại các hồ sơ mức '1 - CAO' cho NV chuyên trách; các hồ sơ mức '2 - TRUNG BÌNH' và '3 - THẤP' sẽ tự động gán cho nhân sự có Backlog_Chua_Goi thấp nhất trong Bảng 2 (như NV08, NV02) hoặc đẩy thông báo xin ý kiến TPKD.", bold_prefix="Luật 3 – Kiểm soát Tải trọng & San tải Tràn (Trần 30 Leads): ")
add_bullet(doc, "Cài đặt 3 quy tắc Automation có sẵn trên Lark Base: (a) Khi cột NVKD_Phu_Trach vừa được điền -> Tự động ghi thời gian hiện tại vào cột Gio_Giao_Viec. (b) Khi Trang_Thai_SLA chuyển thành '⏰ Cảnh báo 4h' (tức quá 4 tiếng chưa có Gio_Goi_Dau) -> Bot Lark tự nhắn tin riêng nhắc NVKD gọi ngay. (c) Khi Trang_Thai_SLA chuyển thành '🚨 Quá hạn 24h' -> Bot tự ghi tên NV đó vào cột NVKD_Bi_Thu_Hoi, xóa tên ở cột NVKD_Phu_Trach để trả hồ sơ về Kho chung, đồng thời bắn cảnh báo đỏ cho Trưởng phòng.", bold_prefix="Luật 4 – Đồng hồ Đếm ngược SLA 4h & Thu hồi 24h: ")

# BƯỚC 3
add_heading_2(doc, "Bước 3: Đóng gói 2 Giao diện Đầu ra Chuẩn Chuyên môn")

add_heading_3(doc, "3.1. Giao diện cho 8 NVKD (View 'Việc Hôm Nay Của Tôi')")
add_bullet(doc, "Tạo một Grid View mới trên Bảng 1 đặt tên là 'Danh Sách Gọi Hôm Nay'.", bold_prefix="Bước 3.1.1: ")
add_bullet(doc, "Bấm Filter (Lọc): NVKD_Phu_Trach là 'Current User' (Chính người đang đăng nhập) VÀ Trạng thái là 'ĐÃ PHÂN CÔNG'.", bold_prefix="Bước 3.1.2: ")
add_bullet(doc, "Bấm Sort (Sắp xếp): Thứ tự 1 theo 'Mức Ưu Tiên' (A -> Z để 1 - CAO nằm trên cùng); Thứ tự 2 theo 'Ngân Sách' (Giảm dần).", bold_prefix="Bước 3.1.3: ")
add_bullet(doc, "Vào Setting -> Advanced Permissions (Phân quyền nâng cao): Chọn 'Chỉ được xem và sửa các dòng có tên mình ở cột NVKD_Phu_Trach'. Từ đây, 8 nhân viên không bao giờ nhìn thấy hay gọi trùng khách của nhau.", bold_prefix="Bước 3.1.4 (Chống đụng độ): ")

add_heading_3(doc, "3.2. Dashboard Giám sát Toàn cảnh cho Trưởng phòng Kinh doanh")
add_body(doc, "Bấm mục Dashboard trên thanh công cụ Lark Base, bấm 'Add Chart' để kéo thả 4 khối biểu đồ sau từ Bảng 1 và Bảng 2:", indent=True)
add_bullet(doc, "Kéo 4 thẻ Metric Card hiển thị số đếm theo thời gian thực: (1) Tổng Lead trong tuần; (2) Số Lead Đã gộp trùng; (3) Số Lead Từ chối; (4) Số Lead VIP/Ngoại lệ đang chờ TPKD duyệt.", bold_prefix="Khối 1 – Thẻ số liệu tổng quan: ")
add_bullet(doc, "Chọn Stacked Bar Chart: Trục ngang chọn NVKD_Phu_Trach, Trục dọc đếm số bản ghi, Chia màu (Stack) theo Mức_Ưu_Tiên. Giúp TPKD nhìn một giây biết ngay ai đang quá tải gần mức trần 30.", bold_prefix="Khối 2 – Biểu đồ Phân bổ Tải trọng 8 NVKD: ")
add_bullet(doc, "Chọn Bar Chart lấy từ Bảng 2: Trục ngang chọn Ten_NVKD, Trục dọc chọn So_Lan_Tre_SLA_24h. Xếp hạng từ cao xuống thấp để nhắc nhở nhân sự hay ngâm Lead.", bold_prefix="Khối 3 – Bảng Xếp hạng Vi phạm SLA 24h: ")
add_bullet(doc, "Kéo một khung List View nhỏ lọc riêng các hồ sơ có trạng thái 'CẦN TPKD DUYỆT', 'VIP ≥ 500TR' và 'KHO CHUNG - QUÁ HẠN 24H' để Trưởng phòng bấm nút phân công trực tiếp ngay trên Dashboard.", bold_prefix="Khối 4 – Danh sách Cần TPKD Duyệt & Cứu hộ: ")

# BƯỚC 4
add_heading_2(doc, "Bước 4: Quy trình Kiểm tra & Vận hành Hàng tuần (Checklist 5 Phút)")
add_callout(doc, "LỊCH TRÌNH VẬN HÀNH CHUẨN CHO ĐỘI NGŨ NAM VIỆT", [
    ("Mỗi sáng 08:30 (Dành cho 8 NVKD)", "Mở Lark Base -> Bấm đúng Tab 'Danh Sách Gọi Hôm Nay' -> Gọi từ trên xuống dưới theo thứ tự 1-CAO trước -> Gọi xong chuyển trạng thái sang 'Đã gọi' để hệ thống tự ghi giờ và tắt đồng hồ SLA."),
    ("Trong ngày (Tự động 24/7)", "Hệ thống tự gộp trùng, lọc rác, chuyển hướng quá tải khi có NV chạm mốc 30, nhắn tin nhắc ở mốc 4h và thu hồi về kho chung ở mốc 24h."),
    ("09:00 và 15:00 hàng ngày (Dành cho TPKD)", "Mở Pipeline Dashboard -> Xử lý nhanh mục 'Khối 4 - Cần TPKD Duyệt' (duyệt khách VIP ≥500tr, xử lý ca khiếu nại như L013, điều phối lại ca quá hạn 24h) -> Kiểm tra biểu đồ tải trọng để điều chỉnh định mức tuần tiếp theo.")
])

doc.save(OUT_WS)
doc.save(OUT_DL)
print("Saved Docx Report successfully!")
print(f"Size: {os.path.getsize(OUT_WS):,} bytes")

