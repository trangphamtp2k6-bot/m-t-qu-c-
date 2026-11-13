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

OUT_WS = r'c:\Users\HP\Downloads\MÔN AI AGENT\Bao_Cao_Giai_Phap_Google_Sheets_Nam_Viet.docx'
OUT_DL = r'C:\Users\HP\Downloads\Bao_Cao_Giai_Phap_Google_Sheets_Nam_Viet.docx'

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

# ==================== HEADER BLOCK ====================
tbl_hdr = doc.add_table(rows=1, cols=1)
tbl_hdr.alignment = WD_TABLE_ALIGNMENT.CENTER
tbl_hdr.autofit = False
cell = tbl_hdr.cell(0, 0)
cell.width = Cm(16.0)
set_cell_background(cell, HEX_LIGHT_BG)
set_cell_margins(cell, top=140, bottom=140, left=180, right=180)
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
r0 = p0.add_run("CÔNG TY CỔ PHẦN PHẦN MỀM NAM VIỆT • TÀI LIỆU TRIỂN KHAI BI & VẬN HÀNH")
format_run(r0, size_pt=11, bold=True, color=COLOR_SECONDARY)

p1 = cell.add_paragraph()
p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
p1.paragraph_format.space_after = Pt(4)
r1 = p1.add_run("GIẢI PHÁP ĐIỀU PHỐI LEADS & GIÁM SÁT SLA BẰNG\nGOOGLE SHEETS + APPS SCRIPT + TELEGRAM / GMAIL")
format_run(r1, size_pt=14, bold=True, color=COLOR_PRIMARY)

p2 = cell.add_paragraph()
p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
p2.paragraph_format.space_after = Pt(0)
r2 = p2.add_run("Phương án 1 (Miễn phí 100% • Thân quen như Excel • Tự động hóa 4 bước & Chia tải 8 NVKD)")
format_run(r2, size_pt=11, italic=True, color=COLOR_TEXT)

doc.add_paragraph().paragraph_format.space_after = Pt(2)

# ==================== PHẦN 1: TÓM TẮT NGẮN GỌN CHẨN ĐOÁN & ẢNH HƯỞNG ====================
add_heading_1(doc, "PHẦN 1: TÓM TẮT CHẨN ĐOÁN LỖI VÀ TÁC ĐỘNG ĐẾN DOANH NGHIỆP")

add_body(doc, "Từ kết quả kiểm toán 100 hồ sơ mẫu (trong quy mô 700 Leads/tuần của Công ty Nam Việt), các bất cập hiện tại bắt nguồn từ hai phía:", indent=True)
add_bullet(doc, "Marketing chuyển danh sách thô chưa gộp trùng (7% trùng lặp như An Phước, Golden Coffee); NVKD tự đọc danh sách và tự ý 'gắp' hồ sơ ngân sách lớn dù ngoài gói SMB; Ma trận phân quyền của Quản lý bỏ sót quy định ranh giới Miền Trung (4 ca giao sai vùng: L055, L084, L090, L100) và thiếu định mức trần 30 leads/người.", bold_prefix="• Lỗi Con người & Quy trình (~70%): ")
add_bullet(doc, "Hệ thống đặt sai thứ tự câu lệnh (lọc 'Ngoài phạm vi SMB' trước khi bắt 'VIP ≥ 500 triệu', làm đóng nhầm L044 - 900tr và L034 - 500tr); thiếu cột lưu mốc thời gian giao/gọi nên máy không thể đếm ngược SLA 4h/24h.", bold_prefix="• Lỗi Máy móc & Thiết lập (~30%): ")

add_body(doc, "Nếu không xử lý, doanh nghiệp chịu 4 thiệt hại trực tiếp: (1) Đánh rơi 1,4 tỷ VNĐ cơ hội VIP trong 100 mẫu (~9,8 tỷ/tuần ở quy mô 700 leads); (2) Đốt 14% công suất Sales vào leads trùng/rác và làm mất uy tín do gọi đụng khách; (3) Vỡ trận quá tải khi lên 700 leads/tuần (NV01 gánh ~119 leads, NV06 ~105 leads, NV03 ~98 leads vượt xa trần 30, trong khi NV08 chỉ có 21 leads); (4) TPKD mù thông tin tiến độ gọi.", indent=True)

add_figure(doc, r'c:\Users\HP\Downloads\MÔN AI AGENT\chart_namviet_workload.png', "Hình 1.1: Mất cân bằng tải trọng 8 NVKD & Dự báo vỡ trần 30 Leads ở quy mô 700 Leads/tuần", "Nguồn: Kiểm toán dữ liệu Nam Việt", width=Cm(14.5))

# ==================== PHẦN 2: KIẾN TRÚC GIẢI PHÁP PHƯƠNG ÁN 1 ====================
add_heading_1(doc, "PHẦN 2: KIẾN TRÚC GIẢI PHÁP GOOGLE SHEETS + APPS SCRIPT + TELEGRAM / GMAIL")

add_body(doc, "Phương án 1 tận dụng 100% các công cụ miễn phí và quen thuộc (giống Excel và nhắn tin hàng ngày) nhưng giải quyết triệt để cả lỗi con người lẫn máy móc thông qua kiến trúc '1 File Tổng – 8 File Con – 1 Động cơ Script':", indent=True)

arch_headers = ["Thành phần hệ thống", "Ai sử dụng / Vai trò", "Cơ chế giải quyết triệt để vấn đề"]
arch_widths = [Cm(4.0), Cm(4.2), Cm(7.8)]
arch_align = [WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT]
arch_data = [
    ["1. File Google Sheets Tổng\n(NamViet_CRM_Master)", "Chỉ Trưởng phòng (TPKD) & Admin xem", "Chứa 100% dữ liệu nguồn, Ma trận 8 NVKD, Blacklist và Dashboard 4 biểu đồ giám sát toàn cục."],
    ["2. 8 File Google Sheets Con\n(Mỗi NVKD 1 file riêng)", "Mỗi NVKD chỉ mở được file của mình", "Chặn tuyệt đối việc NVKD nhìn trộm hoặc tự gắp hồ sơ của đồng nghiệp. Danh sách mỗi sáng đã xếp sẵn 1-CAO lên đầu."],
    ["3. Google Apps Script\n(Tệp .gs chạy ngầm miễn phí)", "Chạy tự động mỗi 15 phút (24/7)", "Thực thi đúng 4 bước: Lọc trùng/Blacklist -> Đảo luật VIP ≥ 500tr -> San tải Overflow trần 30 -> Đếm ngược SLA 4h/24h."],
    ["4. Bot Telegram & Gmail\n(Kênh cảnh báo tức thì)", "Nhắn tin thẳng vào điện thoại NVKD & TPKD", "Nhắc NVKD khi Lead quá 4h chưa gọi; bắn báo động đỏ cho TPKD khi thu hồi Lead quá 24h hoặc khi có khách VIP ≥ 500tr."]
]
add_table(doc, "Bảng 2.1: Kiến trúc tổng thể Phương án 1 (Google Sheets Ecosystem)", arch_headers, arch_data, arch_widths, arch_align, "Thiết kế kiến trúc giải pháp Nam Việt")

# ==================== PHẦN 3: HƯỚNG DẪN CHI TIẾT TỪNG BƯỚC ====================
add_heading_1(doc, "PHẦN 3: HƯỚNG DẪN THỰC HIỆN CHI TIẾT TỪNG BƯỚC CỤ THỂ NHẤT")

# BƯỚC 1
add_heading_2(doc, "Bước 1: Tạo File Google Sheets Tổng (NamViet_CRM_Master) gồm 4 Tab")
add_body(doc, "Truy cập Google Drive, tạo một Google Sheets mới tên là 'NamViet_CRM_Master' và tạo đúng 4 Tab (Sheet) bên dưới:", indent=True)

add_heading_3(doc, "1.1. Tab 'Leads_Master' (Chứa danh sách khách hàng)")
add_body(doc, "Giữ nguyên 14 cột (A đến N) từ bảng chuẩn hóa hiện có của bạn, và thêm tiếp 5 cột (O đến S) ở cuối bảng để theo dõi đồng hồ SLA:", indent=True)

col_headers = ["Cột", "Tên cột trên Sheet", "Mục đích & Công thức nhập sẵn"]
col_widths = [Cm(1.5), Cm(4.5), Cm(10.0)]
col_align = [WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT]
col_data = [
    ["A - N", "14 cột hồ sơ hiện có", "Mã Hồ Sơ, Đại diện, Doanh nghiệp, MST/Domain, SĐT, Email, Tỉnh, Ngành, Gói SP, Ngân sách, Ưu tiên, NVKD Phụ Trách, Trạng thái, Hành động."],
    ["O", "Gio_Giao_Viec", "Để trống. Apps Script tự động đóng dấu ngày giờ khi phân công cho NVKD."],
    ["P", "Gio_Goi_Dau", "Để trống. Script tự động điền giờ ngay khi NVKD bấm 'Đã gọi' ở file con."],
    ["Q", "So_Gio_Cho (Formula)", "Dán công thức ở Q2 kéo xuống:\n=IF(O2=\"\"; \"\"; IF(P2=\"\"; (NOW()-O2)*24; (P2-O2)*24))"],
    ["R", "NVKD_Bi_Thu_Hoi", "Lưu tên NVKD bị thu hồi nếu quá 24h không gọi (để vẽ bảng xếp hạng vi phạm SLA)."],
    ["S", "Da_Gui_Nhac_4h", "Lưu cờ 'ĐÃ NHẮC 4H' để Bot không gửi tin nhắn spam lặp lại nhiều lần."]
]
add_table(doc, "Bảng 3.1: Cấu trúc cột mở rộng cho Tab Leads_Master", col_headers, col_data, col_widths, col_align, "Cấu trúc chuẩn hóa Google Sheets")

add_heading_3(doc, "1.2. Tab 'Ma_Tran_8_NVKD' (Quy định phạm vi & Đếm tải trọng)")
add_body(doc, "Tạo bảng gồm 8 dòng (dòng 2 đến dòng 9) tương ứng 8 NVKD với 8 cột từ A đến H:", indent=True)

nv_headers = ["Cột", "Tên cột", "Dữ liệu mẫu & Công thức điền vào"]
nv_widths = [Cm(1.5), Cm(4.0), Cm(10.5)]
nv_align = [WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT]
nv_data = [
    ["A", "NVKD_Phu_Trach", "NV01 - Nguyễn Văn An (Chép chính xác tên như cột L bên Leads_Master)."],
    ["B", "Goi_SP_Phu_Trach", "PROD_RETAIL | PROD_FNB | PROD_SERVICE | PROD_MFG | ĐA_NGANH"],
    ["C", "Tinh_Thanh_Phu_Trach", "Ghi rõ tất cả các tỉnh (Vá lỗi Miền Trung: Thanh Hóa/Nghệ An/Hà Tĩnh vào nhóm Bắc; Đà Nẵng/Huế/Quảng Ngãi/Khánh Hòa vào nhóm Nam hoặc ghi 'Toàn quốc')."],
    ["D", "Telegram_Chat_ID", "Điền mã Chat ID Telegram của nhân viên đó (xem Bước 3)."],
    ["E", "Email_NVKD", "Điền Gmail của nhân viên để nhận cảnh báo song song."],
    ["F", "ID_File_Sheet_Con", "Dán đoạn mã ID của file Google Sheets riêng dành cho nhân viên đó (xem Bước 2)."],
    ["G", "Backlog_Hien_Tai", "Công thức đếm tự động số khách chưa gọi:\n=COUNTIFS(Leads_Master!L:L; A2; Leads_Master!M:M; \"ĐÃ PHÂN CÔNG\")"],
    ["H", "So_Lan_Vi_Pham_24h", "Công thức đếm số lần trễ bị thu hồi:\n=COUNTIF(Leads_Master!R:R; A2)"]
]
add_table(doc, "Bảng 3.2: Cấu trúc Tab Ma_Tran_8_NVKD trên Google Sheets", nv_headers, nv_data, nv_widths, nv_align, "Ma trận phân quyền động Nam Việt")

add_heading_3(doc, "1.3. Tab 'Blacklist_CSKH' (Kiểm soát Bước 1)")
add_body(doc, "Tạo 4 cột đơn giản: Cột A (`SĐT`), Cột B (`Email`), Cột C (`Phân Loại`: điền 'Blacklist' | 'Khách cũ' | 'Khiếu nại CSKH'), Cột D (`NVKD Đang Chăm Sóc`). Nhập sẵn SĐT/Email của ca L013 vào đây với nhãn 'Khiếu nại CSKH' để test ngay.", indent=True)

# BƯỚC 2
add_heading_2(doc, "Bước 2: Tạo 8 File Google Sheets Con cho 8 NVKD (Chống gọi trùng 100%)")
add_bullet(doc, "Trên Google Drive, tạo 8 file Google Sheets mới đặt tên lần lượt: `Worklist_NV01_An`, `Worklist_NV02_Binh`, ..., `Worklist_NV08_Nam`.", bold_prefix="1. Tạo 8 file riêng: ")
add_bullet(doc, "Mở từng file lên, bấm nút 'Chia sẻ' (Share) -> chỉ cấp quyền 'Người chỉnh sửa' cho đúng email của nhân viên đó (tuyệt đối không share cho 7 người còn lại).", bold_prefix="2. Phân quyền chặt chẽ: ")
add_bullet(doc, "Nhìn lên thanh địa chỉ trình duyệt của mỗi file, copy đoạn mã nằm giữa `/d/` và `/edit` (ví dụ: `1AbCdEfG...xyz`) rồi dán vào Cột F (`ID_File_Sheet_Con`) ở Tab `Ma_Tran_8_NVKD` trong File Tổng.", bold_prefix="3. Lấy ID File dán về File Tổng: ")
add_bullet(doc, "Trong mỗi file con, tại cột K ('Trạng Thái Gọi'), tạo Data Validation (Xác thực dữ liệu) dạng danh sách thả xuống gồm 2 lựa chọn: `Chưa gọi` và `Đã gọi` để nhân viên chỉ cần bấm chọn không phải gõ tay.", bold_prefix="4. Tạo nút bấm cho NVKD: ")

# BƯỚC 3
add_heading_2(doc, "Bước 3: Tạo Bot Telegram Cảnh báo Miễn phí (Trong 2 phút)")
add_body(doc, "Để hệ thống tự động nhắn tin nhắc nhở sau 4 giờ và báo đỏ TPKD sau 24 giờ miễn phí (thay thế hoàn toàn cho Bot trả phí):", indent=True)
add_bullet(doc, "Mở Telegram, tìm kiếm tài khoản có tích xanh `@BotFather` -> Bấm Start -> Gõ lệnh `/newbot` -> Đặt tên Bot (VD: `NamViet CRM Bot`) -> Copy đoạn `HTTP API Token` mà BotFather trả về.", bold_prefix="1. Tạo Bot lấy Token: ")
add_bullet(doc, "Yêu cầu 8 NVKD và TPKD mở con Bot vừa tạo bấm `/start`, sau đó tìm `@userinfobot` bấm Start để lấy dãy số `Id` cá nhân (VD: `123456789`). Điền các mã ID này vào Cột D (`Telegram_Chat_ID`) trong Tab `Ma_Tran_8_NVKD`.", bold_prefix="2. Lấy Chat ID của nhân sự: ")
add_bullet(doc, "Lưu ý: Trong kịch bản code đi kèm, ngoài Telegram, hệ thống cũng tự động gửi Gmail song song qua hàm `MailApp` có sẵn của Google, nên kể cả khi không dùng Telegram thì email cảnh báo vẫn chạy bình thường.", bold_prefix="3. Kênh dự phòng Gmail: ")

# BƯỚC 4
add_heading_2(doc, "Bước 4: Cài đặt Động cơ Tự động hóa (Google Apps Script)")
add_body(doc, "Tôi đã viết sẵn toàn bộ tệp mã nguồn chuẩn hóa mang tên `NamViet_AutoRouting_SLA.gs` (đính kèm cùng thư mục báo cáo). Bạn chỉ cần thực hiện 3 thao tác copy-paste:", indent=True)
add_bullet(doc, "Tại File Google Sheets Tổng (`NamViet_CRM_Master`), trên thanh menu bấm: Tiện ích mở rộng (Extensions) -> Apps Script.", bold_prefix="1. Mở trình cài đặt: ")
add_bullet(doc, "Xóa hết chữ mẫu trong màn hình, copy toàn bộ nội dung file `NamViet_AutoRouting_SLA.gs` dán vào. Thay thế 3 dòng đầu tiên ở mục `CONFIG` (Bot Token, Telegram ID của TPKD, Email TPKD) rồi bấm biểu tượng Đĩa mềm (Save).", bold_prefix="2. Dán mã nguồn & Điền Token: ")
add_bullet(doc, "Ở cột menu bên trái của Apps Script, bấm biểu tượng Đồng hồ (Triggers / Trình kích hoạt) -> Bấm '+ Thêm trình kích hoạt' ở góc dưới bên phải -> Chọn hàm: `runLeadAutomation` | Nguồn sự kiện: `Theo thời gian (Time-driven)` | Loại trình kích hoạt: `Bộ hẹn giờ theo phút` | Khoảng thời gian: `15 phút một lần` -> Bấm Lưu (Lưu ý cấp quyền truy cập Google Sheets/Gmail ở lần chạy đầu tiên).", bold_prefix="3. Bật đồng hồ chạy tự động 24/7: ")

# BƯỚC 5
add_heading_2(doc, "Bước 5: Dựng Tab 'Dashboard_TPKD' ngay trên Google Sheets")
add_body(doc, "Để Trưởng phòng Kinh doanh có góc nhìn toàn cảnh thời gian thực ngay trong File Tổng, tạo Tab thứ 4 tên `Dashboard_TPKD` và dùng 4 thao tác Insert -> Chart:", indent=True)

dash_headers = ["Vị trí Dashboard", "Công thức tổng hợp dữ liệu nguồn", "Loại biểu đồ chọn trong Google Sheets"]
dash_widths = [Cm(4.0), Cm(7.5), Cm(4.5)]
dash_align = [WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT]
dash_data = [
    ["Khối 1: Thẻ Chỉ số Phễu\n(Đầu vào / Gộp / Từ chối)",
     "Dùng hàm COUNTIF đếm trực tiếp từ cột Trạng Thái (M) bên Leads_Master (Tổng Lead, Đã phân công, Đã gộp, Từ chối, Cần TPKD duyệt).",
     "Scorecard Chart (Thẻ số lớn)."],
    ["Khối 2: Biểu đồ Tải trọng 8 NVKD",
     "Tạo bảng phụ dùng `COUNTIFS` đếm theo từng NVKD x Mức Ưu Tiên (1-Cao, 2-TB, 3-Thấp) + cột Backlog hiện tại.",
     "Stacked Column Chart (Biểu đồ cột chồng) để thấy ngay ai chạm trần 30."],
    ["Khối 3: Xếp hạng Vi phạm SLA 24h",
     "Lấy thẳng 2 cột A (`NVKD`) và H (`So_Lan_Vi_Pham_24h`) từ Tab `Ma_Tran_8_NVKD`, sắp xếp giảm dần.",
     "Bar Chart (Biểu đồ thanh ngang xếp hạng từ trên xuống)."],
    ["Khối 4: Danh sách TPKD Cần Duyệt",
     "Dùng hàm `=FILTER(Leads_Master!A2:N; Leads_Master!M2:M=\"CẦN TPKD DUYỆT\")` để tự động hiện lên danh sách VIP ≥500tr, ca khiếu nại L013 và ca bị thu hồi 24h.",
     "Bảng danh sách động (Dynamic Table) hiển thị ngay bên dưới biểu đồ."]
]
add_table(doc, "Bảng 3.3: Hướng dẫn dựng 4 khối Dashboard cho Trưởng phòng trên Google Sheets", dash_headers, dash_data, dash_widths, dash_align, "Thiết kế Pipeline Dashboard Google Sheets")

# BƯỚC 6
add_heading_2(doc, "Bước 6: Quy trình Vận hành Thực tế Hàng ngày")
add_callout(doc, "CÁCH VẬN HÀNH TRƠN TRU SAU KHI CÀI ĐẶT XONG", [
    ("Dành cho 8 NVKD (Mỗi sáng 08:30)", "Mở duy nhất File Sheet con của mình. Danh sách khách đã được Script lọc sạch trùng và xếp sẵn các ca '1 - CAO' lên đầu. Gọi xong khách nào, bấm chuyển cột K từ 'Chưa gọi' sang 'Đã gọi'. Trong vòng 15 phút, Script tự đồng bộ giờ gọi về File Tổng và tắt đồng hồ SLA."),
    ("Khi quên gọi (Tự động 100%)", "Quá 4 tiếng không bấm 'Đã gọi' -> Bot Telegram/Gmail tự động nhắn tin nhắc nhở NVKD. Quá 24 tiếng -> Lead tự động biến mất khỏi File con của nhân viên đó, trả về Kho chung và bắn cảnh báo đỏ cho Trưởng phòng."),
    ("Dành cho Trưởng phòng (TPKD)", "Mở Tab 'Dashboard_TPKD' trên File Tổng để theo dõi biến động tải trọng của 8 nhân sự và xử lý danh sách VIP ≥ 500 triệu hoặc ca khiếu nại CSKH hiện sẵn ở Khối 4.")
])

doc.save(OUT_WS)
doc.save(OUT_DL)
print("Saved PA1 Solution Docx successfully!")
print(f"Size: {os.path.getsize(OUT_WS):,} bytes")

