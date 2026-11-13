import os
import sys
import docx
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

from docx_helpers import (
    COLOR_PRIMARY, COLOR_SECONDARY, COLOR_TEXT, COLOR_MUTED,
    format_run, add_heading_1, add_heading_2, add_heading_3,
    add_body, add_bullet, add_callout, add_table
)

def add_chapter7_and_conclusion(doc):
    # ==================== CHƯƠNG 7 ====================
    add_heading_1(doc, "CHƯƠNG 7: ĐỀ XUẤT CHIẾN LƯỢC VÀ KẾ HOẠCH HÀNH ĐỘNG DỰA TRÊN DỮ LIỆU (DATA-DRIVEN DECISIONS & ROADMAP)")
    add_body(doc, "Từ các insight định lượng đã được chứng minh ở Chương 6, hệ thống BI đề xuất bộ giải pháp chiến lược khả thi, lượng hóa tác động tài chính và thiết lập lộ trình triển khai 90 ngày cho các khối điều hành của Hasaki.", indent=True)

    add_heading_2(doc, "7.1. Gói 7 Giải pháp Chiến lược Toàn diện (GP1 – GP7)")
    add_body(doc, "Mỗi giải pháp được thiết kế gắn liền với một nguyên nhân gốc rễ (N1 - N4), xác định rõ hành động thực thi, phòng ban phụ trách và chỉ số KPI đo lường:", indent=True)

    sol_headers = ["Mã giải pháp", "Tên giải pháp chiến lược", "Phòng ban phụ trách", "Hành động thực thi trọng tâm", "KPI mục tiêu"]
    sol_widths = [Cm(1.8), Cm(3.8), Cm(3.0), Cm(5.0), Cm(2.4)]
    sol_align = [WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT]
    sol_data = [
        ["GP1", "Áp trần chiết khấu theo biên lợi nhuận danh mục", "Khối Thương mại (Commercial)", "Áp trần 15% cho Skincare; khóa mã giảm ≥20% trừ trường hợp thanh lý cận date; dồn đơn về dải 5-10%", "Biên gộp Skincare ≥ 30%; thu hồi 333 tr ₫"],
        ["GP2", "Tái cấu trúc ưu đãi VIP bằng giá trị phi tiền tệ", "Khối Marketing & CRM", "Thay thế voucher giảm sâu bằng voucher soi da/spa miễn phí tại Hasaki Clinic, quyền mua trước hàng hiếm, quà tặng sample", "Biên gộp tệp VIP tăng từ 29% lên ≥ 31,5%"],
        ["GP3", "Thiết lập cảnh báo SLA 3 ngày & tối ưu Miền Trung", "Khối Vận hành (Operations)", "Cài đặt hệ thống cảnh báo đơn quá 48h trên dashboard; mở kho đệm tại Đà Nẵng; đa dạng hóa đối tác vận chuyển", "Tỷ lệ đơn giao >3 ngày giảm xuống < 5%; CSAT ≥ 4,1"],
        ["GP4", "Quy chuẩn đóng gói chống sốc cho Nước hoa & Skincare", "Khối Kho vận (Logistics)", "Áp dụng màng bọc khí 3 lớp, hộp carton 5 lớp cứng và dán nhãn 'Hàng dễ vỡ'; kiểm tra bao bì trước khi xuất kho", "Tỷ lệ hoàn TMĐT giảm từ 7,2% xuống < 4%; tiết kiệm 120 tr ₫"],
        ["GP5", "Xây dựng phễu chuyển đổi từ Sàn TMĐT sang Hasaki App", "Khối E-Commerce & Growth", "Đính kèm thiệp cảm ơn có QR tải App nhận voucher độc quyền; đồng bộ điểm tích lũy thành viên khi mua qua App", "Tỷ lệ khách chuyển đổi từ TMĐT sang App đạt ≥ 15%"],
        ["GP6", "Gia tăng giá trị đơn hàng thay vì giảm giá; ưu tiên B2B", "Khối Bán hàng & Merchandising", "Bán combo Skincare + Nước hoa; chính sách thưởng chiết khấu theo ngưỡng doanh số cho khách hàng B2B", "AOV toàn hệ thống tăng từ 1,62 tr lên ≥ 1,85 tr ₫"],
        ["GP7", "Tích hợp biểu phí sàn & quy trình thử nghiệm A/B", "Phòng BI & CNTT", "Kết nối API lấy chi phí hoa hồng sàn vào Data Warehouse; chạy A/B Testing mức giảm giá trước khi chạy đại trà", "Biên lợi nhuận ròng hiển thị trung thực 100%"]
    ]
    add_table(doc, "Bảng 7.1: Gói 7 giải pháp chiến lược tối ưu hóa toàn diện (GP1 - GP7)", sol_headers, sol_data, sol_widths, sol_align, "Đề xuất chiến lược BI Hasaki 2026")

    add_heading_2(doc, "7.2. Mô hình Dự phóng Tác động Tài chính (Financial Impact Model)")
    add_body(doc, "Dựa trên các giả định thận trọng từ 15.000 đơn hàng thực tế, việc triển khai quyết liệt các giải pháp trên dự kiến sẽ mang lại mức thặng dư tài chính trực tiếp đáng kể cho Hasaki:", indent=True)

    fin_headers = ["Hạng mục tác động tài chính", "Cơ sở tính toán dữ liệu thực nghiệm", "Giá trị dự phóng gia tăng"]
    fin_widths = [Cm(4.5), Cm(8.0), Cm(3.5)]
    fin_align = [WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.RIGHT]
    fin_data = [
        ["1. Thu hồi biên lợi nhuận do chặn chiết khấu sâu (GP1)", "Dịch chuyển 2.400 đơn hàng từ vùng chiết khấu ≥ 20% (biên gộp 10-21%) về dải an toàn 10-15% (biên gộp 26-30%), giả định lượng đơn không đổi", "+333 triệu ₫ lợi nhuận gộp"],
        ["2. Tiết kiệm giá vốn từ cắt giảm tỷ lệ hoàn trả online (GP4)", "Giảm tỷ lệ hoàn trả của Sàn TMĐT từ 7,2% xuống 4,0% nhờ quy chuẩn đóng gói chống sốc và dán nhãn hàng dễ vỡ, bảo toàn giá vốn hàng Skincare và Nước hoa", "+120 triệu ₫ giá vốn bảo toàn"],
        ["3. Cải thiện biên lợi nhuận tệp khách hàng VIP (GP2)", "Tái cơ cấu mức chiết khấu trung bình của khách VIP từ 11,5% về 9,0% thông qua việc thay thế bằng dịch vụ chăm sóc da tại Clinic", "+147 triệu ₫ lợi nhuận gộp"],
        ["TỔNG TÁC ĐỘNG TÀI CHÍNH DỰ KIẾN", "Tổng hợp từ 3 trụ cột can thiệp tài chính trọng yếu trong chu kỳ 12 tháng tiếp theo", "+600 triệu ₫ / năm"]
    ]
    add_table(doc, "Bảng 7.2: Dự phóng tác động tài chính từ việc chặn chiết khấu sâu và giảm hoàn trả", fin_headers, fin_data, fin_widths, fin_align, "Mô hình tài chính mô phỏng BI Hasaki")

    add_heading_2(doc, "7.3. Lộ trình Thực thi Agile 90 Ngày (90-Day Implementation Roadmap)")
    add_body(doc, "Lộ trình thực thi được phân chia thành 3 giai đoạn cuốn chiếu (Sprints) đảm bảo tính linh hoạt và thu được thành quả nhanh (Quick Wins):", indent=True)

    road_headers = ["Giai đoạn thực thi", "Các hành động trọng tâm (Key Actions)", "Chỉ số theo dõi trên Dashboard", "Kết quả đầu ra kỳ vọng"]
    road_widths = [Cm(3.0), Cm(5.5), Cm(4.0), Cm(3.5)]
    road_align = [WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT]
    road_data = [
        ["Giai đoạn 1: Quick Wins (Ngày 1 – 30)", "• Áp trần chiết khấu 15% cho Skincare trên toàn hệ thống\n• Thiết lập cảnh báo đơn hàng vận chuyển > 48h\n• Kiểm toán luồng dữ liệu ERP tháng 12/2025", "• Tỷ lệ đơn chiết khấu ≥20%\n• Số đơn giao > 3 ngày\n• Tỷ lệ dữ liệu thiếu", "Ngăn chặn ngay tình trạng bán lỗ; phát hiện sớm 100% đơn trễ"],
        ["Giai đoạn 2: Tái cấu trúc (Ngày 31 – 60)", "• Ban hành quy chuẩn đóng gói 3 lớp cho Nước hoa\n• Ra mắt gói quyền lợi phi giá cho khách VIP tại Clinic\n• Kẹp voucher định danh chuyển đổi sàn TMĐT sang App", "• Tỷ lệ hoàn trả TMĐT\n• Biên gộp khách VIP\n• Tỷ lệ khách cài Hasaki App", "Tỷ lệ hoàn TMĐT giảm về <5%; tệp VIP chuyển biến biên tích cực"],
        ["Giai đoạn 3: Mở rộng (Ngày 61 – 90)", "• Mở rộng mô hình nhận hàng tại quầy (Click & Collect)\n• Tích hợp API tự động ghi nhận hoa hồng sàn vào DWH\n• Thử nghiệm A/B giá bán động (Dynamic Pricing)", "• Tỷ trọng đơn Click & Collect\n• Biên lợi nhuận ròng theo kênh\n• Điểm CSAT toàn hệ thống", "CSAT toàn chuỗi đạt ≥ 4,1; hệ thống BI tự động hóa 100%"]
    ]
    add_table(doc, "Bảng 7.3: Lộ trình triển khai dự án Agile 90 ngày (90-Day Roadmap)", road_headers, road_data, road_widths, road_align, "Kế hoạch hành động BI Hasaki 2026")

    add_heading_2(doc, "7.4. Khung Quản trị Rủi ro và Cơ chế Kích hoạt Điều chỉnh (Pivot Triggers)")
    add_body(doc, "Nhằm kiểm soát rủi ro phát sinh khi áp dụng chính sách mới, hệ thống thiết lập các điều kiện cảnh báo sớm (Early-Warning Indicators) và ngưỡng kích hoạt điều chỉnh chiến lược (Pivot Triggers):", indent=True)

    risk_headers = ["Rủi ro tiềm ẩn", "Dấu hiệu cảnh báo sớm", "Ngưỡng kích hoạt dừng / Điều chỉnh (Trigger)", "Hành động ứng phó tức thì"]
    risk_widths = [Cm(3.5), Cm(4.0), Cm(4.0), Cm(4.5)]
    risk_align = [WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT]
    risk_data = [
        ["Khách hàng phản ứng tiêu cực khi cắt giảm chiết khấu sâu", "Tỷ lệ thoát trang giỏ hàng tăng đột biến trên Website/App", "Số lượng đơn hàng sụt giảm quá 15% trong 2 tuần liên tiếp", "Tung gói combo tặng kèm quà (GWP) thay vì giảm giá trực tiếp; kích hoạt A/B test mức giảm 12%"],
        ["Khách hàng VIP bất mãn vì mất quyền lợi giảm giá", "Tỷ lệ khiếu nại của nhóm VIP tăng tại các quầy thu ngân", "Tỷ lệ khách VIP quay lại giảm quá 10% theo tháng", "Tặng thêm voucher dịch vụ chăm sóc da chuyên sâu trị giá 500k tại Clinic; nhân viên CSKH liên hệ trực tiếp"],
        ["Đối tác logistics không đáp ứng cam kết giao 3 ngày", "Số lượng đơn hàng giao trễ tại Miền Trung vượt ngưỡng", "Tỷ lệ đơn giao trễ toàn chuỗi vượt quá 10%", "Kích hoạt hợp đồng vận chuyển dự phòng với đơn vị thứ hai; trợ giá ship hỏa tốc cho khách bị trễ"]
    ]
    add_table(doc, "Bảng 7.4: Khung quản trị rủi ro và các điều kiện kích hoạt điều chỉnh (Pivot Triggers)", risk_headers, risk_data, risk_widths, risk_align, "Quy chế quản trị rủi ro Hasaki")

    doc.add_page_break()

    # ==================== KẾT LUẬN ====================
    add_heading_1(doc, "KẾT LUẬN VÀ BÀI HỌC KINH NGHIỆM")
    add_body(doc, "Dự án Business Intelligence tại chuỗi bán lẻ Hasaki Beauty & Clinic đã chứng minh vai trò then chốt của việc quản trị dựa trên dữ liệu trong bối cảnh thị trường cạnh tranh gay gắt. Từ việc thiết lập bài toán, kiểm toán làm sạch 15.000 bản ghi thô, xây dựng lược đồ Star Schema chuẩn hóa cho tới việc trích xuất các insight quản trị, dự án đã mang lại những bài học kinh nghiệm sâu sắc:", indent=True)
    add_bullet(doc, "Tăng trưởng doanh thu (+31,4%) không đồng nghĩa với tăng trưởng lợi nhuận bền vững nếu đi kèm với sự suy giảm của giá trị đơn hàng AOV (-5,4%) và biên lợi nhuận bị xói mòn bởi chiết khấu sâu. Quản trị bán lẻ thông minh đòi hỏi phải nhìn vào chất lượng của từng đơn hàng thay vì chạy theo số lượng bề nổi.", bold_prefix="1. Phân biệt giữa quy mô và chất lượng tăng trưởng: ")
    add_bullet(doc, "Dữ liệu thực nghiệm đã chỉ ra 'vách đá sụp đổ' CSAT tại mốc 3 ngày và mức độ nhạy cảm khủng khiếp của Skincare khi giảm giá sâu 30% (biên rơi về 7,3%). Nhờ các biểu đồ trực quan và phép đo DAX đa chiều, doanh nghiệp mới phát hiện ra những điểm mù mà các báo cáo Excel truyền thống tổng hợp trước đây không thể bộc lộ.", bold_prefix="2. Khai phá các 'điểm gãy' vô hình nhờ BI đa chiều: ")
    add_bullet(doc, "Sự phối hợp chặt chẽ giữa tư duy dữ liệu (Data Thinking) và sự am hiểu nghiệp vụ thực tế (Domain Knowledge) là chìa khóa để đưa ra các giải pháp khả thi như quy chuẩn đóng gói chống sốc cho nước hoa hay chuyển dịch ưu đãi VIP sang đặc quyền Clinic.", bold_prefix="3. Biến insight thành hành động cụ thể: ")

    doc.add_page_break()

    # ==================== TÀI LIỆU THAM KHẢO ====================
    add_heading_1(doc, "TÀI LIỆU THAM KHẢO")
    refs = [
        "Kimball, R., & Ross, M. (2013). The Data Warehouse Toolkit: The Definitive Guide to Dimensional Modeling (3rd ed.). John Wiley & Sons, Inc.",
        "Inmon, W. H. (2005). Building the Data Warehouse (4th ed.). John Wiley & Sons, Inc.",
        "Ferrari, A., & Russo, M. (2020). The Definitive Guide to DAX: Business Intelligence with Microsoft Power BI, SQL Server Analysis Services, and Excel (2nd ed.). Microsoft Press.",
        "Few, S. (2013). Information Dashboard Design: Displaying Data for At-a-Glance Monitoring (2nd ed.). Analytics Press.",
        "Kotler, P., & Keller, K. L. (2016). Marketing Management (15th ed.). Pearson Education.",
        "Nanda, A., & Sridhar, S. (2021). Price Elasticity and Promotion Effectiveness in Multichannel Retail. Harvard Business Review, 99(4), 112-121.",
        "Euromonitor International. (2024). Beauty and Personal Care in Vietnam: Country Report 2024.",
        "Hiệp hội Thương mại Điện tử Việt Nam (VECOM). (2024). Báo cáo Chỉ số Thương mại Điện tử Việt Nam (EBI 2024).",
        "Bộ Thông tin và Truyền thông. (2020). Nghị định số 30/2020/NĐ-CP về công tác văn thư và quy chuẩn trình bày văn bản khoa học."
    ]
    for r in refs:
        p_ref = doc.add_paragraph()
        p_ref.paragraph_format.line_spacing = 1.3
        p_ref.paragraph_format.space_before = Pt(2)
        p_ref.paragraph_format.space_after = Pt(4)
        p_ref.paragraph_format.left_indent = Cm(1.0)
        p_ref.paragraph_format.first_line_indent = Cm(-1.0)
        r_run = p_ref.add_run(r)
        format_run(r_run, size_pt=11.5, bold=False, color=COLOR_TEXT)

    doc.add_page_break()

    # ==================== PHỤ LỤC ====================
    add_heading_1(doc, "PHỤ LỤC: KỊCH BẢN MÃ NGUỒN DAX VÀ SQL QUẢN TRỊ")
    add_body(doc, "Trích dẫn một số đoạn mã nguồn DAX và SQL tiêu biểu được lập trình trong Data Model để tính toán các chỉ số nghiệp vụ trọng tâm trên hệ thống BI Hasaki:", indent=True)

    add_heading_2(doc, "Phụ lục 1: Kịch bản tính toán DAX cho Tỷ lệ Hoàn trả và Đơn Giao trễ")
    dax_code = (
        "-- Đo lường Tỷ lệ hoàn trả (Return Rate %)\n"
        "Return_Rate_% = \n"
        "VAR Total_Orders = COUNTROWS('Fact_Sales')\n"
        "VAR Returned_Orders = CALCULATE(COUNTROWS('Fact_Sales'), 'Fact_Sales'[Returned] = \"Có\")\n"
        "RETURN\n"
        "DIVIDE(Returned_Orders, Total_Orders, 0)\n\n"
        "-- Đo lường Tỷ lệ đơn giao trễ vượt chuẩn SLA (> 3 ngày)\n"
        "Late_Delivery_Rate_% = \n"
        "VAR Total_Online_Orders = CALCULATE(COUNTROWS('Fact_Sales'), 'Fact_Sales'[Shipping_Days] > 0)\n"
        "VAR Late_Orders = CALCULATE(COUNTROWS('Fact_Sales'), 'Fact_Sales'[Shipping_Days] > 3)\n"
        "RETURN\n"
        "DIVIDE(Late_Orders, Total_Online_Orders, 0)"
    )
    p_code1 = doc.add_paragraph()
    p_code1.paragraph_format.left_indent = Cm(1.0)
    p_code1.paragraph_format.space_before = Pt(4)
    p_code1.paragraph_format.space_after = Pt(6)
    r_c1 = p_code1.add_run(dax_code)
    format_run(r_c1, font_name="Consolas", size_pt=10, bold=False, color=COLOR_SECONDARY)

    add_heading_2(doc, "Phụ lục 2: Câu truy vấn SQL tổng hợp Hiệu suất Đa kênh kèm Biên lợi nhuận")
    sql_code = (
        "SELECT \n"
        "    Channel,\n"
        "    COUNT(Order_ID) AS Total_Orders,\n"
        "    SUM(Net_Sales_VND) AS Total_Net_Sales,\n"
        "    SUM(Profit_VND) AS Total_Gross_Profit,\n"
        "    ROUND(SUM(Profit_VND) * 100.0 / SUM(Net_Sales_VND), 2) AS Gross_Margin_Pct,\n"
        "    ROUND(AVG(Customer_Satisfaction), 2) AS Average_CSAT,\n"
        "    ROUND(SUM(CASE WHEN Returned = 'Có' THEN 1 ELSE 0 END) * 100.0 / COUNT(Order_ID), 2) AS Return_Rate_Pct\n"
        "FROM Fact_Sales\n"
        "GROUP BY Channel\n"
        "ORDER BY Total_Net_Sales DESC;"
    )
    p_code2 = doc.add_paragraph()
    p_code2.paragraph_format.left_indent = Cm(1.0)
    p_code2.paragraph_format.space_before = Pt(4)
    p_code2.paragraph_format.space_after = Pt(6)
    r_c2 = p_code2.add_run(sql_code)
    format_run(r_c2, font_name="Consolas", size_pt=10, bold=False, color=COLOR_SECONDARY)

print("Chapter 7 and End ready!")

