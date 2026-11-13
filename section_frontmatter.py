import os
import sys
import docx
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

from docx_helpers import (
    COLOR_PRIMARY, COLOR_SECONDARY, COLOR_TEXT, COLOR_MUTED,
    format_run, add_heading_1, add_heading_2, add_heading_3,
    add_body, add_bullet, add_table
)

def add_frontmatter(doc):
    # LỜI CAM ĐOAN
    add_heading_1(doc, "LỜI CAM ĐOAN")
    add_body(doc, "Tôi xin cam đoan đây là công trình nghiên cứu và thực hành báo cáo Business Intelligence độc lập do chính tôi thực hiện dưới sự định hướng, hướng dẫn chuyên môn của Giảng viên ThS. Hoàng Thị Kim Oanh trong khuôn khổ học phần Business Intelligence (BI301).", indent=True)
    add_body(doc, "Toàn bộ các số liệu thống kê, mô hình dữ liệu Star Schema, công thức đo lường DAX/SQL, bảng biểu và biểu đồ trực quan hóa được kết xuất trực tiếp và trung thực từ tập dữ liệu bán lẻ gồm 15.000 giao dịch thực nghiệm của chuỗi bán lẻ Hasaki Beauty & Clinic trong giai đoạn từ ngày 02/01/2024 đến ngày 27/12/2025. Các thông tin trích dẫn lý thuyết, khung năng lực và phương pháp luận quốc tế từ Ralph Kimball, Bill Inmon và các tổ chức nghiên cứu đều được ghi rõ nguồn gốc tài liệu tham khảo theo đúng chuẩn mực học thuật quốc tế APA.", indent=True)
    
    p_sign = doc.add_paragraph()
    p_sign.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_sign.paragraph_format.space_before = Pt(18)
    p_sign.paragraph_format.space_after = Pt(4)
    r_s1 = p_sign.add_run("Hà Nội, tháng 11 năm 2026\nSinh viên thực hiện\n\n\n\n")
    format_run(r_s1, size_pt=12, italic=True)
    r_s2 = p_sign.add_run("Phạm Thuỳ Trang\nMSV: B25DTQT066")
    format_run(r_s2, size_pt=12.5, bold=True, color=COLOR_PRIMARY)
    
    doc.add_page_break()

    # LỜI CẢM ƠN
    add_heading_1(doc, "LỜI CẢM ƠN")
    add_body(doc, "Để hoàn thành báo cáo chuyên sâu Business Intelligence này, tôi xin bày tỏ lòng biết ơn sâu sắc và chân thành nhất tới Ban Giám hiệu Học viện Công nghệ Bưu chính Viễn thông (PTIT), Ban Đào tạo Viện Rikkei Education cùng các Thầy Cô giáo thuộc Khoa Quản trị Kinh doanh đã tận tâm truyền đạt những nền tảng kiến thức vô cùng quý báu về Khoa học dữ liệu, Quản trị chiến lược kinh doanh và Trí tuệ kinh doanh số.", indent=True)
    add_body(doc, "Đặc biệt, tôi xin gửi lời cảm ơn trân trọng và sâu sắc nhất tới cô giáo - ThS. Hoàng Thị Kim Oanh, người đã luôn dành nhiều thời gian quý báu để định hướng phương pháp luận nghiên cứu, gợi mở tư duy phân tích đa chiều (Dimensional Thinking), phản biện sắc bén các giả thuyết kinh doanh và hướng dẫn tôi hoàn thiện từng bước chuẩn hóa dữ liệu, thiết kế mô hình lược đồ hình sao và trích xuất các insight mang tính chiến lược thực tiễn cao.", indent=True)
    add_body(doc, "Mặc dù đã nỗ lực hết mình với tinh thần nghiêm túc và cầu thị của một chuyên viên phân tích BI, song do giới hạn về mặt thời gian và phạm vi dữ liệu quan sát, báo cáo khó tránh khỏi những thiếu sót nhất định. Tôi rất mong nhận được những đóng góp quý báu của Quý Thầy Cô và các chuyên gia trong ngành để công trình được hoàn thiện hơn nữa.", indent=True)
    
    doc.add_page_break()

    # TÓM TẮT DỰ ÁN (EXECUTIVE SUMMARY)
    add_heading_1(doc, "TÓM TẮT DỰ ÁN (EXECUTIVE SUMMARY)")
    add_body(doc, "Trong bối cảnh thị trường bán lẻ mỹ phẩm và chăm sóc cá nhân (Beauty & Personal Care) tại Việt Nam cạnh tranh ngày càng khốc liệt giữa các chuỗi bán lẻ hiện đại (Hasaki, Guardian, Watsons) và các gian hàng chính hãng trên sàn thương mại điện tử (Shopee Mall, TikTok Shop, LazMall), việc tối ưu hóa biên lợi nhuận và nâng cao trải nghiệm khách hàng đa kênh trở thành yếu tố sống còn của doanh nghiệp.", indent=True)
    add_body(doc, "Báo cáo này đóng vai trò là một tài liệu phân tích quản trị cấp cao (Executive Business Intelligence Report), vận dụng quy trình 6 bước chuẩn hóa của một dự án BI từ khâu xác định yêu cầu nghiệp vụ, thu thập ETL, làm sạch dữ liệu, mô hình hóa Star Schema đến trực quan hóa tương tác và trích xuất quyết định dựa trên dữ liệu (Data-Driven Decision Making). Nghiên cứu được thực thi trên tập dữ liệu gồm 15.000 đơn hàng thực tế của chuỗi Hasaki Beauty & Clinic trong giai đoạn 2024 - 2025 với 24 thuộc tính kinh doanh quan trọng.", indent=True)
    add_body(doc, "Kết quả nghiên cứu đã phát hiện ra 5 insight chiến lược then chốt mang tính đột phá:", bold_prefix="Các phát hiện cốt lõi: ", indent=True)
    add_bullet(doc, "Năm 2025 ghi nhận doanh thu thuần đạt 13,81 tỷ ₫ (tăng trưởng +31,4% so với năm 2024 đạt 10,51 tỷ ₫) và số lượng đơn hàng tăng +38,6% (từ 6.286 lên 8.714 đơn). Tuy nhiên, giá trị đơn hàng trung bình (AOV) lại suy giảm từ 1,67 triệu ₫ xuống 1,58 triệu ₫ (-5,4%) và biên lợi nhuận gộp đi ngang ở mức 30,6%, phản ánh hiện tượng tăng trưởng theo chiều rộng nhưng suy giảm chất lượng trên từng đơn vị giao dịch.", bold_prefix="Ảo tưởng tăng trưởng doanh thu: ")
    add_bullet(doc, "Phân tích độ co giãn giá cho thấy biên lợi nhuận gộp giảm tuyến tính 5 điểm phần trăm cho mỗi nấc tăng 5% chiết khấu. 'Điểm ngọt' (Sweet Spot) tạo ra hiệu quả tối ưu nằm ở mức 5% - 10% (chiếm 51,1% số đơn và đóng góp tới 53,2% lợi nhuận gộp toàn hệ thống). Ngược lại, các mức giảm sâu từ 20% trở lên chỉ mang lại 8,2% lợi nhuận dù chiếm 16% số đơn hàng. Đặc biệt, danh mục Skincare bị xói mòn lợi nhuận nghiêm trọng nhất: biên gộp sụp đổ từ 35,0% (khi không giảm) xuống chỉ còn 7,3% (khi giảm 30%).", bold_prefix="Bẫy chiết khấu sâu (Discount Elasticity Trap): ")
    add_bullet(doc, "Kênh cửa hàng vật lý (In-Store) đóng vai trò là 'trụ cột' vững chắc khi đóng góp 45,6% doanh thu và 45,8% lợi nhuận với tỷ lệ hoàn trả thấp nhất (1,9%) và CSAT cao nhất (4,21). Sàn TMĐT đóng góp 19,1% doanh thu nhưng tỷ lệ hoàn trả lên tới 7,2% (gấp gần 4 lần cửa hàng), bào mòn nghiêm trọng lợi nhuận ròng sau chi phí vận hành.", bold_prefix="Nghịch lý kênh bán hàng: ")
    add_bullet(doc, "Dữ liệu chỉ ra ngưỡng đứt gãy trải nghiệm khách hàng nằm chính xác ở mốc 3 ngày. Đơn mua tại chỗ đạt CSAT 4,21; đơn giao 1-3 ngày duy trì CSAT 4,07 - 4,09; nhưng khi thời gian giao vượt quá 3 ngày, CSAT rơi tự do xuống 2,70 (43,5% bị chấm điểm 1-2 sao). Nhóm giao trễ chiếm 13,7% tổng đơn (2.060 đơn), tập trung chủ yếu ở Website/App và sàn TMĐT.", bold_prefix="Ngưỡng gãy logistics chặng cuối: ")
    add_bullet(doc, "Phương thức thanh toán số (VNPay, Thẻ, MoMo, ZaloPay) đã chiếm tỷ trọng áp đảo 81,6% và giữ ổn định suốt 2 năm. Phân tích đối chiếu cho thấy COD không phải là nguyên nhân gốc rễ gây hoàn trả (với đơn giao từ xa, tỷ lệ hoàn trả của COD là 5,4% tương đương mức 5,5% của thanh toán số). Nguyên nhân cốt lõi của hoàn hàng nằm ở tính chất kênh từ xa và rủi ro bao bì hư hỏng trong quá trình vận chuyển.", bold_prefix="Bản chất của hoàn hàng và thanh toán: ")
    add_body(doc, "Trên cơ sở đó, báo cáo đề xuất gói 7 giải pháp chiến lược (GP1 - GP7) kèm mô hình tài chính dự phóng giúp Hasaki thu hồi thêm 333 triệu ₫ lợi nhuận gộp từ việc chặn chiết khấu sâu và tiết kiệm hơn 120 triệu ₫ từ việc cắt giảm rủi ro hư hỏng đóng gói hàng Skincare/Nước hoa, triển khai theo lộ trình Agile 90 ngày rõ ràng.", indent=True)
    
    doc.add_page_break()

    # DANH MỤC TỪ VIẾT TẮT
    add_heading_1(doc, "DANH MỤC CÁC TỪ VIẾT TẮT")
    abbr_headers = ["Ký hiệu viết tắt", "Tên tiếng Anh đầy đủ", "Giải nghĩa thuật ngữ"]
    abbr_widths = [Cm(3.0), Cm(6.5), Cm(6.5)]
    abbr_align = [WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT]
    abbr_data = [
        ["BI", "Business Intelligence", "Trí tuệ kinh doanh / Phân tích dữ liệu hỗ trợ quyết định"],
        ["ETL", "Extract - Transform - Load", "Trích xuất, Chuyển đổi và Nạp dữ liệu vào kho"],
        ["DWH", "Data Warehouse", "Kho lưu trữ dữ liệu tập trung của doanh nghiệp"],
        ["KPI", "Key Performance Indicator", "Chỉ số đánh giá hiệu suất kinh doanh cốt lõi"],
        ["AOV", "Average Order Value", "Giá trị trung bình trên một đơn hàng"],
        ["CSAT", "Customer Satisfaction Score", "Chỉ số đo lường mức độ hài lòng của khách hàng (thang 1-5 sao)"],
        ["POS", "Point of Sale", "Hệ thống điểm bán hàng vật lý tại các chi nhánh"],
        ["CRM", "Customer Relationship Management", "Hệ thống quản lý thông tin và quan hệ khách hàng"],
        ["ERP", "Enterprise Resource Planning", "Hệ thống hoạch định và quản trị nguồn lực doanh nghiệp"],
        ["TMĐT", "Thương mại điện tử (E-Commerce)", "Kênh bán lẻ trực tuyến trên sàn (Shopee, TikTok, Lazada)"],
        ["DAX", "Data Analysis Expressions", "Ngôn ngữ công thức và tính toán phân tích dữ liệu đa chiều"],
        ["SQL", "Structured Query Language", "Ngôn ngữ truy vấn dữ liệu có cấu trúc"],
        ["COD", "Cash On Delivery", "Phương thức thanh toán bằng tiền mặt khi nhận hàng"],
        ["O2O", "Online-to-Offline / Omnichannel", "Mô hình kết hợp kinh doanh trực tuyến và cửa hàng trực tiếp"],
        ["YoY", "Year-over-Year", "Tốc độ tăng trưởng so với cùng kỳ năm trước"],
        ["MoM", "Month-over-Month", "Tốc độ tăng trưởng so với tháng liền kề trước đó"],
        ["SLA", "Service Level Agreement", "Cam kết chất lượng dịch vụ vận hành và thời gian giao hàng"]
    ]
    add_table(doc, "Danh mục các ký hiệu và thuật ngữ viết tắt sử dụng trong báo cáo", abbr_headers, abbr_data, abbr_widths, abbr_align, "Quy chuẩn thuật ngữ nghiệp vụ BI301")

    doc.add_page_break()

    # DANH MỤC BẢNG BIỂU
    add_heading_1(doc, "DANH MỤC BẢNG BIỂU")
    tbl_list_headers = ["Ký hiệu bảng", "Tên bảng biểu chi tiết", "Trang"]
    tbl_list_widths = [Cm(3.0), Cm(11.5), Cm(1.5)]
    tbl_list_align = [WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.CENTER]
    tbl_list_data = [
        ["Bảng 1.1", "Khung chỉ số hiệu suất kinh doanh cốt lõi (KPI Framework) của Hasaki", "08"],
        ["Bảng 1.2", "Tổng hợp 5 Hướng phân tích và 21 Câu hỏi nghiệp vụ trọng tâm", "11"],
        ["Bảng 2.1", "Kiến trúc hệ thống ETL tích hợp đa nguồn bán lẻ Hasaki", "14"],
        ["Bảng 3.1", "Kết quả kiểm toán chất lượng dữ liệu trên 15.000 bản ghi thô", "18"],
        ["Bảng 3.2", "Từ điển dữ liệu chuẩn hóa chi tiết (Data Dictionary - 24 thuộc tính)", "21"],
        ["Bảng 4.1", "Cấu trúc trường và mối quan hệ trong mô hình Star Schema", "26"],
        ["Bảng 4.2", "Thư viện các hàm tính toán và độ đo nâng cao (DAX / SQL Measures)", "29"],
        ["Bảng 5.1", "Bảng đối soát các chỉ số tổng hợp toàn hệ thống giai đoạn 2024 - 2025", "33"],
        ["Bảng 6.1", "Hiệu suất tài chính và cơ cấu theo danh mục sản phẩm", "37"],
        ["Bảng 6.2", "Hiệu suất theo phân khúc khách hàng (Cá nhân, VIP, B2B)", "39"],
        ["Bảng 6.3", "Số đơn, lợi nhuận gộp, biên gộp và tỷ lệ hoàn trả theo mức chiết khấu", "42"],
        ["Bảng 6.4", "Đánh giá toàn diện hiệu suất và chất lượng vận hành theo kênh bán hàng", "46"],
        ["Bảng 6.5", "Mức độ hài lòng (CSAT) và tỷ lệ hoàn trả theo số ngày giao hàng", "50"],
        ["Bảng 6.6", "Tổn thất tài chính do hoàn trả theo danh mục hàng hóa", "52"],
        ["Bảng 6.7", "Cơ cấu phương thức thanh toán và tỷ lệ hoàn trả tương ứng", "55"],
        ["Bảng 6.8", "Ma trận phân tích nguyên nhân gốc rễ (Root Cause Analysis - 5 Whys)", "58"],
        ["Bảng 6.9", "Đối chiếu giả thuyết quản trị ban đầu với dữ liệu thực chứng", "61"],
        ["Bảng 7.1", "Gói 7 giải pháp chiến lược tối ưu hóa toàn diện (GP1 - GP7)", "65"],
        ["Bảng 7.2", "Dự phóng tác động tài chính từ việc chặn chiết khấu sâu và giảm hoàn trả", "68"],
        ["Bảng 7.3", "Lộ trình triển khai dự án Agile 90 ngày (90-Day Roadmap)", "71"],
        ["Bảng 7.4", "Khung quản trị rủi ro và các điều kiện kích hoạt điều chỉnh (Pivot Triggers)", "73"]
    ]
    add_table(doc, "Danh mục các bảng biểu trình bày trong báo cáo", tbl_list_headers, tbl_list_data, tbl_list_widths, tbl_list_align)

    doc.add_page_break()

    # DANH MỤC HÌNH ẢNH
    add_heading_1(doc, "DANH MỤC HÌNH VẼ VÀ BIỂU ĐỒ")
    fig_list_headers = ["Ký hiệu hình", "Tên hình ảnh / Biểu đồ trực quan", "Trang"]
    fig_list_widths = [Cm(3.0), Cm(11.5), Cm(1.5)]
    fig_list_align = [WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.CENTER]
    fig_list_data = [
        ["Hình 4.1", "Mô hình dữ liệu Lược đồ hình sao (Star Schema) cho hệ thống BI Hasaki", "25"],
        ["Hình 5.1", "Dashboard Tổng quan Điều hành (Executive Overview Dashboard)", "32"],
        ["Hình 6.1", "Biểu đồ xu hướng Doanh thu thuần và Lợi nhuận gộp theo tháng kèm đường xu thế", "35"],
        ["Hình 6.2", "Biểu đồ so sánh Doanh thu thuần và Lợi nhuận gộp theo Danh mục sản phẩm", "38"],
        ["Hình 6.3", "Biểu đồ cơ cấu Doanh thu và Lợi nhuận gộp theo Phân khúc khách hàng", "40"],
        ["Hình 6.4", "Biểu đồ phân bổ Doanh thu thuần và Lợi nhuận gộp theo Khu vực địa lý", "41"],
        ["Hình 6.5", "Biểu đồ phân tích Tỷ trọng số đơn và Biên lợi nhuận gộp theo mức Chiết khấu", "43"],
        ["Hình 6.6", "Đồ thị đường cong Độ co giãn Biên lợi nhuận của từng Danh mục theo Chiết khấu", "45"],
        ["Hình 6.7", "Biểu đồ so sánh Doanh thu thuần và Lợi nhuận gộp theo Kênh bán hàng", "47"],
        ["Hình 6.8", "Biểu đồ phân tích Điểm hài lòng (CSAT) và Tỷ lệ hoàn trả theo Thời gian giao hàng", "51"],
        ["Hình 6.9", "Biểu đồ phân bổ Tỷ lệ hoàn trả theo Kênh bán hàng và theo Danh mục hàng hóa", "53"],
        ["Hình 6.10", "Biểu đồ cơ cấu Phương thức thanh toán và Tỷ lệ hoàn trả tương ứng", "56"]
    ]
    add_table(doc, "Danh mục các hình vẽ, sơ đồ và biểu đồ trực quan hóa", fig_list_headers, fig_list_data, fig_list_widths, fig_list_align)

    doc.add_page_break()

print("Frontmatter module ready!")

