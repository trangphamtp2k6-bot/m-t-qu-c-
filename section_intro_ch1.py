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

def add_intro_and_chapter1(doc):
    # ==================== MỞ ĐẦU ====================
    add_heading_1(doc, "MỞ ĐẦU: TÍNH CẤP THIẾT CỦA ĐỀ TÀI VÀ PHƯƠNG PHÁP NGHIÊN CỨU")
    add_body(doc, "Trong kỷ nguyên chuyển đổi số và bùng nổ dữ liệu lớn (Big Data), ngành bán lẻ hiện đại (Modern Trade & E-Commerce) đang chứng kiến sự chuyển dịch căn bản từ việc ra quyết định dựa trên trực giác, kinh nghiệm cá nhân sang mô hình quản trị thông minh dựa trên dữ liệu thực chứng (Data-Driven Decision Making). Đối với ngành bán lẻ mỹ phẩm và sản phẩm chăm sóc cá nhân (Beauty & Personal Care) – một ngành hàng có tốc độ luân chuyển nhanh, danh mục sản phẩm (SKU) đa dạng và vòng đời xu hướng ngắn – khả năng theo dõi liên tục hiệu quả tài chính, độ nhạy giá của khách hàng và chất lượng dịch vụ vận hành chặng cuối là yếu tố quyết định năng lực cạnh tranh bền vững của doanh nghiệp.", indent=True)
    add_body(doc, "Chuỗi bán lẻ Hasaki Beauty & Clinic với hơn 140 chi nhánh vật lý trên toàn quốc kết hợp cùng các kênh bán hàng số (Website/App, Shopee, TikTok Shop, Lazada) đang đứng trước bài toán cân bằng chiến lược: Làm thế nào để vừa duy trì tốc độ tăng trưởng doanh thu ấn tượng, vừa bảo vệ biên lợi nhuận gộp trước áp lực chiết khấu sâu của các kỳ Mega Sale trực tuyến, đồng thời nâng cao chỉ số thỏa mãn khách hàng (CSAT) trước thách thức vận chuyển hàng dễ vỡ?", indent=True)
    add_body(doc, "Đề tài 'Xây dựng Hệ thống Business Intelligence Tối ưu hóa Chiến lược Khuyến mãi, Kênh Bán hàng và Quản trị Lợi nhuận Danh mục Skincare tại Chuỗi Bán lẻ Hasaki Beauty & Clinic' được triển khai nhằm mục tiêu giải quyết trọn vẹn bài toán quản trị nêu trên thông qua quy trình 6 bước chuẩn mực quốc tế của một dự án Business Intelligence. Báo cáo vận dụng phương pháp phân tích định lượng trên tập dữ liệu thực nghiệm 15.000 giao dịch, kết hợp mô hình hóa đa chiều Kimball, kỹ thuật trực quan hóa tương tác và phân tích nguyên nhân gốc rễ (Root Cause Analysis).", indent=True)
    
    doc.add_page_break()

    # ==================== CHƯƠNG 1 ====================
    add_heading_1(doc, "CHƯƠNG 1: XÁC ĐỊNH YÊU CẦU KINH DOANH VÀ CÂY CHỈ SỐ KPIs (BUSINESS REQUIREMENTS & KPIs)")
    
    add_heading_2(doc, "1.1. Bối cảnh thị trường và Tổng quan Chuỗi bán lẻ Hasaki Beauty & Clinic")
    add_body(doc, "Hasaki Beauty & Clinic được thành lập từ năm 2016, hiện là một trong những hệ thống phân phối mỹ phẩm chính hãng và cung cấp dịch vụ chăm sóc sắc đẹp chuyên sâu (Hasaki Clinic & Spa) hàng đầu tại Việt Nam. Định vị thương hiệu của Hasaki được xây dựng dựa trên ba trụ cột giá trị cốt lõi:", indent=True)
    add_bullet(doc, "Cam kết 100% hàng chính hãng nhập khẩu, đầy đủ hóa đơn chứng từ VAT và tem phụ tiếng Việt, tạo dựng niềm tin tuyệt đối với người tiêu dùng trong một thị trường còn nhiều vấn nạn hàng xách tay trôi nổi.", bold_prefix="Chất lượng sản phẩm cam kết: ")
    add_bullet(doc, "Chiến lược định giá cạnh tranh ở mọi thời điểm ('Giá luôn luôn tốt'), giúp chuỗi thu hút tệp khách hàng trẻ đông đảo.", bold_prefix="Chính sách giá cạnh tranh: ")
    add_bullet(doc, "Dịch vụ 'Hasaki 2H' (giao nhanh trong vòng 2 giờ đối với khu vực nội thành Hà Nội và TP. Hồ Chí Minh), biến tốc độ thành lợi thế cạnh tranh độc nhất (USP) trên thị trường bán lẻ mỹ phẩm.", bold_prefix="Logistics chặng cuối siêu tốc: ")
    add_body(doc, "Hệ sinh thái phân phối của Hasaki vận hành theo mô hình bán lẻ hợp kênh toàn diện (Omnichannel / O2O), bao gồm: Kênh Cửa hàng vật lý (In-Store) với mạng lưới hơn 140 điểm bán trải dài khắp 3 miền; Kênh sở hữu số (Website hasaki.vn và Ứng dụng Hasaki App); và các Gian hàng chính hãng (Mall) trên các sàn TMĐT hàng đầu (Shopee, TikTok Shop, Lazada).", indent=True)

    add_heading_2(doc, "1.2. Thách thức Chiến lược và Ba Áp lực Nghiệp vụ Trọng tâm")
    add_body(doc, "Mặc dù ngành mỹ phẩm có biên lợi nhuận gộp định danh tương đối hấp dẫn (trên 30%), ban điều hành Hasaki đang phải đối mặt với ba áp lực kinh doanh cốt lõi định hình bài toán phân tích BI:", indent=True)
    add_bullet(doc, "Khách hàng thế hệ Gen Z và Millennials cực kỳ nhạy cảm về giá, có xu hướng chờ đợi các đợt Flash Sale, Mega Sale (11.11, 12.12) và áp mã giảm giá tối đa. Việc tung ra các mức chiết khấu sâu (từ 20% đến 30%) tạo ra tăng trưởng ảo về mặt số lượng đơn nhưng trực tiếp ăn mòn lợi nhuận gộp, đặc biệt khi chi phí cố định (mặt bằng, nhân sự, đóng gói) trên mỗi đơn không đổi.", bold_prefix="1. Cuộc chiến giá và bẫy chiết khấu sâu (The Promotion Dilemma): ")
    add_bullet(doc, "Các sàn TMĐT mang lại lưu lượng truy cập (Traffic) khổng lồ nhưng bào mòn lợi nhuận thực tế bởi biểu phí sàn ngày càng tăng (phí cố định, phí thanh toán, phí Freeship Extra lên tới 10-15%). Trong khi đó, kênh cửa hàng và App nội bộ giữ lại toàn bộ biên lợi nhuận nhưng lại đòi hỏi chi phí thu hút khách hàng (CAC) và chi phí vận hành nền tảng cao.", bold_prefix="2. Xung đột đa kênh và chi phí ẩn (Omnichannel Friction): ")
    add_bullet(doc, "Đặc thù của sản phẩm chăm sóc da (Skincare) và Nước hoa là bao bì thủy tinh dễ vỡ, dung dịch nhạy cảm với nhiệt độ và va đập. Mọi sự cố phát sinh như giao trễ, móp méo hộp hay bể vỡ đều dẫn đến khiếu nại, tỷ lệ hoàn trả (Return/Refund) tăng cao, làm phát sinh thiệt hại kép về giá vốn hàng hóa và niềm tin thương hiệu.", bold_prefix="3. Rủi ro bao bì hư hỏng trong logistics chặng cuối (Last-Mile Fragility): ")

    add_heading_2(doc, "1.3. Cây Chỉ số Hiệu suất Chính (KPI Architecture Framework)")
    add_body(doc, "Để đo lường chính xác hiệu quả vận hành và đưa ra cảnh báo kịp thời, hệ thống BI Hasaki thiết lập cây chỉ số hiệu suất (KPI Framework) chuẩn hóa phân cấp theo 4 nhóm trọng yếu:", indent=True)
    
    kpi_headers = ["Nhóm chỉ số", "Tên KPI nghiệp vụ", "Công thức / Phương pháp tính", "Mục tiêu quản trị cốt lõi"]
    kpi_widths = [Cm(3.0), Cm(4.0), Cm(4.5), Cm(4.5)]
    kpi_align = [WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT]
    kpi_data = [
        ["Hiệu suất Tài chính", "Doanh thu gộp (Gross Sales)", "SUM(Quantity * Unit_Price)", "Đo lường quy mô dòng tiền bán hàng"],
        ["Hiệu suất Tài chính", "Doanh thu thuần (Net Sales)", "SUM(Sales_VND - Discount_Amount)", "Đo lường quy mô doanh thu thực nhận"],
        ["Hiệu suất Tài chính", "Lợi nhuận gộp (Gross Profit)", "SUM(Net_Sales_VND - Cost_VND)", "Đo lường giá trị thặng dư tạo ra"],
        ["Hiệu suất Tài chính", "Biên lợi nhuận gộp (Gross Margin %)", "(Gross Profit / Net Sales) * 100", "Đo lường hiệu quả sinh lời trên mỗi đồng doanh thu"],
        ["Hiệu quả Đơn hàng", "Giá trị đơn trung bình (AOV)", "Net Sales / Total Orders", "Theo dõi chất lượng giỏ hàng và sức mua"],
        ["Hiệu quả Đơn hàng", "Tỷ lệ chiết khấu bình quân", "(SUM(Discount_Amount) / SUM(Sales)) * 100", "Kiểm soát ngân sách khuyến mãi"],
        ["Vận hành & Dịch vụ", "Thời gian giao hàng trung bình", "AVERAGE(Shipping_Days)", "Theo dõi cam kết dịch vụ giao nhận (SLA)"],
        ["Vận hành & Dịch vụ", "Tỷ lệ đơn hàng giao trễ", "(Orders với Shipping_Days > 3) / Total Orders", "Cảnh báo nguy cơ đứt gãy trải nghiệm khách"],
        ["Khách hàng & Trải nghiệm", "Điểm hài lòng trung bình (CSAT)", "AVERAGE(Customer_Satisfaction)", "Đo lường chỉ số hài lòng (thang 1-5 sao)"],
        ["Khách hàng & Trải nghiệm", "Tỷ lệ đơn hàng hoàn trả (Return %)", "(Orders Returned = 'Có') / Total Orders", "Kiểm soát rủi ro từ chối nhận hàng và tổn thất vốn"]
    ]
    add_table(doc, "Bảng 1.1: Khung chỉ số hiệu suất kinh doanh cốt lõi (KPI Framework) của Hasaki", kpi_headers, kpi_data, kpi_widths, kpi_align, "Đề xuất chuẩn hóa BI301")

    add_heading_2(doc, "1.4. Thiết kế 5 Hướng Phân tích và 21 Câu hỏi Nghiệp vụ Trọng tâm")
    add_body(doc, "Nhằm phục vụ bài toán ra quyết định cấp cao, hệ thống BI được cấu trúc thành 5 hướng phân tích chuyên đề với 21 câu hỏi nghiệp vụ chi tiết, đòi hỏi kỹ thuật cắt lớp dữ liệu đa chiều (Slicing & Dicing):", indent=True)
    
    q_headers = ["Hướng phân tích", "Mã câu hỏi", "Nội dung câu hỏi nghiệp vụ cần trả lời", "Chiều dữ liệu cắt lớp"]
    q_widths = [Cm(3.5), Cm(2.2), Cm(6.8), Cm(3.5)]
    q_align = [WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT]
    q_data = [
        ["Hướng A: Hiệu suất & Cơ cấu Lợi nhuận", "A1", "Doanh thu và lợi nhuận gộp năm 2025 tăng trưởng bao nhiêu % so với 2024?", "Thời gian (Năm)"],
        ["Hướng A: Hiệu suất & Cơ cấu Lợi nhuận", "A2", "Tháng nào đạt đỉnh doanh thu và xu hướng các tháng cuối năm diễn biến ra sao?", "Thời gian (Tháng)"],
        ["Hướng A: Hiệu suất & Cơ cấu Lợi nhuận", "A3", "Danh mục nào đóng góp lợi nhuận gộp cao nhất và cơ cấu tỷ trọng ra sao?", "Danh mục sản phẩm"],
        ["Hướng A: Hiệu suất & Cơ cấu Lợi nhuận", "A4", "Phân khúc khách hàng nào có AOV và biên lợi nhuận gộp tối ưu nhất?", "Phân khúc khách hàng"],
        ["Hướng A: Hiệu suất & Cơ cấu Lợi nhuận", "A5", "Sự khác biệt về doanh thu và biên gộp giữa 3 miền (Bắc, Trung, Nam) là gì?", "Khu vực địa lý"],
        ["Hướng B: Khuyến mãi & Độ nhạy Giá", "B1", "Biên lợi nhuận gộp biến thiên ra sao theo từng nấc chiết khấu (0% - 30%)?", "Mức chiết khấu (Bin)"],
        ["Hướng B: Khuyến mãi & Độ nhạy Giá", "B2", "'Điểm ngọt' (Sweet Spot) tạo ra sự cân bằng giữa lượng đơn và biên gộp là mức nào?", "Mức chiết khấu"],
        ["Hướng B: Khuyến mãi & Độ nhạy Giá", "B3", "Chiết khấu sâu (≥ 20%) đóng góp bao nhiêu % lợi nhuận và tác động gì tới AOV?", "Chiết khấu × AOV"],
        ["Hướng B: Khuyến mãi & Độ nhạy Giá", "B4", "Danh mục nào nhạy cảm nhất về biên lợi nhuận khi tăng mức giảm giá?", "Danh mục × Chiết khấu"],
        ["Hướng C: Hiệu suất Bán hàng Đa kênh", "C1", "Tỷ trọng doanh thu và lợi nhuận gộp của từng kênh (Cửa hàng, Web/App, TMĐT)?", "Kênh bán hàng"],
        ["Hướng C: Hiệu suất Bán hàng Đa kênh", "C2", "Biên lợi nhuận gộp trên số liệu giữa các kênh có khác biệt hay không?", "Kênh bán hàng"],
        ["Hướng C: Hiệu suất Bán hàng Đa kênh", "C3", "Kênh nào đạt điểm hài lòng khách hàng (CSAT) cao nhất và tỷ lệ hoàn trả thấp nhất?", "Kênh × CSAT × Hoàn"],
        ["Hướng C: Hiệu suất Bán hàng Đa kênh", "C4", "Hành vi mua sắm của các phân khúc khách hàng phân bổ qua các kênh ra sao?", "Phân khúc × Kênh"],
        ["Hướng D: Vận hành, Hài lòng & Hoàn trả", "D1", "Thời gian giao hàng bao nhiêu ngày thì điểm CSAT bắt đầu sụt giảm nghiêm trọng?", "Số ngày giao × CSAT"],
        ["Hướng D: Vận hành, Hài lòng & Hoàn trả", "D2", "Nhóm đơn hàng giao trễ (> 3 ngày) chiếm bao nhiêu % và tập trung ở kênh nào?", "Giao trễ × Kênh"],
        ["Hướng D: Vận hành, Hài lòng & Hoàn trả", "D3", "Tổn thất lợi nhuận và giá vốn hàng hoàn trên toàn hệ thống là bao nhiêu?", "Thiệt hại hoàn trả"],
        ["Hướng D: Vận hành, Hài lòng & Hoàn trả", "D4", "Giao trễ có phải nguyên nhân trực tiếp làm tăng tỷ lệ hoàn trả hay không?", "Số ngày giao × Hoàn"],
        ["Hướng D: Vận hành, Hài lòng & Hoàn trả", "D5", "Danh mục và khu vực nào chịu thiệt hại vốn hoàn trả nặng nề nhất?", "Danh mục × Khu vực"],
        ["Hướng E: Xu hướng Phương thức Thanh toán", "E1", "Cơ cấu giữa thanh toán số và tiền mặt (COD) có sự thay đổi giữa 2024 và 2025?", "Phương thức × Năm"],
        ["Hướng E: Xu hướng Phương thức Thanh toán", "E2", "Tỷ lệ hoàn trả của đơn hàng COD có cao hơn các đơn thanh toán số hay không?", "Thanh toán × Hoàn trả"],
        ["Hướng E: Xu hướng Phương thức Thanh toán", "E3", "Khi loại trừ yếu tố mua tại quầy, hoàn trả của COD và thanh toán trực tuyến ra sao?", "Kênh Online × Thanh toán"]
    ]
    add_table(doc, "Bảng 1.2: Tổng hợp 5 Hướng phân tích và 21 Câu hỏi nghiệp vụ trọng tâm", q_headers, q_data, q_widths, q_align, "Đặc tả yêu cầu nghiệp vụ BI Hasaki")

    doc.add_page_break()

print("Intro and Chapter 1 ready!")

