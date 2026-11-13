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

def add_chapter2_and_3(doc):
    # ==================== CHƯƠNG 2 ====================
    add_heading_1(doc, "CHƯƠNG 2: THU THẬP VÀ TÍCH HỢP DỮ LIỆU ĐA NGUỒN (DATA COLLECTION & ETL)")
    
    add_heading_2(doc, "2.1. Kiến trúc Tổng thể Hệ thống Kho Dữ liệu (Data Warehouse Architecture)")
    add_body(doc, "Để đáp ứng yêu cầu quản trị chuỗi bán lẻ đa kênh với khối lượng giao dịch phát sinh liên tục theo thời gian thực từ hơn 140 chi nhánh và các nền tảng trực tuyến, kiến trúc luồng dữ liệu của Hasaki được thiết kế theo mô hình Kho dữ liệu phân tầng hiện đại (Modern Layered Data Warehouse Architecture) tuân thủ lý thuyết của Ralph Kimball và Bill Inmon.", indent=True)
    add_body(doc, "Hệ thống bao gồm ba tầng xử lý dữ liệu chính:", indent=True)
    add_bullet(doc, "Lưu trữ dữ liệu thô (Raw Data) được trích xuất từ các hệ thống tác nghiệp nguồn theo chu kỳ định kỳ (Batch processing hàng đêm hoặc Micro-batching mỗi 15 phút). Tầng này không thực hiện biến đổi logic nghiệp vụ nhằm đảm bảo khả năng truy vết và kiểm toán dữ liệu gốc.", bold_prefix="1. Tầng Thu nạp & Đệm dữ liệu (Staging Area): ")
    add_bullet(doc, "Dữ liệu sau khi làm sạch, khử trùng lặp và đồng nhất mã định danh sẽ được chuyển đổi thành các bảng Chiều (Dimensions) và bảng Sự kiện (Facts). Tại đây áp dụng mô hình Lược đồ hình sao (Star Schema) tối ưu hóa cho các truy vấn phân tích tổng hợp OLAP.", bold_prefix="2. Tầng Kho dữ liệu Doanh nghiệp (Enterprise Data Warehouse - EDW): ")
    add_bullet(doc, "Các bảng dữ liệu phân tích chuyên biệt (Data Marts) được phân phối tới công cụ trực quan hóa (Power BI, Tableau, Excel Data Model), tích hợp các phép đo DAX/SQL phức tạp phục vụ trực tiếp cho Ban Giám đốc và các nhà quản lý nghiệp vụ.", bold_prefix="3. Tầng Phục vụ & Trực quan hóa (Presentation / Semantic Layer): ")

    add_heading_2(doc, "2.2. Quy trình Trích xuất (Extract) từ Đa Nguồn Bán lẻ")
    add_body(doc, "Dữ liệu kinh doanh của Hasaki xuất phát từ bốn nguồn hệ thống nghiệp vụ phân tán, đòi hỏi các giao thức trích xuất chuyên biệt:", indent=True)
    add_bullet(doc, "Ghi nhận toàn bộ các giao dịch thanh toán trực tiếp tại quầy thu ngân của 140+ chi nhánh. Dữ liệu bao gồm mã đơn hàng, ngày giờ bán, mã sản phẩm, số lượng, phương thức thanh toán quầy (tiền mặt, máy POS quẹt thẻ, QR Code). Dữ liệu được đồng bộ về máy chủ trung tâm qua cơ chế Change Data Capture (CDC).", bold_prefix="Hệ thống POS tại cửa hàng: ")
    add_bullet(doc, "Lưu trữ hồ sơ khách hàng, phân hạng thành viên (Cá nhân, Khách VIP, Đối tác B2B) và lịch sử tích lũy điểm thưởng. Dữ liệu trích xuất thông qua các câu lệnh SQL tự động trên cơ sở dữ liệu PostgreSQL/SQL Server nguồn.", bold_prefix="Hệ thống CRM và ERP: ")
    add_bullet(doc, "Thu thập tự động các đơn hàng thương mại điện tử thông qua các Webhooks và RESTful APIs chính thức của Shopee Open Platform, TikTok Shop Partner API và Lazada Open Platform. Dữ liệu JSON thô được bóc tách để lấy thông tin đơn hàng, chiết khấu voucher sàn, mã vận đơn và trạng thái giao nhận.", bold_prefix="Kênh Sàn Thương mại điện tử (Shopee, TikTok, Lazada): ")
    add_bullet(doc, "Ghi nhận nhật ký giỏ hàng, thông tin thanh toán qua cổng trực tuyến (VNPay, MoMo, ZaloPay) và trạng thái đơn hàng dịch vụ giao nhanh 2h.", bold_prefix="Website Hasaki.vn và Mobile App: ")

    add_heading_2(doc, "2.3. Quy trình Chuyển đổi (Transform) và Nạp Dữ liệu (Load)")
    add_body(doc, "Quá trình ETL đóng vai trò 'xương sống' để biến các dòng dữ liệu thô, phân mảnh thành tập dữ liệu chuẩn mực sẵn sàng cho phân tích quản trị:", indent=True)
    
    etl_headers = ["Giai đoạn ETL", "Công việc thực hiện", "Công nghệ / Công cụ áp dụng", "Kết quả đầu ra chuẩn hóa"]
    etl_widths = [Cm(3.0), Cm(5.0), Cm(4.0), Cm(4.0)]
    etl_align = [WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT]
    etl_data = [
        ["Extract (Trích xuất)", "Trích xuất log giao dịch hàng ngày từ POS, ERP, Web App và TMĐT APIs", "Python Scripts, SQL Queries, REST APIs", "Raw staging files (CSV/JSON) tập trung tại Data Lake/Staging"],
        ["Transform (Làm sạch)", "Kiểm tra kiểu dữ liệu, khử giá trị rỗng, loại bỏ trùng lặp, chuẩn hóa địa danh", "Power Query (M Code), Python Pandas", "Tập dữ liệu sạch, đồng nhất kiểu dữ liệu cho 24 trường"],
        ["Transform (Tính toán)", "Tính toán các cột nghiệp vụ dẫn xuất: Sales, Discount_Amount, Net_Sales, Profit", "Power Query Calculated Columns, SQL", "Các số đo tài chính logic kế toán hoàn chỉnh"],
        ["Load (Nạp kho)", "Nạp dữ liệu vào mô hình đa chiều Star Schema trong Data Model", "Power BI Data Model, VertiPaq Engine", "Mô hình Star Schema sẵn sàng cho truy vấn DAX"],
        ["Semantic Layer", "Tạo các chỉ số đo lường nâng cao (KPI Measures): YoY, Margin %, CSAT TB", "DAX Formulas, Calculated Measures", "Kho dữ liệu ngữ nghĩa phục vụ báo cáo Dashboard tương tác"]
    ]
    add_table(doc, "Bảng 2.1: Kiến trúc hệ thống ETL tích hợp đa nguồn bán lẻ Hasaki", etl_headers, etl_data, etl_widths, etl_align, "Quy chuẩn quy trình ETL BI301")

    doc.add_page_break()

    # ==================== CHƯƠNG 3 ====================
    add_heading_1(doc, "CHƯƠNG 3: LÀM SẠCH VÀ CHUẨN BỊ DỮ LIỆU (DATA CLEANING & PROCESSING)")
    
    add_heading_2(doc, "3.1. Tổng quan Khám phá Tập dữ liệu Thực nghiệm")
    add_body(doc, "Bộ dữ liệu thực nghiệm được sử dụng trong dự án phản ánh hoạt động bán lẻ thực tế của chuỗi Hasaki Beauty & Clinic trong giai đoạn 24 tháng (từ ngày 02/01/2024 đến ngày 27/12/2025). Tập dữ liệu bao gồm chính xác 15.000 bản ghi giao dịch đơn hàng và 24 trường thông tin nghiệp vụ, bao quát đầy đủ các chiều không gian (địa lý, kênh bán), thời gian, sản phẩm, hành vi khách hàng và tài chính kế toán.", indent=True)

    add_heading_2(doc, "3.2. Quy trình Kiểm toán Chất lượng Dữ liệu (Data Quality Audit)")
    add_body(doc, "Trước khi đưa vào mô hình phân tích, tập dữ liệu thô đã được tiến hành kiểm toán chất lượng nghiêm ngặt (Data Quality Audit) trên toàn bộ 15.000 dòng theo các tiêu chuẩn toàn vẹn dữ liệu:", indent=True)
    
    audit_headers = ["Tiêu chí kiểm toán chất lượng", "Nội dung kiểm tra thực tế", "Kết quả kiểm toán", "Quy tắc xử lý chuẩn hóa"]
    audit_widths = [Cm(3.5), Cm(5.0), Cm(3.5), Cm(4.0)]
    audit_align = [WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT]
    audit_data = [
        ["Độ đầy đủ (Completeness)", "Kiểm tra giá trị Null/Empty trên 24 trường", "1.785 dòng thiếu CSAT; các trường khác 100% đầy đủ", "Gán nhãn 'No Rating' thay vì điền số ngẫu nhiên để tránh lệch điểm"],
        ["Tính duy nhất (Uniqueness)", "Kiểm tra trùng lặp mã đơn hàng Order_ID", "15.000 Order_ID hoàn toàn duy nhất, không trùng lặp", "Mỗi dòng đại diện cho một đơn hàng độc lập"],
        ["Tính hợp lệ (Validity)", "Kiểm tra dải giá trị của Discount, Rating, Shipping", "Chiết khấu từ 0.00 đến 0.30; CSAT từ 1 đến 5; Shipping 0 - 5+ ngày", "Hợp lệ 100%, không phát hiện giá trị âm hoặc vượt trần"],
        ["Tính toàn vẹn logic kế toán", "Kiểm tra công thức: Sales = Unit_Price * Quantity", "Khớp chính xác 100% trên toàn bộ 15.000 bản ghi", "Không có sai lệch số học giữa đơn giá và số lượng"],
        ["Logic chiết khấu và doanh thu", "Kiểm tra: Net_Sales = Sales - Discount_Amount", "Khớp chính xác 100% trên toàn bộ bản ghi", "Công thức doanh thu thuần tuân thủ chuẩn kế toán"],
        ["Logic giá vốn và lợi nhuận", "Kiểm tra: Profit = Net_Sales - Cost", "Khớp chính xác 100% trên toàn bộ bản ghi", "Lợi nhuận gộp phản ánh trung thực chênh lệch doanh thu và giá vốn"],
        ["Tính nhất quán đa kênh", "Kiểm tra số ngày giao hàng với kênh In-Store", "100% đơn mua tại cửa hàng có Shipping_Days = 0", "Chuẩn hóa đúng nghiệp vụ mua và nhận hàng tại chỗ"]
    ]
    add_table(doc, "Bảng 3.1: Kết quả kiểm toán chất lượng dữ liệu trên 15.000 bản ghi thô", audit_headers, audit_data, audit_widths, audit_align, "Báo cáo kiểm toán dữ liệu Hasaki 2026")

    add_heading_2(doc, "3.3. Quy tắc Làm sạch và Chuẩn hóa Dữ liệu Nghiệp vụ")
    add_body(doc, "Để chuẩn bị cho việc xây dựng mô hình Star Schema, các quy tắc làm sạch dữ liệu (Data Cleaning & Standardization Rules) sau đây đã được áp dụng triệt để:", indent=True)
    add_bullet(doc, "Trường Order_Date gốc được chuyển đổi đồng nhất về kiểu dữ liệu Date chuẩn định dạng ISO 'YYYY-MM-DD', loại bỏ phần giờ phút không cần thiết để tối ưu hóa việc liên kết với bảng chiều thời gian.", bold_prefix="1. Đồng nhất định dạng thời gian: ")
    add_bullet(doc, "Hệ số chiết khấu từ dạng chuỗi văn bản (VD: '5%', '10%') được chuyển đổi thành số thập phân Decimal (0.05, 0.10, ..., 0.30). Đồng thời, tạo thêm trường phân khoảng chiết khấu Discount_Bin ('0%', '1-10%', '11-20%', '>20%') để phục vụ phân tích nhóm trên dashboard.", bold_prefix="2. Chuẩn hóa tỷ lệ chiết khấu: ")
    add_bullet(doc, "Đối với các đơn hàng khách hàng không để lại đánh giá (1.785 đơn), hệ thống không thực hiện kỹ thuật thay thế bằng trung bình (Mean Imputation) vì sẽ làm sai lệch bản chất phân phối CSAT. Thay vào đó, tạo trường phân loại cờ báo riêng biệt và chỉ tính CSAT trên các đơn hàng có đánh giá thực tế.", bold_prefix="3. Xử lý giá trị rỗng đánh giá hài lòng: ")
    add_bullet(doc, "Toàn bộ tên tỉnh thành, khu vực địa lý, danh mục hàng hóa và kênh bán hàng được chuẩn hóa bảng mã tiếng Việt Unicode dựng sẵn, loại bỏ khoảng trắng thừa (Trim) và đồng nhất danh pháp (Ví dụ: 'Sàn TMĐT (Shopee/Lazada/TikTok)', 'Khách hàng cá nhân').", bold_prefix="4. Chuẩn hóa văn bản và danh mục: ")

    add_heading_2(doc, "3.4. Từ điển Dữ liệu Chuẩn hóa Chi tiết (Comprehensive Data Dictionary)")
    add_body(doc, "Bảng từ điển dữ liệu cung cấp thông số kỹ thuật chi tiết của 24 thuộc tính sau khi đã được làm sạch và chuẩn hóa hoàn thiện:", indent=True)
    
    dict_headers = ["STT", "Tên thuộc tính", "Kiểu dữ liệu", "Mô tả ý nghĩa nghiệp vụ", "Giá trị mẫu điển hình"]
    dict_widths = [Cm(1.0), Cm(3.8), Cm(2.2), Cm(6.0), Cm(3.0)]
    dict_align = [WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT]
    dict_data = [
        ["1", "Order_ID", "String (Varchar)", "Mã định danh duy nhất của mỗi giao dịch đơn hàng", "HSK-ORD-111344"],
        ["2", "Order_Date", "Date (YYYY-MM-DD)", "Ngày khách hàng đặt mua hàng thành công", "2024-01-02"],
        ["3", "Customer_ID", "String (Varchar)", "Mã định danh duy nhất của khách hàng", "KH-DEP-00265"],
        ["4", "Region", "String (Varchar)", "Khu vực địa lý phụ trách đơn hàng", "Miền Bắc, Miền Nam, Miền Trung"],
        ["5", "City", "String (Varchar)", "Tỉnh hoặc Thành phố nơi phát sinh đơn hàng", "Hà Nội, TP.HCM, Đà Nẵng..."],
        ["6", "Segment", "String (Varchar)", "Phân khúc khách hàng phân loại theo CRM", "Cá nhân, VIP, B2B"],
        ["7", "Channel", "String (Varchar)", "Kênh bán hàng phân phối giao dịch", "Cửa hàng, Website/App, Sàn TMĐT"],
        ["8", "Category", "String (Varchar)", "Danh mục ngành hàng chính (4 danh mục)", "Chăm sóc da, Nước hoa, Makeup..."],
        ["9", "Subcategory", "String (Varchar)", "Nhóm sản phẩm con chi tiết (8 nhóm)", "Kem dưỡng, Serum, Son môi..."],
        ["10", "Product_ID", "String (Varchar)", "Mã quản lý sản phẩm nội bộ (SKU)", "HSK-P0014"],
        ["11", "Product_Name", "String (Varchar)", "Tên thương mại chi tiết của sản phẩm", "Phấn phủ Innisfree No-Sebum"],
        ["12", "Quantity", "Integer (Số nguyên)", "Số lượng đơn vị sản phẩm mua trong đơn", "1, 2, 3, 5..."],
        ["13", "Unit_Price_VND", "Currency (VND)", "Đơn giá niêm yết bán lẻ của một sản phẩm", "115.000 ₫, 540.000 ₫..."],
        ["14", "Discount", "Decimal (Tỷ lệ)", "Tỷ lệ chiết khấu giảm giá áp dụng (0 - 0.30)", "0.00, 0.05, 0.10, ..., 0.30"],
        ["15", "Sales_VND", "Currency (VND)", "Tổng doanh thu bán hàng gộp (Chưa trừ giảm)", "345.000 ₫"],
        ["16", "Discount_Amount_VND", "Currency (VND)", "Tổng số tiền chiết khấu giảm giá cho đơn", "17.000 ₫"],
        ["17", "Net_Sales_VND", "Currency (VND)", "Doanh thu thuần thực tế sau giảm giá", "328.000 ₫"],
        ["18", "Cost_VND", "Currency (VND)", "Tổng giá vốn hàng bán xuất kho (COGS)", "186.000 ₫"],
        ["19", "Profit_VND", "Currency (VND)", "Lợi nhuận gộp phát sinh từ đơn hàng", "142.000 ₫"],
        ["20", "Shipping_Days", "Integer (Số nguyên)", "Thời gian từ lúc đặt hàng đến khi nhận hàng", "0 (Tại quầy), 1, 2, 3, 4, 5+ ngày"],
        ["21", "Returned", "String (Binary)", "Trạng thái đơn hàng có bị hoàn trả hay không", "'Có' hoặc 'Không'"],
        ["22", "Customer_Satisfaction", "Integer / Null", "Điểm đánh giá mức độ hài lòng của khách", "1, 2, 3, 4, 5 sao (hoặc Null)"],
        ["23", "Payment_Method", "String (Varchar)", "Phương thức thanh toán được khách lựa chọn", "VNPay QR, Thẻ, MoMo, COD..."],
        ["24", "Discount_Bin", "String (Nhóm)", "Phân nhóm dải chiết khấu phục vụ phân tích", "'0%', '1-10%', '11-20%', '>20%'"]
    ]
    add_table(doc, "Bảng 3.2: Từ điển dữ liệu chuẩn hóa chi tiết (Data Dictionary - 24 thuộc tính)", dict_headers, dict_data, dict_widths, dict_align, "Đặc tả kỹ thuật dữ liệu BI Hasaki")

    doc.add_page_break()

print("Chapters 2 & 3 ready!")

