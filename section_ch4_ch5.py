import os
import sys
import docx
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

from docx_helpers import (
    COLOR_PRIMARY, COLOR_SECONDARY, COLOR_TEXT, COLOR_MUTED,
    format_run, add_heading_1, add_heading_2, add_heading_3,
    add_body, add_bullet, add_callout, add_figure, add_table
)

IMG_DIR = r'C:\Users\HP\Downloads\extracted_hasaki_images'

def add_chapter4_and_5(doc):
    # ==================== CHƯƠNG 4 ====================
    add_heading_1(doc, "CHƯƠNG 4: MÔ HÌNH HÓA DỮ LIỆU ĐA CHIỀU (DATA MODELING - STAR SCHEMA & DAX/SQL)")
    
    add_heading_2(doc, "4.1. Cơ sở Lý thuyết Mô hình Chiều Kimball & Lựa chọn Kiến trúc Star Schema")
    add_body(doc, "Trong lĩnh vực Kho dữ liệu và Trí tuệ kinh doanh (BI), mô hình hóa chiều (Dimensional Modeling) do Ralph Kimball đề xuất là phương pháp luận tiêu chuẩn vàng để xây dựng các hệ thống phục vụ phân tích. So với mô hình Snowflake (chuẩn hóa cao, tách nhỏ các bảng chiều thành nhiều bảng con), mô hình Lược đồ hình sao (Star Schema) phi chuẩn hóa (Denormalized) mang lại những ưu thế vượt trội mang tính quyết định đối với dữ liệu bán lẻ Hasaki:", indent=True)
    add_bullet(doc, "Mô hình Star Schema giảm thiểu số lượng phép nối bảng (JOIN), cho phép công cụ BI thực thi câu lệnh SQL/DAX nhanh hơn gấp nhiều lần khi quét qua 15.000 dòng giao dịch và hàng triệu phép tính lặp.", bold_prefix="1. Tối ưu hóa hiệu năng truy vấn (Query Performance): ")
    add_bullet(doc, "Cấu trúc hình sao trực quan giúp người làm phân tích và các cấp quản lý dễ dàng thấu hiểu mô hình, giảm thiểu rủi ro sai sót trong việc tạo các bảng tổng hợp PivotTable hay viết biểu thức DAX.", bold_prefix="2. Tính trực quan và dễ tiếp cận người dùng (Simplicity): ")
    add_bullet(doc, "Cho phép người dùng tùy biến kéo thả bất kỳ trường chiều nào (Tháng, Kênh, Khu vực, Danh mục, Khách hàng) vào trục phân tích mà không làm vỡ logic tính toán của bảng sự kiện trung tâm.", bold_prefix="3. Khả năng cắt lát dữ liệu đa chiều linh hoạt (Slicing & Dicing): ")

    add_heading_2(doc, "4.2. Thiết kế Lược đồ Hình sao (Star Schema Architecture)")
    add_body(doc, "Hệ thống BI Hasaki được thiết kế chuẩn mực với Bảng sự kiện bán hàng trung tâm (Fact_Sales) kết nối với 4 Bảng chiều vệ tinh (Dimension Tables):", indent=True)
    
    schema_headers = ["Tên bảng", "Loại bảng", "Khóa chính (PK) / Khóa ngoại (FK)", "Các trường thuộc tính chính", "Vai trò trong phân tích"]
    schema_widths = [Cm(3.0), Cm(2.2), Cm(3.8), Cm(4.5), Cm(2.5)]
    schema_align = [WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT]
    schema_data = [
        ["Fact_Sales", "Fact Table", "FK: Order_Date, Customer_ID, Product_ID, Context_ID", "Quantity, Unit_Price, Sales, Discount_Amount, Net_Sales, Cost, Profit, Shipping_Days, Return_Flag, CSAT", "Lưu trữ số đo tài chính và sự kiện giao dịch"],
        ["Dim_Date", "Dimension", "PK: Date_Key", "Date, Day, Month, Quarter, Year, DayOfWeek, IsWeekend, Holiday_Campaign (MegaSale, 8/3, 20/10)", "Hỗ trợ phân tích thời gian, mùa vụ và xu hướng"],
        ["Dim_Product", "Dimension", "PK: Product_ID", "Product_Name, Category, Subcategory, Standard_Cost, List_Price", "Hỗ trợ phân tích cơ cấu danh mục và biên lợi nhuận"],
        ["Dim_Customer", "Dimension", "PK: Customer_ID", "Customer_Name, Segment (Cá nhân, VIP, B2B), Join_Date, Region", "Hỗ trợ phân tích hành vi và phân khúc khách hàng"],
        ["Dim_OrderContext", "Dimension", "PK: Context_ID", "Channel (Cửa hàng, Web, TMĐT), Region (Bắc, Trung, Nam), City, Payment_Method, Discount_Bin", "Cung cấp ngữ cảnh phân phối và trải nghiệm mua"]
    ]
    add_table(doc, "Bảng 4.1: Cấu trúc trường và mối quan hệ trong mô hình Star Schema", schema_headers, schema_data, schema_widths, schema_align, "Đặc tả kỹ thuật Star Schema Hasaki")

    # Insert Figure 4.1 (Star Schema)
    img_star = os.path.join(IMG_DIR, '4ad2157b308ef30a8adf0921e2c764ccac083625.png')
    add_figure(doc, img_star, "Hình 4.1: Mô hình dữ liệu Lược đồ hình sao (Star Schema) cho hệ thống BI Hasaki", "Nguồn: Kiến trúc Data Model Hasaki BI 2024–2025", width=Cm(15.2))

    add_heading_2(doc, "4.3. Thiết lập Mối quan hệ và Tính Toàn vẹn Tham chiếu")
    add_body(doc, "Trong mô hình dữ liệu, toàn bộ mối quan hệ (Relationships) giữa bảng Fact_Sales và các bảng chiều đều được thiết lập theo tỷ lệ Một - Nhiều (1-to-Many Relationship, 1:*), với hướng lọc dữ liệu một chiều (Single Cross-Filter Direction) từ bảng chiều sang bảng sự kiện. Điều này đảm bảo triệt để tính toàn vẹn tham chiếu (Referential Integrity), loại trừ hoàn toàn rủi ro xuất hiện bản ghi mồ côi (Orphan records) hay hiện tượng đếm trùng số liệu (Double counting) khi áp dụng các bộ lọc Slicer.", indent=True)

    add_heading_2(doc, "4.4. Xây dựng Thư viện Độ đo Phân tích Nâng cao (DAX & SQL Measures Library)")
    add_body(doc, "Để đo lường các nghiệp vụ phức tạp, hệ thống BI Hasaki đã xây dựng một thư viện các hàm tính toán và độ đo chuẩn hóa bằng ngôn ngữ DAX (Data Analysis Expressions) và truy vấn SQL:", indent=True)
    
    dax_headers = ["Tên độ đo (Measure)", "Công thức DAX chuẩn tắc", "Mã SQL tương đương", "Mục đích nghiệp vụ"]
    dax_widths = [Cm(3.0), Cm(4.8), Cm(4.2), Cm(4.0)]
    dax_align = [WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT]
    dax_data = [
        ["Total Orders", "COUNTROWS('Fact_Sales')", "SELECT COUNT(*) FROM Fact_Sales", "Đếm tổng số đơn hàng"],
        ["Gross Sales", "SUM('Fact_Sales'[Sales_VND])", "SELECT SUM(Sales_VND) FROM Fact_Sales", "Tính tổng doanh thu bán hàng gộp"],
        ["Net Sales", "SUM('Fact_Sales'[Net_Sales_VND])", "SELECT SUM(Net_Sales_VND) FROM Fact_Sales", "Tính tổng doanh thu thuần thực nhận"],
        ["Gross Profit", "SUM('Fact_Sales'[Profit_VND])", "SELECT SUM(Profit_VND) FROM Fact_Sales", "Tính tổng lợi nhuận gộp toàn hệ thống"],
        ["Gross Margin %", "DIVIDE([Gross Profit], [Net Sales], 0)", "SELECT SUM(Profit_VND)*1.0 / SUM(Net_Sales_VND)", "Đo lường biên lợi nhuận gộp"],
        ["AOV (Giá trị đơn TB)", "DIVIDE([Net Sales], [Total Orders], 0)", "SELECT AVG(Net_Sales_VND) FROM Fact_Sales", "Đo lường quy mô giá trị trung bình giỏ hàng"],
        ["Return Rate %", "DIVIDE(CALCULATE([Total Orders], 'Fact_Sales'[Returned]=\"Có\"), [Total Orders], 0)", "SELECT AVG(CASE WHEN Returned='Có' THEN 1.0 ELSE 0 END)", "Tỷ lệ đơn hàng bị hoàn trả"],
        ["Delay Rate %", "DIVIDE(CALCULATE([Total Orders], 'Fact_Sales'[Shipping_Days] > 3), [Total Orders], 0)", "SELECT AVG(CASE WHEN Shipping_Days > 3 THEN 1.0 ELSE 0 END)", "Tỷ lệ đơn hàng giao trễ vượt SLA"],
        ["Average CSAT", "AVERAGE('Fact_Sales'[Customer_Satisfaction])", "SELECT AVG(Customer_Satisfaction) FROM Fact_Sales WHERE CSAT IS NOT NULL", "Điểm hài lòng trung bình (thang 1-5 sao)"],
        ["YoY Net Sales Growth %", "DIVIDE([Net Sales] - CALCULATE([Net Sales], SAMEPERIODLASTYEAR('Dim_Date'[Date])), CALCULATE([Net Sales], SAMEPERIODLASTYEAR('Dim_Date'[Date])), 0)", "SELECT (SUM(Net_2025) - SUM(Net_2024)) / SUM(Net_2024)", "Tốc độ tăng trưởng doanh thu so với cùng kỳ năm trước"]
    ]
    add_table(doc, "Bảng 4.2: Thư viện các hàm tính toán và độ đo nâng cao (DAX / SQL Measures)", dax_headers, dax_data, dax_widths, dax_align, "Thư viện mã nguồn DAX Hasaki BI")

    doc.add_page_break()

    # ==================== CHƯƠNG 5 ====================
    add_heading_1(doc, "CHƯƠNG 5: THIẾT KẾ HỆ THỐNG TRỰC QUAN HÓA DỮ LIỆU (DATA VISUALIZATION & DASHBOARDS)")
    
    add_heading_2(doc, "5.1. Nguyên lý Thiết kế Dashboard Trực quan Cấp Điều hành")
    add_body(doc, "Hệ thống Dashboard của Hasaki được thiết kế dựa trên các nguyên tắc công thái học trực quan (Visual Ergonomics) và tư duy thiết kế lấy người dùng làm trung tâm (User-Centered Design):", indent=True)
    add_bullet(doc, "Các chỉ số sức khỏe tài chính cốt lõi (KPI Cards: Doanh thu thuần, Lợi nhuận gộp, Số đơn hàng, Số khách hàng) được bố trí trang trọng ở góc trên cùng bên trái – nơi mắt người quan sát đầu tiên theo thói quen đọc hình chữ F/Z.", bold_prefix="Cấu trúc phân tầng thị giác (Visual Hierarchy): ")
    add_bullet(doc, "Sử dụng tông màu Deep Navy (#1B365D) cho yếu tố chủ đạo, kết hợp màu Slate Blue (#234975) cho các đồ thị thứ cấp. Màu đỏ được quy ước nghiêm ngặt chỉ sử dụng cho các cảnh báo nguy hiểm (Tỷ lệ hoàn trả cao, biên lợi nhuận dưới 10%, điểm CSAT dưới 3.0).", bold_prefix="Quy chuẩn màu sắc trực quan (Color Palette Standards): ")
    add_bullet(doc, "Bố trí các bộ lọc Slicer linh hoạt (Danh mục sản phẩm, Phân khúc khách hàng, Kênh phân phối) cùng thanh trượt thời gian (Timeline) cho phép người dùng tùy biến phân tích chỉ với một thao tác nhấp chuột.", bold_prefix="Tương tác phân tích đa chiều (Interactive Slicers & Timelines): ")

    add_heading_2(doc, "5.2. Kiến trúc Dashboard Tổng quan Điều hành (Executive Overview Dashboard)")
    add_body(doc, "Dashboard Tổng quan Điều hành cung cấp bức tranh toàn cảnh 360 độ về hiệu suất tài chính và vận hành của toàn bộ chuỗi Hasaki trong suốt 24 tháng quan sát:", indent=True)

    # Insert Figure 5.1 (Executive Overview Dashboard)
    img_dash = os.path.join(IMG_DIR, '2d3900aad2a5ea3e797d162f3bf54021759482ad.png')
    add_figure(doc, img_dash, "Hình 5.1: Dashboard Tổng quan Điều hành: Doanh thu thuần và Lợi nhuận gộp theo kênh, tháng, khu vực, danh mục và phân khúc", "Nguồn: Giao diện tương tác hệ thống Hasaki BI Dashboard 2024–2025", width=Cm(15.5))

    add_heading_2(doc, "5.3. Bảng Đối soát Các Chỉ số Tổng hợp Toàn Bộ Dữ liệu")
    add_body(doc, "Các con số đo lường tài chính và nghiệp vụ cốt lõi trích xuất trực tiếp từ giao diện Dashboard tổng quan được tổng hợp chi tiết trong bảng dưới đây:", indent=True)
    
    sum_headers = ["Chỉ số tổng hợp cốt lõi", "Giá trị ghi nhận toàn hệ thống", "Ý nghĩa và ghi chú quản trị"]
    sum_widths = [Cm(4.5), Cm(4.5), Cm(7.0)]
    sum_align = [WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.LEFT]
    sum_data = [
        ["Tổng Doanh thu thuần (Net Sales)", "24,32 tỷ ₫", "Doanh thu thực tế sau khi đã trừ toàn bộ số tiền chiết khấu giảm giá"],
        ["Tổng Lợi nhuận gộp (Gross Profit)", "7,44 tỷ ₫", "Lợi nhuận gộp toàn chuỗi (chưa trừ chi phí hoa hồng sàn và chi phí marketing)"],
        ["Biên lợi nhuận gộp trung bình (Gross Margin %)", "30,60%", "Giữ mức ổn định gần như tuyệt đối giữa hai năm (2024: 30,57%; 2025: 30,63%)"],
        ["Tổng số đơn hàng thực hiện (Total Orders)", "15.000 đơn hàng", "Tăng trưởng +38,6% trong năm 2025 (từ 6.286 đơn lên 8.714 đơn)"],
        ["Tổng số khách hàng giao dịch (Unique Customers)", "3.253 khách hàng", "Phân bổ trên 3 phân khúc: Khách cá nhân, Khách VIP và Khách hàng B2B"],
        ["Giá trị trung bình một đơn hàng (AOV)", "1,62 triệu ₫", "Có xu hướng suy giảm nhẹ từ 1,67 triệu ₫ (năm 2024) xuống 1,58 triệu ₫ (năm 2025)"],
        ["Tổng số đơn hàng bị hoàn trả (Returned Orders)", "583 đơn hàng (3,89%)", "Gây thất thoát trực tiếp 0,25 tỷ ₫ lợi nhuận gộp và đọng 0,58 tỷ ₫ giá vốn hàng bán"]
    ]
    add_table(doc, "Bảng 5.1: Bảng đối soát các chỉ số tổng hợp toàn hệ thống giai đoạn 2024 - 2025", sum_headers, sum_data, sum_widths, sum_align, "Kết xuất từ Dashboard BI Hasaki")

    doc.add_page_break()

print("Chapters 4 & 5 ready!")

