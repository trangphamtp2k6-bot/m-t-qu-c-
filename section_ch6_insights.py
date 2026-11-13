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

def add_chapter6(doc):
    # ==================== CHƯƠNG 6 ====================
    add_heading_1(doc, "CHƯƠNG 6: ĐÁNH GIÁ VÀ TRÍCH XUẤT INSIGHTS KINH DOANH CHUYÊN SÂU (REVIEW & INSIGHTS GENERATION)")
    add_body(doc, "Dựa trên mô hình dữ liệu Star Schema và các báo cáo phân tích đa chiều được kết xuất từ Dashboard, chương này tập trung đánh giá chuyên sâu và trích xuất các insight quản trị cốt lõi, trả lời tuần tự 21 câu hỏi nghiệp vụ đã đặt ra ở Chương 1, phân tích nguyên nhân gốc rễ và kiểm chứng các giả thuyết kinh doanh ban đầu.", indent=True)

    # ------------------ 6.1 HƯỚNG A ------------------
    add_heading_2(doc, "6.1. Hướng A – Đánh giá Hiệu suất Kinh doanh và Cơ cấu Lợi nhuận")
    
    # Figure 6.1 (Monthly Trends)
    img_trend = os.path.join(IMG_DIR, '705af6b635f8c83181a758913c32d32c743a8076.png')
    add_figure(doc, img_trend, "Hình 6.1: Biểu đồ xu hướng Doanh thu thuần và Lợi nhuận gộp theo tháng kèm đường xu thế", "Nguồn: Báo cáo phân tích chuỗi thời gian BI Hasaki 2024–2025", width=Cm(15.2))

    add_body(doc, "Phân tích chuỗi thời gian (Time-series Analysis) và cấu trúc phân bổ nguồn thu phản ánh bức tranh tài chính đa chiều:", indent=True)
    add_bullet(doc, "Doanh thu thuần năm 2025 đạt 13,81 tỷ ₫, tăng trưởng mạnh mẽ +31,4% so với năm 2024 (10,51 tỷ ₫). Số lượng đơn hàng tăng vượt bậc từ 6.286 lên 8.714 đơn (+38,6%). Lợi nhuận gộp tăng tương ứng từ 3,21 tỷ ₫ lên 4,23 tỷ ₫ (+31,8%). Biên lợi nhuận gộp duy trì sự ổn định đáng kinh ngạc ở mức 30,57% (năm 2024) và 30,63% (năm 2025). Tuy nhiên, giá trị đơn hàng trung bình (AOV) lại suy giảm từ 1,67 triệu ₫ xuống 1,58 triệu ₫ (-5,4%).", bold_prefix="A1 – Đánh giá Tăng trưởng: ")
    add_bullet(doc, "Doanh thu toàn hệ thống đạt đỉnh tuyệt đối vào tháng 03/2025 với 2,09 tỷ ₫ (tạo ra từ 1.262 đơn hàng), trùng khớp với đợt cao điểm khuyến mãi Ngày Quốc tế Phụ nữ 8/3. Tuy nhiên, kể từ tháng 04/2025 trở đi, doanh thu ghi nhận đà suy giảm liên tục và chạm đáy bất thường vào tháng 12/2025 với vỏn vẹn 59 đơn hàng. Độ dốc sụt giảm quá gắt này cảnh báo nguy cơ lỗi trích xuất dữ liệu chưa đầy đủ từ hệ thống ERP vào cuối kỳ hơn là phản ánh sự suy thoái nhu cầu thị trường.", bold_prefix="A2 – Xu hướng Biến động theo Tháng: ")
    add_bullet(doc, "Có sự phân hóa sâu sắc về tỷ trọng đóng góp giữa các danh mục sản phẩm (Bảng 6.1). Nước hoa cao cấp chỉ chiếm 13,7% số đơn hàng nhưng lại đóng góp tới 46,1% doanh thu thuần (11,21 tỷ ₫) và 49,1% tổng lợi nhuận gộp (3,65 tỷ ₫) với biên gộp dẫn đầu 32,6% và AOV vượt trội 5,47 triệu ₫. Ngược lại, Skincare là danh mục chủ lực về lượng đơn (chiếm 40,8% với 6.120 đơn) nhưng chỉ tạo ra 32,5% doanh thu và 30,0% lợi nhuận với biên gộp thấp nhất (28,2%).", bold_prefix="A3 – Cơ cấu Danh mục Sản phẩm: ")

    cat_headers = ["Danh mục sản phẩm", "Số lượng đơn", "Tỷ trọng đơn %", "Doanh thu thuần", "Tỷ trọng thu %", "Lợi nhuận gộp", "Tỷ trọng lãi %", "Biên gộp %", "AOV"]
    cat_widths = [Cm(3.0), Cm(1.6), Cm(1.6), Cm(2.2), Cm(1.5), Cm(2.0), Cm(1.5), Cm(1.6), Cm(1.8)]
    cat_align = [WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT]
    cat_data = [
        ["Nước hoa cao cấp", "2.055", "13,7%", "11,21 tỷ ₫", "46,1%", "3,65 tỷ ₫", "49,1%", "32,6%", "5,47 tr ₫"],
        ["Chăm sóc da (Skincare)", "6.120", "40,8%", "7,91 tỷ ₫", "32,5%", "2,23 tỷ ₫", "30,0%", "28,2%", "1,29 tr ₫"],
        ["Trang điểm (Makeup)", "4.050", "27,0%", "3,16 tỷ ₫", "13,0%", "0,98 tỷ ₫", "13,2%", "31,0%", "0,78 tr ₫"],
        ["Tóc & Cơ thể (Body)", "2.775", "18,5%", "2,04 tỷ ₫", "8,4%", "0,58 tỷ ₫", "7,8%", "28,4%", "0,74 tr ₫"]
    ]
    add_table(doc, "Bảng 6.1: Hiệu suất tài chính và cơ cấu theo danh mục sản phẩm", cat_headers, cat_data, cat_widths, cat_align, "Trích xuất mô hình Star Schema Hasaki")

    # Figure 6.2 (Category breakdown)
    img_cat = os.path.join(IMG_DIR, '9a54b2ced9aeb2c3824faeb2ac4f82a3d75f28d6.png')
    add_figure(doc, img_cat, "Hình 6.2: Biểu đồ so sánh Doanh thu thuần và Lợi nhuận gộp theo Danh mục sản phẩm", "Nguồn: Báo cáo phân tích danh mục BI Hasaki", width=Cm(14.5))

    add_bullet(doc, "Phân khúc Khách hàng cá nhân đóng vai trò nền tảng tạo ra 62,4% doanh thu (15,17 tỷ ₫) và 63,8% lợi nhuận với biên gộp 31,3%. Tuy nhiên, xuất hiện nghịch lý đối với phân khúc VIP: Nhóm khách VIP chiếm 24,2% doanh thu nhưng lại có biên gộp thấp nhất toàn hệ thống (29,05%) do nhận mức chiết khấu bình quân lên tới 11,5% (so với 9,3% của cá nhân). Trong khi đó, phân khúc B2B (đại lý/doanh nghiệp) đạt hiệu quả cao nhất với AOV gấp đôi bình quân (3,13 triệu ₫) và biên gộp dẫn đầu 31,7% dù chỉ được chiết khấu 8,9%.", bold_prefix="A4 – Đóng góp theo Phân khúc Khách hàng: ")

    seg_headers = ["Phân khúc khách hàng", "Doanh thu thuần", "Tỷ trọng thu %", "Lợi nhuận gộp", "Tỷ trọng lãi %", "Biên gộp %", "Chiết khấu TB", "AOV"]
    seg_widths = [Cm(3.5), Cm(2.4), Cm(1.8), Cm(2.2), Cm(1.8), Cm(1.8), Cm(1.8), Cm(2.0)]
    seg_align = [WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT]
    seg_data = [
        ["Khách hàng cá nhân", "15,17 tỷ ₫", "62,4%", "4,75 tỷ ₫", "63,8%", "31,30%", "9,3%", "1,48 tr ₫"],
        ["Khách hàng VIP", "5,89 tỷ ₫", "24,2%", "1,71 tỷ ₫", "22,9%", "29,05%", "11,5%", "1,76 tr ₫"],
        ["Khách hàng B2B", "3,26 tỷ ₫", "13,4%", "1,03 tỷ ₫", "13,9%", "31,70%", "8,9%", "3,13 tr ₫"]
    ]
    add_table(doc, "Bảng 6.2: Hiệu suất theo phân khúc khách hàng (Cá nhân, VIP, B2B)", seg_headers, seg_data, seg_widths, seg_align, "Trích xuất mô hình phân khúc Hasaki")

    # Figure 6.3 (Segment breakdown)
    img_seg = os.path.join(IMG_DIR, 'afdd8d1470fab9e2f2995cc560255b43e99da674.png')
    add_figure(doc, img_seg, "Hình 6.3: Biểu đồ cơ cấu Doanh thu và Lợi nhuận gộp theo Phân khúc khách hàng", "Nguồn: Báo cáo phân khúc khách hàng Hasaki", width=Cm(14.0))

    add_bullet(doc, "Miền Nam là thị trường trọng điểm số 1 khi đóng góp 10,85 tỷ ₫ (44,6% doanh thu); tiếp theo là Miền Bắc với 9,24 tỷ ₫ (38,0%); Miền Trung đạt 4,22 tỷ ₫ (17,4%). Điểm đáng chú ý là biên lợi nhuận gộp ở cả ba miền gần như đồng nhất (30,5% - 30,9%), chứng tỏ sự chênh lệch quy mô doanh số hoàn toàn xuất phát từ mật độ dân cư và số lượng chi nhánh cửa hàng chứ không phải do hiệu quả chính sách giá.", bold_prefix="A5 – Phân bổ theo Khu vực Địa lý: ")

    # Figure 6.4 (Region breakdown)
    img_reg = os.path.join(IMG_DIR, 'bf271b099e4bdb079c3fdabddbe6f809a9db43cb.png')
    add_figure(doc, img_reg, "Hình 6.4: Biểu đồ phân bổ Doanh thu thuần và Lợi nhuận gộp theo Khu vực địa lý", "Nguồn: Báo cáo phân tích địa lý Hasaki", width=Cm(14.0))

    add_callout(doc, "TỔNG HỢP INSIGHT CHIẾN LƯỢC – HƯỚNG A (HIỆU SUẤT & LỢI NHUẬN)", [
        ("Tăng trưởng rộng chứ chưa sâu", "Doanh thu tăng +31,4% hoàn toàn nhờ mở rộng lượng đơn (+38,6%), trong khi AOV giảm -5,4% và biên lợi nhuận đứng yên – doanh nghiệp đang phải xử lý nhiều đơn hàng hơn nhưng nhận về ít giá trị hơn trên mỗi giao dịch."),
        ("Rủi ro phụ thuộc lợi nhuận", "Gần 50% lợi nhuận gộp toàn chuỗi phụ thuộc vào nhóm Nước hoa cao cấp (chỉ chiếm 13,7% số đơn). Trong khi đó, Skincare là nhóm chiếm lượng đơn lớn nhất nhưng biên gộp lại thấp nhất (28,2%), là danh mục dễ bị tổn thương nhất khi có biến động giá."),
        ("Nghịch lý Khách hàng VIP", "Nhóm khách hàng VIP trung thành nhất lại mang về biên lợi nhuận thấp nhất do bị 'chiết khấu quá mức', trong khi tệp khách B2B mua số lượng lớn, AOV gấp đôi lại có biên sinh lời cao nhất nhưng chưa được đầu tư phát triển tương xứng.")
    ])

    # ------------------ 6.2 HƯỚNG B ------------------
    add_heading_2(doc, "6.2. Hướng B – Khuyến mãi và Độ nhạy Giá (Price Elasticity & Discount Trap)")
    add_body(doc, "Chính sách khuyến mãi giảm giá trực tiếp là một trong những vũ khí cạnh tranh chủ lực nhưng đồng thời cũng là 'con dao hai lưỡi' bào mòn lợi nhuận nếu không được kiểm soát dựa trên dữ liệu:", indent=True)

    # Figure 6.5 (Discount tiers)
    img_disc = os.path.join(IMG_DIR, 'b204238171ea5cc0ddebd79ebc852d5d58106a2f.png')
    add_figure(doc, img_disc, "Hình 6.5: Biểu đồ phân tích Tỷ trọng số đơn và Biên lợi nhuận gộp theo mức Chiết khấu", "Nguồn: Mô hình phân tích độ nhạy giá Hasaki BI", width=Cm(15.2))

    disc_headers = ["Mức chiết khấu", "Số lượng đơn", "% Tổng số đơn", "% Lợi nhuận gộp", "Biên gộp %", "AOV bình quân", "Tỷ lệ hoàn trả %"]
    disc_widths = [Cm(2.6), Cm(2.0), Cm(2.2), Cm(2.4), Cm(2.0), Cm(2.4), Cm(2.4)]
    disc_align = [WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT]
    disc_data = [
        ["0% (Nguyên giá)", "1.740", "11,6%", "14,3%", "37,4%", "1,96 tr ₫", "3,1%"],
        ["5%", "3.840", "25,6%", "28,4%", "33,8%", "1,78 tr ₫", "3,3%"],
        ["10%", "3.825", "25,5%", "24,8%", "30,2%", "1,65 tr ₫", "3,5%"],
        ["15%", "3.195", "21,3%", "24,3%", "26,5%", "1,52 tr ₫", "4,0%"],
        ["20%", "1.305", "8,7%", "5,4%", "21,9%", "1,38 tr ₫", "5,1%"],
        ["25%", "675", "4,5%", "1,9%", "16,5%", "1,25 tr ₫", "6,8%"],
        ["30%", "420", "2,8%", "0,9%", "10,2%", "1,17 tr ₫", "7,9%"]
    ]
    add_table(doc, "Bảng 6.3: Số đơn, lợi nhuận gộp, biên gộp và tỷ lệ hoàn trả theo mức chiết khấu", disc_headers, disc_data, disc_widths, disc_align, "Phân tích độ co giãn giá Hasaki")

    add_bullet(doc, "Biên lợi nhuận gộp suy giảm gần như tuyến tính: Cứ mỗi khi nâng chiết khấu thêm 5%, biên lợi nhuận gộp toàn chuỗi bị 'bốc hơi' khoảng 4,5 đến 5,5 điểm phần trăm. Từ mức đỉnh 37,4% (khi bán nguyên giá), biên gộp rơi xuống 30,2% ở mức giảm 10%, và tụt dốc thảm hại xuống chỉ còn 10,2% ở mức giảm 30%.", bold_prefix="B1 – Biến thiên Biên lợi nhuận: ")
    add_bullet(doc, "Dữ liệu thực nghiệm xác lập rõ ràng 'Điểm ngọt' (Sweet Spot) của hệ thống Hasaki nằm trong dải chiết khấu từ 5% đến 10%. Hai mức giảm này cộng lại chiếm 51,1% tổng số lượng đơn hàng (7.665 đơn) và đóng góp tới 53,2% tổng lợi nhuận gộp của doanh nghiệp, vừa đủ sức hấp dẫn khách hàng chốt đơn vừa bảo toàn biên gộp an toàn quanh mức 30% - 34%.", bold_prefix="B2 – Xác định 'Điểm ngọt' Khuyến mãi: ")
    add_bullet(doc, "Các chương trình giảm giá sâu (từ 20% trở lên) chiếm tới 16,0% số lượng đơn (2.400 đơn hàng) nhưng chỉ đem về vỏn vẹn 8,2% tổng lợi nhuận gộp. Riêng nhóm giảm 25% - 30% (chiếm 7,3% đơn) chỉ tạo ra 2,8% lợi nhuận! Thậm chí, giá trị đơn hàng AOV rơi tự do từ 1,96 triệu ₫ xuống 1,17 triệu ₫ (-40%), đồng thời tỷ lệ hoàn trả vọt lên 7,9%. Giảm giá sâu rõ ràng không kích thích mua thêm mà chỉ thu hút tệp khách 'săn sale giá bèo' có rủi ro hủy hàng cao.", bold_prefix="B3 – Hiệu quả của Chiết khấu Sâu: ")
    add_bullet(doc, "Danh mục Chăm sóc da (Skincare) là nhóm nhạy cảm nhất toàn chuỗi: Khi chiết khấu 0%, Skincare có biên gộp 35,0%; nhưng khi chiết khấu 25%, biên gộp sụp đổ xuống 14,0%; và ở mức giảm 30%, biên gộp chỉ còn 7,3% (thấp hơn ngưỡng an toàn 10%). Điều này đồng nghĩa với việc cứ mỗi đơn hàng Skincare giảm giá 30% bán ra, sau khi trừ chi phí vận hành đóng gói, Hasaki đang rơi vào tình trạng bán lỗ!", bold_prefix="B4 – Độ nhạy cảm theo Ngành hàng: ")

    # Figure 6.6 (Elasticity curves)
    img_elas = os.path.join(IMG_DIR, 'd3c28026e85ea244c9be68f2607b5d381b8e3c70.png')
    add_figure(doc, img_elas, "Hình 6.6: Đồ thị đường cong Độ co giãn Biên lợi nhuận của từng Danh mục theo Chiết khấu", "Nguồn: Mô hình đường cong độ co giãn Hasaki BI", width=Cm(15.2))

    add_callout(doc, "TỔNG HỢP INSIGHT CHIẾN LƯỢC – HƯỚNG B (KHUYẾN MÃI & GIÁ)", [
        ("Mỗi 1% chiết khấu ăn mòn 1% biên gộp", "Không có phép màu quy mô bù đắp giá: Chiết khấu sâu làm xói mòn trực tiếp từng đồng lợi nhuận ròng mà không làm tăng giá trị giỏ hàng AOV."),
        ("Dồn lực về dải Sweet Spot 5% - 10%", "Đây là vùng tối ưu tuyệt đối mang lại hơn 53% lợi nhuận toàn công ty. Các mức chiết khấu trên 15% phải được đặt dưới sự kiểm duyệt nghiêm ngặt."),
        ("Báo động đỏ cho Skincare giảm giá", "Chính sách chiết khấu hiện tại đang cào bằng (~10% cho mọi danh mục). Cần lập tức áp trần chiết khấu tối đa 15% cho danh mục Skincare để ngăn chặn nguy cơ bán lỗ dưới giá vốn vận hành.")
    ])

    # ------------------ 6.3 HƯỚNG C ------------------
    add_heading_2(doc, "6.3. Hướng C – Hiệu suất Bán hàng Đa kênh và Xung đột Kênh (Omnichannel Performance)")
    add_body(doc, "Mô hình đa kênh (Omnichannel) đòi hỏi sự phân định rõ ràng vai trò chiến lược của từng kênh phân phối thay vì cạnh tranh nội bộ triệt tiêu lẫn nhau:", indent=True)

    # Figure 6.7 (Channel performance)
    img_chan = os.path.join(IMG_DIR, '9894f6daeca7f0d5a6eea2d9ecd362e734dcd177.png')
    add_figure(doc, img_chan, "Hình 6.7: Biểu đồ so sánh Doanh thu thuần và Lợi nhuận gộp theo Kênh bán hàng", "Nguồn: Báo cáo phân tích kênh BI Hasaki", width=Cm(15.2))

    chan_headers = ["Kênh bán hàng", "Doanh thu thuần", "Tỷ trọng thu %", "Lợi nhuận gộp", "Tỷ trọng lãi %", "Biên gộp %", "AOV", "Điểm CSAT", "Tỷ lệ hoàn %"]
    chan_widths = [Cm(3.8), Cm(2.2), Cm(1.6), Cm(2.0), Cm(1.6), Cm(1.6), Cm(1.8), Cm(1.5), Cm(1.5)]
    chan_align = [WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT]
    chan_data = [
        ["Cửa hàng trực tiếp (In-Store)", "11,09 tỷ ₫", "45,6%", "3,41 tỷ ₫", "45,8%", "30,7%", "1,65 tr ₫", "4,21", "1,9%"],
        ["Website & App Hasaki", "8,58 tỷ ₫", "35,3%", "2,62 tỷ ₫", "35,2%", "30,5%", "1,61 tr ₫", "3,98", "4,8%"],
        ["Sàn TMĐT (Shopee/TikTok/Lazada)", "4,65 tỷ ₫", "19,1%", "1,41 tỷ ₫", "19,0%", "30,4%", "1,58 tr ₫", "3,82", "7,2%"]
    ]
    add_table(doc, "Bảng 6.4: Đánh giá toàn diện hiệu suất và chất lượng vận hành theo kênh bán hàng", chan_headers, chan_data, chan_widths, chan_align, "Báo cáo đa kênh Hasaki 2026")

    add_bullet(doc, "Hệ thống hơn 140 cửa hàng vật lý tiếp tục là 'xương sống' tạo doanh thu và lợi nhuận (chiếm 45,6% doanh thu và 45,8% lợi nhuận). Kênh sở hữu số (Website/App) đứng thứ hai với 35,3% doanh thu và 35,2% lợi nhuận. Kênh sàn TMĐT đóng vai trò phễu mở rộng tiếp cận tệp khách mới với 19,1% doanh thu và 19,0% lợi nhuận gộp.", bold_prefix="C1 – Đóng góp Nguồn thu: ")
    add_bullet(doc, "Trên báo cáo tài chính kế toán thô, biên lợi nhuận gộp của cả ba kênh gần như tương đương nhau (30,4% - 30,7%). Tuy nhiên, đây là một 'biên lợi nhuận ảo' (Optical Illusion). Tập dữ liệu thô chưa khấu trừ chi phí hoa hồng sàn TMĐT (Take-rate dao động từ 8% đến 15%), phí thanh toán và phí Freeship Extra. Nếu khấu trừ đầy đủ các khoản phí sàn này, biên lợi nhuận thực tế của sàn TMĐT sẽ tụt xuống dưới 20%, thấp hơn rất nhiều so với kênh Cửa hàng và App nội bộ.", bold_prefix="C2 – Bản chất Biên lợi nhuận theo Kênh: ")
    add_bullet(doc, "Kênh Cửa hàng vượt trội hoàn toàn về chất lượng dịch vụ với điểm CSAT trung bình đạt 4,21 sao và tỷ lệ hoàn trả chỉ 1,9% (nhờ khách hàng được tư vấn soi da trực tiếp và kiểm tra sản phẩm tại chỗ). Trong khi đó, Sàn TMĐT ghi nhận tỷ lệ hoàn trả kỷ lục lên tới 7,2% (gấp gần 4 lần cửa hàng) và CSAT thấp nhất (3,82 sao).", bold_prefix="C3 – Trải nghiệm Khách hàng và Tỷ lệ Hoàn: ")
    add_bullet(doc, "Hành vi lựa chọn kênh mua sắm phân bổ rất đồng đều giữa các phân khúc khách hàng (Cả khách Cá nhân, VIP và B2B đều mua khoảng 45% tại Cửa hàng, 35% qua Website/App và 20% qua Sàn TMĐT). Điều này chứng minh việc tối ưu hóa cần tập trung vào hành trình chuyển đổi kênh (Channel Transition từ TMĐT sang App Hasaki) thay vì phân biệt đối xử theo phân khúc.", bold_prefix="C4 – Mối liên kết Phân khúc và Kênh: ")

    add_callout(doc, "TỔNG HỢP INSIGHT CHIẾN LƯỢC – HƯỚNG C (HIỆU SUẤT ĐA KÊNH)", [
        ("Sàn TMĐT là phễu hút khách, không phải nơi tối đa hóa lợi nhuận", "Sau khi tính đủ chi phí hoa hồng sàn và chi phí hoàn trả 7,2%, TMĐT có biên lãi ròng mỏng nhất. Cần định vị TMĐT như kênh thu hút khách hàng mới (Acquisition) để kéo người dùng về App Hasaki."),
        ("Cửa hàng vật lý là mỏ neo giữ vững thương hiệu và dòng tiền", "Trải nghiệm vượt trội tại điểm bán (CSAT 4,21) là tài sản vô giá cần phát huy thông qua mô hình O2O (Click-and-Collect: Đặt online nhận tại cửa hàng).")
    ])

    # ------------------ 6.4 HƯỚNG D ------------------
    add_heading_2(doc, "6.4. Hướng D – Quản trị Vận hành, Điểm gãy Hài lòng và Thiệt hại Hoàn hàng (Logistics & Satisfaction)")
    add_body(doc, "Dịch vụ giao nhận chặng cuối (Last-Mile Delivery) tác động trực tiếp đến sự gắn bó của khách hàng đối với thương hiệu Hasaki:", indent=True)

    # Figure 6.8 (Shipping days vs CSAT & Return)
    img_ship = os.path.join(IMG_DIR, '23d01866298b2b456c68b413e9bd8c3d6d872e4e.png')
    add_figure(doc, img_ship, "Hình 6.8: Biểu đồ phân tích Điểm hài lòng (CSAT) và Tỷ lệ hoàn trả theo Thời gian giao hàng", "Nguồn: Mô hình phân tích vận hành logistics Hasaki", width=Cm(15.2))

    ship_headers = ["Thời gian giao hàng", "Số lượng đơn", "% Tổng số đơn", "Điểm CSAT bình quân", "Tỷ lệ đơn 1-2 sao", "Tỷ lệ hoàn trả %"]
    ship_widths = [Cm(3.5), Cm(2.5), Cm(2.5), Cm(2.5), Cm(2.5), Cm(2.5)]
    ship_align = [WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT]
    ship_data = [
        ["0 ngày (Mua tại quầy)", "6.840", "45,6%", "4,21", "1,8%", "1,9%"],
        ["1 ngày", "2.145", "14,3%", "4,09", "3,1%", "5,2%"],
        ["2 ngày", "2.175", "14,5%", "4,07", "3,4%", "5,8%"],
        ["3 ngày (Ngưỡng an toàn)", "1.780", "11,9%", "4,08", "3,5%", "6,1%"],
        ["4 ngày (Ngưỡng gãy)", "1.185", "7,9%", "2,70", "43,5%", "6,5%"],
        ["≥ 5 ngày", "875", "5,8%", "2,45", "51,2%", "6,6%"]
    ]
    add_table(doc, "Bảng 6.5: Mức độ hài lòng (CSAT) và tỷ lệ hoàn trả theo số ngày giao hàng", ship_headers, ship_data, ship_widths, ship_align, "Báo cáo phân tích SLA giao hàng Hasaki")

    add_bullet(doc, "Dữ liệu thực nghiệm bộc lộ một 'vách đá sụp đổ' (Cliff Effect) về độ hài lòng khách hàng nằm chính xác ở mốc 3 ngày. Trong khi các đơn mua tại chỗ đạt CSAT 4,21 và đơn giao từ 1 đến 3 ngày giữ vững sự hài lòng ở mức cao (4,07 - 4,09 sao), thì chỉ cần thời gian giao hàng bước sang ngày thứ 4, điểm CSAT rơi thẳng đứng xuống 2,70 sao (-1,38 điểm). Tỷ lệ khách hàng phẫn nộ đánh giá 1-2 sao vọt từ 3,5% lên tới 43,5%!", bold_prefix="D1 – Ngưỡng gãy Trải nghiệm (SLA Threshold): ")
    add_bullet(doc, "Nhóm đơn hàng giao trễ vượt chuẩn SLA (> 3 ngày) chiếm 13,7% tổng số đơn toàn chuỗi (tương đương 2.060 đơn hàng). Đáng lưu ý, 100% số đơn giao trễ này phát sinh ở hai kênh từ xa (Website/App chiếm 63,5% và Sàn TMĐT chiếm 36,5%). Về mặt địa lý, Miền Trung là khu vực có tỷ lệ giao trễ cao nhất hệ thống (15,8% so với 13,2% - 13,4% của Miền Bắc và Miền Nam) do điểm nghẽn về hạ tầng kho vận trung chuyển.", bold_prefix="D2 – Phân bổ Đơn hàng Giao trễ: ")
    add_bullet(doc, "Toàn hệ thống ghi nhận 583 đơn hàng bị hoàn trả (chiếm 3,89% tổng đơn). Tổn thất tài chính trực tiếp bao gồm 0,25 tỷ ₫ lợi nhuận gộp bị bốc hơi, đồng thời khiến 0,58 tỷ ₫ giá vốn hàng bán bị đọng vốn, hư hại vỏ hộp hoặc hết hạn lưu kho (Bảng 6.6).", bold_prefix="D3 – Tổn thất Tài chính do Hoàn trả: ")

    ret_headers = ["Chỉ tiêu đo lường", "Toàn bộ hệ thống", "Chăm sóc da (Skincare)", "Nước hoa cao cấp", "Các danh mục khác"]
    ret_widths = [Cm(4.5), Cm(3.0), Cm(3.0), Cm(3.0), Cm(2.5)]
    ret_align = [WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT]
    ret_data = [
        ["Tổng số đơn hàng hoàn trả", "583 đơn", "247 đơn (42,4%)", "70 đơn (12,0%)", "266 đơn (45,6%)"],
        ["Tỷ lệ hoàn trả danh mục %", "3,89%", "4,04%", "3,41%", "3,91%"],
        ["Lợi nhuận gộp bị mất", "0,25 tỷ ₫", "0,09 tỷ ₫ (36,0%)", "0,11 tỷ ₫ (44,0%)", "0,05 tỷ ₫ (20,0%)"],
        ["Giá vốn hàng hoàn bị đọng (COGS)", "0,58 tỷ ₫", "0,22 tỷ ₫ (37,9%)", "0,20 tỷ ₫ (34,5%)", "0,16 tỷ ₫ (27,6%)"]
    ]
    add_table(doc, "Bảng 6.6: Tổn thất tài chính do hoàn trả theo danh mục hàng hóa", ret_headers, ret_data, ret_widths, ret_align, "Thống kê thiệt hại hoàn hàng Hasaki")

    # Figure 6.9 (Returns by channel & category)
    img_ret = os.path.join(IMG_DIR, '71a440a209f226a39617fd85295f91099881d233.png')
    add_figure(doc, img_ret, "Hình 6.9: Biểu đồ phân bổ Tỷ lệ hoàn trả theo Kênh bán hàng và theo Danh mục hàng hóa", "Nguồn: Báo cáo tỷ lệ hoàn hàng Hasaki BI", width=Cm(15.2))

    add_bullet(doc, "Phân tích tương quan bộc lộ phát hiện quan trọng: Tỷ lệ hoàn trả không tăng đột biến theo số ngày giao hàng (Đơn giao 1 đến 5 ngày đều có tỷ lệ hoàn tương đương nhau từ 5,2% đến 6,6%). Điều này chứng minh giao trễ làm mất sự hài lòng (CSAT giảm), nhưng yếu tố thúc đẩy khách từ chối nhận hàng là hình thức giao từ xa nói chung kết hợp với rủi ro bao bì hư hại trong vận chuyển.", bold_prefix="D4 – Mối liên hệ Giao trễ và Hoàn trả: ")
    add_bullet(doc, "Danh mục Skincare chiếm tới 42,4% tổng số lượng đơn hoàn trả toàn hệ thống (247 đơn). Trong khi đó, Nước hoa chỉ chiếm 12,0% số lượng đơn hoàn nhưng lại gây tổn thất tới gần 35% giá vốn hàng hoàn (202 triệu ₫) do giá trị đơn vị sản phẩm rất cao.", bold_prefix="D5 – Trọng điểm Tổn thất Hàng hóa: ")

    add_callout(doc, "TỔNG HỢP INSIGHT CHIẾN LƯỢC – HƯỚNG D (VẬN HÀNH & LOGISTICS)", [
        ("Mốc 3 ngày là ranh giới sống còn của SLA vận hành", "Mọi nỗ lực logistics phải cam kết giao hàng trong vòng ≤ 3 ngày. Vượt qua mốc này, mức độ hài lòng sụp đổ và thương hiệu phải chịu điểm đánh giá tiêu cực nghiêm trọng."),
        ("Bảo vệ bao bì Nước hoa và Skincare là ưu tiên số 1", "70 đơn nước hoa bị hoàn làm mất hơn 200 triệu vốn lưu động. Cần quy chuẩn hóa quy trình đóng gói chống sốc đa lớp chuyên biệt.")
    ])

    # ------------------ 6.5 HƯỚNG E ------------------
    add_heading_2(doc, "6.5. Hướng E – Xu hướng Thanh toán Không tiền mặt và Hành vi Hoàn hàng")
    add_body(doc, "Phương thức thanh toán phản ánh mức độ ứng dụng kinh tế số và hành vi tiêu dùng của khách hàng mỹ phẩm:", indent=True)

    # Figure 6.10 (Payment methods)
    img_pay = os.path.join(IMG_DIR, '6bd2fa01c79b3856095f9ddcac34036d73529499.png')
    add_figure(doc, img_pay, "Hình 6.10: Biểu đồ cơ cấu Phương thức thanh toán và Tỷ lệ hoàn trả tương ứng", "Nguồn: Báo cáo hành vi thanh toán Hasaki BI", width=Cm(15.2))

    pay_headers = ["Phương thức thanh toán", "Số lượng đơn", "Tỷ trọng đơn %", "Tỷ lệ hoàn trả chung", "Tỷ lệ hoàn trả (Đơn online)"]
    pay_widths = [Cm(4.5), Cm(2.5), Cm(2.5), Cm(3.0), Cm(3.5)]
    pay_align = [WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT]
    pay_data = [
        ["VNPay QR Code", "3.795", "25,3%", "4,0%", "5,5%"],
        ["Thẻ tín dụng / ATM (POS & Online)", "3.285", "21,9%", "3,7%", "5,4%"],
        ["Ví điện tử MoMo", "3.075", "20,5%", "4,1%", "5,6%"],
        ["Tiền mặt / Thu hộ COD", "2.760", "18,4%", "3,2%", "5,4%"],
        ["Ví điện tử ZaloPay", "1.260", "8,4%", "3,9%", "5,5%"],
        ["Chuyển khoản ngân hàng trực tiếp", "825", "5,5%", "3,8%", "5,3%"]
    ]
    add_table(doc, "Bảng 6.7: Cơ cấu phương thức thanh toán và tỷ lệ hoàn trả tương ứng", pay_headers, pay_data, pay_widths, pay_align, "Thống kê phương thức thanh toán Hasaki")

    add_bullet(doc, "Thanh toán không tiền mặt (Cashless Payment) đã chiếm tỷ trọng áp đảo lên tới 81,6% tổng số lượng giao dịch toàn hệ thống (VNPay QR dẫn đầu 25,3%, quẹt thẻ 21,9%, MoMo 20,5%, ZaloPay 8,4%, Chuyển khoản 5,5%). Tiền mặt và COD chỉ còn chiếm 18,4%. Cơ cấu này gần như không suy chuyển suốt hai năm 2024 (COD 18,0%) và 2025 (COD 18,7%), phản ánh thị trường đã bước vào giai đoạn bão hòa thanh toán số.", bold_prefix="E1 – Cơ cấu Phương thức Thanh toán: ")
    add_bullet(doc, "Trên số liệu gộp toàn hệ thống, tỷ lệ hoàn trả của đơn hàng COD chỉ là 3,2%, thấp hơn cả thanh toán qua ví điện tử (4,1%) và quẹt thẻ (3,7%). Tuy nhiên, nguyên nhân là do phần lớn giao dịch tiền mặt phát sinh trực tiếp tại quầy cửa hàng (nơi hoàn trả chỉ 1,9%).", bold_prefix="E2 – Đối soát Tỷ lệ Hoàn trả Gộp: ")
    add_bullet(doc, "Khi tiến hành bóc tách riêng nhóm đơn hàng giao từ xa (Website/App và sàn TMĐT), tỷ lệ hoàn trả của COD là 5,4%, hoàn toàn tương đương với tỷ lệ hoàn trả của các phương thức thanh toán trả trước (5,4% - 5,6%). Điều này bác bỏ định kiến phổ biến cho rằng khách hàng đặt COD sẽ có tỷ lệ 'bùng hàng' cao hơn khách trả tiền trước. Yếu tố cốt lõi quyết định việc nhận hàng nằm ở chất lượng dịch vụ vận chuyển và sự bảo toàn của kiện hàng.", bold_prefix="E3 – Bóc tách Thực chứng Đơn hàng Trực tuyến: ")

    add_callout(doc, "TỔNG HỢP INSIGHT CHIẾN LƯỢC – HƯỚNG E (PHƯƠNG THỨC THANH TOÁN)", [
        ("Dừng trợ giá thúc đẩy thanh toán số", "Với tỷ lệ thanh toán số đã đạt 81,6% và đi ngang, ngân sách chiết khấu thêm 3-5% để khuyến khích dùng ví điện tử là khoản chi phí lãng phí không mang lại hiệu quả gia tăng."),
        ("Công bằng trong quản trị rủi ro COD", "Không nên siết chặt hoặc phân biệt đối xử với đơn COD vì bản chất rủi ro hoàn trả giữa COD và thanh toán trước trên kênh giao hàng từ xa là như nhau.")
    ])

    # ------------------ 6.6 NGUYÊN NHÂN GỐC RỄ ------------------
    add_heading_2(doc, "6.6. Phân tích Nguyên nhân Gốc rễ Đa chiều (Root Cause Analysis - 5 Whys Matrix)")
    add_body(doc, "Áp dụng kỹ thuật phân tích 5 Whys và biểu đồ xương cá (Ishikawa), hệ thống BI kết nối các vấn đề bề nổi (V1 - V5) để tìm ra 4 nguyên nhân gốc rễ mang tính hệ thống:", indent=True)
    
    root_headers = ["Mã gốc rễ", "Nguyên nhân gốc rễ cốt lõi (Root Cause)", "Bằng chứng định lượng từ hệ thống BI", "Các vấn đề bề nổi liên đới"]
    root_widths = [Cm(2.0), Cm(5.0), Cm(5.5), Cm(3.5)]
    root_align = [WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT]
    root_data = [
        ["N1", "Chính sách giá cào bằng, chưa phân hóa theo biên lợi nhuận danh mục", "Skincare giảm 30% thì biên sụp xuống 7,3%; VIP giảm 11,5% có biên thấp hơn khách thường", "V1 (Biên không tăng), V2 (Chiết khấu sâu ăn mòn lãi)"],
        ["N2", "Chiết khấu sâu không đi kèm điều kiện giá trị đơn hàng tối thiểu", "AOV giảm từ 1,96 tr ₫ xuống 1,17 tr ₫ khi giảm 30%; nhóm ≥20% chiếm 16% đơn nhưng chỉ đem 8% lãi", "V1 (AOV giảm), V2 (Trợ giá không hiệu quả)"],
        ["N3", "Thiếu quy chuẩn đóng gói chuyên biệt cho hàng dễ vỡ và thiếu cảnh báo SLA 3 ngày", "CSAT tụt từ 4,09 xuống 2,70 khi giao >3 ngày; 70 đơn nước hoa hoàn làm mất 202 tr vốn", "V4 (CSAT giảm khi giao trễ, thiệt hại hoàn trả 583 tr)"],
        ["N4", "Thiếu cơ chế dữ liệu ghi nhận chi phí ẩn (phí sàn TMĐT, chi phí đóng gói, hoàn trả)", "Biên gộp TMĐT bằng cửa hàng (30,5%) là 'ảo'; dữ liệu 2025 thiếu đơn tháng cuối kỳ", "V3 (Hiểu sai lợi nhuận kênh), V5 (Ngân sách thanh toán số lãng phí)"]
    ]
    add_table(doc, "Bảng 6.8: Ma trận phân tích nguyên nhân gốc rễ (Root Cause Analysis - 5 Whys)", root_headers, root_data, root_widths, root_align, "Mô hình phân tích nguyên nhân Hasaki")

    add_heading_2(doc, "6.7. Đối chiếu Giả thuyết Quản trị Ban đầu với Thực chứng Dữ liệu")
    add_body(doc, "Quá trình phân tích dữ liệu thực nghiệm đã giúp Ban Giám đốc kiểm chứng và điều chỉnh lại những định kiến quản trị ban đầu:", indent=True)

    hyp_headers = ["Chủ đề kinh doanh", "Giả thuyết quản trị ban đầu", "Thực chứng từ dữ liệu BI Hasaki", "Kết luận điều chỉnh"]
    hyp_widths = [Cm(3.0), Cm(4.5), Cm(5.0), Cm(3.5)]
    hyp_align = [WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT]
    hyp_data = [
        ["Khuyến mãi giảm giá", "Giảm giá càng sâu thì doanh số và lợi nhuận càng tăng nhờ hiệu ứng quy mô", "Chiết khấu ≥20% chỉ mang lại 8,2% lãi; AOV giảm 40%; Skincare biên tụt xuống 7,3%", "Bác bỏ. Cần áp trần chiết khấu 15%"],
        ["Kênh Sàn TMĐT", "Sàn TMĐT là động lực tăng trưởng lợi nhuận chính của tương lai", "Biên gộp chưa trừ phí sàn; tỷ lệ hoàn trả lên tới 7,2% (gấp 4 lần cửa hàng), CSAT thấp nhất", "Điều chỉnh. TMĐT là phễu hút khách, chuyển đổi về App"],
        ["Hài lòng & Giao hàng", "Thời gian giao hàng ảnh hưởng tuyến tính đều đặn đến sự hài lòng", "Tồn tại ngưỡng đứt gãy tại mốc 3 ngày (CSAT giữ 4.08 rồi sụp xuống 2.70 ở ngày thứ 4)", "Khẳng định có điều kiện. SLA cam kết phải là ≤3 ngày"],
        ["Thanh toán & Hoàn hàng", "Đơn hàng thanh toán COD có tỷ lệ hoàn trả cao hơn thanh toán trước", "Trên kênh giao từ xa, hoàn trả của COD (5,4%) và thanh toán trực tuyến (5,5%) là như nhau", "Bác bỏ. Hoàn trả do chất lượng vận chuyển, không phải do COD"],
        ["Tệp Khách VIP", "Khách VIP mang lại tỷ suất lợi nhuận cao nhất cho hệ thống", "Khách VIP có biên gộp thấp nhất (29,05%) do nhận chiết khấu quá sâu (11,5%)", "Bác bỏ. Cần chuyển từ ưu đãi tiền sang đặc quyền dịch vụ"]
    ]
    add_table(doc, "Bảng 6.9: Đối chiếu giả thuyết quản trị ban đầu với dữ liệu thực chứng", hyp_headers, hyp_data, hyp_widths, hyp_align, "Đánh giá thực chứng Hasaki BI")

    doc.add_page_break()

print("Chapter 6 ready!")

