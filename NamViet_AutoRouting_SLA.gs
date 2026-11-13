/**
 * HỆ THỐNG TỰ ĐỘNG HÓA ĐIỀU PHỐI LEADS & GIÁM SÁT SLA 4H/24H - CÔNG TY NAM VIỆT
 * Nền tảng: Google Sheets + Google Apps Script + Telegram Bot / Gmail (Miễn phí 100%)
 * 
 * HƯỚNG DẪN TRIỂN KHAI (DEPLOY):
 * 1. Mở Google Sheets của bạn.
 * 2. Vào menu: Tiện ích mở rộng (Extensions) -> Apps Script.
 * 3. Xóa code cũ, dán toàn bộ code này vào file Code.gs.
 * 4. Chạy hàm `khoiTaoHeThong()` một lần đầu tiên để tạo cấu trúc bảng chuẩn.
 * 5. Cài đặt Triggers (Đồng hồ kích hoạt) chạy hàm `runLeadAutomation` mỗi 15 phút.
 */

// ================= CẤU HÌNH HỆ THỐNG =================
const CONFIG = {
  TELEGRAM_BOT_TOKEN: "NHAP_TOKEN_BOT_TELEGRAM_CUA_BAN", // Lấy từ @BotFather (tùy chọn)
  TPKD_TELEGRAM_ID: "NHAP_CHAT_ID_TRUONG_PHONG",
  TPKD_EMAIL: "tpkd@namviet.vn",
  VIP_BUDGET_THRESHOLD: 500000000, // 500 triệu VNĐ (Luồng VIP đẩy TPKD)
  MAX_BACKLOG_CAP: 30,             // Ngưỡng trần 30 leads chưa gọi/NVKD (Luồng Overflow)
  SLA_WARN_HOURS: 4,               // Nhắc nhở sau 4 giờ
  SLA_RECLAIM_HOURS: 24            // Thu hồi sau 24 giờ
};

/**
 * HÀM KHỞI TẠO CẤU TRÚC HỆ THỐNG (Chạy 1 lần khi Deploy mới)
 */
function khoiTaoHeThong() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Tạo sheet Ma_Tran_8_NVKD nếu chưa có
  let wsMatrix = ss.getSheetByName("Ma_Tran_8_NVKD");
  if (!wsMatrix) {
    wsMatrix = ss.insertSheet("Ma_Tran_8_NVKD");
    const matrixInit = [
      ["Mã & Tên NVKD", "Gói Sản Phẩm", "Tỉnh Thành / Khu Vực", "Telegram Chat ID", "Email", "URL_Hoac_ID_File_Rieng"],
      ["NV01 - Nguyễn Văn An", "PROD_RETAIL", "Hà Nội, Hải Phòng, Quảng Ninh, Bắc Ninh, Hải Dương, Hưng Yên", "", "an.nv@namviet.vn", ""],
      ["NV02 - Trần Thị Bình", "PROD_RETAIL", "TP. Hồ Chí Minh, Bình Dương, Đồng Nai, Long An, Cần Thơ", "", "binh.tt@namviet.vn", ""],
      ["NV03 - Lê Hoàng Cường", "PROD_FNB", "Hà Nội, Đà Nẵng, Quảng Nam, Thừa Thiên Huế, Khánh Hòa, Hải Phòng", "", "cuong.lh@namviet.vn", ""],
      ["NV04 - Phạm Minh Dũng", "PROD_FNB", "TP. Hồ Chí Minh, Bình Dương, Đồng Nai, Bà Rịa - Vũng Tàu, Cần Thơ", "", "dung.pm@namviet.vn", ""],
      ["NV05 - Vũ Thị Hoa", "PROD_SERVICE", "Toàn quốc", "", "hoa.vt@namviet.vn", ""],
      ["NV06 - Đỗ Tuấn Kiệt", "PROD_MFG", "Toàn quốc", "", "kiet.dt@namviet.vn", ""],
      ["NV07 - Bùi Thu Lan", "ĐA_NGANH", "Hà Nội, Hải Phòng, Bắc Ninh, Thái Nguyên, Vĩnh Phúc, Thanh Hóa", "", "lan.bt@namviet.vn", ""],
      ["NV08 - Ngô Quang Nam", "ĐA_NGANH", "TP. Hồ Chí Minh, Bình Dương, Đồng Nai, Long An, Tiền Giang, Bình Thuận", "", "nam.nq@namviet.vn", ""]
    ];
    wsMatrix.getRange(1, 1, matrixInit.length, matrixInit[0].length).setValues(matrixInit);
    wsMatrix.getRange(1, 1, 1, matrixInit[0].length).setBackground("#1B365D").setFontColor("#FFFFFF").setFontWeight("bold");
  }

  // 2. Tạo sheet Blacklist_CSKH nếu chưa có
  let wsBlack = ss.getSheetByName("Blacklist_CSKH");
  if (!wsBlack) {
    wsBlack = ss.insertSheet("Blacklist_CSKH");
    const blackInit = [
      ["Số Điện Thoại", "Địa Chỉ Email", "Loại Ràng Buộc", "Nhân Viên Chăm Sóc Gốc / Ghi Chú"],
      ["0909999000", "blacklist@spam.com", "Blacklist", "Do Not Contact - Khách từ chối tiếp cận"],
      ["0988000111", "oldclient@abccorp.vn", "Khách cũ", "NV03 - Lê Hoàng Cường (HĐ 2025)"],
      ["0918000222", "khieunai@badfeedback.com", "Khiếu nại CSKH", "Từng khiếu nại chất lượng - Chuyển TPKD"]
    ];
    wsBlack.getRange(1, 1, blackInit.length, blackInit[0].length).setValues(blackInit);
    wsBlack.getRange(1, 1, 1, blackInit[0].length).setBackground("#B42318").setFontColor("#FFFFFF").setFontWeight("bold");
  }

  SpreadsheetApp.getUi().alert("Khởi tạo cấu trúc bảng thành công! Hệ thống đã sẵn sàng điều phối.");
}

/**
 * HÀM CHÍNH: Chạy tự động mỗi 15 phút (Cài đặt trong Triggers)
 * Thực hiện trọn vẹn 4 bước AI/Logic + Đồng bộ 8 File NVKD
 */
function runLeadAutomation() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Tự động nhận diện Sheet dữ liệu Leads (hỗ trợ nhiều tên sheet khác nhau)
  const wsLeads = ss.getSheetByName("Leads_Master") 
               || ss.getSheetByName("2_Data_100_Leads_Audit") 
               || ss.getSheetByName("Leads_Clean_100") 
               || ss.getSheetByName("100 Hồ Sơ Chuẩn Hóa")
               || ss.getSheets()[0];
               
  let wsMatrix = ss.getSheetByName("Ma_Tran_8_NVKD");
  let wsBlack = ss.getSheetByName("Blacklist_CSKH");

  if (!wsLeads) {
    Logger.log("Lỗi: Không tìm thấy sheet dữ liệu Leads!");
    return;
  }

  // Nếu chưa có bảng ma trận hoặc blacklist, dùng dữ liệu mặc định trong code
  const leadsData = wsLeads.getDataRange().getValues();
  if (leadsData.length < 2) return;

  // Lấy dữ liệu Blacklist
  const blackMap = {};
  if (wsBlack) {
    const blackData = wsBlack.getDataRange().getValues();
    for (let i = 1; i < blackData.length; i++) {
      const phone = String(blackData[i][0]).replace(/\s+/g, "");
      const email = String(blackData[i][1]).trim().toLowerCase();
      const type = blackData[i][2];
      const owner = blackData[i][3];
      if (phone) blackMap[phone] = { type, owner };
      if (email) blackMap[email] = { type, owner };
    }
  }

  // Lấy dữ liệu 8 NVKD
  const nvList = [];
  if (wsMatrix) {
    const matrixData = wsMatrix.getDataRange().getValues();
    for (let i = 1; i < matrixData.length; i++) {
      nvList.push({
        idAndName: matrixData[i][0],
        pkg: String(matrixData[i][1]),
        provinces: String(matrixData[i][2]),
        tgId: String(matrixData[i][3]),
        email: String(matrixData[i][4]),
        sheetId: String(matrixData[i][5] || "").trim(),
        backlog: 0
      });
    }
  } else {
    // Default 8 NVKD if sheet not yet created
    const defaults = [
      ["NV01 - Nguyễn Văn An", "PROD_RETAIL", "Hà Nội, Hải Phòng, Bắc Ninh", "an.nv@namviet.vn"],
      ["NV02 - Trần Thị Bình", "PROD_RETAIL", "TP. Hồ Chí Minh, Bình Dương, Đồng Nai", "binh.tt@namviet.vn"],
      ["NV03 - Lê Hoàng Cường", "PROD_FNB", "Hà Nội, Đà Nẵng, Quảng Nam", "cuong.lh@namviet.vn"],
      ["NV04 - Phạm Minh Dũng", "PROD_FNB", "TP. Hồ Chí Minh, Bình Dương, Đồng Nai", "dung.pm@namviet.vn"],
      ["NV05 - Vũ Thị Hoa", "PROD_SERVICE", "Toàn quốc", "hoa.vt@namviet.vn"],
      ["NV06 - Đỗ Tuấn Kiệt", "PROD_MFG", "Toàn quốc", "kiet.dt@namviet.vn"],
      ["NV07 - Bùi Thu Lan", "ĐA_NGANH", "Hà Nội, Hải Phòng, Thanh Hóa", "lan.bt@namviet.vn"],
      ["NV08 - Ngô Quang Nam", "ĐA_NGANH", "TP. Hồ Chí Minh, Bình Dương, Đồng Nai", "nam.nq@namviet.vn"]
    ];
    defaults.forEach(d => nvList.push({ idAndName: d[0], pkg: d[1], provinces: d[2], tgId: "", email: d[3], sheetId: "", backlog: 0 }));
  }

  // 1. Đồng bộ kết quả gọi điện từ các file con (nếu có cấu hình)
  if (wsMatrix) syncBackFromSalesSheets(wsLeads, wsMatrix);

  // Đếm backlog hiện tại & chỉ mục trùng lặp
  const existingKeys = {};
  for (let i = 1; i < leadsData.length; i++) {
    const id = leadsData[i][0];
    const mst = String(leadsData[i][3]).trim();
    const phone = String(leadsData[i][4]).replace(/\s+/g, "");
    const status = String(leadsData[i][12]).trim();
    const owner = String(leadsData[i][11]).trim();

    if (status === "ĐÃ PHÂN CÔNG") {
      const nv = nvList.find(n => n.idAndName === owner);
      if (nv) nv.backlog++;
    }
    if (status !== "ĐÃ GỘP HỒ SƠ" && id) {
      if (mst && mst !== "CHƯA XÁC ĐỊNH") existingKeys[mst] = id;
      if (phone && phone !== "CHƯA CÓ") existingKeys[phone] = id;
    }
  }

  const now = new Date();

  // Duyệt từng dòng Lead
  for (let r = 1; r < leadsData.length; r++) {
    const rowIdx = r + 1;
    const leadId = leadsData[r][0];
    const repName = leadsData[r][1];
    const company = leadsData[r][2];
    const mst = String(leadsData[r][3]).trim();
    const phone = String(leadsData[r][4]).replace(/\s+/g, "");
    const email = String(leadsData[r][5]).trim().toLowerCase();
    const city = String(leadsData[r][6]).trim();
    const pkg = String(leadsData[r][8]).trim();
    const budget = Number(leadsData[r][9]) || 0;
    const priority = String(leadsData[r][10]).trim();
    let owner = String(leadsData[r][11]).trim();
    let status = String(leadsData[r][12]).trim();
    let assignedAt = leadsData[r][14];
    let firstCallAt = leadsData[r][15];
    let warnSent = leadsData[r][18];

    // XỬ LÝ LEAD MỚI CHƯA PHÂN LOẠI
    if (status === "MỚI" || status === "" || status === "CHƯA PHÂN BỔ") {
      
      // BƯỚC 1: KIỂM TRA BLACKLIST, CSKH CŨ & TRÙNG LẶP
      const historyHit = blackMap[phone] || blackMap[email];
      if (historyHit) {
        if (historyHit.type === "Blacklist") {
          wsLeads.getRange(rowIdx, 13).setValue("TỪ CHỐI / ĐÓNG LEAD");
          wsLeads.getRange(rowIdx, 14).setValue("Do Not Contact - Khách yêu cầu không làm phiền.");
          continue;
        } else if (historyHit.type === "Khiếu nại CSKH") {
          wsLeads.getRange(rowIdx, 12).setValue("CHỜ TPKD CHỈ ĐỊNH");
          wsLeads.getRange(rowIdx, 13).setValue("CẦN TPKD DUYỆT");
          wsLeads.getRange(rowIdx, 14).setValue("Cảnh báo CSKH: Từng khiếu nại. Trình TPKD.");
          sendAlert(CONFIG.TPKD_TELEGRAM_ID, CONFIG.TPKD_EMAIL, `🔍 CẢNH BÁO CSKH: Lead ${leadId} (${company}) từng khiếu nại, cần TPKD duyệt!`);
          continue;
        } else if (historyHit.type === "Khách cũ" && historyHit.owner) {
          assignLeadToNV(wsLeads, rowIdx, historyHit.owner, now, "Khách cũ Upsell - Giao thẳng NV chăm sóc.");
          continue;
        }
      }

      // Kiểm tra trùng MST hoặc SĐT
      const masterId = existingKeys[mst] || existingKeys[phone];
      if (masterId && masterId !== leadId) {
        wsLeads.getRange(rowIdx, 11).setValue("0 - ĐÃ GỘP");
        wsLeads.getRange(rowIdx, 12).setValue("Theo Master " + masterId);
        wsLeads.getRange(rowIdx, 13).setValue("ĐÃ GỘP HỒ SƠ");
        wsLeads.getRange(rowIdx, 14).setValue(`Gộp vào hồ sơ ${masterId}. Lưu ${repName} làm liên hệ phụ.`);
        continue;
      }

      // BƯỚC 2: LUỒNG VIP >= 500 TRIỆU
      if (budget >= CONFIG.VIP_BUDGET_THRESHOLD) {
        wsLeads.getRange(rowIdx, 11).setValue("TPKD DUYỆT");
        wsLeads.getRange(rowIdx, 12).setValue("CHỜ TPKD CHỈ ĐỊNH");
        wsLeads.getRange(rowIdx, 13).setValue("CẦN TPKD DUYỆT");
        wsLeads.getRange(rowIdx, 14).setValue(`VIP Ngân sách ${(budget/1e6).toFixed(0)}tr (≥500tr). Đẩy TPKD chỉ định KAM.`);
        sendAlert(CONFIG.TPKD_TELEGRAM_ID, CONFIG.TPKD_EMAIL, `🚨 CƠ HỘI VIP ≥500TR: ${leadId} - ${company} (${(budget/1e6).toFixed(0)} triệu). Cần TPKD chỉ định KAM ngay!`);
        continue;
      }

      // Kiểm tra ngoài phạm vi SMB
      if (pkg === "NGOÀI PHẠM VI") {
        wsLeads.getRange(rowIdx, 13).setValue("TỪ CHỐI / ĐÓNG LEAD");
        wsLeads.getRange(rowIdx, 14).setValue("Ngoài gói SMB. Tự động gửi email từ chối lịch sự.");
        continue;
      }

      // Kiểm tra thiếu thông tin
      if (pkg === "CHƯA XÁC ĐỊNH" || company === "CHƯA XÁC ĐỊNH") {
        wsLeads.getRange(rowIdx, 13).setValue("CHỜ BỔ SUNG THÔNG TIN");
        wsLeads.getRange(rowIdx, 14).setValue("Gửi Form khảo sát bổ sung tự động.");
        continue;
      }

      // Tìm NVKD phù hợp
      let targetNV = nvList.find(n => n.pkg.includes(pkg) && (n.provinces.includes("Toàn quốc") || n.provinces.includes(city)));
      if (!targetNV) {
        const multiNVs = nvList.filter(n => n.pkg.includes("ĐA_NGANH"));
        multiNVs.sort((a, b) => a.backlog - b.backlog);
        targetNV = multiNVs[0];
      }

      // BƯỚC 3: KIỂM TRA TẢI TRỌNG TRẦN 30 LEADS (OVERFLOW)
      if (targetNV && targetNV.backlog >= CONFIG.MAX_BACKLOG_CAP) {
        if (priority !== "1 - CAO") {
          const availableNVs = nvList.filter(n => n.backlog < CONFIG.MAX_BACKLOG_CAP).sort((a, b) => a.backlog - b.backlog);
          if (availableNVs.length > 0) {
            targetNV = availableNVs[0];
          } else {
            wsLeads.getRange(rowIdx, 13).setValue("CẦN TPKD DUYỆT");
            wsLeads.getRange(rowIdx, 14).setValue("Toàn bộ team vượt trần 30 leads. Chờ TPKD điều phối.");
            continue;
          }
        }
      }

      if (targetNV) {
        assignLeadToNV(wsLeads, rowIdx, targetNV.idAndName, now, `Tư vấn giải pháp ${pkg} trong 24h.`);
        targetNV.backlog++;
      }
    }

    // BƯỚC 4: GIÁM SÁT ĐỒNG HỒ SLA 4H & 24H
    if (status === "ĐÃ PHÂN CÔNG" && assignedAt instanceof Date && !firstCallAt) {
      const hoursElapsed = (now - assignedAt) / (1000 * 60 * 60);
      const nvInfo = nvList.find(n => n.idAndName === owner);

      if (hoursElapsed >= CONFIG.SLA_RECLAIM_HOURS) {
        wsLeads.getRange(rowIdx, 12).setValue("KHO CHUNG (QUÁ HẠN 24H)");
        wsLeads.getRange(rowIdx, 13).setValue("CẦN TPKD DUYỆT");
        wsLeads.getRange(rowIdx, 18).setValue(owner);
        sendAlert(CONFIG.TPKD_TELEGRAM_ID, CONFIG.TPKD_EMAIL,
          `🚨 CẢNH BÁO ĐỎ SLA 24H: Lead ${leadId} (${company}) giao cho ${owner} đã quá hạn 24h không gọi. Đã THU HỒI về kho chung!`);
      } else if (hoursElapsed >= CONFIG.SLA_WARN_HOURS && warnSent !== "ĐÃ NHẮC 4H") {
        if (nvInfo) {
          sendAlert(nvInfo.tgId, nvInfo.email,
            `⏰ NHẮC NHỞ SLA 4H: Lead ${leadId} (${company} - SĐT: ${phone}) đã giao quá 4 giờ chưa cập nhật gọi. Vui lòng xử lý ngay!`);
        }
        wsLeads.getRange(rowIdx, 19).setValue("ĐÃ NHẮC 4H");
      }
    }
  }

  // Đồng bộ file NVKD nếu có
  if (wsMatrix) syncToSalesSheets(wsLeads, wsMatrix);
}

function assignLeadToNV(ws, rowIdx, nvName, timestamp, actionNote) {
  ws.getRange(rowIdx, 12).setValue(nvName);
  ws.getRange(rowIdx, 13).setValue("ĐÃ PHÂN CÔNG");
  ws.getRange(rowIdx, 14).setValue(actionNote);
  ws.getRange(rowIdx, 15).setValue(timestamp);
}

function syncToSalesSheets(wsLeads, wsMatrix) {
  const leads = wsLeads.getDataRange().getValues();
  const matrix = wsMatrix.getDataRange().getValues();

  for (let i = 1; i < matrix.length; i++) {
    const nvName = matrix[i][0];
    const sheetId = String(matrix[i][5]).trim();
    if (!sheetId) continue;

    try {
      const childSS = SpreadsheetApp.openById(sheetId);
      const childWs = childSS.getSheets()[0];
      const myLeads = leads.slice(1).filter(r => r[11] === nvName && r[12] === "ĐÃ PHÂN CÔNG");
      myLeads.sort((a, b) => String(a[10]).localeCompare(String(b[10])) || (Number(b[9]) - Number(a[9])));

      const output = [["Mã Hồ Sơ", "Ưu Tiên", "Doanh Nghiệp", "Người Đại Diện", "SĐT", "Email", "Tỉnh/TP", "Gói SP", "Ngân Sách", "Hành Động", "Trạng Thái Gọi (Chọn 'Đã gọi')"]];
      myLeads.forEach(r => {
        output.push([r[0], r[10], r[2], r[1], r[4], r[5], r[6], r[8], r[9], r[13], r[15] ? "Đã gọi" : "Chưa gọi"]);
      });

      childWs.clearContents();
      childWs.getRange(1, 1, output.length, output[0].length).setValues(output);
    } catch (e) {
      Logger.log("Lỗi đồng bộ file NV " + nvName + ": " + e.message);
    }
  }
}

function syncBackFromSalesSheets(wsLeads, wsMatrix) {
  const matrix = wsMatrix.getDataRange().getValues();
  const calledIds = new Set();

  for (let i = 1; i < matrix.length; i++) {
    const sheetId = String(matrix[i][5]).trim();
    if (!sheetId) continue;
    try {
      const childWs = SpreadsheetApp.openById(sheetId).getSheets()[0];
      const rows = childWs.getDataRange().getValues();
      for (let j = 1; j < rows.length; j++) {
        if (String(rows[j][10]).trim().toLowerCase() === "đã gọi") {
          calledIds.add(rows[j][0]);
        }
      }
    } catch (e) {}
  }

  const now = new Date();
  const leads = wsLeads.getDataRange().getValues();
  for (let r = 1; r < leads.length; r++) {
    if (calledIds.has(leads[r][0]) && !leads[r][15]) {
      wsLeads.getRange(r + 1, 16).setValue(now);
      wsLeads.getRange(r + 1, 13).setValue("ĐÃ LIÊN HỆ");
    }
  }
}

function sendAlert(chatId, email, message) {
  if (CONFIG.TELEGRAM_BOT_TOKEN && !CONFIG.TELEGRAM_BOT_TOKEN.includes("NHAP_") && chatId) {
    try {
      const url = `https://api.telegram.org/bot${CONFIG.TELEGRAM_BOT_TOKEN}/sendMessage`;
      UrlFetchApp.fetch(url, {
        method: "post",
        contentType: "application/json",
        payload: JSON.stringify({ chat_id: chatId, text: message })
      });
    } catch (e) {
      Logger.log("Telegram error: " + e.message);
    }
  }
  if (email && email.includes("@")) {
    try {
      MailApp.sendEmail(email, "[Nam Việt CRM] Cảnh báo Điều phối Leads & SLA", message);
    } catch (e) {}
  }
}
