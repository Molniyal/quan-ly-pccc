function setupPCCC() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss.getSheetByName("DanhSachNV")) {
    var sNV = ss.insertSheet("DanhSachNV");
    sNV.appendRow(["Mã NV", "Tên Nhân Viên", "Quyền"]);
    sNV.appendRow(["ADMIN", "Quản Trị Viên", "ADMIN"]);
  }
  // CẬP NHẬT SHEET MASTER
  if (!ss.getSheetByName("Master_ViTri")) {
    var sMaster = ss.insertSheet("Master_ViTri");
    sMaster.appendRow(["Mã Thiết Bị", "Xưởng", "Khu Vực", "Tên Quản Lý", "Email Quản Lý", "Link Gốc", "Link Web QR", "Mã QR"]);
  }
  if (!ss.getSheetByName("Log_KiemTra")) {
    var sLog = ss.insertSheet("Log_KiemTra");
    sLog.appendRow(["Thời gian", "Mã NV", "Tên NV", "Mã Thiết Bị", "Ngoại quan", "Áp suất", "Hạn SD", "Ghi chú", "Link Ảnh", "Ảnh thực tế"]);
  }
}

function doGet(e) {
  try {
    var action = e.parameter.action;
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    if (action === "getEmployee") {
      var maNV = e.parameter.maNV;
      if (!maNV) return responseJSON({status: "error"});
      var data = ss.getSheetByName("DanhSachNV").getDataRange().getValues();
      for (var i = 1; i < data.length; i++) {
        if (data[i][0].toString().trim().toUpperCase() === maNV.toUpperCase()) {
          return responseJSON({ status: "success", name: data[i][1], role: data[i][2] || "USER" });
        }
      }
      return responseJSON({status: "error"});
    }

    if (action === "verifyAdmin") {
      var maNV = e.parameter.maNV;
      var data = ss.getSheetByName("DanhSachNV").getDataRange().getValues();
      for (var i = 1; i < data.length; i++) {
        if (data[i][0].toString().trim().toUpperCase() === maNV.toUpperCase() && (data[i][2] || "").toString().trim().toUpperCase() === "ADMIN") {
          return responseJSON({status: "success"});
        }
      }
      return responseJSON({status: "error", message: "Bạn không có quyền quản trị!"});
    }

    // TÍNH TOÁN DASHBOARD MỚI
    if (action === "getDashboardData") {
      var sMaster = ss.getSheetByName("Master_ViTri");
      var sLog = ss.getSheetByName("Log_KiemTra");
      
      var masterData = sMaster.getDataRange().getValues();
      var logData = sLog.getDataRange().getValues();
      var now = new Date();
      var currentMonth = now.getMonth();
      var currentYear = now.getFullYear();

      // 1. Quét các thiết bị đã kiểm tra trong tháng này
      var daKiemTraThisMonth = new Set();
      var logs = [];
      var issues = 0;

      for (var i = logData.length - 1; i >= 1; i--) {
        var row = logData[i];
        var time = new Date(row[0]);
        var isThisMonth = time.getMonth() === currentMonth && time.getFullYear() === currentYear;
        var maTB = row[3].toString().trim().toUpperCase();
        var isIssue = row[4] === "Lỗi" || row[5] === "Lỗi" || row[6] === "Hết hạn";

        if (isThisMonth) {
          daKiemTraThisMonth.add(maTB);
          if (isIssue) issues++;
        }

        // Lấy 50 log gần nhất cho bảng
        if (logs.length < 50) { 
          logs.push({ time: row[0], nv: row[2], tb: row[3], nq: row[4], as: row[5], hsd: row[6], note: row[7], img: row[8] });
        }
      }

      // 2. Thống kê theo Xưởng từ Master Data
      var xuongStats = {};
      var totalEquipments = 0;

      for (var j = 1; j < masterData.length; j++) {
        var tbId = masterData[j][0].toString().trim().toUpperCase();
        var xuong = masterData[j][1].toString().trim(); // Cột B: Xưởng
        if (!tbId || !xuong) continue;
        
        totalEquipments++;
        
        if (!xuongStats[xuong]) {
          xuongStats[xuong] = { total: 0, inspected: 0 };
        }
        
        xuongStats[xuong].total++;
        if (daKiemTraThisMonth.has(tbId)) {
          xuongStats[xuong].inspected++;
        }
      }

      return responseJSON({
        total: totalEquipments,
        issues: issues,
        monthCount: daKiemTraThisMonth.size, // Số bình duy nhất đã kiểm trong tháng
        xuongStats: xuongStats,
        logs: logs
      });
    }

    if (action === "sendAlertEmail") {
      var maTB = e.parameter.maTB;
      var tenNV = e.parameter.tenNV;
      var ghiChu = e.parameter.ghiChu;
      var masterData = ss.getSheetByName("Master_ViTri").getDataRange().getValues();
      var managerEmail = "";
      
      for (var j = 1; j < masterData.length; j++) {
        if (masterData[j][0].toString().trim().toUpperCase() === maTB.toUpperCase()) {
          managerEmail = masterData[j][4]; // Sửa thành cột E (Index 4)
          break;
        }
      }
      
      if (!managerEmail) return responseJSON({status: "error", message: "Không tìm thấy email quản lý!"});
      MailApp.sendEmail(managerEmail, "⚠️ CẢNH BÁO PCCC: Thiết bị " + maTB + " LỖI", "Kính gửi Quản lý,\nPhát hiện LỖI tại: " + maTB + "\n- Nhân viên: " + tenNV + "\n- Ghi chú: " + ghiChu);
      return responseJSON({status: "success"});
    }

    // CHECK TRÙNG LẶP
    var maTB = e.parameter.maTB;
    if (!maTB) return ContentService.createTextOutput("");
    var data = ss.getSheetByName("Log_KiemTra").getDataRange().getValues();
    var now = new Date();
    for (var i = 1; i < data.length; i++) {
      var time = new Date(data[i][0]);
      if (data[i][3].toString().trim().toUpperCase() === maTB.toString().trim().toUpperCase() && time.getMonth() === now.getMonth() && time.getFullYear() === now.getFullYear()) {
        return responseJSON({status: "inspected", message: "ĐÃ ĐƯỢC KIỂM TRA trong tháng này!"});
      }
    }
    return responseJSON({status: "available"});

  } catch(err) { return responseJSON({status: "error", message: err.message}); }
}

function responseJSON(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var nvData = ss.getSheetByName("DanhSachNV").getDataRange().getValues();
    var tenNV = "Unknown";
    for (var i = 1; i < nvData.length; i++) {
        if (nvData[i][0].toString().trim().toUpperCase() === data.maNV.toString().trim().toUpperCase()) {
            tenNV = nvData[i][1]; break;
        }
    }
    var formula = '=IMAGE(INDIRECT("I" & ROW()))'; 
    ss.getSheetByName("Log_KiemTra").appendRow([new Date(), data.maNV.toUpperCase(), tenNV, data.maTB, data.ngoaiQuan, data.apSuat, data.hanSD, data.ghiChu, data.linkAnh, formula]);
    return responseJSON({status: "success", name: tenNV});
  } catch(err) { return responseJSON({status: "error", message: err.message}); }
}
function doOptions(e) { return ContentService.createTextOutput("").setMimeType(ContentService.MimeType.TEXT); }