// ==========================================
// 1. HÀM TẠO SHEET MẪU (Giữ nguyên)
// ==========================================
function setupPCCC() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss.getSheetByName("DanhSachNV")) {
    var sNV = ss.insertSheet("DanhSachNV");
    sNV.appendRow(["Mã NV", "Tên Nhân Viên"]);
    sNV.appendRow(["NV001", "Nguyễn Văn A"]); 
  }
  if (!ss.getSheetByName("Master_ViTri")) {
    var sMaster = ss.insertSheet("Master_ViTri");
    sMaster.appendRow(["Mã Thiết Bị", "Khu Vực", "Tên Quản Lý", "Email Quản Lý", "Link Gốc", "Link Web QR", "Mã QR"]);
  }
  if (!ss.getSheetByName("Log_KiemTra")) {
    var sLog = ss.insertSheet("Log_KiemTra");
    sLog.appendRow(["Thời gian", "Mã NV", "Tên NV", "Mã Thiết Bị", "Ngoại quan", "Áp suất", "Hạn SD", "Ghi chú", "Link Ảnh", "Ảnh thực tế"]);
  }
}

// ==========================================
// 2. HÀM XỬ LÝ GET REQUEST (MỞ RỘNG)
// ==========================================
function doGet(e) {
  try {
    var action = e.parameter.action;
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // --- ACTION A: LẤY THÔNG TIN NHÂN VIÊN (BỔ SUNG VAI TRÒ) ---
    if (action === "getEmployee") {
      var maNV = e.parameter.maNV;
      if (!maNV) return responseJSON({status: "error"});
      
      var sNV = ss.getSheetByName("DanhSachNV");
      var data = sNV.getDataRange().getValues();
      for (var i = 1; i < data.length; i++) {
        if (data[i][0].toString().trim().toUpperCase() === maNV.toUpperCase()) {
          return responseJSON({
            status: "success", 
            name: data[i][1], 
            role: data[i][2] || "User" // Lấy vai trò ở cột C
          });
        }
      }
      return responseJSON({status: "error"});
    }

    // --- ACTION E: XÁC THỰC QUYỀN ADMIN ---
    if (action === "verifyAdmin") {
      var maNV = e.parameter.maNV;
      var sNV = ss.getSheetByName("DanhSachNV");
      var data = sNV.getDataRange().getValues();
      for (var i = 1; i < data.length; i++) {
        var rowMa = data[i][0].toString().trim().toUpperCase();
        var rowRole = data[i][2] ? data[i][2].toString().trim().toUpperCase() : "USER";
        
        if (rowMa === maNV.toUpperCase() && rowRole === "ADMIN") {
          return responseJSON({status: "success"});
        }
      }
      return responseJSON({status: "error", message: "Bạn không có quyền quản trị!"});
    }

    // --- ACTION B: LẤY DỮ LIỆU DASHBOARD CHO ADMIN ---
    if (action === "getDashboardData") {
      var sLog = ss.getSheetByName("Log_KiemTra");
      var data = sLog.getDataRange().getValues();
      var now = new Date();
      
      var logs = [];
      var issues = 0;
      var monthCount = 0;

      for (var i = data.length - 1; i >= 1; i--) {
        var row = data[i];
        var time = new Date(row[0]);
        var isMonth = time.getMonth() === now.getMonth() && time.getFullYear() === now.getFullYear();
        
        if (isMonth) monthCount++;
        
        var isIssue = row[4] === "Lỗi" || row[5] === "Lỗi" || row[6] === "Hết hạn";
        if (isIssue) issues++;

        if (logs.length < 50) { 
          logs.push({
            time: row[0],
            nv: row[2],
            tb: row[3],
            nq: row[4],
            as: row[5],
            hsd: row[6],
            note: row[7],
            img: row[8]
          });
        }
      }

      return responseJSON({
        total: data.length - 1,
        issues: issues,
        monthCount: monthCount,
        logs: logs
      });
    }

    // --- ACTION D: GỬI EMAIL CẢNH BÁO ---
    if (action === "sendAlertEmail") {
      var maTB = e.parameter.maTB;
      var tenNV = e.parameter.tenNV;
      var ghiChu = e.parameter.ghiChu;
      
      var sMaster = ss.getSheetByName("Master_ViTri");
      var masterData = sMaster.getDataRange().getValues();
      var managerEmail = "";
      
      for (var j = 1; j < masterData.length; j++) {
        if (masterData[j][0].toString().trim().toUpperCase() === maTB.toUpperCase()) {
          managerEmail = masterData[j][3]; // Cột D: Email Quản Lý
          break;
        }
      }
      
      if (!managerEmail) return responseJSON({status: "error", message: "Không tìm thấy email quản lý thiết bị này!"});
      
      var subject = "⚠️ CẢNH BÁO PCCC: Thiết bị " + maTB + " phát hiện LỖI";
      var body = "Kính gửi Quản lý,\n\n" +
                 "Hệ thống PCCC Digital phát hiện một báo cáo có LỖI tại vị trí: " + maTB + "\n" +
                 "- Nhân viên báo cáo: " + tenNV + "\n" +
                 "- Nội dung lỗi/Ghi chú: " + (ghiChu || "Không có ghi chú") + "\n\n" +
                 "Vui lòng kiểm tra và xử lý ngay để đảm bảo an toàn.\n" +
                 "Trân trọng,\nBệ thống PCCC Digital.";
      
      MailApp.sendEmail(managerEmail, subject, body);
      return responseJSON({status: "success"});
    }

    // --- ACTION C [MẶC ĐỊNH]: KIỂM TRA TRÙNG LẶP ---
    var maTB = e.parameter.maTB;
    if (!maTB) return ContentService.createTextOutput("").setMimeType(ContentService.MimeType.TEXT);

    var sLog = ss.getSheetByName("Log_KiemTra");
    var data = sLog.getDataRange().getValues();
    var now = new Date();

    for (var i = 1; i < data.length; i++) {
      var thoiGian = new Date(data[i][0]);
      if (data[i][3].toString().trim().toUpperCase() === maTB.toString().trim().toUpperCase() &&
          thoiGian.getMonth() === now.getMonth() &&
          thoiGian.getFullYear() === now.getFullYear()) {
        return responseJSON({status: "inspected", message: "Vị trí này ĐÃ ĐƯỢC KIỂM TRA trong tháng này!"});
      }
    }
    return responseJSON({status: "available"});

  } catch(err) {
    return responseJSON({status: "error", message: err.message});
  }
}

// Hàm bổ trợ trả về JSON
function responseJSON(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ==========================================
// 3. HÀM LƯU DỮ LIỆU (GIỮ NGUYÊN POS)
// ==========================================
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    var sNV = ss.getSheetByName("DanhSachNV");
    var nvData = sNV.getDataRange().getValues();
    var tenNV = "Unknown";
    
    for (var i = 1; i < nvData.length; i++) {
        if (nvData[i][0].toString().trim().toUpperCase() === data.maNV.toString().trim().toUpperCase()) {
            tenNV = nvData[i][1];
            break;
        }
    }
    
    var sLog = ss.getSheetByName("Log_KiemTra");
    var formula = '=IMAGE(INDIRECT("I" & ROW()))'; 
    sLog.appendRow([new Date(), data.maNV.toUpperCase(), tenNV, data.maTB, data.ngoaiQuan, data.apSuat, data.hanSD, data.ghiChu, data.linkAnh, formula]);
    
    return responseJSON({status: "success", name: tenNV});
  } catch(err) {
    return responseJSON({status: "error", message: err.message});
  }
}

function doOptions(e) { return ContentService.createTextOutput("").setMimeType(ContentService.MimeType.TEXT); }