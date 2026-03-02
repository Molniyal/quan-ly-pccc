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
// 2. [MỚI] HÀM KIỂM TRA TRẠNG THÁI KHI VỪA QUÉT MÃ
// ==========================================
function doGet(e) {
  try {
    var maTB = e.parameter.maTB;
    if (!maTB) return ContentService.createTextOutput("").setMimeType(ContentService.MimeType.TEXT);

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sLog = ss.getSheetByName("Log_KiemTra");
    var data = sLog.getDataRange().getValues();

    var now = new Date();
    var currentMonth = now.getMonth();
    var currentYear = now.getFullYear();

    // Dò xem mã này đã quét trong tháng này chưa
    for (var i = 1; i < data.length; i++) {
      var thoiGian = new Date(data[i][0]);
      if (data[i][3].toString().trim() === maTB.toString().trim() &&
          thoiGian.getMonth() === currentMonth &&
          thoiGian.getFullYear() === currentYear) {
        // Đã quét rồi -> Báo lỗi
        return ContentService.createTextOutput(JSON.stringify({status: "inspected", message: "Vị trí này ĐÃ ĐƯỢC KIỂM TRA trong tháng này!"})).setMimeType(ContentService.MimeType.JSON);
      }
    }
    // Chưa quét -> Cho phép làm
    return ContentService.createTextOutput(JSON.stringify({status: "available"})).setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({status: "error", message: err.message})).setMimeType(ContentService.MimeType.JSON);
  }
}

// ==========================================
// 3. HÀM LƯU DỮ LIỆU & EMAIL (Giữ nguyên)
// ==========================================
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    var sNV = ss.getSheetByName("DanhSachNV");
    var nvData = sNV.getDataRange().getValues();
    var tenNV = "";
    var isValid = false;
    
    for (var i = 1; i < nvData.length; i++) {
      if (nvData[i][0].toString().trim().toUpperCase() === data.maNV.toString().trim().toUpperCase()) {
        tenNV = nvData[i][1];
        isValid = true;
        break;
      }
    }
    
    if (!isValid) return ContentService.createTextOutput(JSON.stringify({status: "error", message: "Sai Mã Nhân Viên!"})).setMimeType(ContentService.MimeType.JSON);
    
    var sLog = ss.getSheetByName("Log_KiemTra");
    var formula = '=IMAGE(INDIRECT("I" & ROW()))'; 
    sLog.appendRow([new Date(), data.maNV.toUpperCase(), tenNV, data.maTB, data.ngoaiQuan, data.apSuat, data.hanSD, data.ghiChu, data.linkAnh, formula]);
    
    return ContentService.createTextOutput(JSON.stringify({status: "success", name: tenNV})).setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({status: "error", message: err.message})).setMimeType(ContentService.MimeType.JSON);
  }
}

function doOptions(e) { return ContentService.createTextOutput("").setMimeType(ContentService.MimeType.TEXT); }