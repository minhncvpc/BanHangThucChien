/**
 * Code.gs - Server-side logic for Data Entry App
 * [CHANGE_LOG] 2026-09-11 17:55:00 | Editor: AI - Antigravity (Gemini 3.8 Flash) | Mục đích: Tối ưu hoá toàn diện cơ chế kết nối Sheet (ưu tiên ActiveSpreadsheet), tự động tạo Config/Data nếu thiếu, chống crash ngày tháng và xử lý khoảng trắng tên sheet
 */

// ID dự phòng nếu triển khai dưới dạng Standalone Script (không gắn trực tiếp vào Google Sheet)
var SPREADSHEET_ID = '1w21WXrW4geKQCegd_VTlmZp-1-_ks_n2q4a9Ge-9IIk';

/**
 * Serves the web app HTML
 */
function doGet(e) {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Ứng Dụng Nhập Liệu Bán Hàng')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Include helper for HTML templates (if needed for splitting CSS/JS later)
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Helper to get the spreadsheet
 * Ưu tiên: 
 * 1. Sheet hiện tại (Container-bound Script - khi mở từ Tiện ích mở rộng > Apps Script)
 * 2. Mở theo SPREADSHEET_ID (nếu là Standalone Script)
 */
function getSpreadsheet() {
  try {
    var activeSs = SpreadsheetApp.getActiveSpreadsheet();
    if (activeSs) {
      return activeSs;
    }
  } catch (err) {
    console.warn("Không thể lấy ActiveSpreadsheet: " + err.message);
  }

  if (typeof SPREADSHEET_ID !== 'undefined' && SPREADSHEET_ID && SPREADSHEET_ID.trim() !== '') {
    try {
      return SpreadsheetApp.openById(SPREADSHEET_ID.trim());
    } catch (e) {
      console.error("Lỗi khi mở theo SPREADSHEET_ID (" + SPREADSHEET_ID + "): " + e.message);
      throw new Error("Không thể mở Google Sheet theo ID: " + SPREADSHEET_ID + ". Lỗi: " + e.message + ". Vui lòng kiểm tra quyền truy cập bảng tính.");
    }
  }

  throw new Error("Không xác định được Google Sheet! Nếu gắn script vào Google Sheet, vui lòng mở trực tiếp từ Tiện ích mở rộng > Apps Script của Sheet đó.");
}

/**
 * Tìm sheet theo tên (không phân biệt hoa thường và tự động trim khoảng trắng)
 */
function findSheetByName(ss, targetName) {
  if (!ss) return null;
  var exact = ss.getSheetByName(targetName);
  if (exact) return exact;

  var sheets = ss.getSheets();
  var normTarget = targetName.trim().toLowerCase();
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getName().trim().toLowerCase() === normTarget) {
      return sheets[i];
    }
  }
  return null;
}

/**
 * Lấy hoặc tự động tạo Sheet Data
 */
function getDataSheet(ss) {
  var dataSheet = findSheetByName(ss, 'Data');
  if (!dataSheet) {
    dataSheet = ss.insertSheet('Data');
    dataSheet.appendRow(['Thời gian nhập', 'Email người nhập', 'Tên đơn vị', 'Mã nhân viên', 'Số thuê bao', 'Dịch vụ', 'Giá cước', 'Ghi chú', 'Tên nhân viên']);
  }
  return dataSheet;
}

/**
 * Get configuration data from the 'Config' sheet
 * Returns list of unique units and list of all employees with their metadata
 */
function getConfigData() {
  var ss = getSpreadsheet();
  var configSheet = findSheetByName(ss, 'Config');

  if (!configSheet) {
    // Tự động khởi tạo Sheet Config với dữ liệu mẫu nếu chưa có
    configSheet = ss.insertSheet('Config');
    configSheet.appendRow(['Đơn vị', 'Mã nhân viên', 'Tên nhân viên']);
    configSheet.appendRow(['VNPT Cẩm Khê', 'CKH001', 'Nguyễn Văn A']);
    configSheet.appendRow(['VNPT Việt Trì', 'VTR001', 'Trần Thị B']);
  }

  var lastRow = configSheet.getLastRow();
  if (lastRow < 2) {
    return { units: [], employees: [] };
  }

  var data = configSheet.getDataRange().getValues();
  var units = new Set();
  var employees = [];

  var currentUnit = '';
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var unit = String(row[0] || '').trim();
    var code = String(row[1] || '').trim();
    var name = row[2] ? String(row[2]).trim() : '';

    if (unit) {
      currentUnit = unit;
      units.add(currentUnit);
    }

    if (code) {
      var assignedUnit = currentUnit || 'Đơn vị khác';
      units.add(assignedUnit);
      employees.push({
        unit: assignedUnit,
        code: code,
        name: name
      });
    }
  }

  return {
    units: Array.from(units),
    employees: employees
  };
}

/**
 * Process the form submission
 * @param {Object} formObject - The form data object
 */
function processForm(formObject) {
  var ss = getSpreadsheet();
  var dataSheet = getDataSheet(ss);

  var timestamp = new Date();
  var userEmail = '';
  try {
    userEmail = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail() || '';
  } catch (e) {
    userEmail = 'N/A';
  }

  // Normalize data
  var unit = formObject.unitName || '';
  var empCode = formObject.empCode || '';
  var rawPhone = formObject.phoneNumber ? String(formObject.phoneNumber).trim() : '';
  var phone = "'" + rawPhone; // Force string for phone numbers
  var serviceType = formObject.serviceType || '';
  var price = formObject.price || 0;
  var note = formObject.note || '';
  var empName = formObject.empName || '';

  // DUPLICATE CHECK
  var data = dataSheet.getDataRange().getValues();
  var normRawPhone = rawPhone.replace(/^(84|0)/, '');

  for (var i = 1; i < data.length; i++) { // Skip header
    var rowDate = data[i][0] ? new Date(data[i][0]) : null;
    var rowPhone = String(data[i][4] || '').trim();
    if (!rowPhone) continue;

    var normRowPhone = rowPhone.replace(/^(84|0)/, '');

    // Compare exact or normalized phone
    if (rowPhone === rawPhone || (normRawPhone && normRowPhone.length >= 9 && normRowPhone === normRawPhone)) {
      var dupTime = (rowDate && !isNaN(rowDate.getTime())) 
        ? Utilities.formatDate(rowDate, Session.getScriptTimeZone(), "HH:mm:ss dd/MM/yyyy") 
        : "trước đó";
      var dupEmail = data[i][1] || "người dùng khác";
      var dupEmp = data[i][3] || "";

      return {
        success: false,
        message: "Số thuê bao " + rawPhone + " đã được nhập lúc " + dupTime + " bởi " + dupEmail + (dupEmp ? (" cho nhân viên " + dupEmp) : "")
      };
    }
  }

  dataSheet.appendRow([
    timestamp,
    userEmail,
    unit,
    empCode,
    phone,
    serviceType,
    price,
    note,
    empName
  ]);

  return { success: true, message: "Đã nhập dữ liệu thành công!" };
}

/**
 * Get report data based on filters
 */
function getReportData(startDateStr, endDateStr, unitName, empCode) {
  var ss = getSpreadsheet();
  var dataSheet = findSheetByName(ss, 'Data');

  if (!dataSheet) return [];

  var data = dataSheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  var results = [];
  var targetStartDate = startDateStr ? new Date(startDateStr) : null;
  if (targetStartDate && !isNaN(targetStartDate.getTime())) {
    targetStartDate.setHours(0, 0, 0, 0);
  } else {
    targetStartDate = null;
  }

  var targetEndDate = endDateStr ? new Date(endDateStr) : null;
  if (targetEndDate && !isNaN(targetEndDate.getTime())) {
    targetEndDate.setHours(23, 59, 59, 999);
  } else {
    targetEndDate = null;
  }

  for (var i = 1; i < data.length; i++) { // Skip header
    var row = data[i];
    if (!row || !row[0]) continue;

    var rowDate = new Date(row[0]);
    if (isNaN(rowDate.getTime())) continue; // Skip corrupted date rows

    var rowUnit = String(row[2] || '').trim();
    var rowEmpCode = String(row[3] || '').trim();

    // Check Date Range
    if (targetStartDate && rowDate.getTime() < targetStartDate.getTime()) {
      continue;
    }
    if (targetEndDate && rowDate.getTime() > targetEndDate.getTime()) {
      continue;
    }

    // Check Unit
    if (unitName && rowUnit !== unitName.trim()) {
      continue;
    }

    // Check EmpCode
    if (empCode && rowEmpCode !== empCode.trim()) {
      continue;
    }

    var formattedDate = "";
    try {
      formattedDate = Utilities.formatDate(rowDate, Session.getScriptTimeZone(), "HH:mm dd/MM/yyyy");
    } catch (e) {
      formattedDate = String(row[0]);
    }

    results.push({
      time: formattedDate,
      email: row[1] || '',
      unit: row[2] || '',
      empCode: row[3] || '',
      phone: row[4] || '',
      service: row[5] || '',
      price: row[6] || 0,
      note: row[7] || '',
      empName: row[8] || ''
    });
  }

  return results;
}
