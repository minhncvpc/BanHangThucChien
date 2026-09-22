/**
 * Code.gs - Server-side logic for Data Entry App
 * [CHANGE_LOG] 2026-09-22 10:25:00 | Editor: AI - Antigravity (Gemini 3.8 Flash) | Mục đích: Cập nhật giá cước cho phép không bắt buộc nhập hoặc nhập số >= 0 (mặc định 0 nếu để trống)
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
 * Safely parse a date from Google Sheet value (handles Date objects, ISO strings, and dd/MM/yyyy HH:mm:ss)
 * [CHANGE_LOG] 2026-09-15 10:15:00 | Editor: AI - Antigravity (Gemini 3.8 Flash) | Mục đích: Chuẩn hóa parseDate an toàn mọi định dạng (Date object, dd/MM/yyyy, ISO) và formatDayKey theo múi giờ script
 */
function parseDate(val) {
  if (!val) return null;
  if (val instanceof Date) {
    return isNaN(val.getTime()) ? null : val;
  }
  if (typeof val === 'string') {
    var str = val.trim();
    if (!str) return null;
    // Hỗ trợ dd/MM/yyyy hoặc dd/MM/yyyy HH:mm:ss
    var dmyMatch = str.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s+(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?)?/);
    if (dmyMatch) {
      var d = parseInt(dmyMatch[1], 10);
      var m = parseInt(dmyMatch[2], 10) - 1;
      var y = parseInt(dmyMatch[3], 10);
      var hh = dmyMatch[4] ? parseInt(dmyMatch[4], 10) : 0;
      var mm = dmyMatch[5] ? parseInt(dmyMatch[5], 10) : 0;
      var ss = dmyMatch[6] ? parseInt(dmyMatch[6], 10) : 0;
      var parsed = new Date(y, m, d, hh, mm, ss);
      if (!isNaN(parsed.getTime())) return parsed;
    }
    // Hỗ trợ định dạng ISO yyyy-MM-dd
    var parsedIso = new Date(str);
    if (!isNaN(parsedIso.getTime())) return parsedIso;
  }
  return null;
}

/**
 * Format a Date object to yyyy-MM-dd in the local script timezone (Asia/Bangkok)
 */
function formatDayKey(dateObj, timeZone) {
  if (!dateObj || isNaN(dateObj.getTime())) return '';
  var tz = timeZone || (typeof Session !== 'undefined' && Session.getScriptTimeZone ? Session.getScriptTimeZone() : 'Asia/Bangkok');
  if (typeof Utilities !== 'undefined' && Utilities.formatDate) {
    try {
      return Utilities.formatDate(dateObj, tz, 'yyyy-MM-dd');
    } catch (e) {}
  }
  var y = dateObj.getFullYear();
  var m = ('0' + (dateObj.getMonth() + 1)).slice(-2);
  var d = ('0' + dateObj.getDate()).slice(-2);
  return y + '-' + m + '-' + d;
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
  // [CHANGE_LOG] 2026-09-15 13:40:00 | Editor: AI - Antigravity (Gemini 3.8 Flash) | Mục đích: Chuẩn hóa số di động bỏ tiền tố 84; chỉ cho phép 9 số đầu 8xxx, 9xxx hoặc 10 số đầu 128xxx, 138xxx; tự động làm sạch khoảng trắng và tiền tố thừa
  var unit = formObject.unitName || '';
  var empCode = formObject.empCode || '';
  var serviceType = formObject.serviceType || '';
  var rawPhone = '';
  if (serviceType === 'DDTT' || serviceType === 'DDTS') {
    rawPhone = formObject.phoneNumber ? String(formObject.phoneNumber).trim().replace(/[\s.\-_()]/g, '') : '';
    if (rawPhone.startsWith('+84')) {
      rawPhone = rawPhone.substring(3);
    }
    if (rawPhone.startsWith('84') && rawPhone.length >= 11) {
      rawPhone = rawPhone.substring(2);
    }
    if (rawPhone.startsWith('0')) {
      rawPhone = rawPhone.replace(/^0+/, '');
    }

    // Kiểm tra định dạng: chỉ bao gồm 9 số với đầu 8xxx, 9xxx hoặc 10 số với đầu 128xxx, 138xxx (không nhập 84 ở đầu)
    var mobileRegex = /^(?:[89][0-9]{8}|(?:128|138)[0-9]{7})$/;
    if (!mobileRegex.test(rawPhone)) {
      return {
        success: false,
        message: "Số thuê bao di động không hợp lệ (" + rawPhone + "). Quy định chỉ bao gồm 9 số đầu 8xxx, 9xxx hoặc 10 số đầu 128xxx, 138xxx (không nhập mã 84 ở đầu)."
      };
    }
  } else {
    // Fiber Khôi phục hoặc dịch vụ khác: Cho phép nhập text tự do, bảo lưu dấu _, -, ký tự đặc biệt
    rawPhone = formObject.phoneNumber ? String(formObject.phoneNumber).trim() : '';
  }
  var phone = "'" + rawPhone; // Force string for phone numbers and account codes
  // Chuẩn hóa giá cước: cho phép không nhập (mặc định 0) hoặc nhập số >= 0
  var rawPrice = formObject.price;
  var price = 0;
  if (rawPrice !== undefined && rawPrice !== null && String(rawPrice).trim() !== '') {
    var parsedPrice = Number(rawPrice);
    if (isNaN(parsedPrice) || parsedPrice < 0) {
      return {
        success: false,
        message: "Giá cước không hợp lệ (" + rawPrice + "). Giá cước phải là kiểu số lớn hơn hoặc bằng 0."
      };
    }
    price = parsedPrice;
  }
  var note = formObject.note || '';
  var empName = formObject.empName || '';

  // DUPLICATE CHECK - Chỉ kiểm tra trùng lặp trong CÙNG 1 NGÀY BÁN HÀNG
  var data = dataSheet.getDataRange().getValues();
  var scriptTz = (typeof Session !== 'undefined' && Session.getScriptTimeZone) ? Session.getScriptTimeZone() : 'Asia/Bangkok';
  var todayKey = formatDayKey(timestamp, scriptTz);
  var isMobile = (serviceType === 'DDTT' || serviceType === 'DDTS');

  for (var i = 1; i < data.length; i++) { // Skip header
    var rowDate = parseDate(data[i][0]);
    if (!rowDate) continue;

    // 1. Kiểm tra ngày bán hàng: Nếu KHÔNG cùng ngày bán hàng với thời điểm hiện tại -> Bỏ qua (cho phép nhập)
    var rowDayKey = formatDayKey(rowDate, scriptTz);
    if (rowDayKey !== todayKey) {
      continue;
    }

    // 2. Cùng ngày bán hàng -> Kiểm tra xem có trùng số thuê bao/mã tài khoản hay không
    var rowPhoneRaw = String(data[i][4] || '').trim();
    if (!rowPhoneRaw) continue;

    var rowService = String(data[i][5] || '').trim();
    var isRowMobile = (rowService === 'DDTT' || rowService === 'DDTS');
    var isDuplicate = false;

    if (isMobile && isRowMobile) {
      var rowPhone = rowPhoneRaw.replace(/[\s.\-_()]/g, '');
      var normRowPhone = rowPhone;
      if (normRowPhone.startsWith('+84')) {
        normRowPhone = normRowPhone.substring(3);
      } else if (normRowPhone.startsWith('84') && normRowPhone.length >= 11) {
        normRowPhone = normRowPhone.substring(2);
      }
      if (normRowPhone.startsWith('0')) {
        normRowPhone = normRowPhone.replace(/^0+/, '');
      }

      if (rowPhone === rawPhone || normRowPhone === rawPhone) {
        isDuplicate = true;
      }
    } else if (!isMobile && !isRowMobile) {
      // Fiber Khôi phục: so sánh chính xác không phân biệt hoa thường, bảo lưu nguyên vẹn ký tự gạch dưới _ và ký tự đặc biệt
      if (rowPhoneRaw.toLowerCase() === rawPhone.toLowerCase()) {
        isDuplicate = true;
      }
    } else if (rowService === serviceType) {
      if (rowPhoneRaw.toLowerCase() === rawPhone.toLowerCase()) {
        isDuplicate = true;
      }
    }

    if (isDuplicate) {
      var dupTime = (typeof Utilities !== 'undefined' && Utilities.formatDate) 
        ? Utilities.formatDate(rowDate, scriptTz, "HH:mm:ss") 
        : "trước đó";
      var dupDateStr = (typeof Utilities !== 'undefined' && Utilities.formatDate)
        ? Utilities.formatDate(rowDate, scriptTz, "dd/MM/yyyy")
        : "";
      var dupEmail = data[i][1] || "người dùng khác";
      var dupEmp = data[i][3] || "";

      return {
        success: false,
        message: (isMobile ? "Số thuê bao " : "Mã/Số thuê bao ") + rawPhone + " đã được nhập trong ngày hôm nay" + (dupDateStr ? (" (" + dupDateStr + ")") : "") + " lúc " + dupTime + " bởi " + dupEmail + (dupEmp ? (" cho nhân viên " + dupEmp) : "")
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

    var rowDate = parseDate(row[0]);
    if (!rowDate) continue; // Skip corrupted date rows

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
