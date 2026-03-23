/**
 * Code.gs - Server-side logic for Data Entry App
 */

/**
 * Serves the web app HTML
 */
function doGet(e) {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Ứng Dụng Nhập Liệu')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Include helper for HTML templates (if needed for splitting CSS/JS later)
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename)
    .getContent();
}

// ID của Google Sheet chứa dữ liệu (được cung cấp bởi người dùng)
var SPREADSHEET_ID = '1w21WXrW4geKQCegd_VTlmZp-1-_ks_n2q4a9Ge-9IIk';

/**
 * Helper to get the spreadsheet
 */
function getSpreadsheet() {
  try {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
  } catch (e) {
    // Fallback if ID is invalid or script is bound to the sheet itself
    return SpreadsheetApp.getActiveSpreadsheet();
  }
}

/**
 * Get configuration data from the 'Config' sheet
 * Returns list of unique units and list of all employees with their metadata
 */
function getConfigData() {
  var ss = getSpreadsheet();
  var configSheet = ss.getSheetByName('Config');

  if (!configSheet) {
    throw new Error("Không tìm thấy sheet 'Config'. Vui lòng tạo sheet 'Config' trước.");
  }

  var lastRow = configSheet.getLastRow();
  if (lastRow < 2) {
    return { units: [], employees: [] };
  }

  // New Config Structure:
  // Col A: Unit (Đơn vị)
  // Col B: Code (Mã NV)
  // Col C: Name (Tên NV)
  // Get data from A2:C
  var data = configSheet.getRange(2, 1, lastRow - 1, 3).getValues();

  var units = new Set();
  var employees = [];

  for (var i = 0; i < data.length; i++) {
    var unit = data[i][0]; // Column A
    var code = data[i][1]; // Column B
    var name = data[i][2]; // Column C

    if (unit) {
      units.add(unit);

      if (code) {
        employees.push({
          unit: unit,
          code: code,
          name: name || '' // Handle empty name
        });
      }
    }
  }

  return {
    units: [...units],
    employees: employees
  };
}

/**
 * Process the form submission
 * @param {Object} formObject - The form data object
 */
function processForm(formObject) {
  var ss = getSpreadsheet();
  var dataSheet = ss.getSheetByName('Data');

  if (!dataSheet) {
    // Attempt to create Data sheet if it doesn't exist
    dataSheet = ss.insertSheet('Data');
    dataSheet.appendRow(['Thời gian nhập', 'Email người nhập', 'Tên đơn vị', 'Mã nhân viên', 'Số thuê bao', 'Dịch vụ', 'Giá cước', 'Ghi chú', 'Tên nhân viên']);
  }

  var timestamp = new Date();
  var userEmail = Session.getActiveUser().getEmail();

  // Normalize data
  var unit = formObject.unitName;
  var empCode = formObject.empCode;
  var phone = "'" + formObject.phoneNumber; // Force string for phone numbers
  var serviceType = formObject.serviceType;
  var price = formObject.price;
  var note = formObject.note;
  var empName = formObject.empName || '';

  // DUPLICATE CHECK
  // Check if same phone number already exists in the system
  var data = dataSheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) { // Skip header
    var rowDate = new Date(data[i][0]);
    var rowPhone = String(data[i][4]); // Column 5 is Phone (0-based 4)
    // Note: getValues() returns value. If stored as '123, value is 123.
    // formObject.phoneNumber is just string "123".

    // Compare
    if (rowPhone === formObject.phoneNumber) {
      // Duplicate found
      var dupTime = Utilities.formatDate(rowDate, Session.getScriptTimeZone(), "HH:mm:ss dd/MM/yyyy");
      var dupEmail = data[i][1];
      var dupEmp = data[i][3];

      return {
        success: false,
        message: "Số thuê bao " + formObject.phoneNumber + " đã được nhập lúc " + dupTime + " bởi " + dupEmail + " cho nhân viên " + dupEmp
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
  var dataSheet = ss.getSheetByName('Data');

  if (!dataSheet) return [];

  var data = dataSheet.getDataRange().getValues();
  // Remove header
  if (data.length > 0) data.shift();

  var results = [];
  var targetStartDate = startDateStr ? new Date(startDateStr) : null;
  if (targetStartDate) targetStartDate.setHours(0, 0, 0, 0);

  var targetEndDate = endDateStr ? new Date(endDateStr) : null;
  if (targetEndDate) targetEndDate.setHours(0, 0, 0, 0);

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var rowDate = new Date(row[0]); // Col 0 is Timestamp
    var rowUnit = row[2]; // Col 2 is Unit

    var rowDay = new Date(rowDate);
    rowDay.setHours(0, 0, 0, 0);

    // Check Date Range
    if (targetStartDate && rowDay.getTime() < targetStartDate.getTime()) {
      continue;
    }
    if (targetEndDate && rowDay.getTime() > targetEndDate.getTime()) {
      continue;
    }

    // Check Unit
    if (unitName && rowUnit !== unitName) {
      continue;
    }

    // Check EmpCode
    var rowEmpCode = row[3];
    if (empCode && rowEmpCode !== empCode) {
      continue;
    }

    // Format date for display
    var formattedDate = Utilities.formatDate(rowDate, Session.getScriptTimeZone(), "HH:mm dd/MM/yyyy");

    results.push({
      time: formattedDate,
      email: row[1],
      unit: row[2],
      empCode: row[3],
      phone: row[4],
      service: row[5],
      price: row[6],
      note: row[7],
      empName: row[8] || ''
    });
  }

  return results;
}
