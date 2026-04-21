/**
 * 初始化入口：在空白 Google Sheet 執行。
 */
function setupWorkbook() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetTimeZone(APP_CONFIG.timezone);

  APP_CONFIG.sheets.forEach(function (cfg) {
    upsertSheetWithHeaders_(ss, cfg.name, cfg.headers);
  });

  seedEnums_();
  seedServiceTypes_();
  seedDemoSessions_();
  setupDataValidation_();

  SpreadsheetApp.flush();
  Logger.log('✅ setupWorkbook 完成。');
}

function resetWorkbookDangerous() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  APP_CONFIG.sheets.forEach(function (cfg) {
    var sh = ss.getSheetByName(cfg.name);
    if (sh) {
      sh.clearContents();
      sh.clearFormats();
      sh.getRange(1, 1, 1, cfg.headers.length).setValues([cfg.headers]);
      styleHeader_(sh, cfg.headers.length);
      sh.setFrozenRows(1);
      sh.autoResizeColumns(1, cfg.headers.length);
    }
  });
  seedEnums_();
  seedServiceTypes_();
  seedDemoSessions_();
  setupDataValidation_();
  Logger.log('⚠️ resetWorkbookDangerous 完成（已清空資料）。');
}

function upsertSheetWithHeaders_(ss, name, headers) {
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);

  var existingLastCol = Math.max(sh.getLastColumn(), headers.length);
  if (existingLastCol > 0) {
    sh.getRange(1, 1, 1, existingLastCol).clearContent();
  }

  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeader_(sh, headers.length);
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, headers.length);
}

function styleHeader_(sheet, colCount) {
  sheet
    .getRange(1, 1, 1, colCount)
    .setFontWeight('bold')
    .setBackground('#eff6ff')
    .setFontColor('#1d4ed8')
    .setHorizontalAlignment('center');
}

function seedEnums_() {
  var sheet = SpreadsheetApp.getActive().getSheetByName('Enums');
  if (!sheet) throw new Error('Enums sheet 不存在');

  var headers = APP_CONFIG.sheets.filter(function (s) { return s.name === 'Enums'; })[0].headers;
  var data = APP_CONFIG.enums;

  sheet.getRange(2, 1, Math.max(sheet.getLastRow() - 1, 0), headers.length).clearContent();
  if (data.length > 0) {
    sheet.getRange(2, 1, data.length, headers.length).setValues(data);
  }
}

function seedServiceTypes_() {
  var sheet = SpreadsheetApp.getActive().getSheetByName('ServiceTypes');
  if (!sheet) throw new Error('ServiceTypes sheet 不存在');

  var baseRows = APP_CONFIG.seedServiceTypes.map(function (row) {
    var now = new Date();
    return row.concat([now, now]);
  });

  sheet.getRange(2, 1, Math.max(sheet.getLastRow() - 1, 0), sheet.getLastColumn()).clearContent();
  if (baseRows.length > 0) {
    sheet.getRange(2, 1, baseRows.length, baseRows[0].length).setValues(baseRows);
  }
}

function seedDemoSessions_() {
  var sheet = SpreadsheetApp.getActive().getSheetByName('Sessions');
  if (!sheet) throw new Error('Sessions sheet 不存在');

  var now = new Date();
  var rows = [
    ['session_1', '診1', 'clinic_room', '09:00', '12:00', 15, 1, 'doctor,consultant', '1,2,3,4,5,6', true, 5, now, now],
    ['session_2', '診2', 'clinic_room', '14:00', '17:30', 30, 1, 'doctor,dietitian', '1,2,3,4,5,6', true, 5, now, now],
    ['session_3', '診3', 'consult_room', '18:00', '21:00', 30, 1, 'consultant,dietitian', '1,2,3,4,5', true, 0, now, now]
  ];

  sheet.getRange(2, 1, Math.max(sheet.getLastRow() - 1, 0), sheet.getLastColumn()).clearContent();
  sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
}

function setupDataValidation_() {
  var ss = SpreadsheetApp.getActive();
  var apptSheet = ss.getSheetByName('Appointments');
  var segSheet = ss.getSheetByName('AppointmentSegments');

  var statusList = getEnumValues_('appointment_status');
  var segStatusList = getEnumValues_('segment_status');

  if (apptSheet) {
    var rule = SpreadsheetApp.newDataValidation().requireValueInList(statusList, true).setAllowInvalid(false).build();
    applyValidationByHeader_(apptSheet, 'status', rule, 2, 5000);
  }

  if (segSheet) {
    var segRule = SpreadsheetApp.newDataValidation().requireValueInList(segStatusList, true).setAllowInvalid(false).build();
    applyValidationByHeader_(segSheet, 'segment_status', segRule, 2, 5000);
  }
}

function applyValidationByHeader_(sheet, headerName, rule, startRow, rowCount) {
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var col = headers.indexOf(headerName) + 1;
  if (col <= 0) throw new Error('找不到欄位: ' + headerName + ' @ ' + sheet.getName());
  sheet.getRange(startRow, col, rowCount, 1).setDataValidation(rule);
}

function getEnumValues_(category) {
  var sheet = SpreadsheetApp.getActive().getSheetByName('Enums');
  if (!sheet) return [];

  var data = sheet.getDataRange().getValues();
  var result = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === category && data[i][3] === true) {
      result.push(data[i][1]);
    }
  }
  return result;
}
