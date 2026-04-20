/**
 * 舊專案 / 修正版資料整合工具
 * - 提供欄位映射匯入
 * - 匯入前後檢查錯
 * - 可輸出錯誤報表到 ValidationReport
 */

function importRowsWithMapping(sheetName, sourceHeaders, sourceRows, headerMap) {
  if (!sheetName) throw new Error('sheetName 必填');
  if (!sourceHeaders || !sourceHeaders.length) throw new Error('sourceHeaders 必填');
  if (!sourceRows || !sourceRows.length) return { imported: 0, skipped: 0, errors: [] };

  var targetSheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
  if (!targetSheet) throw new Error('目標分頁不存在: ' + sheetName);

  var targetHeaders = targetSheet.getRange(1, 1, 1, targetSheet.getLastColumn()).getValues()[0];
  var sourceIndex = indexFromArray_(sourceHeaders);
  var targetIndex = indexFromArray_(targetHeaders);

  var errors = [];
  var outputRows = [];

  for (var i = 0; i < sourceRows.length; i++) {
    var srcRow = sourceRows[i];
    var out = new Array(targetHeaders.length);

    targetHeaders.forEach(function (targetHeader) {
      var mapped = headerMap && headerMap[targetHeader] ? headerMap[targetHeader] : targetHeader;
      if (sourceIndex[mapped] !== undefined) {
        out[targetIndex[targetHeader]] = srcRow[sourceIndex[mapped]];
      } else {
        out[targetIndex[targetHeader]] = '';
      }
    });

    outputRows.push(out);
  }

  if (outputRows.length) {
    targetSheet.getRange(targetSheet.getLastRow() + 1, 1, outputRows.length, targetHeaders.length).setValues(outputRows);
  }

  return {
    imported: outputRows.length,
    skipped: 0,
    errors: errors
  };
}

function verifyWorkbookIntegrity() {
  var report = {
    missingSheets: [],
    missingHeaders: [],
    invalidStatus: [],
    orphanSegments: [],
    staffOverlaps: [],
    sessionOverCapacity: []
  };

  var ss = SpreadsheetApp.getActive();
  APP_CONFIG.sheets.forEach(function (cfg) {
    var sh = ss.getSheetByName(cfg.name);
    if (!sh) {
      report.missingSheets.push(cfg.name);
      return;
    }

    var actual = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    var actualSet = indexFromArray_(actual);
    cfg.headers.forEach(function (h) {
      if (actualSet[h] === undefined) {
        report.missingHeaders.push(cfg.name + '.' + h);
      }
    });
  });

  validateStatusWhitelist_(report);
  validateOrphanSegments_(report);
  validateStaffOverlap_(report);
  validateSessionCapacity_(report);

  writeValidationReport_(report);
  return report;
}

function validateStatusWhitelist_(report) {
  var apptSheet = SpreadsheetApp.getActive().getSheetByName('Appointments');
  var segSheet = SpreadsheetApp.getActive().getSheetByName('AppointmentSegments');
  if (!apptSheet || !segSheet) return;

  var allowedAppt = getEnumValues_('appointment_status');
  var allowedSeg = getEnumValues_('segment_status');

  var appt = apptSheet.getDataRange().getValues();
  if (appt.length > 1) {
    var apptHeader = indexFromArray_(appt[0]);
    for (var i = 1; i < appt.length; i++) {
      var status = appt[i][apptHeader.status];
      if (status && allowedAppt.indexOf(status) === -1) {
        report.invalidStatus.push('Appointments row ' + (i + 1) + ' status=' + status);
      }
    }
  }

  var seg = segSheet.getDataRange().getValues();
  if (seg.length > 1) {
    var segHeader = indexFromArray_(seg[0]);
    for (var j = 1; j < seg.length; j++) {
      var segStatus = seg[j][segHeader.segment_status];
      if (segStatus && allowedSeg.indexOf(segStatus) === -1) {
        report.invalidStatus.push('AppointmentSegments row ' + (j + 1) + ' segment_status=' + segStatus);
      }
    }
  }
}

function validateOrphanSegments_(report) {
  var apptSheet = SpreadsheetApp.getActive().getSheetByName('Appointments');
  var segSheet = SpreadsheetApp.getActive().getSheetByName('AppointmentSegments');
  if (!apptSheet || !segSheet) return;

  var apptData = apptSheet.getDataRange().getValues();
  var segData = segSheet.getDataRange().getValues();
  if (apptData.length < 2 || segData.length < 2) return;

  var aIdx = indexFromArray_(apptData[0]);
  var sIdx = indexFromArray_(segData[0]);

  var apptIds = {};
  for (var i = 1; i < apptData.length; i++) {
    apptIds[String(apptData[i][aIdx.appointment_id])] = true;
  }

  for (var j = 1; j < segData.length; j++) {
    var aid = String(segData[j][sIdx.appointment_id]);
    if (aid && !apptIds[aid]) {
      report.orphanSegments.push('AppointmentSegments row ' + (j + 1) + ' appointment_id=' + aid);
    }
  }
}

function validateStaffOverlap_(report) {
  var segSheet = SpreadsheetApp.getActive().getSheetByName('AppointmentSegments');
  if (!segSheet) return;

  var data = segSheet.getDataRange().getValues();
  if (data.length < 3) return;

  var h = indexFromArray_(data[0]);
  var byStaff = {};

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var staff = row[h.staff_id];
    if (!staff) continue;
    if (row[h.segment_status] === 'cancelled') continue;

    if (!byStaff[staff]) byStaff[staff] = [];
    byStaff[staff].push({
      rowNo: i + 1,
      start: row[h.start_at],
      end: row[h.end_at]
    });
  }

  Object.keys(byStaff).forEach(function (staff) {
    var arr = byStaff[staff];
    arr.sort(function (a, b) { return new Date(a.start) - new Date(b.start); });
    for (var k = 1; k < arr.length; k++) {
      if (rangesOverlap_(arr[k - 1].start, arr[k - 1].end, arr[k].start, arr[k].end)) {
        report.staffOverlaps.push('staff=' + staff + ' rows=' + arr[k - 1].rowNo + ',' + arr[k].rowNo);
      }
    }
  });
}

function validateSessionCapacity_(report) {
  var sessionSheet = SpreadsheetApp.getActive().getSheetByName('Sessions');
  var segSheet = SpreadsheetApp.getActive().getSheetByName('AppointmentSegments');
  if (!sessionSheet || !segSheet) return;

  var sessions = sessionSheet.getDataRange().getValues();
  var segs = segSheet.getDataRange().getValues();
  if (sessions.length < 2 || segs.length < 2) return;

  var sh = indexFromArray_(sessions[0]);
  var gh = indexFromArray_(segs[0]);

  var capBySession = {};
  for (var i = 1; i < sessions.length; i++) {
    capBySession[sessions[i][sh.session_id]] = Number(sessions[i][sh.capacity] || 1);
  }

  var groups = {};
  for (var j = 1; j < segs.length; j++) {
    if (segs[j][gh.segment_status] === 'cancelled') continue;
    var sessionId = segs[j][gh.session_id];
    if (!sessionId) continue;

    if (!groups[sessionId]) groups[sessionId] = [];
    groups[sessionId].push({
      rowNo: j + 1,
      start: segs[j][gh.start_at],
      end: segs[j][gh.end_at]
    });
  }

  Object.keys(groups).forEach(function (sessionId) {
    var cap = capBySession[sessionId] || 1;
    var items = groups[sessionId];

    for (var x = 0; x < items.length; x++) {
      var overlapCount = 0;
      for (var y = 0; y < items.length; y++) {
        if (rangesOverlap_(items[x].start, items[x].end, items[y].start, items[y].end)) {
          overlapCount++;
        }
      }
      if (overlapCount > cap) {
        report.sessionOverCapacity.push('session=' + sessionId + ' row=' + items[x].rowNo + ' overlap=' + overlapCount + ' cap=' + cap);
      }
    }
  });
}

function writeValidationReport_(report) {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName('ValidationReport');
  if (!sh) sh = ss.insertSheet('ValidationReport');

  sh.clearContents();
  sh.getRange(1, 1, 1, 4).setValues([['category', 'message', 'count', 'checked_at']]);

  var now = new Date();
  var rows = [];
  Object.keys(report).forEach(function (category) {
    var arr = report[category] || [];
    if (!arr.length) {
      rows.push([category, 'OK', 0, now]);
    } else {
      arr.forEach(function (msg) {
        rows.push([category, msg, arr.length, now]);
      });
    }
  });

  if (rows.length) {
    sh.getRange(2, 1, rows.length, 4).setValues(rows);
  }
  sh.autoResizeColumns(1, 4);
}

function indexFromArray_(headers) {
  var map = {};
  for (var i = 0; i < headers.length; i++) {
    map[String(headers[i])] = i;
  }
  return map;
}
