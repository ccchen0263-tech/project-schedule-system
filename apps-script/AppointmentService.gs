/**
 * 預約建立與檢錯核心服務。
 */
function createAppointment(payload) {
  validateCreatePayload_(payload);

  var lock = LockService.getDocumentLock();
  lock.waitLock(20000);

  try {
    var availability = checkAvailability(payload);
    if (!availability.ok) {
      throw new Error('不可預約：' + availability.reasons.join('；'));
    }

    var now = new Date();
    var appointmentId = payload.appointment_id || genId_('appt');

    writeAppointmentRow_(Object.assign({}, payload, {
      appointment_id: appointmentId,
      created_at: now,
      updated_at: now
    }));

    payload.segments.forEach(function (seg, idx) {
      writeSegmentRow_(Object.assign({}, seg, {
        segment_id: seg.segment_id || genId_('seg'),
        appointment_id: appointmentId,
        sequence: idx + 1,
        created_at: now,
        updated_at: now
      }));
    });

    writeAuditLog_('create_appointment', appointmentId, payload);

    return {
      ok: true,
      appointment_id: appointmentId
    };
  } finally {
    lock.releaseLock();
  }
}

function checkAvailability(payload) {
  var reasons = [];

  payload.segments.forEach(function (seg) {
    if (hasStaffOverlap_(seg.staff_id, seg.start_at, seg.end_at)) {
      reasons.push('staff 重疊: ' + seg.staff_id + ' @ ' + seg.start_at + '~' + seg.end_at);
    }
    if (hasSessionOverCapacity_(seg.session_id, seg.start_at, seg.end_at)) {
      reasons.push('session 滿載: ' + seg.session_id + ' @ ' + seg.start_at + '~' + seg.end_at);
    }
  });

  return { ok: reasons.length === 0, reasons: reasons };
}

function validateCreatePayload_(payload) {
  if (!payload) throw new Error('payload 不可為空');
  if (!payload.client_id) throw new Error('client_id 必填');
  if (!payload.appointment_date) throw new Error('appointment_date 必填');
  if (!payload.status) throw new Error('status 必填');
  if (!payload.segments || !payload.segments.length) throw new Error('segments 至少 1 筆');

  var allowed = getEnumValues_('appointment_status');
  if (allowed.indexOf(payload.status) === -1) {
    throw new Error('status 不在白名單: ' + payload.status);
  }
}

function writeAppointmentRow_(payload) {
  var sheet = SpreadsheetApp.getActive().getSheetByName('Appointments');
  var map = indexByHeader_(sheet);

  var row = [];
  row[map.appointment_id] = payload.appointment_id;
  row[map.client_id] = payload.client_id;
  row[map.appointment_date] = payload.appointment_date;
  row[map.source] = payload.source || 'manual';
  row[map.status] = payload.status;
  row[map.notes] = payload.notes || '';
  row[map.created_by] = payload.created_by || Session.getActiveUser().getEmail();
  row[map.created_at] = payload.created_at;
  row[map.updated_at] = payload.updated_at;
  row[map.followup_batch] = payload.followup_batch || '';
  row[map.reminder_status] = payload.reminder_status || '';

  sheet.appendRow(normalizeRowLength_(row, sheet.getLastColumn()));
}

function writeSegmentRow_(segment) {
  var sheet = SpreadsheetApp.getActive().getSheetByName('AppointmentSegments');
  var map = indexByHeader_(sheet);

  var row = [];
  row[map.segment_id] = segment.segment_id;
  row[map.appointment_id] = segment.appointment_id;
  row[map.sequence] = segment.sequence;
  row[map.service_type_id] = segment.service_type_id;
  row[map.role] = segment.role;
  row[map.staff_id] = segment.staff_id;
  row[map.session_id] = segment.session_id;
  row[map.start_at] = segment.start_at;
  row[map.end_at] = segment.end_at;
  row[map.duration_min] = segment.duration_min;
  row[map.segment_status] = segment.segment_status || 'booked';
  row[map.notes] = segment.notes || '';
  row[map.created_at] = segment.created_at;
  row[map.updated_at] = segment.updated_at;

  sheet.appendRow(normalizeRowLength_(row, sheet.getLastColumn()));
}

function hasStaffOverlap_(staffId, startAt, endAt) {
  if (!staffId) return false;
  var sheet = SpreadsheetApp.getActive().getSheetByName('AppointmentSegments');
  var values = sheet.getDataRange().getValues();
  var headers = values[0];

  var iStaff = headers.indexOf('staff_id');
  var iStart = headers.indexOf('start_at');
  var iEnd = headers.indexOf('end_at');
  var iStatus = headers.indexOf('segment_status');

  for (var i = 1; i < values.length; i++) {
    if (values[i][iStaff] !== staffId) continue;
    if (values[i][iStatus] === 'cancelled') continue;
    if (rangesOverlap_(values[i][iStart], values[i][iEnd], startAt, endAt)) return true;
  }
  return false;
}

function hasSessionOverCapacity_(sessionId, startAt, endAt) {
  if (!sessionId) return false;
  var sessions = SpreadsheetApp.getActive().getSheetByName('Sessions').getDataRange().getValues();
  var sHeaders = sessions[0];
  var iId = sHeaders.indexOf('session_id');
  var iCap = sHeaders.indexOf('capacity');

  var cap = 1;
  for (var i = 1; i < sessions.length; i++) {
    if (sessions[i][iId] === sessionId) {
      cap = Number(sessions[i][iCap] || 1);
      break;
    }
  }

  var segSheet = SpreadsheetApp.getActive().getSheetByName('AppointmentSegments');
  var segs = segSheet.getDataRange().getValues();
  var h = segs[0];
  var idxSession = h.indexOf('session_id');
  var idxStart = h.indexOf('start_at');
  var idxEnd = h.indexOf('end_at');
  var idxStatus = h.indexOf('segment_status');

  var overlapping = 0;
  for (var j = 1; j < segs.length; j++) {
    if (segs[j][idxSession] !== sessionId) continue;
    if (segs[j][idxStatus] === 'cancelled') continue;
    if (rangesOverlap_(segs[j][idxStart], segs[j][idxEnd], startAt, endAt)) overlapping++;
  }

  return overlapping >= cap;
}

function rangesOverlap_(aStart, aEnd, bStart, bEnd) {
  var s1 = new Date(aStart).getTime();
  var e1 = new Date(aEnd).getTime();
  var s2 = new Date(bStart).getTime();
  var e2 = new Date(bEnd).getTime();
  return s1 < e2 && s2 < e1;
}

function indexByHeader_(sheet) {
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var index = {};
  headers.forEach(function (h, i) { index[h] = i; });
  return index;
}

function normalizeRowLength_(row, length) {
  var normalized = [];
  for (var i = 0; i < length; i++) {
    normalized[i] = row[i] !== undefined ? row[i] : '';
  }
  return normalized;
}

function genId_(prefix) {
  return prefix + '_' + Utilities.getUuid().split('-')[0] + '_' + new Date().getTime();
}

function writeAuditLog_(action, targetId, payload) {
  var sheet = SpreadsheetApp.getActive().getSheetByName('AuditLog');
  var now = new Date();
  sheet.appendRow([
    genId_('audit'),
    action,
    Session.getActiveUser().getEmail() || 'system',
    'appointment',
    targetId,
    JSON.stringify(payload),
    now
  ]);
}
