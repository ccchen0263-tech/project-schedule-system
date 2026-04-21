/**
 * AppointmentService
 */
const AppointmentService = (() => {
  function create(payload) {
    PermissionService.assertRoleAtLeast('staff');
    validatePayload(payload);

    const lock = LockService.getDocumentLock();
    lock.waitLock(20000);
    try {
      const appointmentId = payload.appointment_id || LookupService.genId('appt');
      payload.segments.forEach((seg) => {
        const check = ScheduleService.checkSlotAvailability(seg);
        if (!check.ok) throw new Error(`slot unavailable: ${JSON.stringify(check)}`);
      });

      writeAppointment(appointmentId, payload);
      writeSegments(appointmentId, payload.segments);
      AdminService.writeLog('create_appointment', 'appointment', appointmentId, payload);
      return { ok: true, appointment_id: appointmentId };
    } finally {
      lock.releaseLock();
    }
  }

  function listByDate(dateText) {
    const date = dateText || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const rows = LookupService.readObjects(LookupService.mustSheet('Appointments'));
    return rows.filter((r) => String(r.appointment_date) === date);
  }

  function cancel(appointmentId, reason) {
    PermissionService.assertRoleAtLeast('staff');
    const appSheet = LookupService.mustSheet('Appointments');
    const segSheet = LookupService.mustSheet('AppointmentSegments');

    updateStatusById(appSheet, 'appointment_id', appointmentId, 'status', 'cancelled');
    updateStatusById(segSheet, 'appointment_id', appointmentId, 'segment_status', 'cancelled');

    AdminService.writeLog('cancel_appointment', 'appointment', appointmentId, { reason: reason || '' });
    return { ok: true };
  }

  function updateStatusById(sheet, idCol, idValue, statusCol, statusValue) {
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return;
    const idx = LookupService.headerIndex(data[0]);

    for (let i = 1; i < data.length; i += 1) {
      if (String(data[i][idx[idCol]]) === String(idValue)) {
        data[i][idx[statusCol]] = statusValue;
      }
    }
    sheet.getRange(2, 1, data.length - 1, data[0].length).setValues(data.slice(1));
  }

  function validatePayload(payload) {
    if (!payload.client_id) throw new Error('client_id required');
    if (!payload.appointment_date) throw new Error('appointment_date required');
    if (!Array.isArray(payload.segments) || !payload.segments.length) throw new Error('segments required');

    const whitelist = LookupService.enumValues('appointment_status');
    if (!whitelist.includes(payload.status || 'booked')) throw new Error('invalid appointment status');

    payload.segments.forEach((seg) => {
      ['service_type_id', 'role', 'staff_id', 'session_id', 'start_at', 'end_at', 'duration_min'].forEach((k) => {
        if (seg[k] === undefined || seg[k] === '') throw new Error(`segment.${k} required`);
      });
    });
  }

  function writeAppointment(id, payload) {
    const sh = LookupService.mustSheet('Appointments');
    const headers = LookupService.headers(sh);
    const map = LookupService.headerIndex(headers);
    const now = new Date();

    const row = new Array(headers.length).fill('');
    row[map.appointment_id] = id;
    row[map.client_id] = payload.client_id;
    row[map.appointment_date] = payload.appointment_date;
    row[map.status] = payload.status || 'booked';
    row[map.source] = payload.source || 'manual';
    row[map.notes] = payload.notes || '';
    row[map.created_by] = Session.getEffectiveUser().getEmail();
    row[map.created_at] = now;
    row[map.updated_at] = now;
    sh.appendRow(row);
  }

  function writeSegments(appointmentId, segments) {
    const sh = LookupService.mustSheet('AppointmentSegments');
    const headers = LookupService.headers(sh);
    const map = LookupService.headerIndex(headers);

    segments.forEach((seg, i) => {
      const row = new Array(headers.length).fill('');
      row[map.segment_id] = LookupService.genId('seg');
      row[map.appointment_id] = appointmentId;
      row[map.sequence] = i + 1;
      row[map.service_type_id] = seg.service_type_id;
      row[map.role] = seg.role;
      row[map.staff_id] = seg.staff_id;
      row[map.session_id] = seg.session_id;
      row[map.start_at] = seg.start_at;
      row[map.end_at] = seg.end_at;
      row[map.duration_min] = seg.duration_min;
      row[map.segment_status] = seg.segment_status || 'booked';
      sh.appendRow(row);
    });
  }

  return { create, listByDate, cancel };
})();
