/**
 * SetupService
 * 初始化空白 Google Sheet 的主要入口。
 */
const SetupService = (() => {
  const REQUIRED_SHEETS = {
    Clients: ['client_id', 'name', 'phone', 'birthday', 'gender', 'source', 'notes', 'is_blacklisted', 'created_at', 'updated_at'],
    Staff: ['staff_id', 'name', 'roles', 'is_active', 'is_bookable', 'color', 'created_at', 'updated_at'],
    Sessions: ['session_id', 'name', 'weekday', 'start_time', 'end_time', 'slot_min', 'capacity', 'is_active'],
    ServiceTypes: ['service_type_id', 'name', 'default_role', 'duration_min', 'is_active'],
    Appointments: ['appointment_id', 'client_id', 'appointment_date', 'status', 'source', 'notes', 'created_by', 'created_at', 'updated_at'],
    AppointmentSegments: ['segment_id', 'appointment_id', 'sequence', 'service_type_id', 'role', 'staff_id', 'session_id', 'start_at', 'end_at', 'duration_min', 'segment_status'],
    AuditLogs: ['log_id', 'actor', 'action', 'target_type', 'target_id', 'payload_json', 'created_at'],
    AdminAccounts: ['email', 'name', 'role', 'is_active', 'created_at', 'updated_at'],
    Enums: ['category', 'value', 'label', 'is_active', 'sort_order']
  };

  const ENUM_SEEDS = [
    ['appointment_status', 'booked', '已預約', true, 1],
    ['appointment_status', 'checked_in', '已報到', true, 2],
    ['appointment_status', 'completed', '已完成', true, 3],
    ['appointment_status', 'cancelled', '取消', true, 4],
    ['segment_status', 'booked', '已預約', true, 1],
    ['segment_status', 'completed', '已完成', true, 2],
    ['segment_status', 'cancelled', '取消', true, 3],
    ['role', 'doctor', '醫師', true, 1],
    ['role', 'consultant', '諮詢師', true, 2],
    ['role', 'dietitian', '營養師', true, 3]
  ];

  function setupAll() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    Object.keys(REQUIRED_SHEETS).forEach((name) => upsertSheet(ss, name, REQUIRED_SHEETS[name]));
    seedEnums();
    seedDefaults();
    PermissionService.ensureOwnerAdmin();
    return { ok: true, message: 'Setup completed' };
  }

  function upsertSheet(ss, name, headers) {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    sh.clear();
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#eef6ff');
    sh.setFrozenRows(1);
    sh.autoResizeColumns(1, headers.length);
  }

  function seedEnums() {
    const sh = LookupService.mustSheet('Enums');
    sh.getRange(2, 1, Math.max(0, sh.getLastRow() - 1), 5).clearContent();
    sh.getRange(2, 1, ENUM_SEEDS.length, 5).setValues(ENUM_SEEDS);
  }

  function seedDefaults() {
    const now = new Date();
    const serviceSheet = LookupService.mustSheet('ServiceTypes');
    const rows = [
      ['svc_initial', '初診諮詢', 'consultant', 30, true],
      ['svc_doctor', '醫師門診', 'doctor', 15, true],
      ['svc_dietitian', '營養追蹤', 'dietitian', 30, true]
    ];
    serviceSheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);

    const sessions = LookupService.mustSheet('Sessions');
    const sessionRows = [
      ['session_1', '診1', 'Mon,Tue,Wed,Thu,Fri,Sat', '09:00', '12:00', 15, 1, true],
      ['session_2', '診2', 'Mon,Tue,Wed,Thu,Fri,Sat', '14:00', '17:30', 30, 1, true],
      ['session_3', '診3', 'Mon,Tue,Wed,Thu,Fri', '18:00', '21:00', 30, 1, true]
    ];
    sessions.getRange(2, 1, sessionRows.length, sessionRows[0].length).setValues(sessionRows);

    const adminSheet = LookupService.mustSheet('AdminAccounts');
    adminSheet.getRange(2, 1, 1, 6).setValues([[Session.getEffectiveUser().getEmail(), 'Owner', 'owner', true, now, now]]);
  }

  return { setupAll };
})();
