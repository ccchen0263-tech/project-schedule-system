/**
 * 書心預約系統 Apps Script 設定檔
 * 目標：可直接在空白 Google Sheet 執行 setup，建立新版結構。
 */
var APP_CONFIG = {
  timezone: 'Asia/Taipei',
  sheets: [
    {
      name: 'Clients',
      headers: ['client_id', 'name', 'phone', 'birthday', 'gender', 'source', 'first_visit_date', 'notes', 'is_blacklisted', 'alerts', 'created_at', 'updated_at']
    },
    {
      name: 'Staff',
      headers: ['staff_id', 'name', 'roles', 'is_bookable', 'color', 'sort_order', 'is_active', 'service_types', 'created_at', 'updated_at']
    },
    {
      name: 'Sessions',
      headers: ['session_id', 'name', 'resource_type', 'start_time', 'end_time', 'default_duration_min', 'capacity', 'roles', 'weekdays', 'is_active', 'buffer_min', 'created_at', 'updated_at']
    },
    {
      name: 'ServiceTypes',
      headers: ['service_type_id', 'name', 'default_role', 'default_duration_min', 'is_initial_visit', 'require_session', 'require_doctor', 'require_dietitian', 'color', 'is_active', 'created_at', 'updated_at']
    },
    {
      name: 'Appointments',
      headers: ['appointment_id', 'client_id', 'appointment_date', 'source', 'status', 'notes', 'created_by', 'created_at', 'updated_at', 'followup_batch', 'reminder_status']
    },
    {
      name: 'AppointmentSegments',
      headers: ['segment_id', 'appointment_id', 'sequence', 'service_type_id', 'role', 'staff_id', 'session_id', 'start_at', 'end_at', 'duration_min', 'segment_status', 'notes', 'created_at', 'updated_at']
    },
    {
      name: 'Closures',
      headers: ['closure_id', 'date', 'start_at', 'end_at', 'target_type', 'target_id', 'reason', 'block_booking', 'created_at', 'updated_at']
    },
    {
      name: 'Waitlist',
      headers: ['waitlist_id', 'client_id', 'preferred_date', 'preferred_time_range', 'preferred_role', 'priority', 'notify_method', 'status', 'notes', 'created_at', 'updated_at']
    },
    {
      name: 'AuditLog',
      headers: ['audit_id', 'action', 'actor', 'target_type', 'target_id', 'payload_json', 'created_at']
    },
    {
      name: 'Enums',
      headers: ['category', 'value', 'label', 'is_active', 'sort_order']
    },
    {
      name: 'ValidationReport',
      headers: ['category', 'message', 'count', 'checked_at']
    }
  ],
  enums: [
    ['appointment_status', 'booked', '已預約', true, 1],
    ['appointment_status', 'confirmed', '已確認', true, 2],
    ['appointment_status', 'checked_in', '已報到', true, 3],
    ['appointment_status', 'in_progress', '進行中', true, 4],
    ['appointment_status', 'completed', '已完成', true, 5],
    ['appointment_status', 'cancelled', '取消', true, 6],
    ['appointment_status', 'rescheduled', '改期', true, 7],
    ['appointment_status', 'no_show', '爽約', true, 8],
    ['segment_status', 'booked', '已預約', true, 1],
    ['segment_status', 'completed', '已完成', true, 2],
    ['segment_status', 'cancelled', '取消', true, 3],
    ['role', 'doctor', '醫師', true, 1],
    ['role', 'consultant', '諮詢師', true, 2],
    ['role', 'dietitian', '營養師', true, 3],
    ['waitlist_status', 'waiting', '候補中', true, 1],
    ['waitlist_status', 'notified', '已通知', true, 2],
    ['waitlist_status', 'fulfilled', '已補位', true, 3],
    ['waitlist_status', 'closed', '已關閉', true, 4]
  ],
  seedServiceTypes: [
    ['svc_initial_consult', '初診諮詢', 'consultant', 30, true, true, false, false, '#22c55e', true],
    ['svc_doctor_followup', '醫師複診', 'doctor', 15, false, true, true, false, '#2563eb', true],
    ['svc_dietitian_followup', '營養追蹤', 'dietitian', 30, false, true, false, true, '#f59e0b', true]
  ]
};
