/**
 * StaffService
 */
const StaffService = (() => {
  function listActive() {
    const sheet = LookupService.mustSheet('Staff');
    const rows = LookupService.readObjects(sheet);
    return rows.filter((r) => r.is_active === true && r.is_bookable === true);
  }

  function upsertStaff(payload) {
    PermissionService.assertRoleAtLeast('manager');
    if (!payload.staff_id || !payload.name) throw new Error('staff_id/name required');

    const sheet = LookupService.mustSheet('Staff');
    const headers = LookupService.headers(sheet);
    const map = LookupService.headerIndex(headers);
    const values = sheet.getDataRange().getValues();
    const now = new Date();

    for (let i = 1; i < values.length; i += 1) {
      if (values[i][map.staff_id] === payload.staff_id) {
        values[i][map.name] = payload.name;
        values[i][map.roles] = payload.roles || '';
        values[i][map.is_active] = payload.is_active !== false;
        values[i][map.is_bookable] = payload.is_bookable !== false;
        values[i][map.color] = payload.color || '#2563eb';
        values[i][map.updated_at] = now;
        sheet.getRange(i + 1, 1, 1, headers.length).setValues([values[i]]);
        return { ok: true, mode: 'updated' };
      }
    }

    const row = new Array(headers.length).fill('');
    row[map.staff_id] = payload.staff_id;
    row[map.name] = payload.name;
    row[map.roles] = payload.roles || '';
    row[map.is_active] = payload.is_active !== false;
    row[map.is_bookable] = payload.is_bookable !== false;
    row[map.color] = payload.color || '#2563eb';
    row[map.created_at] = now;
    row[map.updated_at] = now;
    sheet.appendRow(row);
    return { ok: true, mode: 'created' };
  }

  return { listActive, upsertStaff };
})();
