/**
 * AdminService
 */
const AdminService = (() => {
  function writeLog(action, targetType, targetId, payload) {
    const sh = LookupService.mustSheet('AuditLogs');
    sh.appendRow([
      LookupService.genId('log'),
      Session.getEffectiveUser().getEmail(),
      action,
      targetType,
      targetId,
      JSON.stringify(payload || {}),
      new Date()
    ]);
  }

  function listLogs(limit = 200) {
    PermissionService.assertRoleAtLeast('manager');
    const rows = LookupService.readObjects(LookupService.mustSheet('AuditLogs'));
    return rows.slice(-limit).reverse();
  }

  function listAdmins() {
    PermissionService.assertRoleAtLeast('admin');
    return LookupService.readObjects(LookupService.mustSheet('AdminAccounts'));
  }

  function upsertAdmin(payload) {
    PermissionService.assertRoleAtLeast('owner');
    if (!payload.email || !payload.role) throw new Error('email/role required');

    const sh = LookupService.mustSheet('AdminAccounts');
    const data = sh.getDataRange().getValues();
    const idx = LookupService.headerIndex(data[0]);
    const now = new Date();

    for (let i = 1; i < data.length; i += 1) {
      if (String(data[i][idx.email]).toLowerCase() === String(payload.email).toLowerCase()) {
        data[i][idx.name] = payload.name || data[i][idx.name];
        data[i][idx.role] = payload.role;
        data[i][idx.is_active] = payload.is_active !== false;
        data[i][idx.updated_at] = now;
        sh.getRange(i + 1, 1, 1, data[0].length).setValues([data[i]]);
        return { ok: true, mode: 'updated' };
      }
    }

    sh.appendRow([payload.email, payload.name || '', payload.role, payload.is_active !== false, now, now]);
    return { ok: true, mode: 'created' };
  }

  return { writeLog, listLogs, listAdmins, upsertAdmin };
})();
