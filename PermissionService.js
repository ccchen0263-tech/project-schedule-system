/**
 * PermissionService
 */
const PermissionService = (() => {
  const ROLE_WEIGHT = { owner: 100, admin: 80, manager: 60, staff: 40, viewer: 10 };

  function ensureOwnerAdmin() {
    const sheet = LookupService.mustSheet('AdminAccounts');
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) {
      const now = new Date();
      sheet.appendRow([Session.getEffectiveUser().getEmail(), 'Owner', 'owner', true, now, now]);
    }
  }

  function getCurrentUserRole() {
    const email = Session.getEffectiveUser().getEmail();
    return getRoleByEmail(email);
  }

  function getRoleByEmail(email) {
    const sheet = LookupService.mustSheet('AdminAccounts');
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return 'viewer';

    const h = LookupService.headerIndex(data[0]);
    for (let i = 1; i < data.length; i += 1) {
      if (String(data[i][h.email]).toLowerCase() === String(email).toLowerCase() && data[i][h.is_active] === true) {
        return data[i][h.role] || 'viewer';
      }
    }
    return 'viewer';
  }

  function assertRoleAtLeast(minRole) {
    const current = getCurrentUserRole();
    if ((ROLE_WEIGHT[current] || 0) < (ROLE_WEIGHT[minRole] || 999)) {
      throw new Error(`Permission denied. Require ${minRole}, current ${current}`);
    }
  }

  return { ensureOwnerAdmin, getCurrentUserRole, getRoleByEmail, assertRoleAtLeast };
})();
