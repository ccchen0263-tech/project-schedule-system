/**
 * AuthService
 */
const AuthService = (() => {
  function me() {
    const email = Session.getEffectiveUser().getEmail();
    const role = PermissionService.getRoleByEmail(email);
    return { email, role, now: new Date().toISOString() };
  }

  function canManage() {
    const role = PermissionService.getCurrentUserRole();
    return ['owner', 'admin', 'manager'].includes(role);
  }

  return { me, canManage };
})();
