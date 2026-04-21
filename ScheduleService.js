/**
 * ScheduleService
 */
const ScheduleService = (() => {
  function getDaySchedule(dateText) {
    const date = dateText || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const appts = LookupService.readObjects(LookupService.mustSheet('Appointments'));
    const segs = LookupService.readObjects(LookupService.mustSheet('AppointmentSegments'));

    const apptMap = {};
    appts.forEach((a) => { apptMap[a.appointment_id] = a; });

    return segs
      .filter((s) => String(s.start_at || '').startsWith(date))
      .map((s) => ({ ...s, appointment: apptMap[s.appointment_id] || null }))
      .sort((a, b) => new Date(a.start_at) - new Date(b.start_at));
  }

  function checkSlotAvailability({ staff_id, session_id, start_at, end_at }) {
    const segments = LookupService.readObjects(LookupService.mustSheet('AppointmentSegments'));
    const sessionCap = getSessionCapacity(session_id);

    let staffOverlap = false;
    let overlapCount = 0;

    segments.forEach((s) => {
      if (s.segment_status === 'cancelled') return;
      if (!isOverlap(s.start_at, s.end_at, start_at, end_at)) return;
      if (s.staff_id === staff_id) staffOverlap = true;
      if (s.session_id === session_id) overlapCount += 1;
    });

    return {
      ok: !staffOverlap && overlapCount < sessionCap,
      staffOverlap,
      overlapCount,
      sessionCap
    };
  }

  function getSessionCapacity(sessionId) {
    const sessions = LookupService.readObjects(LookupService.mustSheet('Sessions'));
    const found = sessions.find((s) => s.session_id === sessionId);
    return Number(found?.capacity || 1);
  }

  function isOverlap(aStart, aEnd, bStart, bEnd) {
    const as = new Date(aStart).getTime();
    const ae = new Date(aEnd).getTime();
    const bs = new Date(bStart).getTime();
    const be = new Date(bEnd).getTime();
    return as < be && bs < ae;
  }

  return { getDaySchedule, checkSlotAvailability };
})();
