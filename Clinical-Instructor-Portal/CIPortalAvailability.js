/**
 * Clinical Instructor Portal — Submit Availability.
 * Writes availability blocks into the shared CI availability Google Calendar
 * (owned by the program; the Portal runs as the owner). Each event carries the
 * instructor's stable Evaluator ID tag so the Portal can list and remove only
 * that instructor's own entries. New entries are Pending (yellow); the
 * coordinator approves in Google Calendar by changing the colour to green or
 * adding "Status: Approved" to the description.
 */
const CI_AVAIL_DEFAULT_CAL = 'c_9a1de097bee716e5e56fc1fa11bcc4b13b56ed7c351a60f27766afd1c14afe25@group.calendar.google.com';
const CI_AVAIL_GREEN = ['2', '10']; // CalendarApp.EventColor.PALE_GREEN, GREEN
const CI_AVAIL_RED = ['4', '11'];   // PALE_RED, RED → declined

function ciAvailCalendarId_() {
  return String(PropertiesService.getScriptProperties().getProperty('CI_PORTAL_AVAILABILITY_CALENDAR_ID') || CI_AVAIL_DEFAULT_CAL).trim();
}
function ciAvailCalendar_() {
  const cal = CalendarApp.getCalendarById(ciAvailCalendarId_());
  if (!cal) throw new Error('The shared availability calendar is not available to the Portal owner account.');
  return cal;
}
function ciAvailTag_(ev) {return '[CI:' + ev.evaluatorId + ']';}
function ciAvailStatus_(e) {
  const color = String(e.getColor() || ''), d = String(e.getDescription() || '');
  if (/Status:\s*Approved/i.test(d) || CI_AVAIL_GREEN.includes(color)) return 'Approved';
  if (/Status:\s*(Declined|Not approved)/i.test(d) || CI_AVAIL_RED.includes(color)) return 'Declined';
  return 'Pending';
}
function ciAvailOut_(e) {
  const tz = Session.getScriptTimeZone(), allDay = e.isAllDayEvent();
  const s = allDay ? e.getAllDayStartDate() : e.getStartTime();
  const en = allDay ? e.getAllDayEndDate() : e.getEndTime();
  const note = (/^Note:\s*(.*)$/m.exec(String(e.getDescription() || '')) || [])[1] || '';
  return {id: e.getId(), title: e.getTitle(), allDay: allDay,
    date: Utilities.formatDate(s, tz, 'yyyy-MM-dd'),
    start: allDay ? '' : Utilities.formatDate(s, tz, 'HH:mm'),
    end: allDay ? '' : Utilities.formatDate(en, tz, 'HH:mm'),
    status: ciAvailStatus_(e), note: note,
    coordinatorNote: ((/^Coordinator note:\s*(.*)$/m.exec(String(e.getDescription() || '')) || [])[1] || '').trim()};
}
function ciAvailMine_(ev, from, to) {
  const tag = ciAvailTag_(ev);
  return ciAvailCalendar_().getEvents(from, to).filter(e => String(e.getDescription() || '').indexOf(tag) >= 0);
}
function ciR1AvailabilityList(token) {
  const ev = ciR1Require_(token);
  const id = ciAvailCalendarId_();
  const from = new Date(); from.setHours(0, 0, 0, 0); from.setDate(from.getDate() - 7);
  const to = new Date(from.getTime()); to.setDate(to.getDate() + 190);
  const events = ciAvailMine_(ev, from, to).map(ciAvailOut_).sort((a, b) => (a.date + a.start).localeCompare(b.date + b.start));
  return {success: true, events: events, calendarId: id,
    embedUrl: 'https://calendar.google.com/calendar/embed?src=' + encodeURIComponent(id) + '&ctz=' + encodeURIComponent(Session.getScriptTimeZone()) + '&mode=MONTH&showPrint=0&showTabs=1&showCalendars=0',
    openUrl: 'https://calendar.google.com/calendar/u/0/r?cid=' + encodeURIComponent(id)};
}
function ciR1AvailabilityAdd(token, request) {
  const ev = ciR1Require_(token);
  if (String(ev.evaluatorType || '').trim().toLowerCase() !== 'college ci') throw new Error('Availability is for college clinical instructors.');
  const req = request || {};
  const dates = Array.from(new Set((Array.isArray(req.dates) ? req.dates : [req.date]).map(d => String(d || '').trim()))).filter(Boolean);
  if (!dates.length) throw new Error('Choose at least one date.');
  if (dates.length > 31) throw new Error('Add up to 31 dates at a time.');
  const allDay = req.allDay !== false;
  const hm = v => {const m = /^([01]\d|2[0-3]):([0-5]\d)$/.exec(String(v || '')); if (!m) throw new Error('Enter start and end times.'); return [Number(m[1]), Number(m[2])];};
  let st = null, en = null;
  if (!allDay) {st = hm(req.start); en = hm(req.end); if (st[0] * 60 + st[1] >= en[0] * 60 + en[1]) throw new Error('End time must be after start time.');}
  const note = String(req.note || '').replace(/\s+/g, ' ').trim().slice(0, 500);
  const today = new Date(); today.setHours(0, 0, 0, 0);
  const limit = new Date(today.getTime()); limit.setDate(limit.getDate() + 366);
  const parsed = dates.map(d => {
    const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(d); if (!m) throw new Error('Invalid date: ' + d);
    const dt = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
    if (dt < today || dt > limit) throw new Error('Dates must be between today and one year from now: ' + d);
    return dt;
  });
  const cal = ciAvailCalendar_();
  const title = 'Available — ' + ev.evaluatorName + (allDay ? '' : ' (' + req.start + '–' + req.end + ')');
  const desc = ['Clinical instructor availability', 'Evaluator: ' + ev.evaluatorName + ' (' + ev.evaluatorId + ')',
    'Status: Pending coordinator approval', 'Note: ' + note,
    'Submitted via Clinical Instructor Portal ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'),
    ciAvailTag_(ev)].join('\n');
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  const created = [];
  try {
    parsed.forEach(dt => {
      const dayEnd = new Date(dt.getTime()); dayEnd.setDate(dayEnd.getDate() + 1);
      const dup = ciAvailMine_(ev, dt, dayEnd).some(e => e.isAllDayEvent() === allDay &&
        (allDay || (e.getStartTime().getHours() === st[0] && e.getStartTime().getMinutes() === st[1])));
      if (dup) return;
      let e;
      if (allDay) e = cal.createAllDayEvent(title, dt, {description: desc});
      else {
        const s = new Date(dt.getTime()); s.setHours(st[0], st[1], 0, 0);
        const t = new Date(dt.getTime()); t.setHours(en[0], en[1], 0, 0);
        e = cal.createEvent(title, s, t, {description: desc});
      }
      e.setColor(CalendarApp.EventColor.YELLOW);
      created.push(ciAvailOut_(e));
    });
  } finally {lock.releaseLock();}
  if (created.length) {
    // Tell the coordinator (Coordinator Dashboard bell / inbox).
    try {
      const tz = Session.getScriptTimeZone();
      const when = created.map(c => Utilities.formatDate(new Date(c.date + 'T12:00:00'), tz, 'EEE MMM d') + (c.allDay ? '' : ' ' + c.start + '–' + c.end)).join('; ');
      const row = {'Notification ID': 'NT-' + Utilities.getUuid().replace(/-/g, '').slice(0, 14).toUpperCase(), 'Created At': new Date(),
        'Recipient Type': 'Coordinator', 'Recipient': 'Clinical Coordinator', 'Type': 'Availability Submitted',
        'Item Kind': 'availability', 'Item ID': created.map(c => c.id).join(' '),
        'Title': 'Availability submitted — ' + ev.evaluatorName + ' (' + created.length + ' date' + (created.length === 1 ? '' : 's') + ')',
        'Message': when + (note ? '\nNote: ' + note : ''), 'From': ev.evaluatorName + ' (' + ev.evaluatorId + ')', 'Status': 'Unread'};
      ciNtSheet_().appendRow(CI_NT_HEADERS.map(h => row[h] == null ? '' : row[h]));
    } catch (err) {console.warn('Coordinator availability notification failed: ' + (err.message || err));}
  }
  return {success: true, created: created, skipped: parsed.length - created.length};
}
function ciR1AvailabilityRemove(token, eventId) {
  const ev = ciR1Require_(token);
  const e = ciAvailCalendar_().getEventById(String(eventId || ''));
  if (!e || String(e.getDescription() || '').indexOf(ciAvailTag_(ev)) < 0) throw new Error('Availability entry not found.');
  if (ciAvailStatus_(e) === 'Approved') throw new Error('This availability was already approved. Contact the coordinator to change it.');
  e.deleteEvent();
  return {success: true};
}
/** Run once from the editor after deploying: grants the Calendar permission and confirms access. */
function ciAvailabilityAuthorize() {
  const cal = ciAvailCalendar_();
  console.log('Availability calendar: ' + cal.getName() + ' · owned by me: ' + cal.isOwnedByMe());
  return cal.getName();
}
