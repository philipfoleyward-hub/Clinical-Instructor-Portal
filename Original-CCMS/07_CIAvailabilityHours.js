/**
 * 07_CIAvailabilityHours.gs — Coordinator: CI availability, assignments, hours.
 *
 * Simple model (coordinator decision 2026-11-24):
 *   - A CI says "I am available on these dates for N hours" (shared calendar,
 *     written by the CI Portal). Coordinator approves / declines with a note.
 *   - Approved availability can be assigned to a site for that day, which
 *     writes the authoritative CI Assignments row (feeds the Portal's
 *     "Today's Clinical Site").
 *   - Hours: approved-available vs. assigned hours per CI over a period,
 *     compared with optional per-CI target / maximum (Evaluators columns).
 * The calendar is mirrored into the CI Availability sheet on every load.
 */
const CP_HOURS = {
  TARGET_COL: 'Semester Hours Target',
  MAX_COL: 'Semester Hours Max',
  AVAIL_EXTRA: ['Hours', 'Notes'],
  PORTAL_URL_DEFAULT: 'https://script.google.com/macros/s/AKfycbyBKBxKHozIu9zDq0euDLbwvvCjamp9v0TWuiGQE8j2bv92Fbvc_LKbDvVLo4bbzyPPQQ/exec'
};
function cpHours_portalUrl_() {
  return String(PropertiesService.getScriptProperties().getProperty('CI_PORTAL_URL') || CP_HOURS.PORTAL_URL_DEFAULT);
}
function cpHours_d_(v) {
  const d = v instanceof Date ? v : new Date(String(v || '') + (/^\d{4}-\d{2}-\d{2}$/.test(String(v || '')) ? 'T12:00:00' : ''));
  return isNaN(d) ? null : d;
}
function cpHours_key_(d) {return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');}
function cpHours_minutes_(v) {
  const m = typeof scheduling_ciTimeValue_ === 'function' ? scheduling_ciTimeValue_(v) : NaN;
  return Number.isFinite(m) && m !== Number.MAX_SAFE_INTEGER ? m : null;
}
function cpHours_span_(a, b) {
  const s = cpHours_minutes_(a), e = cpHours_minutes_(b);
  return s != null && e != null && e > s ? Math.round((e - s) / 6) / 10 : null;
}
/** Hours declared on a calendar entry: "Hours: N" line, else the time window, else null. */
function cpHours_eventHours_(e) {
  const m = /^Hours:\s*([\d.]+)/m.exec(String(e.getDescription() || ''));
  if (m) return Number(m[1]);
  if (!e.isAllDayEvent()) return Math.round((e.getEndTime() - e.getStartTime()) / 360000) / 10;
  return null;
}
function cpHours_evaluators_(ss) {
  const sh = appGetSheet_(ss, APP_CONFIG.SHEETS.EVALUATORS);
  return appSheetToObjects_(sh).filter(function(r) {
    return appIsTrue_(r['Active']) && appNormalize_(r['Evaluator Type']) === 'college ci';
  }).map(function(r) {
    const n = function(v) {return v === '' || v == null || isNaN(Number(v)) ? null : Number(v);};
    return {evaluatorId: cpLink_t_(r['Evaluator ID']), evaluator: cpLink_t_(r['Evaluator']), email: cpLink_t_(r['E-mail']),
      portal: appIsTrue_(r['Portal Access']), target: n(r[CP_HOURS.TARGET_COL]), max: n(r[CP_HOURS.MAX_COL])};
  }).filter(function(x) {return x.evaluatorId;});
}

/* ---------------- Calendar → CI Availability mirror ---------------- */
function cpHours_syncAvailability_(from, to) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = appGetSheet_(ss, APP_CONFIG.SHEETS.CI_AVAILABILITY);
  let head = sh.getRange(1, 1, 1, Math.max(1, sh.getLastColumn())).getValues()[0].map(String);
  CP_HOURS.AVAIL_EXTRA.forEach(function(h) {
    if (head.indexOf(h) < 0) {sh.getRange(1, head.length + 1).setValue(h).setFontWeight('bold'); head.push(h);}
  });
  const col = function(h) {return head.indexOf(h);};
  const last = sh.getLastRow();
  const rows = last > 1 ? sh.getRange(2, 1, last - 1, head.length).getValues() : [];
  const byEvent = {};
  rows.forEach(function(r, i) {const id = cpLink_t_(r[col('Calendar Event ID')]); if (id) byEvent[id] = i;});
  const events = cpAvail_calendar_().getEvents(from, to);
  const seen = {}, now = new Date(), out = [];
  events.forEach(function(e) {
    const d = String(e.getDescription() || ''), tag = /\[CI:([^\]]+)\]/.exec(d);
    if (!tag) return;
    const id = e.getId(); seen[id] = true;
    const all = e.isAllDayEvent(), tz = Session.getScriptTimeZone();
    const start = all ? '' : Utilities.formatDate(e.getStartTime(), tz, 'h:mm a');
    const end = all ? '' : Utilities.formatDate(e.getEndTime(), tz, 'h:mm a');
    const date = all ? e.getAllDayStartDate() : e.getStartTime();
    const note = ((/^Note:\s*(.*)$/m.exec(d) || [])[1] || '').trim();
    const rec = {'Evaluator ID': tag[1].trim(), 'Date': new Date(date.getFullYear(), date.getMonth(), date.getDate()),
      'Start Time': start, 'End Time': end, 'Availability Status': cpAvail_status_(e), 'Calendar Event ID': id,
      'Last Synced': now, 'Hours': cpHours_eventHours_(e), 'Notes': note};
    const i = byEvent[id];
    if (i == null) {
      rec['Availability ID'] = 'AV-' + Utilities.getUuid().replace(/-/g, '').slice(0, 10).toUpperCase();
      out.push(head.map(function(h) {return rec[h] == null ? '' : rec[h];}));
    } else {
      Object.keys(rec).forEach(function(h) {if (col(h) >= 0) rows[i][col(h)] = rec[h] == null ? '' : rec[h];});
    }
  });
  // Entries deleted from the calendar inside the window → Removed.
  rows.forEach(function(r) {
    const id = cpLink_t_(r[col('Calendar Event ID')]), d = cpHours_d_(r[col('Date')]);
    if (id && !seen[id] && d && d >= from && d < to && cpLink_t_(r[col('Availability Status')]) !== 'Removed') {
      r[col('Availability Status')] = 'Removed'; r[col('Last Synced')] = now;
    }
  });
  if (rows.length) sh.getRange(2, 1, rows.length, head.length).setValues(rows);
  if (out.length) sh.getRange(sh.getLastRow() + 1, 1, out.length, head.length).setValues(out);
}

/* ---------------- Page data ---------------- */
function dashboardV2CiHours(request) {
  const r = request || {};
  const from = cpHours_d_(r.from), to0 = cpHours_d_(r.to);
  if (!from || !to0 || to0 < from) throw new Error('Choose a valid date range.');
  from.setHours(0, 0, 0, 0);
  const to = new Date(to0.getFullYear(), to0.getMonth(), to0.getDate() + 1);
  if ((to - from) / 864e5 > 400) throw new Error('Choose a range of 13 months or less.');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {cpHours_syncAvailability_(from, to);} finally {lock.releaseLock();}
  const evs = cpHours_evaluators_(ss), byId = {};
  evs.forEach(function(e) {byId[e.evaluatorId] = e;});
  const sites = {};
  appSheetToObjects_(appGetSheet_(ss, APP_CONFIG.SHEETS.CLINICAL_SITES)).forEach(function(s) {
    sites[cpLink_t_(s['Clinical Site ID'])] = {siteId: cpLink_t_(s['Clinical Site ID']), site: cpLink_t_(s['Clinical Site']), active: appIsTrue_(s['Active'])};
  });
  const inRange = function(d) {return d && d >= from && d < to;};
  const assignments = appSheetToObjects_(appGetSheet_(ss, APP_CONFIG.SHEETS.CI_ASSIGNMENTS)).filter(function(a) {
    return /^current$/i.test(cpLink_t_(a['Status'])) && inRange(cpHours_d_(a['Date']));
  }).map(function(a) {
    const d = cpHours_d_(a['Date']), site = sites[cpLink_t_(a['Site ID'])] || {};
    return {assignmentId: cpLink_t_(a['CI Assignment ID']), evaluatorId: cpLink_t_(a['Evaluator ID']),
      evaluator: (byId[cpLink_t_(a['Evaluator ID'])] || {}).evaluator || cpLink_t_(a['Evaluator ID']),
      date: cpHours_key_(d), siteId: cpLink_t_(a['Site ID']), site: site.site || cpLink_t_(a['Site ID']),
      start: a['Start Time'] instanceof Date ? Utilities.formatDate(a['Start Time'], Session.getScriptTimeZone(), 'h:mm a') : cpLink_t_(a['Start Time']),
      end: a['End Time'] instanceof Date ? Utilities.formatDate(a['End Time'], Session.getScriptTimeZone(), 'h:mm a') : cpLink_t_(a['End Time']),
      hours: cpHours_span_(a['Start Time'], a['End Time']), type: cpLink_t_(a['Assignment Type']), notes: cpLink_t_(a['Notes'])};
  }).sort(function(a, b) {return a.date.localeCompare(b.date);});
  const avail = appSheetToObjects_(appGetSheet_(ss, APP_CONFIG.SHEETS.CI_AVAILABILITY)).filter(function(a) {
    return cpLink_t_(a['Availability Status']) !== 'Removed' && inRange(cpHours_d_(a['Date']));
  }).map(function(a) {
    const d = cpHours_key_(cpHours_d_(a['Date'])), id = cpLink_t_(a['Evaluator ID']);
    const asg = assignments.filter(function(x) {return x.evaluatorId === id && x.date === d;});
    return {eventId: cpLink_t_(a['Calendar Event ID']), evaluatorId: id, evaluator: (byId[id] || {}).evaluator || id,
      date: d, start: cpLink_t_(a['Start Time']), end: cpLink_t_(a['End Time']),
      hours: a['Hours'] === '' ? null : Number(a['Hours']), status: cpLink_t_(a['Availability Status']) || 'Pending',
      note: cpLink_t_(a['Notes']), assigned: asg.map(function(x) {return x.site;})};
  }).sort(function(a, b) {return (a.date + a.evaluator).localeCompare(b.date + b.evaluator);});
  const totals = evs.map(function(e) {
    const mine = avail.filter(function(a) {return a.evaluatorId === e.evaluatorId;});
    const sum = function(list) {return Math.round(list.reduce(function(s, x) {return s + (Number(x.hours) || 0);}, 0) * 10) / 10;};
    const asg = assignments.filter(function(a) {return a.evaluatorId === e.evaluatorId;});
    return {evaluatorId: e.evaluatorId, evaluator: e.evaluator, email: !!e.email, portal: e.portal, target: e.target, max: e.max,
      pendingHours: sum(mine.filter(function(a) {return a.status === 'Pending';})),
      approvedHours: sum(mine.filter(function(a) {return a.status === 'Approved';})),
      assignedHours: sum(asg), assignedDays: asg.length, untimedAssignments: asg.filter(function(a) {return a.hours == null;}).length,
      lastSubmitted: mine.length ? mine[mine.length - 1].date : ''};
  });
  return {from: cpHours_key_(from), to: cpHours_key_(to0), availability: avail, assignments: assignments, totals: totals,
    sites: Object.keys(sites).map(function(k) {return sites[k];}).filter(function(s) {return s.active;})
      .sort(function(a, b) {return a.site.localeCompare(b.site);})};
}

/* ---------------- Decisions (by calendar event) ---------------- */
function cpHours_email_(evaluatorId, subject, lines) {
  const ev = cpHours_evaluators_(SpreadsheetApp.getActiveSpreadsheet()).filter(function(e) {return e.evaluatorId === evaluatorId;})[0];
  if (!ev || !ev.email) return false;
  const body = lines.concat(['', 'Open the Clinical Instructor Portal: ' + cpHours_portalUrl_(), '', 'Radiologic Technology Program · SUNY Orange']).join('\n');
  MailApp.sendEmail({to: ev.email, subject: subject, body: body, name: 'Clinical Evaluation System'});
  return true;
}
/** Approve / decline availability entries (any instructors), one notification per instructor. */
function dashboardV2AvailabilityDecideEvents(request) {
  const r = request || {};
  const decision = r.decision === 'Declined' ? 'Declined' : r.decision === 'Approved' ? 'Approved' : '';
  if (!decision) throw new Error('Choose Approve or Decline.');
  const notes = cpLink_requireText_(r.notes, 'Notes', 2000);
  if (decision === 'Declined' && !notes) throw new Error('Add a note so the instructor knows why.');
  const ids = (Array.isArray(r.eventIds) ? r.eventIds : []).map(cpLink_t_).filter(Boolean);
  if (!ids.length) throw new Error('Select at least one entry.');
  const cal = cpAvail_calendar_(), who = cpLink_user_();
  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
  const byCi = {};
  ids.forEach(function(id) {
    const e = cpAvail_event_(cal, id); if (!e) return;
    let d = String(e.getDescription() || '');
    const tag = /\[CI:([^\]]+)\]/.exec(d); if (!tag) return;
    d = /^Status:.*$/m.test(d) ? d.replace(/^Status:.*$/m, 'Status: ' + decision + ' by ' + who + ' ' + stamp) : d + '\nStatus: ' + decision + ' by ' + who + ' ' + stamp;
    d = /^Coordinator note:.*$/m.test(d) ? d.replace(/^Coordinator note:.*$/m, 'Coordinator note: ' + notes.replace(/\n/g, ' ')) :
      d.replace(tag[0], 'Coordinator note: ' + notes.replace(/\n/g, ' ') + '\n' + tag[0]);
    e.setDescription(d);
    e.setColor(decision === 'Approved' ? CalendarApp.EventColor.GREEN : CalendarApp.EventColor.RED);
    (byCi[tag[1].trim()] = byCi[tag[1].trim()] || []).push(cpAvail_out_(e));
  });
  const ciIds = Object.keys(byCi);
  if (!ciIds.length) throw new Error('None of the selected entries are still on the calendar.');
  const evs = cpHours_evaluators_(SpreadsheetApp.getActiveSpreadsheet());
  let emailed = 0;
  ciIds.forEach(function(ci) {
    const list = byCi[ci], name = (evs.filter(function(e) {return e.evaluatorId === ci;})[0] || {}).evaluator || ci;
    const dates = list.map(function(x) {return x.label.replace(/, \d{4}/, '');}).join('; ');
    cpLink_add_({'Recipient Type': 'CI', 'Recipient ID': ci, 'Recipient': name, 'Type': 'Availability ' + decision,
      'Item Kind': 'availability', 'Item ID': list.map(function(x) {return x.id;}).join(' '),
      'Title': 'Availability ' + decision.toLowerCase() + ': ' + dates, 'Message': notes, 'From': who});
    if (r.email !== false) {
      try {
        if (cpHours_email_(ci, 'Your clinical availability was ' + decision.toLowerCase(),
          ['Hello ' + name + ',', '', 'The clinical coordinator ' + decision.toLowerCase() + ' your availability for:', dates]
            .concat(notes ? ['', 'Note from the coordinator:', notes] : []))) emailed++;
      } catch (e) {console.warn('Availability email failed for ' + ci + ': ' + (e.message || e));}
    }
  });
  // Mark matching coordinator inbox requests read.
  const sh = cpLink_sheet_();
  cpLink_rows_().forEach(function(n) {
    if (cpLink_t_(n['Recipient Type']) === 'Coordinator' && cpLink_t_(n['Item Kind']) === 'availability' && cpLink_t_(n['Status']) === 'Unread' &&
        cpLink_t_(n['Item ID']).split(/\s+/).some(function(x) {return ids.indexOf(x) >= 0;}))
      sh.getRange(n._row, CP_LINK.HEADERS.indexOf('Status') + 1, 1, 2).setValues([['Read', new Date()]]);
  });
  return {success: true, instructors: ciIds.length, entries: ids.length, emailed: emailed};
}

/* ---------------- Assign a CI to a site for a day ---------------- */
function dashboardV2AssignCi(request) {
  const r = request || {};
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const evId = cpLink_t_(r.evaluatorId), siteId = cpLink_t_(r.siteId), d = cpHours_d_(r.date);
  if (!evId || !siteId || !d) throw new Error('Instructor, site and date are required.');
  const ev = cpHours_evaluators_(ss).filter(function(e) {return e.evaluatorId === evId;})[0];
  if (!ev) throw new Error('Instructor is not an active college CI.');
  const site = appSheetToObjects_(appGetSheet_(ss, APP_CONFIG.SHEETS.CLINICAL_SITES)).filter(function(s) {
    return cpLink_t_(s['Clinical Site ID']) === siteId && appIsTrue_(s['Active']);
  })[0];
  if (!site) throw new Error('Clinical site is not active.');
  const fmt = function(t) {
    const m = /^([01]\d|2[0-3]):([0-5]\d)$/.exec(cpLink_t_(t)); if (!m) return '';
    const h = Number(m[1]); return ((h + 11) % 12 + 1) + ':' + m[2] + ' ' + (h < 12 ? 'AM' : 'PM');
  };
  const start = fmt(r.start), end = fmt(r.end);
  if ((r.start || r.end) && (!start || !end || cpHours_span_(start, end) == null)) throw new Error('Enter a valid start and end time.');
  const sh = appGetSheet_(ss, APP_CONFIG.SHEETS.CI_ASSIGNMENTS);
  const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  let id;
  try {
    const existing = appSheetToObjects_(sh);
    const dup = existing.filter(function(a) {
      return cpLink_t_(a['Evaluator ID']) === evId && cpLink_t_(a['Site ID']) === siteId && /^current$/i.test(cpLink_t_(a['Status'])) &&
        cpHours_d_(a['Date']) && cpHours_key_(cpHours_d_(a['Date'])) === cpHours_key_(d);
    })[0];
    if (dup) throw new Error(ev.evaluator + ' is already assigned to ' + cpLink_t_(site['Clinical Site']) + ' on that day.');
    const m = d.getMonth();
    const prefix = 'CIA-' + (m >= 7 ? 'F' : m >= 5 ? 'U' : 'S') + String(d.getFullYear()).slice(2) + '-';
    const max = existing.reduce(function(n, a) {
      const x = cpLink_t_(a['CI Assignment ID']); return x.indexOf(prefix) === 0 ? Math.max(n, Number(x.slice(prefix.length)) || 0) : n;
    }, 0);
    id = prefix + String(max + 1).padStart(3, '0');
    const rec = {'CI Assignment ID': id, 'Evaluator ID': evId, 'Date': new Date(d.getFullYear(), d.getMonth(), d.getDate()),
      'Site ID': siteId, 'Start Time': start, 'End Time': end, 'Assignment Type': 'Clinical Coverage', 'Status': 'Current',
      'Published At': new Date(), 'Published By': cpLink_user_(), 'Notes': cpLink_requireText_(r.notes, 'Notes', 500) || 'Assigned from Coordinator Dashboard'};
    sh.appendRow(head.map(function(h) {return rec[h] == null ? '' : rec[h];}));
  } finally {lock.releaseLock();}
  const dayLabel = Utilities.formatDate(d, Session.getScriptTimeZone(), 'EEE, MMM d');
  const siteName = cpLink_t_(site['Clinical Site']);
  cpLink_add_({'Recipient Type': 'CI', 'Recipient ID': evId, 'Recipient': ev.evaluator, 'Type': 'Site Assignment',
    'Item Kind': 'assignment', 'Item ID': id, 'Title': 'Assigned: ' + siteName + ' · ' + dayLabel + (start ? ' · ' + start + '–' + end : ''),
    'Message': cpLink_t_(r.notes), 'From': cpLink_user_()});
  let emailed = false;
  if (r.email !== false) {
    try {emailed = cpHours_email_(evId, 'Clinical site assignment: ' + siteName + ' on ' + dayLabel,
      ['Hello ' + ev.evaluator + ',', '', 'You are assigned to ' + siteName + ' on ' + dayLabel + (start ? ', ' + start + '–' + end : '') + '.']
        .concat(cpLink_t_(r.notes) ? ['', 'Note: ' + cpLink_t_(r.notes)] : []));} catch (e) {console.warn(e);}
  }
  return {success: true, assignmentId: id, emailed: emailed};
}
function dashboardV2RemoveAssignment(assignmentId) {
  const sh = appGetSheet_(SpreadsheetApp.getActiveSpreadsheet(), APP_CONFIG.SHEETS.CI_ASSIGNMENTS);
  const values = sh.getDataRange().getValues(), head = values[0].map(String);
  const i = values.findIndex(function(r, k) {return k > 0 && cpLink_t_(r[head.indexOf('CI Assignment ID')]) === cpLink_t_(assignmentId);});
  if (i < 1) throw new Error('Assignment not found.');
  sh.getRange(i + 1, head.indexOf('Status') + 1).setValue('Cancelled');
  const notes = head.indexOf('Notes');
  if (notes >= 0) sh.getRange(i + 1, notes + 1).setValue(cpLink_t_(values[i][notes]) + ' · Cancelled by ' + cpLink_user_() + ' ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'));
  return {success: true};
}

/* ---------------- Hours target / maximum ---------------- */
function dashboardV2SetCiHours(request) {
  const r = request || {}, id = cpLink_t_(r.evaluatorId);
  const num = function(v) {
    if (v === '' || v == null) return '';
    const n = Number(v); if (!Number.isFinite(n) || n < 0 || n > 2000) throw new Error('Hours must be between 0 and 2000.'); return n;
  };
  const target = num(r.target), max = num(r.max);
  if (target !== '' && max !== '' && target > max) throw new Error('Target cannot exceed the maximum.');
  const sh = appGetSheet_(SpreadsheetApp.getActiveSpreadsheet(), APP_CONFIG.SHEETS.EVALUATORS);
  let head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  [CP_HOURS.TARGET_COL, CP_HOURS.MAX_COL].forEach(function(h) {
    if (head.indexOf(h) < 0) {sh.getRange(1, head.length + 1).setValue(h).setFontWeight('bold'); head.push(h);}
  });
  const ids = sh.getRange(2, head.indexOf('Evaluator ID') + 1, Math.max(1, sh.getLastRow() - 1), 1).getValues().map(function(x) {return cpLink_t_(x[0]);});
  const i = ids.indexOf(id); if (i < 0) throw new Error('Instructor not found.');
  sh.getRange(i + 2, head.indexOf(CP_HOURS.TARGET_COL) + 1).setValue(target);
  sh.getRange(i + 2, head.indexOf(CP_HOURS.MAX_COL) + 1).setValue(max);
  return {success: true};
}

/* ---------------- Reminder to submit availability ---------------- */
function dashboardV2SendAvailabilityReminder(request) {
  const r = request || {};
  const ids = (Array.isArray(r.evaluatorIds) ? r.evaluatorIds : []).map(cpLink_t_).filter(Boolean);
  if (!ids.length) throw new Error('Select at least one instructor.');
  const period = cpLink_requireText_(r.period, 'Period', 200);
  const extra = cpLink_requireText_(r.message, 'Message', 2000);
  const evs = cpHours_evaluators_(SpreadsheetApp.getActiveSpreadsheet()).filter(function(e) {return ids.indexOf(e.evaluatorId) >= 0;});
  let emailed = 0, notified = 0;
  const failures = [];
  evs.forEach(function(e) {
    cpLink_add_({'Recipient Type': 'CI', 'Recipient ID': e.evaluatorId, 'Recipient': e.evaluator, 'Type': 'Availability Reminder',
      'Item Kind': 'reminder', 'Item ID': '', 'Title': 'Please submit your availability' + (period ? ' for ' + period : ''),
      'Message': extra, 'From': cpLink_user_()});
    notified++;
    try {
      if (cpHours_email_(e.evaluatorId, 'Reminder: submit your clinical availability' + (period ? ' for ' + period : ''),
        ['Hello ' + e.evaluator + ',', '', 'Please submit your clinical availability' + (period ? ' for ' + period : '') +
          ' in the Clinical Instructor Portal (My CI Workspace → Submit Availability).'].concat(extra ? ['', extra] : []))) emailed++;
      else failures.push(e.evaluator + ' (no e-mail on file)');
    } catch (err) {failures.push(e.evaluator + ' (' + (err.message || err) + ')');}
  });
  return {success: true, notified: notified, emailed: emailed, failures: failures, remainingDailyEmail: MailApp.getRemainingDailyQuota()};
}
