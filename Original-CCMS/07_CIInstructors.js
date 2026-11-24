/**
 * 07_CIInstructors.gs — Coordinator: Clinical Instructors page.
 *
 * One card per instructor (stable Evaluator ID), assembled read-only from the
 * authoritative sheets: Evaluators (contact, hours target/max), CI Assignments
 * (site days and hours), CI Availability (mirrored from the shared calendar),
 * Clinical Tracking + Grade Records (evaluations submitted), Evaluator
 * Dosimeters (wear periods), and the mileage sheets. Plus a coverage view of
 * who is assigned where over the next two weeks. The only write here is
 * recording a new dosimeter wear period (appends a row; history is kept).
 */
function cpCi_d_(v) {
  if (v instanceof Date && !isNaN(v)) return v;
  const s = cpLink_t_(v);
  let m = /^(\d{4})-(\d{2})-(\d{2})/.exec(s);
  if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]), 12);
  m = /^(\d{1,2})\/(\d{1,2})\/(\d{2}|\d{4})$/.exec(s);
  if (m) return new Date(m[3].length === 2 ? 2000 + Number(m[3]) : Number(m[3]), Number(m[1]) - 1, Number(m[2]), 12);
  return null;
}
function cpCi_key_(d) {return d ? Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd') : '';}
function cpCi_time_(v) {return v instanceof Date ? Utilities.formatDate(v, Session.getScriptTimeZone(), 'h:mm a') : cpLink_t_(v);}

function dashboardV2Instructors(request) {
  const r = request || {};
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const term = cpMile_termRange_(r.term && /^(Spring|Summer|Fall) \d{4}$/.test(r.term) ? r.term : cpMile_terms_()[0]);
  const today = cpCi_key_(new Date());
  const in2w = cpCi_key_(new Date(Date.now() + 13 * 864e5));
  const inTerm = function(k) {return k && k >= term.start && k <= term.end;};

  // Evaluators (all types), with hours target / max when present.
  const hoursBy = {};
  cpHours_evaluators_(ss).forEach(function(e) {hoursBy[e.evaluatorId] = e;});
  const evs = appSheetToObjects_(appGetSheet_(ss, APP_CONFIG.SHEETS.EVALUATORS)).map(function(e) {
    const id = cpLink_t_(e['Evaluator ID']), h = hoursBy[id] || {};
    return {evaluatorId: id, evaluator: cpLink_t_(e['Evaluator']), type: cpLink_t_(e['Evaluator Type']), email: cpLink_t_(e['E-mail']),
      active: appIsTrue_(e['Active']), portal: appIsTrue_(e['Portal Access']), target: h.target == null ? null : h.target, max: h.max == null ? null : h.max};
  }).filter(function(e) {return e.evaluatorId;});
  const card = {};
  evs.forEach(function(e) {
    card[e.evaluatorId] = Object.assign(e, {assignedDays: 0, assignedHours: 0, untimed: 0, todaySite: '', next: null, upcoming: [],
      availPending: 0, availApproved: 0, availUpcoming: [], evalCount: 0, evalAwaiting: 0, evalFlagged: 0, lastEval: '', evaluations: [],
      pcCount: 0, dosimeter: null, dosimeters: [], mileage: null, distances: []});
  });
  const sites = {};
  appSheetToObjects_(appGetSheet_(ss, APP_CONFIG.SHEETS.CLINICAL_SITES)).forEach(function(s) {sites[cpLink_t_(s['Clinical Site ID'])] = cpLink_t_(s['Clinical Site']);});

  // Assignments (this term + coverage for the next 14 days).
  const coverage = {};
  appSheetToObjects_(appGetSheet_(ss, APP_CONFIG.SHEETS.CI_ASSIGNMENTS)).forEach(function(a) {
    if (!/^current$/i.test(cpLink_t_(a['Status']))) return;
    const k = cpCi_key_(cpCi_d_(a['Date'])), c = card[cpLink_t_(a['Evaluator ID'])];
    if (!k) return;
    const site = sites[cpLink_t_(a['Site ID'])] || cpLink_t_(a['Site ID']);
    const hrs = cpHours_span_(a['Start Time'], a['End Time']);
    const item = {date: k, site: site, start: cpCi_time_(a['Start Time']), end: cpCi_time_(a['End Time']), hours: hrs, assignmentId: cpLink_t_(a['CI Assignment ID'])};
    if (k >= today && k <= in2w) (coverage[k] = coverage[k] || []).push({evaluatorId: cpLink_t_(a['Evaluator ID']), evaluator: c ? c.evaluator : cpLink_t_(a['Evaluator ID']), site: site, start: item.start, end: item.end});
    if (!c) return;
    if (inTerm(k)) {c.assignedDays++; if (hrs == null) c.untimed++; else c.assignedHours += hrs;}
    if (k === today) c.todaySite = site;
    if (k >= today) c.upcoming.push(item);
  });

  // Availability (calendar mirror for this term; mirror refreshed when possible).
  try {
    const lock = LockService.getScriptLock();
    if (lock.tryLock(5000)) {
      try {cpHours_syncAvailability_(cpCi_d_(term.start), new Date(cpCi_d_(term.end).getTime() + 864e5));} finally {lock.releaseLock();}
    }
  } catch (e) {console.warn('Availability sync skipped: ' + (e.message || e));}
  appSheetToObjects_(appGetSheet_(ss, APP_CONFIG.SHEETS.CI_AVAILABILITY)).forEach(function(a) {
    const c = card[cpLink_t_(a['Evaluator ID'])], k = cpCi_key_(cpCi_d_(a['Date'])), st = cpLink_t_(a['Availability Status']) || 'Pending';
    if (!c || !k || st === 'Removed') return;
    if (inTerm(k)) {if (st === 'Pending') c.availPending++; if (st === 'Approved') c.availApproved++;}
    if (k >= today && st !== 'Declined') c.availUpcoming.push({date: k, status: st, hours: a['Hours'] === '' ? null : Number(a['Hours'])});
  });

  // Evaluations (Clinical Tracking, excluding historical imports) and Professional Competencies.
  appSheetToObjects_(appGetSheet_(ss, APP_CONFIG.SHEETS.CLINICAL_TRACKING || 'Clinical Tracking')).forEach(function(t) {
    const c = card[cpLink_t_(t['Evaluator ID'])];
    if (!c || appNormalize_(t['Submission Method']) === 'historical import') return;
    const k = cpCi_key_(cpCi_d_(t['Date']));
    const flagged = appIsTrue_(t['Coordinator Flag']) || /^(yes|flagged)$/i.test(cpLink_t_(t['Coordinator Flag']));
    const reviewed = appNormalize_(t['Coordinator Review Status']) === 'reviewed';
    if (inTerm(k)) {c.evalCount++; if (!reviewed) c.evalAwaiting++; if (flagged) c.evalFlagged++;}
    if (k > c.lastEval) c.lastEval = k;
    c.evaluations.push({date: k, student: cpLink_t_(t['Student']), examination: cpLink_t_(t['Examination']), type: cpLink_t_(t['Evaluation Type']),
      site: cpLink_t_(t['Clinical Site']), result: cpLink_t_(t['Result']), status: flagged ? 'Returned' : reviewed ? 'Reviewed' : 'Awaiting review',
      submissionId: cpLink_t_(t['Submission ID'])});
  });
  try {
    appSheetToObjects_(appGetSheet_(ss, 'Grade Records')).forEach(function(g) {
      if (appNormalize_(g['Grade Category']) !== 'professional competency' || !/^PC-/.test(cpLink_t_(g['Source ID']))) return;
      const m = /Evaluator ID:\s*([^\s\n]+)/.exec(cpLink_t_(g['Notes']));
      const c = m && card[m[1]];
      if (c && inTerm(cpCi_key_(cpCi_d_(g['Date'])))) c.pcCount++;
    });
  } catch (e) {console.warn('PC count skipped: ' + (e.message || e));}

  // Dosimeters (latest wear period is current; history kept).
  const dsh = ss.getSheetByName('Evaluator Dosimeters');
  if (dsh) appSheetToObjects_(dsh).forEach(function(d, i) {
    const c = card[cpLink_t_(d['Evaluator ID'])];
    if (!c) return;
    const s = cpCi_d_(d['Wear Period Start']), e = cpCi_d_(d['Wear Period End']);
    c.dosimeters.push({row: i + 2, dosimeterId: cpLink_t_(d['Dosimeter ID']), start: cpCi_key_(s), end: cpCi_key_(e), status: cpLink_t_(d['Status'])});
  });

  // Mileage (this term) and approved distances.
  try {
    cpMile_rows_(CP_MILE.SUB, CP_MILE.SUB_HEADERS).forEach(function(s) {
      const c = card[cpLink_t_(s['Evaluator ID'])];
      if (c && cpLink_t_(s['Term']) === term.term) c.mileage = {mileageId: cpLink_t_(s['Mileage ID']), status: cpLink_t_(s['Status']), totalMiles: cpMile_n_(s['Total Miles']), totalAmount: cpMile_n_(s['Total Amount'])};
    });
    cpMile_rows_(CP_MILE.DIST, CP_MILE.DIST_HEADERS).forEach(function(d) {
      const c = card[cpLink_t_(d['Evaluator ID'])];
      if (c && cpLink_t_(d['Status']) !== 'Superseded') c.distances.push({site: cpLink_t_(d['Site']), miles: cpMile_n_(d['Reimbursed Miles']), basis: cpLink_t_(d['Route Basis']), status: cpLink_t_(d['Status'])});
    });
  } catch (e) {console.warn('Mileage skipped: ' + (e.message || e));}

  // Summary reflections (CI survey responses).
  let refl = {};
  try {refl = cpRefl_byEvaluator_();} catch (e) {console.warn('Reflections skipped: ' + (e.message || e));}
  const soon = cpCi_key_(new Date(Date.now() + 14 * 864e5));
  const out = evs.map(function(e) {
    const c = card[e.evaluatorId];
    c.assignedHours = Math.round(c.assignedHours * 10) / 10;
    c.upcoming.sort(function(a, b) {return a.date.localeCompare(b.date);});
    c.next = c.upcoming.filter(function(u) {return u.date > today;})[0] || null;
    c.upcoming = c.upcoming.slice(0, 12);
    c.availUpcoming.sort(function(a, b) {return a.date.localeCompare(b.date);}); c.availUpcoming = c.availUpcoming.slice(0, 12);
    c.evaluations.sort(function(a, b) {return b.date.localeCompare(a.date);}); c.evaluations = c.evaluations.slice(0, 12);
    c.dosimeters.sort(function(a, b) {return (b.end || '').localeCompare(a.end || '') || b.row - a.row;});
    const cur = c.dosimeters[0];
    if (cur) {
      const stated = cur.status;
      const state = /return|exchang|replac|lost/i.test(stated) ? stated : !cur.end ? (stated || 'No end date') :
        cur.end < today ? 'Exchange due' : cur.end <= soon ? 'Exchange soon' : 'Current';
      c.dosimeter = {dosimeterId: cur.dosimeterId, start: cur.start, end: cur.end, state: state, status: stated};
    }
    c.dosimeters = c.dosimeters.slice(0, 6);
    c.reflections = (refl[e.evaluatorId] || []).slice(0, 6);
    c.reflection = c.reflections.filter(function(x) {return x.term === term.term;})[0] || null;
    return c;
  }).sort(function(a, b) {return (b.active ? 1 : 0) - (a.active ? 1 : 0) || a.evaluator.localeCompare(b.evaluator);});
  return {success: true, term: term.term, start: term.start, end: term.end, today: today, terms: cpMile_terms_(),
    instructors: out, coverage: Object.keys(coverage).sort().map(function(k) {return {date: k, items: coverage[k].sort(function(a, b) {return a.evaluator.localeCompare(b.evaluator);})};}),
    workbookUrl: ss.getUrl()};
}

/** Record a new dosimeter wear period for an instructor (appends; earlier periods stay as history). */
function dashboardV2RecordDosimeter(request) {
  const r = request || {};
  const id = cpLink_t_(r.evaluatorId), dos = cpLink_requireText_(r.dosimeterId, 'Dosimeter ID', 60);
  const s = cpCi_d_(r.start), e = cpCi_d_(r.end);
  if (!id) throw new Error('Choose an instructor.');
  if (!dos) throw new Error('Enter the dosimeter ID.');
  if (!s || !e || e <= s) throw new Error('Enter a wear period start and a later end date.');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const known = appSheetToObjects_(appGetSheet_(ss, APP_CONFIG.SHEETS.EVALUATORS)).some(function(x) {return cpLink_t_(x['Evaluator ID']) === id;});
  if (!known) throw new Error('Instructor not found.');
  let sh = ss.getSheetByName('Evaluator Dosimeters');
  if (!sh) {sh = ss.insertSheet('Evaluator Dosimeters'); sh.getRange(1, 1, 1, 5).setValues([['Evaluator ID', 'Dosimeter ID', 'Wear Period Start', 'Wear Period End', 'Status']]).setFontWeight('bold');}
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(cpLink_t_);
    // Close the instructor's previous period (status only) so history reads clearly.
    if (r.markPreviousReturned !== false && sh.getLastRow() > 1) {
      const vals = sh.getRange(2, 1, sh.getLastRow() - 1, head.length).getValues();
      let last = -1, lastEnd = '';
      vals.forEach(function(v, i) {
        if (cpLink_t_(v[head.indexOf('Evaluator ID')]) !== id) return;
        const k = cpCi_key_(cpCi_d_(v[head.indexOf('Wear Period End')]));
        if (k >= lastEnd) {lastEnd = k; last = i;}
      });
      if (last >= 0 && !cpLink_t_(vals[last][head.indexOf('Status')])) sh.getRange(last + 2, head.indexOf('Status') + 1).setValue('Returned');
    }
    const row = head.map(function(h) {
      return h === 'Evaluator ID' ? id : h === 'Dosimeter ID' ? dos : h === 'Wear Period Start' ? s : h === 'Wear Period End' ? e : h === 'Status' ? 'Issued' : '';
    });
    sh.appendRow(row);
  } finally {lock.releaseLock();}
  return {success: true};
}
