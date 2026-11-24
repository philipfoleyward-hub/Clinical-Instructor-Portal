/**
 * Clinical Instructor Portal — Mileage Submission (college clinical instructors).
 *
 * Privacy model (coordinator decision 2026-11-24): no home address is ever
 * stored. The instructor enters only their one-way HOME → SITE miles per site
 * (looked up in Google Maps themselves) and, when home is the shorter route,
 * uploads that Google Maps route PDF once. Reimbursed miles per trip =
 * the shorter of home → site and Middletown campus → site, one way, at the
 * rate in the "Mileage Settings" sheet. Tolls and parking are reimbursed
 * separately and are not handled here. The voucher is prepared by the
 * coordinator and chair outside the system.
 *
 * Shared CCMS workbook sheets (also used by the Coordinator Dashboard):
 *   Mileage Settings, CI Site Distances, Mileage Submissions, Mileage Trips.
 * Trips are pre-filled from the authoritative CI Assignments sheet.
 */
const CI_MILE = {
  SETTINGS: 'Mileage Settings',
  DIST: 'CI Site Distances',
  SUB: 'Mileage Submissions',
  TRIP: 'Mileage Trips',
  DIST_HEADERS: ['Distance ID', 'Evaluator ID', 'Evaluator', 'Site ID', 'Site', 'Home Miles', 'College Miles',
    'Reimbursed Miles', 'Route Basis', 'Route PDF File ID', 'Status', 'Submitted At', 'Reviewed At', 'Reviewed By', 'Coordinator Note'],
  SUB_HEADERS: ['Mileage ID', 'Evaluator ID', 'Evaluator', 'Term', 'Period Start', 'Period End', 'Trips', 'Total Miles',
    'Rate', 'Total Amount', 'Status', 'Submitted At', 'CI Note', 'Reviewed At', 'Reviewed By', 'Coordinator Note',
    'Output Folder ID', 'Mileage Sheet ID', 'Version'],
  TRIP_HEADERS: ['Mileage ID', 'Trip Date', 'Site ID', 'Site', 'Route Basis', 'From', 'To', 'Miles', 'Rate', 'Amount',
    'On Schedule', 'CI Assignment ID', 'Distance ID', 'Route PDF ID'],
  DEFAULTS: {'Rate per Mile': 0.725, 'College Address': '115 South Street, Middletown, NY 10940',
    'College Label': 'SUNY Orange — Middletown campus'},
  FOLDER: 'CCMS — CI Mileage (restricted)',
  MAX_PDF_BYTES: 8 * 1024 * 1024
};
const ciMileT_ = v => v == null ? '' : String(v).trim();
const ciMileNum_ = v => (v === '' || v == null || isNaN(Number(v))) ? null : Number(v);
const ciMileRound_ = (n, p) => Math.round(Number(n) * Math.pow(10, p)) / Math.pow(10, p);

function ciMileSheet_(name, headers) {
  const book = ciDataBook_();
  let sh = book.getSheetByName(name);
  if (!sh) {
    sh = book.insertSheet(name);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}
function ciMileRows_(name, headers) {
  const sh = ciMileSheet_(name, headers), last = sh.getLastRow();
  if (last < 2) return [];
  const values = sh.getRange(1, 1, last, headers.length).getValues();
  const head = values[0].map(String);
  return values.slice(1).map((r, i) => {const o = {_row: i + 2}; head.forEach((h, j) => {o[h] = r[j];}); return o;})
    .filter(o => headers.some(h => o[h] !== '' && o[h] != null));
}
function ciMileWrite_(name, headers, row, obj) {
  ciMileSheet_(name, headers).getRange(row, 1, 1, headers.length).setValues([headers.map(h => obj[h] == null ? '' : obj[h])]);
}
function ciMileAppend_(name, headers, obj) {
  ciMileSheet_(name, headers).appendRow(headers.map(h => obj[h] == null ? '' : obj[h]));
}
/** Settings sheet: one value per row (Setting | Value). Missing rows are created with defaults. */
function ciMileSettings_() {
  const sh = ciMileSheet_(CI_MILE.SETTINGS, ['Setting', 'Value']);
  const vals = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues() : [];
  const out = {};
  vals.forEach(r => {if (ciMileT_(r[0])) out[ciMileT_(r[0])] = r[1];});
  Object.keys(CI_MILE.DEFAULTS).forEach(k => {
    if (!(k in out)) {sh.appendRow([k, CI_MILE.DEFAULTS[k]]); out[k] = CI_MILE.DEFAULTS[k];}
  });
  const rate = Number(out['Rate per Mile']);
  if (!(rate > 0 && rate < 5)) throw new Error('The mileage rate in Mileage Settings is not valid.');
  return {rate: rate, collegeAddress: ciMileT_(out['College Address']), collegeLabel: ciMileT_(out['College Label']) || 'SUNY Orange'};
}

/* ---------------- Terms ---------------- */
function ciMileTermRange_(term) {
  const m = /^(Spring|Summer|Fall) (\d{4})$/.exec(ciMileT_(term));
  if (!m) throw new Error('Choose a semester.');
  const y = m[2];
  const r = {Spring: [y + '-01-01', y + '-05-31'], Summer: [y + '-06-01', y + '-08-14'], Fall: [y + '-08-15', y + '-12-31']}[m[1]];
  return {term: m[1] + ' ' + y, start: r[0], end: r[1], code: m[1][0] + y.slice(2)};
}
function ciMileTermOf_(iso) {
  const y = iso.slice(0, 4), md = iso.slice(5);
  return (md <= '05-31' ? 'Spring ' : md <= '08-14' ? 'Summer ' : 'Fall ') + y;
}
function ciMileTerms_() {
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const order = ['Spring', 'Summer', 'Fall'], cur = ciMileTermOf_(today).split(' ');
  let i = order.indexOf(cur[0]), y = Number(cur[1]);
  const out = [];
  for (let k = 0; k < 3; k++) {
    out.push(order[i] + ' ' + y);
    i--; if (i < 0) {i = 2; y--;}
  }
  return out;
}

/* ---------------- Reference data ---------------- */
function ciMileSites_() {
  const sh = ciDataBook_().getSheetByName('Clinical Sites');
  const values = sh.getDataRange().getValues(), head = values.shift().map(v => ciMileT_(v));
  const col = h => head.indexOf(h);
  const out = {};
  values.forEach(r => {
    const id = ciMileT_(r[col('Clinical Site ID')]);
    if (!id) return;
    out[id] = {siteId: id, site: ciMileT_(r[col('Clinical Site')]), active: ciDataTrue_(r[col('Active')]),
      address: col('Address') >= 0 ? ciMileT_(r[col('Address')]) : '',
      collegeMiles: col('College Miles') >= 0 ? ciMileNum_(r[col('College Miles')]) : null};
  });
  return out;
}
function ciMileAssignments_(ev, range) {
  return ciDataRows_(ciDataBook_(), 'CI Assignments', ['CI Assignment ID', 'Evaluator ID', 'Date', 'Site ID', 'Status'])
    .filter(r => ciMileT_(r['Evaluator ID']) === ev.evaluatorId && /^current$/i.test(ciMileT_(r.Status)))
    .map(r => ({assignmentId: ciMileT_(r['CI Assignment ID']), date: ciDataDate_(r.Date), siteId: ciMileT_(r['Site ID'])}))
    .filter(a => a.date && a.date >= range.start && a.date <= range.end)
    .sort((a, b) => a.date.localeCompare(b.date));
}
function ciMileMyDistances_(ev) {
  return ciMileRows_(CI_MILE.DIST, CI_MILE.DIST_HEADERS).filter(d => ciMileT_(d['Evaluator ID']) === ev.evaluatorId);
}
/** Latest non-superseded distance per site for this instructor. */
function ciMileDistanceBySite_(ev) {
  const by = {};
  ciMileMyDistances_(ev).forEach(d => {
    if (ciMileT_(d.Status) === 'Superseded') return;
    const k = ciMileT_(d['Site ID']);
    if (!by[k] || d._row > by[k]._row) by[k] = d;
  });
  return by;
}
function ciMileDistOut_(d) {
  if (!d) return null;
  return {distanceId: ciMileT_(d['Distance ID']), homeMiles: ciMileNum_(d['Home Miles']), collegeMiles: ciMileNum_(d['College Miles']),
    reimbursedMiles: ciMileNum_(d['Reimbursed Miles']), basis: ciMileT_(d['Route Basis']), hasPdf: !!ciMileT_(d['Route PDF File ID']),
    status: ciMileT_(d.Status) || 'Pending', note: ciMileT_(d['Coordinator Note'])};
}
function ciMileSubmissionFor_(ev, term) {
  const subs = ciMileRows_(CI_MILE.SUB, CI_MILE.SUB_HEADERS)
    .filter(s => ciMileT_(s['Evaluator ID']) === ev.evaluatorId && ciMileT_(s.Term) === term);
  return subs.length ? subs[subs.length - 1] : null;
}
function ciMileNotifyCoordinator_(ev, type, kind, itemId, title, message) {
  try {
    const row = {'Notification ID': 'NT-' + Utilities.getUuid().replace(/-/g, '').slice(0, 14).toUpperCase(), 'Created At': new Date(),
      'Recipient Type': 'Coordinator', 'Recipient': 'Clinical Coordinator', 'Type': type, 'Item Kind': kind, 'Item ID': itemId,
      'Title': title, 'Message': message, 'From': ev.evaluatorName + ' (' + ev.evaluatorId + ')', 'Status': 'Unread'};
    ciNtSheet_().appendRow(CI_NT_HEADERS.map(h => row[h] == null ? '' : row[h]));
  } catch (err) {console.warn('Coordinator mileage notification failed: ' + (err.message || err));}
}
function ciMileRequireCollegeCi_(ev) {
  if (ciMileT_(ev.evaluatorType).toLowerCase() !== 'college ci') throw new Error('Mileage is for college clinical instructors.');
}

/* ---------------- Portal API ---------------- */
function ciR1MileageInit(token, term, addSiteIds) {
  const ev = ciR1Require_(token);
  ciMileRequireCollegeCi_(ev);
  const terms = ciMileTerms_();
  const range = ciMileTermRange_(term && /^(Spring|Summer|Fall) \d{4}$/.test(term) ? term : terms[0]);
  const settings = ciMileSettings_(), sites = ciMileSites_(), dist = ciMileDistanceBySite_(ev);
  const assignments = ciMileAssignments_(ev, range);
  const extra = (Array.isArray(addSiteIds) ? addSiteIds : []).map(ciMileT_).filter(id => sites[id] && sites[id].active).slice(0, 20);
  const siteIds = Array.from(new Set(assignments.map(a => a.siteId).concat(Object.keys(dist), extra)));
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const sub = ciMileSubmissionFor_(ev, range.term);
  let subOut = null;
  if (sub) {
    const id = ciMileT_(sub['Mileage ID']);
    subOut = {mileageId: id, status: ciMileT_(sub.Status), submittedAt: sub['Submitted At'] instanceof Date ? sub['Submitted At'].toISOString() : ciMileT_(sub['Submitted At']),
      totalMiles: ciMileNum_(sub['Total Miles']), totalAmount: ciMileNum_(sub['Total Amount']), note: ciMileT_(sub['CI Note']),
      coordinatorNote: ciMileT_(sub['Coordinator Note']), version: ciMileNum_(sub.Version) || 1,
      trips: ciMileRows_(CI_MILE.TRIP, CI_MILE.TRIP_HEADERS).filter(t => ciMileT_(t['Mileage ID']) === id)
        .map(t => ({date: ciDataDate_(t['Trip Date']), siteId: ciMileT_(t['Site ID']), site: ciMileT_(t.Site), miles: ciMileNum_(t.Miles),
          amount: ciMileNum_(t.Amount), basis: ciMileT_(t['Route Basis']), onSchedule: ciMileT_(t['On Schedule']) !== 'No',
          assignmentId: ciMileT_(t['CI Assignment ID'])}))};
  }
  return {success: true, terms: terms, term: range.term, start: range.start, end: range.end, today: today,
    rate: settings.rate, collegeLabel: settings.collegeLabel,
    sites: siteIds.map(id => {const s = sites[id] || {siteId: id, site: id};
      return {siteId: id, site: s.site, address: s.address || '', collegeMiles: s.collegeMiles, ready: s.collegeMiles != null,
        distance: ciMileDistOut_(dist[id])};})
      .sort((a, b) => a.site.localeCompare(b.site)),
    otherSites: Object.keys(sites).map(k => sites[k]).filter(s => s.active && s.collegeMiles != null && siteIds.indexOf(s.siteId) < 0)
      .map(s => ({siteId: s.siteId, site: s.site})).sort((a, b) => a.site.localeCompare(b.site)),
    assignments: assignments.map(a => ({assignmentId: a.assignmentId, date: a.date, siteId: a.siteId, site: (sites[a.siteId] || {}).site || a.siteId, past: a.date <= today})),
    submission: subOut};
}

/** Save the instructor's one-way home → site miles (and the route PDF when home is shorter). */
function ciR1MileageSaveDistance(token, request) {
  const ev = ciR1Require_(token);
  ciMileRequireCollegeCi_(ev);
  const r = request || {};
  const site = ciMileSites_()[ciMileT_(r.siteId)];
  if (!site) throw new Error('Choose a clinical site.');
  if (site.collegeMiles == null) throw new Error('The coordinator has not set up mileage for ' + site.site + ' yet.');
  const home = ciMileRound_(Number(r.homeMiles), 1);
  if (!(home > 0 && home <= 300)) throw new Error('Enter your one-way miles from home (more than 0, up to 300).');
  const basis = home < site.collegeMiles ? 'Home' : 'College';
  const reimbursed = basis === 'Home' ? home : site.collegeMiles;
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const prev = ciMileDistanceBySite_(ev)[site.siteId];
    let pdfId = prev ? ciMileT_(prev['Route PDF File ID']) : '';
    if (r.pdfBase64) {
      const bytes = Utilities.base64Decode(String(r.pdfBase64));
      if (bytes.length > CI_MILE.MAX_PDF_BYTES) throw new Error('The PDF is larger than 8 MB.');
      if (String.fromCharCode.apply(null, bytes.slice(0, 5)) !== '%PDF-') throw new Error('That file is not a PDF.');
      const blob = Utilities.newBlob(bytes, 'application/pdf', ev.evaluatorId + ' ' + site.siteId + ' home route.pdf');
      pdfId = ciMileFolder_('Route PDFs').createFile(blob).getId();
    }
    if (basis === 'Home' && !pdfId) throw new Error('Home is your shorter route to ' + site.site + ', so attach your Google Maps route PDF (home → ' + site.site + ').');
    const unchanged = prev && ciMileNum_(prev['Home Miles']) === home && ciMileNum_(prev['College Miles']) === site.collegeMiles &&
      !r.pdfBase64 && ciMileT_(prev.Status) !== 'Declined';
    if (unchanged) return {success: true, distance: ciMileDistOut_(prev), unchanged: true};
    if (prev) {prev.Status = 'Superseded'; ciMileWrite_(CI_MILE.DIST, CI_MILE.DIST_HEADERS, prev._row, prev);}
    const row = {'Distance ID': 'DST-' + Utilities.getUuid().replace(/-/g, '').slice(0, 10).toUpperCase(), 'Evaluator ID': ev.evaluatorId,
      'Evaluator': ev.evaluatorName, 'Site ID': site.siteId, 'Site': site.site, 'Home Miles': home, 'College Miles': site.collegeMiles,
      'Reimbursed Miles': reimbursed, 'Route Basis': basis, 'Route PDF File ID': basis === 'Home' ? pdfId : '', 'Status': 'Pending',
      'Submitted At': new Date()};
    ciMileAppend_(CI_MILE.DIST, CI_MILE.DIST_HEADERS, row);
    ciMileNotifyCoordinator_(ev, 'Mileage Distance Submitted', 'mileage-distance', row['Distance ID'],
      'Mileage distance to approve — ' + ev.evaluatorName + ' · ' + site.site,
      'Home → site: ' + home + ' mi · Campus → site: ' + site.collegeMiles + ' mi · Reimbursed: ' + reimbursed + ' mi (' + basis.toLowerCase() + ' route)');
    return {success: true, distance: ciMileDistOut_(row)};
  } finally {lock.releaseLock();}
}

/** Submit (or resubmit after a return) the instructor's trips for one semester. */
function ciR1MileageSubmit(token, request) {
  const ev = ciR1Require_(token);
  ciMileRequireCollegeCi_(ev);
  const r = request || {};
  const range = ciMileTermRange_(r.term);
  const settings = ciMileSettings_(), sites = ciMileSites_();
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const note = ciMileT_(r.note).replace(/\s+/g, ' ').slice(0, 1000);
  const input = Array.isArray(r.trips) ? r.trips : [];
  if (!input.length) throw new Error('Choose at least one trip.');
  if (input.length > 120) throw new Error('Too many trips in one submission.');
  const lock = LockService.getScriptLock(); lock.waitLock(25000);
  try {
    const dist = ciMileDistanceBySite_(ev);
    const assigned = {};
    ciMileAssignments_(ev, range).forEach(a => {assigned[a.date + '|' + a.siteId] = a.assignmentId;});
    const seen = {};
    const trips = input.map(t => {
      const date = ciMileT_(t.date), siteId = ciMileT_(t.siteId);
      if (!/^\d{4}-\d{2}-\d{2}$/.test(date) || date < range.start || date > range.end) throw new Error('Trip date ' + date + ' is not in ' + range.term + '.');
      if (date > today) throw new Error('Trip date ' + date + ' is in the future.');
      if (seen[date + '|' + siteId]) throw new Error('Trip on ' + date + ' is listed twice.');
      seen[date + '|' + siteId] = true;
      const site = sites[siteId];
      if (!site) throw new Error('Unknown site on ' + date + '.');
      const d = dist[siteId];
      if (!d || ciMileT_(d.Status) === 'Declined') throw new Error('Set up your miles for ' + site.site + ' first.');
      const miles = ciMileNum_(d['Reimbursed Miles']), basis = ciMileT_(d['Route Basis']);
      return {'Trip Date': date, 'Site ID': siteId, 'Site': site.site, 'Route Basis': basis,
        'From': basis === 'Home' ? 'Home' : settings.collegeLabel, 'To': site.site, 'Miles': miles, 'Rate': settings.rate,
        'Amount': ciMileRound_(miles * settings.rate, 2), 'On Schedule': assigned[date + '|' + siteId] ? 'Yes' : 'No',
        'CI Assignment ID': assigned[date + '|' + siteId] || '', 'Distance ID': ciMileT_(d['Distance ID'])};
    }).sort((a, b) => a['Trip Date'].localeCompare(b['Trip Date']));
    const totalMiles = ciMileRound_(trips.reduce((s, t) => s + t.Miles, 0), 1);
    const totalAmount = ciMileRound_(trips.reduce((s, t) => s + t.Amount, 0), 2);
    const existing = ciMileSubmissionFor_(ev, range.term);
    if (existing && ciMileT_(existing.Status) === 'Approved') throw new Error('Your ' + range.term + ' mileage was already approved. Contact the coordinator to change it.');
    let id, version = 1;
    if (existing) {
      id = ciMileT_(existing['Mileage ID']); version = (ciMileNum_(existing.Version) || 1) + 1;
      // Replace this submission's trips (rows deleted bottom-up).
      const sh = ciMileSheet_(CI_MILE.TRIP, CI_MILE.TRIP_HEADERS);
      ciMileRows_(CI_MILE.TRIP, CI_MILE.TRIP_HEADERS).filter(t => ciMileT_(t['Mileage ID']) === id)
        .map(t => t._row).sort((a, b) => b - a).forEach(row => sh.deleteRow(row));
    } else {
      const prefix = 'MIL-' + range.code + '-';
      const n = ciMileRows_(CI_MILE.SUB, CI_MILE.SUB_HEADERS).map(s => ciMileT_(s['Mileage ID']))
        .filter(x => x.indexOf(prefix) === 0).map(x => Number(x.slice(prefix.length)) || 0).reduce((a, b) => Math.max(a, b), 0);
      id = prefix + String(n + 1).padStart(3, '0');
    }
    const sub = {'Mileage ID': id, 'Evaluator ID': ev.evaluatorId, 'Evaluator': ev.evaluatorName, 'Term': range.term,
      'Period Start': trips[0]['Trip Date'], 'Period End': trips[trips.length - 1]['Trip Date'], 'Trips': trips.length,
      'Total Miles': totalMiles, 'Rate': settings.rate, 'Total Amount': totalAmount, 'Status': 'Submitted', 'Submitted At': new Date(),
      'CI Note': note, 'Version': version};
    if (existing) {
      ['Reviewed At', 'Reviewed By', 'Coordinator Note', 'Output Folder ID', 'Mileage Sheet ID'].forEach(k => {sub[k] = existing[k];});
      ciMileWrite_(CI_MILE.SUB, CI_MILE.SUB_HEADERS, existing._row, sub);
    } else ciMileAppend_(CI_MILE.SUB, CI_MILE.SUB_HEADERS, sub);
    const tsh = ciMileSheet_(CI_MILE.TRIP, CI_MILE.TRIP_HEADERS);
    const rows = trips.map(t => {t['Mileage ID'] = id; return CI_MILE.TRIP_HEADERS.map(h => t[h] == null ? '' : t[h]);});
    tsh.getRange(tsh.getLastRow() + 1, 1, rows.length, CI_MILE.TRIP_HEADERS.length).setValues(rows);
    const off = trips.filter(t => t['On Schedule'] === 'No').length;
    ciMileNotifyCoordinator_(ev, 'Mileage Submitted', 'mileage', id,
      (version > 1 ? 'Mileage resubmitted — ' : 'Mileage submitted — ') + ev.evaluatorName + ' · ' + range.term,
      trips.length + ' trip' + (trips.length === 1 ? '' : 's') + ' · ' + totalMiles + ' mi · $' + totalAmount.toFixed(2) +
      (off ? '\n' + off + ' trip' + (off === 1 ? ' is' : 's are') + ' not on the CI schedule' : '') + (note ? '\nNote: ' + note : ''));
    return {success: true, mileageId: id, version: version, trips: trips.length, totalMiles: totalMiles, totalAmount: totalAmount};
  } finally {lock.releaseLock();}
}

/* ---------------- Drive (owner account; folder is private) ---------------- */
function ciMileFolder_(sub) {
  const root = DriveApp.getRootFolder();
  const it = root.getFoldersByName(CI_MILE.FOLDER);
  const base = it.hasNext() ? it.next() : root.createFolder(CI_MILE.FOLDER);
  if (!sub) return base;
  const s = base.getFoldersByName(sub);
  return s.hasNext() ? s.next() : base.createFolder(sub);
}
