/**
 * 07_CIMileage.gs — Coordinator: CI mileage (one-way, shorter of home or the
 * Middletown campus to the site), reviewed here and turned into the files the
 * coordinator hands off with the voucher (the voucher itself is completed by
 * the coordinator and chair outside the system).
 *
 * No home address is stored anywhere. Instructors enter only their one-way
 * home → site miles per site (and upload their Google Maps route PDF when home
 * is the shorter route). Campus → site miles are calculated once per site with
 * Google Maps from the Clinical Sites "Address" column and stored in
 * "College Miles". The rate, campus address and label live in "Mileage
 * Settings" (shared with the CI Portal).
 *
 * On approval this builds, in a private Drive folder:
 *   - the filled MILEAGE & TOLLS sheet (address and tolls left blank), and
 *   - one route PDF per trip (campus routes are generated from Google Maps
 *     directions; home routes are copies of the instructor's uploaded PDF).
 */
const CP_MILE = {
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
  FOLDER: 'CCMS — CI Mileage (restricted)'
};
function cpMile_t_(v) {return v == null ? '' : String(v).trim();}
function cpMile_n_(v) {return v === '' || v == null || isNaN(Number(v)) ? null : Number(v);}
function cpMile_r_(n, p) {return Math.round(Number(n) * Math.pow(10, p)) / Math.pow(10, p);}
function cpMile_tz_() {return Session.getScriptTimeZone();}
function cpMile_iso_(v) {
  if (v instanceof Date && !isNaN(v)) return Utilities.formatDate(v, cpMile_tz_(), 'yyyy-MM-dd');
  const s = cpMile_t_(v);
  return /^\d{4}-\d{2}-\d{2}$/.test(s) ? s : '';
}
function cpMile_sheet_(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}
function cpMile_rows_(name, headers) {
  const sh = cpMile_sheet_(name, headers), last = sh.getLastRow();
  if (last < 2) return [];
  const values = sh.getRange(1, 1, last, headers.length).getValues();
  const head = values[0].map(String);
  return values.slice(1).map(function(r, i) {const o = {_row: i + 2}; head.forEach(function(h, j) {o[h] = r[j];}); return o;})
    .filter(function(o) {return headers.some(function(h) {return o[h] !== '' && o[h] != null;});});
}
function cpMile_write_(name, headers, obj) {
  cpMile_sheet_(name, headers).getRange(obj._row, 1, 1, headers.length).setValues([headers.map(function(h) {return obj[h] == null ? '' : obj[h];})]);
}
function cpMile_settings_() {
  const sh = cpMile_sheet_(CP_MILE.SETTINGS, ['Setting', 'Value']);
  const vals = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues() : [];
  const out = {}, rows = {};
  vals.forEach(function(r, i) {if (cpMile_t_(r[0])) {out[cpMile_t_(r[0])] = r[1]; rows[cpMile_t_(r[0])] = i + 2;}});
  Object.keys(CP_MILE.DEFAULTS).forEach(function(k) {
    if (!(k in out)) {sh.appendRow([k, CP_MILE.DEFAULTS[k]]); out[k] = CP_MILE.DEFAULTS[k]; rows[k] = sh.getLastRow();}
  });
  return {rate: Number(out['Rate per Mile']), collegeAddress: cpMile_t_(out['College Address']),
    collegeLabel: cpMile_t_(out['College Label']) || 'SUNY Orange', _rows: rows, _sheet: sh};
}
function cpMile_termRange_(term) {
  const m = /^(Spring|Summer|Fall) (\d{4})$/.exec(cpMile_t_(term));
  if (!m) throw new Error('Choose a semester.');
  const y = m[2];
  const r = {Spring: [y + '-01-01', y + '-05-31'], Summer: [y + '-06-01', y + '-08-14'], Fall: [y + '-08-15', y + '-12-31']}[m[1]];
  return {term: m[1] + ' ' + y, start: r[0], end: r[1]};
}
function cpMile_terms_() {
  const today = Utilities.formatDate(new Date(), cpMile_tz_(), 'yyyy-MM-dd');
  const md = today.slice(5), order = ['Spring', 'Summer', 'Fall'];
  let i = md <= '05-31' ? 0 : md <= '08-14' ? 1 : 2, y = Number(today.slice(0, 4));
  const out = [];
  for (let k = 0; k < 4; k++) {out.push(order[i] + ' ' + y); i--; if (i < 0) {i = 2; y--;}}
  return out;
}
/** Clinical Sites with Address / College Miles columns (created when missing). */
function cpMile_sites_() {
  const sh = appGetSheet_(SpreadsheetApp.getActiveSpreadsheet(), APP_CONFIG.SHEETS.CLINICAL_SITES);
  let head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(cpMile_t_);
  ['Address', 'College Miles'].forEach(function(h) {
    if (head.indexOf(h) < 0) {sh.getRange(1, head.length + 1).setValue(h).setFontWeight('bold'); head.push(h);}
  });
  const last = sh.getLastRow();
  const values = last > 1 ? sh.getRange(2, 1, last - 1, head.length).getValues() : [];
  const c = function(h) {return head.indexOf(h);};
  const out = {};
  values.forEach(function(r, i) {
    const id = cpMile_t_(r[c('Clinical Site ID')]);
    if (!id) return;
    out[id] = {siteId: id, site: cpMile_t_(r[c('Clinical Site')]), active: appIsTrue_(r[c('Active')]),
      address: cpMile_t_(r[c('Address')]), collegeMiles: cpMile_n_(r[c('College Miles')]), _row: i + 2};
  });
  return {sheet: sh, head: head, sites: out};
}
function cpMile_markInboxRead_(itemId) {
  try {
    const sh = cpLink_sheet_();
    cpLink_rows_().forEach(function(n) {
      if (cpLink_t_(n['Recipient Type']) === 'Coordinator' && cpLink_t_(n['Status']) === 'Unread' &&
        cpLink_t_(n['Item ID']) === itemId && /^Mileage/.test(cpLink_t_(n['Type']))) {
        sh.getRange(n._row, CP_LINK.HEADERS.indexOf('Status') + 1, 1, 2).setValues([['Read', new Date()]]);
      }
    });
  } catch (e) {console.warn('Inbox update failed: ' + (e.message || e));}
}

/* ---------------- Google Maps ---------------- */
function cpMile_route_(origin, destination) {
  const res = Maps.newDirectionFinder().setOrigin(origin).setDestination(destination)
    .setMode(Maps.DirectionFinder.Mode.DRIVING).getDirections();
  if (!res || res.status !== 'OK' || !res.routes || !res.routes.length) {
    throw new Error('Google Maps could not find a driving route to ' + destination + (res && res.status ? ' (' + res.status + ')' : ''));
  }
  const route = res.routes[0], leg = route.legs[0];
  return {miles: cpMile_r_(leg.distance.value / 1609.344, 1), distanceText: leg.distance.text, durationText: leg.duration.text,
    summary: route.summary || '', start: leg.start_address, end: leg.end_address, polyline: route.overview_polyline && route.overview_polyline.points,
    steps: (leg.steps || []).map(function(s) {
      return {text: String(s.html_instructions || '').replace(/<div[^>]*>/g, ' — ').replace(/<[^>]+>/g, '').replace(/&nbsp;/g, ' ').replace(/&amp;/g, '&'),
        distance: s.distance ? s.distance.text : ''};
    })};
}
function cpMile_mapBase64_(route) {
  try {
    const map = Maps.newStaticMap().setSize(640, 380).setMapType(Maps.StaticMap.Type.ROADMAP)
      .setMarkerStyle(Maps.StaticMap.MarkerSize.MID, Maps.StaticMap.Color.GREEN, 'A').addMarker(route.start)
      .setMarkerStyle(Maps.StaticMap.MarkerSize.MID, Maps.StaticMap.Color.RED, 'B').addMarker(route.end)
      .setPathStyle(5, '0x1A73E8CC', null);
    if (route.polyline) map.addPath(route.polyline);
    return Utilities.base64Encode(map.getBlob().getBytes());
  } catch (e) {console.warn('Static map failed: ' + (e.message || e)); return '';}
}

/* ---------------- Coordinator page data ---------------- */
function dashboardV2Mileage(request) {
  const r = request || {};
  const terms = cpMile_terms_();
  const range = cpMile_termRange_(r.term && /^(Spring|Summer|Fall) \d{4}$/.test(r.term) ? r.term : terms[0]);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settings = cpMile_settings_(), siteInfo = cpMile_sites_().sites;
  const evs = cpHours_evaluators_(ss), byId = {};
  evs.forEach(function(e) {byId[e.evaluatorId] = e;});
  const assigned = {};
  appSheetToObjects_(appGetSheet_(ss, APP_CONFIG.SHEETS.CI_ASSIGNMENTS)).forEach(function(a) {
    const d = cpMile_iso_(a['Date']);
    if (!/^current$/i.test(cpMile_t_(a['Status'])) || !d || d < range.start || d > range.end) return;
    const id = cpMile_t_(a['Evaluator ID']);
    assigned[id] = (assigned[id] || 0) + 1;
  });
  const distances = cpMile_rows_(CP_MILE.DIST, CP_MILE.DIST_HEADERS).filter(function(d) {return cpMile_t_(d['Status']) !== 'Superseded';})
    .map(function(d) {
      const pdf = cpMile_t_(d['Route PDF File ID']);
      return {distanceId: cpMile_t_(d['Distance ID']), evaluatorId: cpMile_t_(d['Evaluator ID']), evaluator: cpMile_t_(d['Evaluator']),
        siteId: cpMile_t_(d['Site ID']), site: cpMile_t_(d['Site']), homeMiles: cpMile_n_(d['Home Miles']), collegeMiles: cpMile_n_(d['College Miles']),
        reimbursedMiles: cpMile_n_(d['Reimbursed Miles']), basis: cpMile_t_(d['Route Basis']), status: cpMile_t_(d['Status']) || 'Pending',
        pdfUrl: pdf ? 'https://drive.google.com/file/d/' + pdf + '/view' : '', note: cpMile_t_(d['Coordinator Note']),
        submittedAt: cpLink_iso_(d['Submitted At'])};
    }).sort(function(a, b) {return (a.status === 'Pending' ? 0 : 1) - (b.status === 'Pending' ? 0 : 1) || a.evaluator.localeCompare(b.evaluator) || a.site.localeCompare(b.site);});
  const trips = cpMile_rows_(CP_MILE.TRIP, CP_MILE.TRIP_HEADERS);
  const subs = cpMile_rows_(CP_MILE.SUB, CP_MILE.SUB_HEADERS).filter(function(s) {return cpMile_t_(s['Term']) === range.term;})
    .map(function(s) {
      const id = cpMile_t_(s['Mileage ID']), folder = cpMile_t_(s['Output Folder ID']), sheet = cpMile_t_(s['Mileage Sheet ID']);
      return {mileageId: id, evaluatorId: cpMile_t_(s['Evaluator ID']), evaluator: cpMile_t_(s['Evaluator']), status: cpMile_t_(s['Status']),
        trips: cpMile_n_(s['Trips']), totalMiles: cpMile_n_(s['Total Miles']), rate: cpMile_n_(s['Rate']), totalAmount: cpMile_n_(s['Total Amount']),
        submittedAt: cpLink_iso_(s['Submitted At']), note: cpMile_t_(s['CI Note']), coordinatorNote: cpMile_t_(s['Coordinator Note']),
        version: cpMile_n_(s['Version']) || 1, reviewedAt: cpLink_iso_(s['Reviewed At']),
        folderUrl: folder ? 'https://drive.google.com/drive/folders/' + folder : '', sheetUrl: sheet ? 'https://docs.google.com/spreadsheets/d/' + sheet + '/edit' : '',
        tripList: trips.filter(function(t) {return cpMile_t_(t['Mileage ID']) === id;}).map(function(t) {
          return {date: cpMile_iso_(t['Trip Date']), site: cpMile_t_(t['Site']), basis: cpMile_t_(t['Route Basis']), from: cpMile_t_(t['From']),
            miles: cpMile_n_(t['Miles']), amount: cpMile_n_(t['Amount']), onSchedule: cpMile_t_(t['On Schedule']) !== 'No', distanceId: cpMile_t_(t['Distance ID'])};
        }).sort(function(a, b) {return a.date.localeCompare(b.date);})};
    }).sort(function(a, b) {return (a.status === 'Submitted' ? 0 : 1) - (b.status === 'Submitted' ? 0 : 1) || a.evaluator.localeCompare(b.evaluator);});
  const subBy = {};
  subs.forEach(function(s) {subBy[s.evaluatorId] = s;});
  return {success: true, terms: terms, term: range.term, start: range.start, end: range.end,
    settings: {rate: settings.rate, collegeAddress: settings.collegeAddress, collegeLabel: settings.collegeLabel},
    sites: Object.keys(siteInfo).map(function(k) {const s = siteInfo[k]; return {siteId: s.siteId, site: s.site, active: s.active, address: s.address, collegeMiles: s.collegeMiles};})
      .sort(function(a, b) {return (b.active ? 1 : 0) - (a.active ? 1 : 0) || a.site.localeCompare(b.site);}),
    distances: distances, submissions: subs,
    instructors: evs.map(function(e) {
      const s = subBy[e.evaluatorId];
      return {evaluatorId: e.evaluatorId, evaluator: e.evaluator, email: !!e.email, assignedDays: assigned[e.evaluatorId] || 0,
        status: s ? s.status : 'Not submitted', mileageId: s ? s.mileageId : ''};
    }).sort(function(a, b) {return a.evaluator.localeCompare(b.evaluator);})};
}

function dashboardV2MileageSaveSettings(request) {
  const r = request || {};
  const rate = Number(r.rate);
  if (!(rate > 0 && rate < 5)) throw new Error('Enter a rate per mile, e.g. 0.725.');
  const address = cpLink_requireText_(r.collegeAddress, 'Campus address', 300);
  if (!address) throw new Error('Enter the campus address.');
  const label = cpLink_requireText_(r.collegeLabel, 'Campus label', 120) || 'SUNY Orange';
  const s = cpMile_settings_();
  s._sheet.getRange(s._rows['Rate per Mile'], 2).setValue(cpMile_r_(rate, 4));
  s._sheet.getRange(s._rows['College Address'], 2).setValue(address);
  s._sheet.getRange(s._rows['College Label'], 2).setValue(label);
  return {success: true, addressChanged: address !== s.collegeAddress};
}

/** Calculate campus → site miles with Google Maps for the given sites (or all active sites with an address). */
function dashboardV2MileageCollegeMiles(request) {
  const r = request || {};
  const settings = cpMile_settings_(), info = cpMile_sites_();
  const col = info.head.indexOf('College Miles') + 1;
  const ids = Array.isArray(r.siteIds) && r.siteIds.length ? r.siteIds.map(cpMile_t_) :
    Object.keys(info.sites).filter(function(k) {return info.sites[k].active && info.sites[k].address;});
  const results = [];
  ids.forEach(function(id) {
    const s = info.sites[id];
    if (!s) return;
    if (!s.address) {results.push({siteId: id, site: s.site, error: 'No address'}); return;}
    try {
      const route = cpMile_route_(settings.collegeAddress, s.address);
      info.sheet.getRange(s._row, col).setValue(route.miles);
      results.push({siteId: id, site: s.site, miles: route.miles, previous: s.collegeMiles, via: route.summary, duration: route.durationText});
    } catch (e) {results.push({siteId: id, site: s.site, error: e.message || String(e)});}
  });
  return {success: true, results: results};
}
function dashboardV2MileageSetCollegeMiles(request) {
  const r = request || {};
  const info = cpMile_sites_(), s = info.sites[cpMile_t_(r.siteId)];
  if (!s) throw new Error('Site not found.');
  const miles = r.miles === '' || r.miles == null ? '' : cpMile_r_(Number(r.miles), 1);
  if (miles !== '' && !(miles > 0 && miles <= 300)) throw new Error('Enter miles between 0 and 300, or leave blank.');
  info.sheet.getRange(s._row, info.head.indexOf('College Miles') + 1).setValue(miles);
  return {success: true};
}

/* ---------------- Distances ---------------- */
function dashboardV2MileageDistanceDecide(request) {
  const r = request || {};
  const decision = r.decision === 'Approved' ? 'Approved' : r.decision === 'Declined' ? 'Declined' : '';
  if (!decision) throw new Error('Choose approve or decline.');
  const notes = cpLink_requireText_(r.notes, 'Note', 2000);
  if (decision === 'Declined' && !notes) throw new Error('Add a note so the instructor knows what to fix.');
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  let d;
  try {
    d = cpMile_rows_(CP_MILE.DIST, CP_MILE.DIST_HEADERS).filter(function(x) {return cpMile_t_(x['Distance ID']) === cpMile_t_(r.distanceId);})[0];
    if (!d) throw new Error('Distance not found.');
    if (cpMile_t_(d['Status']) === 'Superseded') throw new Error('The instructor has replaced this distance; refresh the page.');
    d['Status'] = decision; d['Reviewed At'] = new Date(); d['Reviewed By'] = cpLink_user_(); d['Coordinator Note'] = notes;
    cpMile_write_(CP_MILE.DIST, CP_MILE.DIST_HEADERS, d);
  } finally {lock.releaseLock();}
  const ev = cpMile_t_(d['Evaluator ID']), site = cpMile_t_(d['Site']);
  cpLink_add_({'Recipient Type': 'CI', 'Recipient ID': ev, 'Recipient': cpMile_t_(d['Evaluator']), 'Type': 'Distance ' + decision,
    'Item Kind': 'mileage-distance', 'Item ID': cpMile_t_(d['Distance ID']),
    'Title': 'Mileage distance ' + decision.toLowerCase() + ' — ' + site,
    'Message': 'Reimbursed one-way miles: ' + cpMile_n_(d['Reimbursed Miles']) + (notes ? '\n' + notes : ''), 'From': cpLink_user_()});
  cpMile_markInboxRead_(cpMile_t_(d['Distance ID']));
  let emailed = false;
  if (r.email !== false) {
    try {
      emailed = cpHours_email_(ev, 'Mileage distance ' + decision.toLowerCase() + ': ' + site,
        ['Your one-way mileage for ' + site + ' was ' + decision.toLowerCase() + ' by the clinical coordinator.',
          'Reimbursed miles per trip: ' + cpMile_n_(d['Reimbursed Miles'])].concat(notes ? ['', 'Coordinator note: ' + notes] : [])
          .concat(decision === 'Declined' ? ['', 'Please update it in the Clinical Instructor Portal (My CI Workspace → Mileage Submission).'] : []));
    } catch (e) {console.warn('Distance e-mail failed: ' + (e.message || e));}
  }
  return {success: true, emailed: emailed};
}

/* ---------------- Submissions ---------------- */
function dashboardV2MileageDecide(request) {
  const r = request || {};
  const decision = r.decision === 'Approved' ? 'Approved' : r.decision === 'Returned' ? 'Returned' : '';
  if (!decision) throw new Error('Choose approve or return.');
  const notes = cpLink_requireText_(r.notes, 'Note', 2000);
  if (decision === 'Returned' && !notes) throw new Error('Add a note so the instructor knows what to change.');
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  let sub, out = {};
  try {
    sub = cpMile_rows_(CP_MILE.SUB, CP_MILE.SUB_HEADERS).filter(function(s) {return cpMile_t_(s['Mileage ID']) === cpMile_t_(r.mileageId);})[0];
    if (!sub) throw new Error('Mileage submission not found.');
    if (cpMile_t_(sub['Status']) === 'Approved') throw new Error('This mileage is already approved.');
    if (cpMile_t_(sub['Status']) === 'Returned') throw new Error('This mileage was returned to the instructor; wait for the resubmission.');
    if (decision === 'Approved') {
      const id = cpMile_t_(sub['Mileage ID']);
      const trips = cpMile_rows_(CP_MILE.TRIP, CP_MILE.TRIP_HEADERS).filter(function(t) {return cpMile_t_(t['Mileage ID']) === id;})
        .sort(function(a, b) {return cpMile_iso_(a['Trip Date']).localeCompare(cpMile_iso_(b['Trip Date']));});
      if (!trips.length) throw new Error('This submission has no trips.');
      // Distances used by the trips are approved together with the submission.
      const used = {};
      trips.forEach(function(t) {used[cpMile_t_(t['Distance ID'])] = true;});
      cpMile_rows_(CP_MILE.DIST, CP_MILE.DIST_HEADERS).forEach(function(d) {
        if (!used[cpMile_t_(d['Distance ID'])]) return;
        if (cpMile_t_(d['Status']) === 'Declined') throw new Error('A distance used by this submission was declined (' + cpMile_t_(d['Site']) + '). Return it to the instructor instead.');
        if (cpMile_t_(d['Status']) === 'Pending' || cpMile_t_(d['Status']) === 'Superseded') {
          if (cpMile_t_(d['Status']) === 'Pending') {
            d['Status'] = 'Approved'; d['Reviewed At'] = new Date(); d['Reviewed By'] = cpLink_user_(); d['Coordinator Note'] = 'Approved with mileage ' + id;
            cpMile_write_(CP_MILE.DIST, CP_MILE.DIST_HEADERS, d);
          }
        }
        if (cpMile_t_(d['Route Basis']) === 'Home' && !cpMile_t_(d['Route PDF File ID'])) throw new Error('The home route PDF for ' + cpMile_t_(d['Site']) + ' is missing.');
        used[cpMile_t_(d['Distance ID'])] = d;
      });
      out = cpMile_generate_(sub, trips, used);
      sub['Output Folder ID'] = out.folderId; sub['Mileage Sheet ID'] = out.sheetId;
    }
    sub['Status'] = decision; sub['Reviewed At'] = new Date(); sub['Reviewed By'] = cpLink_user_(); sub['Coordinator Note'] = notes;
    cpMile_write_(CP_MILE.SUB, CP_MILE.SUB_HEADERS, sub);
  } finally {lock.releaseLock();}
  const ev = cpMile_t_(sub['Evaluator ID']), term = cpMile_t_(sub['Term']), id = cpMile_t_(sub['Mileage ID']);
  const total = '$' + Number(sub['Total Amount']).toFixed(2) + ' · ' + cpMile_n_(sub['Total Miles']) + ' mi · ' + cpMile_n_(sub['Trips']) + ' trips';
  cpLink_add_({'Recipient Type': 'CI', 'Recipient ID': ev, 'Recipient': cpMile_t_(sub['Evaluator']), 'Type': 'Mileage ' + decision,
    'Item Kind': 'mileage', 'Item ID': id, 'Title': 'Mileage ' + (decision === 'Approved' ? 'approved' : 'returned to you') + ' — ' + term,
    'Message': total + (notes ? '\n' + notes : ''), 'From': cpLink_user_()});
  cpMile_markInboxRead_(id);
  let emailed = false;
  if (r.email !== false) {
    try {
      emailed = cpHours_email_(ev, 'Your ' + term + ' mileage was ' + (decision === 'Approved' ? 'approved' : 'returned for changes'),
        (decision === 'Approved' ? ['Your ' + term + ' mileage (' + total + ') was approved and sent on for the reimbursement voucher.']
          : ['Your ' + term + ' mileage (' + total + ') was returned for changes.', '', 'Please update and resubmit it in the Clinical Instructor Portal (My CI Workspace → Mileage Submission).'])
          .concat(notes ? ['', 'Coordinator note: ' + notes] : []));
    } catch (e) {console.warn('Mileage e-mail failed: ' + (e.message || e));}
  }
  return {success: true, emailed: emailed, folderUrl: out.folderId ? 'https://drive.google.com/drive/folders/' + out.folderId : '',
    sheetUrl: out.sheetId ? 'https://docs.google.com/spreadsheets/d/' + out.sheetId + '/edit' : '', routePdfs: out.pdfs || 0, warnings: out.warnings || []};
}

/* ---------------- Output files ---------------- */
function cpMile_folder_(path) {
  const root = DriveApp.getRootFolder();
  const it = root.getFoldersByName(CP_MILE.FOLDER);
  let f = it.hasNext() ? it.next() : root.createFolder(CP_MILE.FOLDER);
  (path || []).forEach(function(name) {const s = f.getFoldersByName(name); f = s.hasNext() ? s.next() : f.createFolder(name);});
  return f;
}
function cpMile_generate_(sub, trips, dists) {
  const settings = cpMile_settings_(), tz = cpMile_tz_();
  const name = cpMile_t_(sub['Evaluator']), term = cpMile_t_(sub['Term']);
  const parent = cpMile_folder_(['Approved', term]);
  const folderName = name + ' — ' + term + ' Mileage';
  const old = parent.getFoldersByName(folderName);
  const folder = old.hasNext() ? old.next().createFolder('Revision ' + Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HHmm')) : parent.createFolder(folderName);
  const fmtDate = function(iso) {return Utilities.formatDate(new Date(iso + 'T12:00:00'), tz, 'MM/dd/yyyy');};

  // 1) MILEAGE & TOLLS sheet (same layout as the program's blank template).
  const ss = SpreadsheetApp.create(name + ' ' + term + ' Mileage');
  DriveApp.getFileById(ss.getId()).moveTo(folder);
  const sh = ss.getSheets()[0];
  sh.setName('Mileage');
  sh.getRange('A1').setValue('MILEAGE & TOLLS').setFontWeight('bold').setFontSize(14);
  sh.getRange('A3').setValue('NAME : ').setFontWeight('bold'); sh.getRange('B3').setValue(name);
  sh.getRange('D3').setValue('Address: ').setFontWeight('bold');
  sh.getRange('H3').setValue('Semester:').setFontWeight('bold'); sh.getRange('I3').setValue(term);
  sh.getRange('A5').setValue('MILEAGE').setFontWeight('bold'); sh.getRange('I5').setValue('TOLLS ').setFontWeight('bold');
  sh.getRange('A7:G7').setValues([['Date ', 'From ', 'To', 'Miles ', 'Rate ', 'Total ', 'Gmaps']]).setFontWeight('bold');
  sh.getRange('I7:J7').setValues([['Date ', 'Amt']]).setFontWeight('bold');
  const first = 9, rows = trips.map(function(t, i) {
    return [fmtDate(cpMile_iso_(t['Trip Date'])), cpMile_t_(t['From']), cpMile_t_(t['To']), cpMile_n_(t['Miles']), cpMile_n_(t['Rate']),
      '=D' + (first + i) + '*E' + (first + i), 'X'];
  });
  sh.getRange(first, 1, rows.length, 7).setValues(rows);
  const lastTrip = first + rows.length - 1;
  sh.getRange(first, 5, rows.length, 2).setNumberFormat('$0.000');
  sh.getRange(first, 6, rows.length, 1).setNumberFormat('$0.00');
  const tollRow = Math.max(27, lastTrip + 2), totRow = tollRow + 5;
  sh.getRange(tollRow, 9, 1, 2).setValues([['TOTAL $:', '=SUM(J8:J' + (tollRow - 1) + ')']]);
  sh.getRange(totRow, 3, 1, 4).setValues([['Total Miles: ', '=SUM(D' + first + ':D' + (totRow - 1) + ')', 'Reimb $:', '=SUM(F' + first + ':F' + (totRow - 1) + ')']]);
  sh.getRange(totRow, 9, 1, 2).setValues([['Total Reimb $ : ', '=SUM(F' + totRow + ',J' + tollRow + ')']]);
  sh.getRange(totRow, 3, 1, 8).setFontWeight('bold');
  sh.getRange('F' + totRow).setNumberFormat('$0.00'); sh.getRange('J' + tollRow).setNumberFormat('$0.00'); sh.getRange('J' + totRow).setNumberFormat('$0.00');
  sh.setColumnWidth(2, 230); sh.setColumnWidth(3, 170);
  sh.getRange(totRow + 2, 1).setValue('One-way miles: the shorter of home → site and ' + settings.collegeLabel + ' → site. Tolls and parking are reimbursed separately. Mileage ID ' + cpMile_t_(sub['Mileage ID']) + '.')
    .setFontStyle('italic').setFontColor('#5e7288');
  SpreadsheetApp.flush();

  // 2) One route PDF per trip.
  const routes = {}, warnings = [];
  let pdfs = 0;
  const tsh = cpMile_sheet_(CP_MILE.TRIP, CP_MILE.TRIP_HEADERS);
  const siteInfo = cpMile_sites_().sites;
  trips.forEach(function(t, i) {
    const iso = cpMile_iso_(t['Trip Date']), site = cpMile_t_(t['Site']), d = dists[cpMile_t_(t['Distance ID'])] || {};
    const fileName = String(i + 1).padStart(2, '0') + ' ' + iso + ' ' + site + ' route.pdf';
    let file = null;
    try {
      if (cpMile_t_(t['Route Basis']) === 'Home') {
        file = DriveApp.getFileById(cpMile_t_(d['Route PDF File ID'])).makeCopy(fileName, folder);
      } else {
        const s = siteInfo[cpMile_t_(t['Site ID'])] || {};
        if (!s.address) throw new Error('No address for ' + site);
        if (!routes[s.siteId]) {const rt = cpMile_route_(settings.collegeAddress, s.address); rt.map = cpMile_mapBase64_(rt); routes[s.siteId] = rt;}
        file = folder.createFile(cpMile_routePdf_(routes[s.siteId], iso, name, site, settings, cpMile_n_(t['Miles'])).setName(fileName));
      }
      pdfs++;
      tsh.getRange(t._row, CP_MILE.TRIP_HEADERS.indexOf('Route PDF ID') + 1).setValue(file.getId());
    } catch (e) {warnings.push(iso + ' ' + site + ': ' + (e.message || e));}
  });
  return {folderId: folder.getId(), sheetId: ss.getId(), pdfs: pdfs, warnings: warnings};
}
function cpMile_esc_(v) {return String(v == null ? '' : v).replace(/[&<>"]/g, function(c) {return {'&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;'}[c];});}
function cpMile_routePdf_(route, iso, instructor, site, settings, miles) {
  const tz = cpMile_tz_();
  const when = Utilities.formatDate(new Date(iso + 'T12:00:00'), tz, 'EEEE, MMMM d, yyyy');
  const html = '<html><head><style>body{font-family:Arial,Helvetica,sans-serif;font-size:11px;color:#202124;margin:24px}' +
    'h1{font-size:16px;margin:0 0 4px}h2{font-size:13px;margin:14px 0 6px}.meta td{padding:2px 10px 2px 0}.k{color:#5f6368}' +
    'ol{padding-left:18px;margin:0}li{margin:3px 0}.d{color:#5f6368}.foot{margin-top:14px;color:#5f6368;font-size:9px}</style></head><body>' +
    '<h1>' + cpMile_esc_(route.start) + ' to ' + cpMile_esc_(route.end) + '</h1>' +
    '<div><b>Drive ' + cpMile_esc_(route.distanceText) + ', ' + cpMile_esc_(route.durationText) + '</b>' + (route.summary ? ' · via ' + cpMile_esc_(route.summary) : '') + '</div>' +
    '<table class="meta" style="margin-top:8px"><tr><td class="k">Trip date</td><td><b>' + cpMile_esc_(when) + '</b></td></tr>' +
    '<tr><td class="k">Instructor</td><td>' + cpMile_esc_(instructor) + '</td></tr>' +
    '<tr><td class="k">From</td><td>' + cpMile_esc_(settings.collegeLabel) + ' — ' + cpMile_esc_(route.start) + '</td></tr>' +
    '<tr><td class="k">To</td><td>' + cpMile_esc_(site) + ' — ' + cpMile_esc_(route.end) + '</td></tr>' +
    '<tr><td class="k">One-way miles claimed</td><td><b>' + cpMile_esc_(miles) + '</b></td></tr></table>' +
    (route.map ? '<div style="margin-top:12px"><img src="data:image/png;base64,' + route.map + '" width="560"></div>' : '') +
    '<h2>Directions</h2><ol>' + route.steps.map(function(s) {return '<li>' + cpMile_esc_(s.text) + ' <span class="d">' + cpMile_esc_(s.distance) + '</span></li>';}).join('') + '</ol>' +
    '<div class="foot">Route and distance from Google Maps directions (driving), generated ' +
    cpMile_esc_(Utilities.formatDate(new Date(), tz, 'MM/dd/yyyy h:mm a')) + '. Map data © Google.</div></body></html>';
  return Utilities.newBlob(html, 'text/html', 'route.html').getAs('application/pdf');
}

/* ---------------- Reminder ---------------- */
function dashboardV2SendMileageReminder(request) {
  const r = request || {};
  const ids = (Array.isArray(r.evaluatorIds) ? r.evaluatorIds : []).map(cpLink_t_).filter(Boolean);
  if (!ids.length) throw new Error('Select at least one instructor.');
  const term = cpMile_termRange_(r.term).term;
  const extra = cpLink_requireText_(r.message, 'Message', 2000);
  const evs = cpHours_evaluators_(SpreadsheetApp.getActiveSpreadsheet()).filter(function(e) {return ids.indexOf(e.evaluatorId) >= 0;});
  let emailed = 0, notified = 0;
  const failures = [];
  evs.forEach(function(e) {
    cpLink_add_({'Recipient Type': 'CI', 'Recipient ID': e.evaluatorId, 'Recipient': e.evaluator, 'Type': 'Mileage Reminder',
      'Item Kind': 'reminder', 'Item ID': '', 'Title': 'Please submit your ' + term + ' mileage', 'Message': extra, 'From': cpLink_user_()});
    notified++;
    try {
      if (cpHours_email_(e.evaluatorId, 'Reminder: submit your ' + term + ' mileage',
        ['Hello ' + e.evaluator + ',', '', 'Please submit your ' + term + ' mileage in the Clinical Instructor Portal (My CI Workspace → Mileage Submission).',
          'Your clinical days are already listed there; tick the days you drove and submit.'].concat(extra ? ['', extra] : []))) emailed++;
      else failures.push(e.evaluator + ' (no e-mail on file)');
    } catch (err) {failures.push(e.evaluator + ' (' + (err.message || err) + ')');}
  });
  return {success: true, notified: notified, emailed: emailed, failures: failures};
}

/** Run once from the editor: confirms Google Maps and Drive access for mileage. */
function cpMileageAuthorize() {
  const s = cpMile_settings_();
  const route = cpMile_route_(s.collegeAddress, '707 East Main Street, Middletown, NY 10940');
  route.map = cpMile_mapBase64_(route);
  // Sample route page (no instructor data) so the layout can be checked once.
  const folder = cpMile_folder_();
  const old = folder.getFilesByName('Route PDF sample.pdf');
  while (old.hasNext()) old.next().setTrashed(true);
  const file = folder.createFile(cpMile_routePdf_(route, Utilities.formatDate(new Date(), cpMile_tz_(), 'yyyy-MM-dd'), 'Sample', 'Garnet Middletown', s, route.miles).setName('Route PDF sample.pdf'));
  console.log('Campus → Garnet Middletown: ' + route.miles + ' mi via ' + route.summary + ' · map ' + (route.map ? 'OK' : 'missing') + ' · sample: ' + file.getUrl());
  return route.miles;
}
