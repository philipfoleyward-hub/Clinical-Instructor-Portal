/**
 * Clinical Resources (READ ONLY): Program Handbook, Clinical Syllabus and the
 * Evaluation & Scoring Guide.
 *
 * Documents stay in the coordinator's Drive. The Portal (running as the owner)
 * reads them on request, so no sharing change is needed and nothing is stored
 * on instructors' devices. Word files are shown through a hidden Google Docs
 * copy that is refreshed only when the source file changes.
 *
 * To point a resource at a different file, set the Script Property named in
 * `prop` to the new Drive file ID (no code change).
 */
const CI_RESOURCE_DOCS_ = Object.freeze({
  handbook: {title: 'Program Handbook', prop: 'CI_PORTAL_RESOURCE_HANDBOOK_ID', fallback: '1umkfabK1TAtTS9MfnDtuAHUgd5rO9iar'},
  syllabus: {title: 'Clinical Syllabus', prop: 'CI_PORTAL_RESOURCE_SYLLABUS_ID', fallback: '15mehzMWHVZqXwbKSfnq3oBbytujC87jY'}
});
const CI_RESOURCE_GDOC_ = 'application/vnd.google-apps.document';

function ciResourceDef_(key) {
  const def = CI_RESOURCE_DOCS_[String(key || '')];
  if (!def) throw new Error('Unknown resource');
  const id = String(PropertiesService.getScriptProperties().getProperty(def.prop) || def.fallback).trim();
  return {key: String(key), title: def.title, sourceId: id};
}

function ciResourceDrive_(method, url, payload) {
  const res = UrlFetchApp.fetch(url, {
    method: method, contentType: 'application/json', muteHttpExceptions: true,
    payload: payload ? JSON.stringify(payload) : undefined,
    headers: {Authorization: 'Bearer ' + ScriptApp.getOAuthToken()}
  });
  if (res.getResponseCode() >= 300) throw new Error('Drive request failed (' + res.getResponseCode() + ')');
  return res;
}

function ciResourceCacheFolderId_() {
  const props = PropertiesService.getScriptProperties();
  const existing = props.getProperty('CI_PORTAL_RESOURCE_CACHE_FOLDER_ID');
  if (existing) {
    try { DriveApp.getFolderById(existing); return existing; } catch (e) {}
  }
  const folder = DriveApp.createFolder('CCMS Portal — Resource viewer copies (automatic, safe to ignore)');
  props.setProperty('CI_PORTAL_RESOURCE_CACHE_FOLDER_ID', folder.getId());
  return folder.getId();
}

/** Google Docs ID to export from; Word files get a cached converted copy. */
function ciResourceDocId_(def) {
  const file = DriveApp.getFileById(def.sourceId);
  const updated = file.getLastUpdated().getTime();
  if (file.getMimeType() === CI_RESOURCE_GDOC_) return {docId: def.sourceId, updated: updated, name: file.getName()};
  const props = PropertiesService.getScriptProperties();
  const key = 'CI_PORTAL_RESOURCE_COPY_' + def.key;
  let cached = null;
  try { cached = JSON.parse(props.getProperty(key) || 'null'); } catch (e) {}
  if (cached && cached.sourceId === def.sourceId && cached.updated === updated) {
    try { DriveApp.getFileById(cached.docId); return {docId: cached.docId, updated: updated, name: file.getName()}; } catch (e) {}
  }
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    const copy = JSON.parse(ciResourceDrive_('post',
      'https://www.googleapis.com/drive/v3/files/' + encodeURIComponent(def.sourceId) + '/copy?fields=id',
      {name: def.title + ' (Portal viewer copy)', mimeType: CI_RESOURCE_GDOC_, parents: [ciResourceCacheFolderId_()]}).getContentText());
    if (cached && cached.docId && cached.docId !== copy.id) {
      try { DriveApp.getFileById(cached.docId).setTrashed(true); } catch (e) {}
    }
    props.setProperty(key, JSON.stringify({sourceId: def.sourceId, updated: updated, docId: copy.id}));
    CacheService.getScriptCache().remove('ciRes:' + def.key + ':n');
    return {docId: copy.id, updated: updated, name: file.getName()};
  } finally { lock.releaseLock(); }
}

function ciResourceSanitize_(html) {
  return String(html)
    .replace(/<script[\s\S]*?<\/script>/gi, '')
    .replace(/\son[a-z]+\s*=\s*("[^"]*"|'[^']*'|[^\s>]+)/gi, '')
    .replace(/javascript:/gi, '');
}

/** Cached HTML (chunked: CacheService items are limited to 100 KB). */
function ciResourceCachedHtml_(cacheKey, stamp) {
  const cache = CacheService.getScriptCache();
  try {
    const meta = JSON.parse(cache.get(cacheKey + ':n') || 'null');
    if (!meta || meta.stamp !== stamp) return null;
    const keys = []; for (let i = 0; i < meta.n; i++) keys.push(cacheKey + ':' + i);
    const parts = cache.getAll(keys);
    if (keys.some(k => typeof parts[k] !== 'string')) return null;
    return keys.map(k => parts[k]).join('');
  } catch (e) { return null; }
}
function ciResourceCacheHtml_(cacheKey, stamp, html) {
  try {
    const size = 90000, chunks = {};
    let n = 0; for (let i = 0; i < html.length; i += size) chunks[cacheKey + ':' + (n++)] = html.slice(i, i + size);
    if (n > 40) return;
    const cache = CacheService.getScriptCache();
    cache.putAll(chunks, 21600); cache.put(cacheKey + ':n', JSON.stringify({stamp: stamp, n: n}), 21600);
  } catch (e) {}
}

function ciR1ResourceDocument(token, key) {
  ciR1Require_(token);
  const def = ciResourceDef_(key);
  const doc = ciResourceDocId_(def);
  const cacheKey = 'ciRes:' + def.key;
  let html = ciResourceCachedHtml_(cacheKey, doc.updated);
  if (html === null) {
    html = ciResourceSanitize_(ciResourceDrive_('get',
      'https://www.googleapis.com/drive/v3/files/' + encodeURIComponent(doc.docId) + '/export?mimeType=text%2Fhtml').getContentText());
    ciResourceCacheHtml_(cacheKey, doc.updated, html);
  }
  return {success: true, key: def.key, title: def.title, fileName: doc.name,
    updated: Utilities.formatDate(new Date(doc.updated), Session.getScriptTimeZone(), 'MMM d, yyyy'), html: html};
}

function ciR1ResourcePdf(token, key) {
  ciR1Require_(token);
  const def = ciResourceDef_(key);
  const doc = ciResourceDocId_(def);
  const blob = ciResourceDrive_('get',
    'https://www.googleapis.com/drive/v3/files/' + encodeURIComponent(doc.docId) + '/export?mimeType=application%2Fpdf').getBlob();
  const name = String(doc.name || def.title).replace(/\.(docx?|pdf)$/i, '') + '.pdf';
  return {success: true, fileName: name, base64: Utilities.base64Encode(blob.getBytes())};
}

/**
 * Evaluation & Scoring Guide: built from the SAME authoritative sources the
 * grading engine uses (Scoring Weights, Image Scoring Rules, official Form),
 * so it cannot drift from how grades are calculated.
 */
function ciR1ScoringGuide(token) {
  ciR1Require_(token);
  const cache = CacheService.getScriptCache();
  const hit = cache.get('ciScoringGuide:v1');
  if (hit) { try { return JSON.parse(hit); } catch (e) {} }
  const book = ciDataBook_();
  const weights = ciPortalAcademic_loadPerformanceWeights_(book);
  const penalties = ciPortalAcademic_loadImagePenalties_(book);
  const items = FormApp.openById(ciPortalOfficialFormId_()).getItems();
  const norm = v => String(v || '').trim().toLowerCase().replace(/\s+/g, ' ');
  const T = FormApp.ItemType;

  // Performance sections in official Form order, with criterion weights.
  const sections = [];
  items.filter(i => i.getType() === T.GRID).forEach(item => {
    const grid = item.asGridItem(), title = item.getTitle();
    const rows = grid.getRows().map(criterion => {
      const w = ciPortalAcademic_findWeight_(weights, title, ciPortalAcademic_criterionKey_(criterion));
      return {criterion: criterion, weight: w ? w.weight : null};
    });
    sections.push({title: title, help: item.getHelpText() || '', scale: grid.getColumns(), criteria: rows});
  });

  // Image evaluation answer descriptions from the first image view.
  const image = [];
  const start = items.findIndex(i => /^image evaluation\s*-\s*view\s*1$/i.test(String(i.getTitle()).trim()));
  if (start >= 0) {
    for (let i = start + 1; i < items.length; i++) {
      const t = items[i].getType();
      if (t === T.PAGE_BREAK || t === T.SECTION_HEADER) break;
      const title = norm(items[i].getTitle());
      if (['positioning', 'equipment', 'radiation protection', 'exposure', 'marker placement', 'marker evaluation'].includes(title) &&
          (t === T.MULTIPLE_CHOICE || t === T.LIST)) {
        const q = t === T.LIST ? items[i].asListItem() : items[i].asMultipleChoiceItem();
        image.push({title: items[i].getTitle().trim(), choices: q.getChoices().map(c => c.getValue())});
      }
    }
  }

  const result = {
    success: true,
    sections: sections,
    image: image,
    penalties: {
      zero: penalties.zeroPenalty, missingMarker: penalties.missingMarkerPenalty,
      wrongMarker: penalties.wrongMarkerPenalty, illegibleMarker: penalties.illegibleMarkerPenalty
    },
    // Fixed rules of the authoritative grading engine (04_ClinicalEvaluationGradingEngine).
    rules: {passEach: 2.0, scaleMax: 3}
  };
  try { cache.put('ciScoringGuide:v1', JSON.stringify(result), 600); } catch (e) {}
  return result;
}

/** One-off owner check from the editor (no sign-in token). Safe to delete. */
function ciResourceSelfTest() {
  const out = [];
  ['handbook', 'syllabus'].forEach(k => {
    const def = ciResourceDef_(k), doc = ciResourceDocId_(def);
    const html = ciResourceDrive_('get', 'https://www.googleapis.com/drive/v3/files/' + encodeURIComponent(doc.docId) + '/export?mimeType=text%2Fhtml').getContentText();
    const pdf = ciResourceDrive_('get', 'https://www.googleapis.com/drive/v3/files/' + encodeURIComponent(doc.docId) + '/export?mimeType=application%2Fpdf').getBlob();
    out.push(k + ': html ' + html.length + ' chars, pdf ' + pdf.getBytes().length + ' bytes, source ' + doc.name);
  });
  const book = ciDataBook_(), p = ciPortalAcademic_loadImagePenalties_(book);
  out.push('penalties ' + JSON.stringify(p));
  out.push('weights rows ' + ciPortalAcademic_loadPerformanceWeights_(book).length);
  Logger.log(out.join('\n'));
  return out;
}
