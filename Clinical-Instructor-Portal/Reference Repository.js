/**
 * CCMS read-only reference retrieval: Examinations and Clinical Sites ONLY.
 * No eligibility, grading, scheduling, active-status or other business decisions.
 * Services call appReferenceRepository_() and decide how to use factual rows.
 * Sheet headers/types/order are preserved by appSheetToObjects_().
 */
function appReferenceRepository_(context) {
  context = context || appDataContext_();
  function examinations() { return context.getObjects(APP_CONFIG.SHEETS.EXAMINATIONS); }
  function clinicalSites() { return context.getObjects(APP_CONFIG.SHEETS.CLINICAL_SITES); }
  function byId(rows, header, id) {
    const target = appNormalize_(id);
    if (!target) return null;
    return rows.find(row => appNormalize_(row[header]) === target) || null;
  }
  return Object.freeze({
    getExaminations: examinations,
    getExaminationById: function(id) { return byId(examinations(), 'Examination ID', id); },
    getClinicalSites: clinicalSites,
    getClinicalSiteById: function(id) { return byId(clinicalSites(), 'Clinical Site ID', id); }
  });
}

/**
 * Administrative invalidation, private from google.script.run.
 * After changing either source sheet, run appReferenceCacheClearAll from editor,
 * or call appReferenceCacheInvalidate_(ss, APP_CONFIG.SHEETS.EXAMINATIONS)
 * (or CLINICAL_SITES) from an authorized server maintenance path.
 * Generation rotation makes concurrent old-generation writes unreachable.
 * Existing request snapshots stay consistent; fresh executions read the Sheet.
 * Automatic fallback freshness is 300 seconds; eviction may occur earlier.
 * No triggers, deployments, source-sheet writes or business rules are added.
 */
function appReferenceCacheInvalidate_(ss, sheetName) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  if (sheetName !== APP_CONFIG.SHEETS.EXAMINATIONS && sheetName !== APP_CONFIG.SHEETS.CLINICAL_SITES) {
    throw new Error('Only Examinations and Clinical Sites have reference caches.');
  }
  const sheet = appGetSheet_(ss, sheetName);
  const props = PropertiesService.getScriptProperties();
  const versionKey = appDataContextVersionKey_(ss.getId(), sheet);
  const oldGeneration = props.getProperty(versionKey) || '0';
  const generation = Utilities.getUuid();
  props.setProperty(versionKey, generation);
  try { CacheService.getScriptCache().remove(appDataContextCacheKey_(ss.getId(), sheet, oldGeneration)); }
  catch (error) { /* Generation already invalidated; old cache expires naturally. */ }
  return {sheet: sheetName, generation: generation};
}

// Editor maintenance entry point. Private underscore methods cannot be called by UI.
function appReferenceCacheClearAll() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return [appReferenceCacheInvalidate_(ss, APP_CONFIG.SHEETS.EXAMINATIONS),
    appReferenceCacheInvalidate_(ss, APP_CONFIG.SHEETS.CLINICAL_SITES)];
}
