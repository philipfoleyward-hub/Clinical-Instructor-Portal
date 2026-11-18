/**
 * CCMS Portal data-access foundation. No business rules or writes.
 * Controller -> Service -> Repository -> DataContext -> optional reference cache -> Sheet.
 * appDataContext_() is the shared lazy context for ONE Apps Script execution.
 * Pass the same context through a request; do not create contexts per repository call.
 * Only Examinations and Clinical Sites may enter CacheService. All other sheets
 * are memoized in execution memory only. Existing callers are not migrated here.
 */
let appDataContextExecution_ = null;

function appDataContext_(ss) {
  if (!appDataContextExecution_) {
    appDataContextExecution_ = appDataContextCreate_(ss || SpreadsheetApp.getActiveSpreadsheet());
  } else if (ss && ss.getId() !== appDataContextExecution_.spreadsheetId) {
    throw new Error('DataContext already belongs to another spreadsheet in this execution.');
  }
  return appDataContextExecution_;
}

/** Explicit factory for isolated tests; production repositories use appDataContext_(). */
function appDataContextCreate_(ss, options) {
  if (!ss) throw new Error('Spreadsheet is required.');
  options = options || {};
  const sheets = Object.create(null);
  const datasets = Object.create(null);
  const spreadsheetId = ss.getId();
  const stats = {sheetLookups: 0, sheetReads: 0, cacheHits: 0, cacheMisses: 0, cacheWrites: 0, cacheErrors: 0};
  function getSheet(sheetName) {
    if (typeof sheetName !== 'string' || !sheetName.trim()) throw new Error('Sheet name is required.');
    sheetName = sheetName.trim();
    if (!Object.prototype.hasOwnProperty.call(sheets, sheetName)) {
      sheets[sheetName] = appGetSheet_(ss, sheetName);
      stats.sheetLookups++;
    }
    return sheets[sheetName];
  }
  function getObjects(sheetName) {
    const sheet = getSheet(sheetName);
    // Sheet ID unifies aliases that resolve to the same authoritative sheet.
    const datasetId = String(sheet.getSheetId());
    if (Object.prototype.hasOwnProperty.call(datasets, datasetId)) return datasets[datasetId];
    const stable = sheet.getName() === APP_CONFIG.SHEETS.EXAMINATIONS ||
      sheet.getName() === APP_CONFIG.SHEETS.CLINICAL_SITES;
    let cache = null, key = null;
    if (stable && !options.disableReferenceCache) {
      try {
        cache = options.cache || CacheService.getScriptCache();
        const version = options.referenceVersion ? options.referenceVersion(sheet) :
          PropertiesService.getScriptProperties().getProperty(appDataContextVersionKey_(spreadsheetId, sheet)) || '0';
        key = appDataContextCacheKey_(spreadsheetId, sheet, version);
        const stored = cache.get(key);
        if (stored !== null) {
          const parsed = JSON.parse(stored, function(k, v) {
            if (v && typeof v === 'object' && v.ccmsDate !== undefined) return new Date(v.ccmsDate);
            return v;
          });
          if (parsed.schema !== 1 || !Array.isArray(parsed.rows) ||
              parsed.rows.some(row => !row || typeof row !== 'object' || Array.isArray(row))) {
            throw new Error('Invalid reference cache payload.');
          }
          datasets[datasetId] = parsed.rows;
          stats.cacheHits++;
          return datasets[datasetId];
        }
        stats.cacheMisses++;
      } catch (error) {
        stats.cacheErrors++;
        // Cache failures are optional infrastructure failures, never empty data.
      }
    }
    const rows = appSheetToObjects_(sheet);
    stats.sheetReads++;
    datasets[datasetId] = rows;
    if (cache && key) {
      try {
        const payload = JSON.stringify({schema: 1, rows: rows}, function(k, v) {
          return this[k] instanceof Date ? {ccmsDate: this[k].getTime()} : v;
        });
        // CacheService's 100 KB item limit: leave room and fall back for large sheets.
        if (Utilities.newBlob(payload).getBytes().length <= 90000) {
          cache.put(key, payload, 300);
          stats.cacheWrites++;
        }
      } catch (error) { stats.cacheErrors++; }
    }
    return rows;
  }
  return Object.freeze({spreadsheetId: spreadsheetId, getSheet: getSheet, getObjects: getObjects,
    getStats: function() { return Object.assign({}, stats); }});
}

function appDataContextVersionKey_(spreadsheetId, sheet) {
  return 'CCMS_REFERENCE_GENERATION:' + spreadsheetId + ':' + sheet.getSheetId();
}

function appDataContextCacheKey_(spreadsheetId, sheet, generation) {
  return 'CCMS:REFERENCE:v1:' + spreadsheetId + ':' + sheet.getSheetId() + ':' + generation;
}
