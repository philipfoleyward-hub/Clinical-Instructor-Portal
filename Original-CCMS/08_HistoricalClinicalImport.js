/**
 * =====================================================================
 * CLINICAL COMPETENCY TRACKER
 * HISTORICAL GRADE IMPORT
 * 07_HistoricalGradeImport.gs
 * =====================================================================
 *
 * Moves approved staging rows from "Historical Grade Import" into the
 * authoritative Grade Records service. This module never writes Clinical
 * Tracking and does not calculate grades.
 */

const HISTORICAL_GRADE_IMPORT_CONFIG = {
  sheetName: 'Historical Grade Import',
  source: 'Historical Grade Import',
  importIdPrefix: 'HGI-',
  headers: [
    'Import ID',
    'Student ID',
    'Student',
    'Record Term',
    'Grade Category',
    'Title',
    'Evaluation Period',
    'Date',
    'Grade',
    'Instructor',
    'Notes',
    'Source Type',
    'Source Document',
    'Source File ID',
    'Ready to Import',
    'Imported',
    'Imported Grade Record ID',
    'Import Error'
  ],
  allowedCategories: [
    'Professional Competency',
    'Clinical Test',
    'Clinical Quiz',
    'Case Presentation',
    'Grade Point Reduction'
  ],
  clinicalTerms: [
    'Clinical 1',
    'Clinical 2',
    'Clinical 3',
    'Clinical 4',
    'Clinical 5'
  ]
};

// After reviewing previewIsabellaHistoricalGradeRepairs(), paste its exact
// approvalToken here, save, and run applyApprovedIsabellaHistoricalGradeRepairs().
// Leave blank until the preview has been explicitly approved.
const ISABELLA_HISTORICAL_REPAIR_APPROVAL_TOKEN = '';


/**
 * Dry-run/preflight entry point.
 *
 * The only write performed by preview is assigning a missing permanent
 * Import ID in the staging/audit sheet. It never writes Grade Records.
 */
function previewHistoricalGradeImport() {
  const context = historicalGradeImport_context_();
  const assigned = historicalGradeImport_assignMissingIds_(context);
  const preview = historicalGradeImport_buildPreview_(context);
  preview.importIdsAssigned = assigned;
  Logger.log(JSON.stringify(preview, null, 2));
  return preview;
}


/**
 * Imports only rows where Ready to Import = TRUE and Imported != TRUE.
 * Invalid rows remain unimported and receive an Import Error.
 */
function importHistoricalGradeRecords() {
  const context = historicalGradeImport_context_();
  const assigned = historicalGradeImport_assignMissingIds_(context);
  // Prevent two importer runs from processing the same staging rows. Dated
  // records still acquire the Grade Records document lock inside
  // gradeRecords_add(); undated historical rows acquire it in the adapter.
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  const result = {
    success: true,
    importIdsAssigned: assigned,
    rowsReady: 0,
    recordsCreated: 0,
    alreadyImported: 0,
    reconciled: 0,
    skipped: 0,
    errors: [],
    created: [],
    reconciledRecords: [],
    clinicalTrackingWrites: 0,
    message: ''
  };

  try {
    // Refresh after ID assignment and build the live idempotency index once.
    const rows = historicalGradeImport_readRows_(context.sheet);
    const gradeIndex = historicalGradeImport_gradeRecordIndex_(context.ss);

    rows.forEach(row => {
      if (!historicalGradeImport_isTrue_(row.values['Ready to Import'])) return;
      result.rowsReady += 1;

      const importId = historicalGradeImport_clean_(row.values['Import ID']);
      const importedRecordId = historicalGradeImport_clean_(
        row.values['Imported Grade Record ID']
      );
      const existing =
        (importedRecordId && gradeIndex.byRecordId[importedRecordId]) ||
        (importId && gradeIndex.bySourceId[importId]) ||
        null;

      if (existing) {
        historicalGradeImport_markImported_(context.sheet, context.headers,
          row.rowNumber, existing.recordId);
        result.alreadyImported += historicalGradeImport_isTrue_(
          row.values.Imported
        ) ? 1 : 0;
        result.reconciled += historicalGradeImport_isTrue_(
          row.values.Imported
        ) ? 0 : 1;
        result.reconciledRecords.push({
          row: row.rowNumber,
          importId: importId,
          gradeRecordId: existing.recordId
        });
        return;
      }

      if (historicalGradeImport_isTrue_(row.values.Imported)) {
        const error = importedRecordId
          ? 'Imported Grade Record ID was not found in Grade Records.'
          : 'Row is marked Imported but has no matching Grade Record association.';
        historicalGradeImport_setError_(context.sheet, context.headers,
          row.rowNumber, error);
        result.skipped += 1;
        result.errors.push({ row: row.rowNumber, importId: importId, error: error });
        return;
      }

      const validation = historicalGradeImport_validateRow_(context.ss, row.values);
      if (!validation.valid) {
        const error = validation.errors.join(' ');
        historicalGradeImport_setError_(context.sheet, context.headers,
          row.rowNumber, error);
        result.skipped += 1;
        result.errors.push({ row: row.rowNumber, importId: importId, error: error });
        return;
      }

      try {
        const payload = historicalGradeImport_payload_(row.values, validation);
        const record = payload.date
          ? gradeRecords_add(payload)
          : historicalGradeImport_addUndatedRecord_(payload);
        const recordId = historicalGradeImport_clean_(record.gradeRecordId);
        if (!recordId) {
          throw new Error('Grade Records service returned no Grade Record ID.');
        }
        historicalGradeImport_markImported_(context.sheet, context.headers,
          row.rowNumber, recordId);
        result.recordsCreated += 1;
        result.created.push({
          row: row.rowNumber,
          importId: importId,
          gradeRecordId: recordId,
          studentId: record.studentId,
          student: record.student,
          recordTerm: record.clinicalTerm,
          gradeCategory: record.gradeCategory,
          percentGrade: record.percentGrade,
          points: record.points
        });
        gradeIndex.byRecordId[recordId] = record;
        gradeIndex.bySourceId[importId] = record;
      } catch (error) {
        const message = historicalGradeImport_errorMessage_(error);
        historicalGradeImport_setError_(context.sheet, context.headers,
          row.rowNumber, message);
        result.skipped += 1;
        result.errors.push({ row: row.rowNumber, importId: importId, error: message });
      }
    });

    SpreadsheetApp.flush();
    result.success = result.errors.length === 0;
    result.message = result.recordsCreated + ' Grade Record(s) created; ' +
      result.reconciled + ' staging row(s) reconciled; ' +
      result.skipped + ' row(s) skipped. Clinical Tracking was not written.';
    Logger.log(JSON.stringify(result, null, 2));
    return result;
  } finally {
    lock.releaseLock();
  }
}


/** Alias retained for discoverability in the Apps Script function selector. */
function runHistoricalGradeImport() {
  return importHistoricalGradeRecords();
}


function historicalGradeImport_context_() {
  historicalGradeImport_assertDependencies_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(HISTORICAL_GRADE_IMPORT_CONFIG.sheetName);
  if (!sheet) {
    throw new Error('Missing staging sheet: ' +
      HISTORICAL_GRADE_IMPORT_CONFIG.sheetName + '.');
  }
  const headers = historicalGradeImport_headerMap_(sheet);
  const missing = HISTORICAL_GRADE_IMPORT_CONFIG.headers.filter(header =>
    !Object.prototype.hasOwnProperty.call(headers, header)
  );
  if (missing.length) {
    throw new Error('Historical Grade Import is missing header(s): ' +
      missing.join(', ') + '.');
  }
  return { ss: ss, sheet: sheet, headers: headers };
}


function historicalGradeImport_assertDependencies_() {
  if (typeof gradeRecords_add !== 'function') {
    throw new Error('gradeRecords_add() is required from GradeRecordsService.');
  }
  if (typeof gradeRecords_resolveStudent_ !== 'function') {
    throw new Error('gradeRecords_resolveStudent_() is required from GradeRecordsService.');
  }
  if (typeof gradeRecords_getSheet_ !== 'function') {
    throw new Error('gradeRecords_getSheet_() is required from GradeRecordsService.');
  }
}


function historicalGradeImport_assignMissingIds_(context) {
  const rows = historicalGradeImport_readRows_(context.sheet);
  const idColumn = context.headers['Import ID'];
  let assigned = 0;
  rows.forEach(row => {
    if (historicalGradeImport_clean_(row.values['Import ID'])) return;
    const importId = HISTORICAL_GRADE_IMPORT_CONFIG.importIdPrefix +
      Utilities.getUuid().replace(/-/g, '').slice(0, 10).toUpperCase();
    context.sheet.getRange(row.rowNumber, idColumn).setValue(importId);
    assigned += 1;
  });
  if (assigned) SpreadsheetApp.flush();
  return assigned;
}


function historicalGradeImport_buildPreview_(context) {
  const rows = historicalGradeImport_readRows_(context.sheet);
  const gradeIndex = historicalGradeImport_gradeRecordIndex_(context.ss);
  const report = {
    success: true,
    dryRun: true,
    gradeRecordsWritten: 0,
    clinicalTrackingWrites: 0,
    rowsReady: 0,
    rowsValid: 0,
    rowsInvalid: 0,
    rowsAlreadyImported: 0,
    recordsWouldBeCreated: [],
    alreadyImported: [],
    validationErrors: []
  };

  rows.forEach(row => {
    if (!historicalGradeImport_isTrue_(row.values['Ready to Import'])) return;
    report.rowsReady += 1;
    const importId = historicalGradeImport_clean_(row.values['Import ID']);
    const importedRecordId = historicalGradeImport_clean_(
      row.values['Imported Grade Record ID']
    );
    const existing =
      (importedRecordId && gradeIndex.byRecordId[importedRecordId]) ||
      (importId && gradeIndex.bySourceId[importId]) || null;

    if (existing || historicalGradeImport_isTrue_(row.values.Imported)) {
      report.rowsAlreadyImported += 1;
      report.alreadyImported.push({
        row: row.rowNumber,
        importId: importId,
        gradeRecordId: existing ? existing.recordId : importedRecordId,
        status: existing ? 'Matched existing Grade Record' :
          'Marked Imported but matching Grade Record was not found'
      });
      return;
    }

    const validation = historicalGradeImport_validateRow_(context.ss, row.values);
    if (!validation.valid) {
      report.rowsInvalid += 1;
      report.validationErrors.push({
        row: row.rowNumber,
        importId: importId,
        errors: validation.errors
      });
      return;
    }

    report.rowsValid += 1;
    const payload = historicalGradeImport_payload_(row.values, validation);
    report.recordsWouldBeCreated.push({
      row: row.rowNumber,
      importId: importId,
      studentId: validation.student.studentId,
      student: validation.student.student,
      recordTerm: validation.recordTerm,
      gradeCategory: validation.category,
      title: payload.assignmentName,
      evaluationPeriod: payload.evaluationPeriod,
      date: validation.date
        ? historicalGradeImport_formatDate_(validation.date)
        : null,
      percentGrade: validation.category === 'Grade Point Reduction'
        ? null : validation.grade,
      pointsDeducted: validation.category === 'Grade Point Reduction'
        ? validation.grade : null,
      source: HISTORICAL_GRADE_IMPORT_CONFIG.source,
      sourceId: importId,
      sourceType: payload.deliveryMethod,
      sourceDocument: payload.sourceDocument,
      sourceFileId: payload.sourceFileId,
      clinicalTrackingWrite: false
    });
  });
  return report;
}


function historicalGradeImport_validateRow_(ss, values) {
  const errors = [];
  const studentId = historicalGradeImport_clean_(values['Student ID']);
  const suppliedName = historicalGradeImport_clean_(values.Student);
  let student = null;
  if (!studentId) {
    errors.push('Student ID is required.');
  } else {
    try {
      student = gradeRecords_resolveStudent_(ss, studentId, suppliedName);
      if (suppliedName && historicalGradeImport_normalize_(suppliedName) !==
          historicalGradeImport_normalize_(student.student)) {
        errors.push('Student name does not match Student ID ' + studentId + '.');
      }
    } catch (error) {
      errors.push(historicalGradeImport_errorMessage_(error));
    }
  }

  const recordTerm = historicalGradeImport_match_(
    values['Record Term'], HISTORICAL_GRADE_IMPORT_CONFIG.clinicalTerms
  );
  if (!recordTerm) {
    errors.push('Record Term must be Clinical 1 through Clinical 5.');
  }

  const category = historicalGradeImport_match_(
    values['Grade Category'], HISTORICAL_GRADE_IMPORT_CONFIG.allowedCategories
  );
  if (!category) {
    errors.push('Grade Category must be one of: ' +
      HISTORICAL_GRADE_IMPORT_CONFIG.allowedCategories.join(', ') + '.');
  }
  if (category === 'Case Presentation' && recordTerm &&
      recordTerm !== 'Clinical 2' && recordTerm !== 'Clinical 5') {
    errors.push('Case Presentation is grade-bearing only in Clinical 2 and Clinical 5.');
  }

  const grade = historicalGradeImport_number_(values.Grade);
  if (grade === null) {
    errors.push('Grade must be numeric.');
  } else if (category === 'Grade Point Reduction') {
    if (!(grade > 0 && grade <= 100)) {
      errors.push('Grade Point Reduction must be a positive number no greater than 100.');
    }
  } else if (grade < 0 || grade > 100) {
    errors.push('Percentage Grade must be between 0 and 100.');
  }

  const suppliedDate = historicalGradeImport_clean_(values.Date);
  const date = historicalGradeImport_date_(values.Date);
  // A blank date is valid only in this Historical Grade Import adapter. If a
  // staging value is supplied, it must still parse as a real date and is
  // preserved unchanged as the resulting Date value.
  if (suppliedDate && !date) {
    errors.push('Date is present but is not a valid date.');
  }

  const importId = historicalGradeImport_clean_(values['Import ID']);
  if (!/^HGI-[A-Z0-9]{10}$/.test(importId)) {
    errors.push('Import ID must use the permanent HGI-XXXXXXXXXX format.');
  }

  return {
    valid: errors.length === 0,
    errors: errors,
    student: student,
    recordTerm: recordTerm,
    category: category,
    grade: grade,
    date: date
  };
}


function historicalGradeImport_payload_(values, validation) {
  const isDeduction = validation.category === 'Grade Point Reduction';
  return {
    studentId: validation.student.studentId,
    student: validation.student.student,
    clinicalTerm: validation.recordTerm,
    // Existing Grade Records names this category Point Deduction.
    gradeCategory: isDeduction ? 'Point Deduction' : validation.category,
    assignmentName: historicalGradeImport_clean_(values.Title) ||
      validation.recordTerm + ' ' + validation.category,
    evaluationPeriod: historicalGradeImport_clean_(values['Evaluation Period']),
    date: validation.date,
    instructor: historicalGradeImport_clean_(values.Instructor),
    percentGrade: isDeduction ? undefined : validation.grade,
    points: isDeduction ? validation.grade : undefined,
    notes: historicalGradeImport_clean_(values.Notes),
    deliveryMethod: historicalGradeImport_clean_(values['Source Type']),
    source: HISTORICAL_GRADE_IMPORT_CONFIG.source,
    sourceId: historicalGradeImport_clean_(values['Import ID']),
    sourceDocument: historicalGradeImport_clean_(values['Source Document']),
    sourceFileId: historicalGradeImport_clean_(values['Source File ID']),
    // Historical rows are discrete source records. Do not supersede a previous
    // test/case row solely because the category normally has one official row.
    replaceExisting: false
  };
}


/**
 * Historical-import-only Grade Records adapter for records whose source
 * workbook does not preserve an exact event date.
 *
 * The general gradeRecords_add() API intentionally continues to require a
 * date for Dashboard/manual entry. The Grade Records sheet itself accepts a
 * blank cell, so this narrowly scoped path reuses the existing service's
 * student/category/value validation, headers, ID convention, append helper,
 * audit fields, and public-record conversion while writing Date = ''.
 */
function historicalGradeImport_addUndatedRecord_(payload) {
  if (historicalGradeImport_normalize_(payload.source) !==
      historicalGradeImport_normalize_(HISTORICAL_GRADE_IMPORT_CONFIG.source)) {
    throw new Error('Blank dates are allowed only for Historical Grade Import.');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = gradeRecords_getOrCreateSheet_(ss);
  gradeRecords_ensureHeaders_(sheet);
  const student = gradeRecords_resolveStudent_(
    ss, payload.studentId, payload.student
  );
  const clinicalTerm = gradeRecords_validateClinicalTerm_(payload.clinicalTerm);
  const categoryKey = gradeRecords_normalizeCategory_(payload.gradeCategory);
  if (!categoryKey) throw new Error('A recognized Grade Category is required.');
  if (categoryKey === 'casePresentation' && clinicalTerm !== 'Clinical 2' &&
      clinicalTerm !== 'Clinical 5') {
    throw new Error(
      'Case Presentation is grade-bearing only in Clinical 2 and Clinical 5.'
    );
  }

  const values = gradeRecords_calculateValues_(categoryKey, payload);
  const categoryLabel = GRADE_RECORDS_CONFIG.categories[categoryKey].label;
  const assignmentName = gradeRecords_clean_(payload.assignmentName) ||
    gradeRecords_defaultAssignmentName_(categoryKey, clinicalTerm);
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);

  try {
    const now = new Date();
    const recordId = 'GR-' + Utilities.getUuid().replace(/-/g, '')
      .slice(0, 16).toUpperCase();
    const record = {
      'Grade Record ID': recordId,
      'Student ID': student.studentId,
      'Student': student.student,
      'Clinical Term': clinicalTerm,
      'Grade Category': categoryLabel,
      'Assignment / Evaluation Name': assignmentName,
      'Evaluation Period': gradeRecords_clean_(payload.evaluationPeriod),
      'Date': '',
      'Instructor': gradeRecords_clean_(payload.instructor),
      'Raw Score': values.rawScore,
      'Maximum Score': values.maximumScore,
      'Percent Grade': values.percentGrade,
      'Points': values.points,
      'Delivery Method': gradeRecords_clean_(payload.deliveryMethod),
      'Source': HISTORICAL_GRADE_IMPORT_CONFIG.source,
      'Source ID': gradeRecords_clean_(payload.sourceId),
      'Source Document': gradeRecords_clean_(payload.sourceDocument),
      'Source File ID': gradeRecords_clean_(payload.sourceFileId),
      'Entered By': gradeRecords_currentUser_(),
      'Entered At': now,
      'Notes': gradeRecords_clean_(payload.notes),
      'Active': true,
      'Supersedes Record ID': ''
    };

    appAppendByHeaders_(sheet, record);
    SpreadsheetApp.flush();
    return gradeRecords_publicRecord_(record);
  } finally {
    lock.releaseLock();
  }
}


function historicalGradeImport_gradeRecordIndex_(ss) {
  const sheet = gradeRecords_getSheet_(ss);
  const byRecordId = {};
  const bySourceId = {};
  gradeRecords_getSheetObjects_(sheet).forEach(row => {
    const recordId = historicalGradeImport_clean_(row['Grade Record ID']);
    const sourceId = historicalGradeImport_clean_(row['Source ID']);
    const source = historicalGradeImport_normalize_(row.Source);
    if (recordId) byRecordId[recordId] = historicalGradeImport_indexRecord_(row);
    if (sourceId && source === historicalGradeImport_normalize_(
      HISTORICAL_GRADE_IMPORT_CONFIG.source
    )) {
      bySourceId[sourceId] = historicalGradeImport_indexRecord_(row);
    }
  });
  return { byRecordId: byRecordId, bySourceId: bySourceId };
}


function historicalGradeImport_indexRecord_(row) {
  return {
    recordId: historicalGradeImport_clean_(row['Grade Record ID']),
    studentId: historicalGradeImport_clean_(row['Student ID']),
    student: historicalGradeImport_clean_(row.Student),
    clinicalTerm: historicalGradeImport_clean_(row['Clinical Term']),
    gradeCategory: historicalGradeImport_clean_(row['Grade Category']),
    percentGrade: historicalGradeImport_number_(row['Percent Grade']),
    points: historicalGradeImport_number_(row.Points)
  };
}


function historicalGradeImport_readRows_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  if (lastRow < 2 || lastColumn < 1) return [];
  const values = sheet.getRange(1, 1, lastRow, lastColumn).getValues();
  const headers = values[0].map(h => historicalGradeImport_clean_(h));
  return values.slice(1).map((row, index) => {
    const object = {};
    headers.forEach((header, column) => { if (header) object[header] = row[column]; });
    return { rowNumber: index + 2, values: object };
  }).filter(row => Object.keys(row.values).some(key =>
    historicalGradeImport_clean_(row.values[key]) !== ''
  ));
}


function historicalGradeImport_headerMap_(sheet) {
  const lastColumn = sheet.getLastColumn();
  if (!lastColumn) return {};
  const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  const map = {};
  headers.forEach((header, index) => {
    const key = historicalGradeImport_clean_(header);
    if (key) map[key] = index + 1;
  });
  return map;
}


function historicalGradeImport_markImported_(sheet, headers, row, recordId) {
  sheet.getRange(row, headers.Imported).setValue(true);
  sheet.getRange(row, headers['Imported Grade Record ID']).setValue(recordId);
  sheet.getRange(row, headers['Import Error']).clearContent();
}


function historicalGradeImport_setError_(sheet, headers, row, error) {
  sheet.getRange(row, headers.Imported).setValue(false);
  sheet.getRange(row, headers['Import Error']).setValue(error);
}


function historicalGradeImport_match_(value, allowed) {
  const target = historicalGradeImport_normalize_(value);
  return allowed.find(item => historicalGradeImport_normalize_(item) === target) || '';
}


function historicalGradeImport_date_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value)) {
    return value;
  }
  if (value === null || value === undefined || value === '') return null;
  const parsed = new Date(value);
  return isNaN(parsed.getTime()) ? null : parsed;
}


function historicalGradeImport_formatDate_(value) {
  return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}


function historicalGradeImport_number_(value) {
  if (value === null || value === undefined || value === '') return null;
  const number = typeof value === 'number'
    ? value : Number(String(value).replace(/%/g, '').trim());
  return isFinite(number) ? number : null;
}


function historicalGradeImport_isTrue_(value) {
  return value === true || historicalGradeImport_normalize_(value) === 'true' ||
    historicalGradeImport_normalize_(value) === 'yes' || String(value).trim() === '1';
}


function historicalGradeImport_clean_(value) {
  return value === null || value === undefined ? '' : String(value).trim();
}


function historicalGradeImport_normalize_(value) {
  return historicalGradeImport_clean_(value).toLowerCase().replace(/\s+/g, ' ');
}


function historicalGradeImport_errorMessage_(error) {
  return error && error.message ? error.message : String(error || 'Unknown import error.');
}


/**
 * Static regression test: validates Record Term, category mapping, deduction
 * sign, duplicate indexing assumptions, and the 0/1/2/3 PC averaging rule.
 * This test does not write either staging or live sheets.
 */
function testHistoricalGradeImportLogic() {
  const pc = [90, 88.75, 91];
  const average = values => values.length
    ? Math.round((values.reduce((sum, value) => sum + value, 0) /
      values.length) * 100) / 100 : null;
  const zero = average([]);
  const one = average(pc.slice(0, 1));
  const two = average(pc.slice(0, 2));
  const three = average(pc);
  const deductionPayload = historicalGradeImport_payload_({
    'Import ID': 'HGI-TEST000001',
    Title: 'Late Paperwork',
    'Evaluation Period': '',
    Instructor: '',
    Notes: 'Late paperwork',
    'Source Type': 'Historical Workbook',
    'Source Document': 'Clinical 1 grades.xlsx',
    'Source File ID': ''
  }, {
    student: { studentId: 'TEST-001', student: 'Test Student' },
    recordTerm: 'Clinical 1',
    category: 'Grade Point Reduction',
    grade: 2,
    date: new Date(2026, 4, 1)
  });
  const result = {
    success: zero === null && one === 90 && two === 89.38 && three === 89.92 &&
      deductionPayload.points === 2 &&
      deductionPayload.gradeCategory === 'Point Deduction' &&
      deductionPayload.clinicalTerm === 'Clinical 1' &&
      deductionPayload.source === 'Historical Grade Import' &&
      deductionPayload.sourceId === 'HGI-TEST000001' &&
      deductionPayload.replaceExisting === false,
    professionalCompetencyAverages: {
      zeroRecords: zero,
      oneRecord: one,
      twoRecords: two,
      threeRecords: three
    },
    deductionInput: deductionPayload.points,
    deductionStoredByExistingService: -2,
    recordTerm: deductionPayload.clinicalTerm,
    clinicalTrackingWrites: 0
  };
  if (!result.success) throw new Error('Historical Grade Import logic test failed.');
  Logger.log(JSON.stringify(result, null, 2));
  return result;
}


/**
 * Live read-only Isabella regression report. Run after import. It does not
 * create, update, void, or delete records.
 */
function inspectIsabellaHistoricalGradeImport() {
  if (typeof gradeEngine_getStudentSummary !== 'function') {
    throw new Error('gradeEngine_getStudentSummary() is required for this inspection.');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // Resolve by the authoritative Students source. Do not reconstruct or
  // normalize Isabella's department ID in the inspector.
  const student = gradeRecords_resolveStudent_(ss, '', 'Isabella Steinberg');
  const studentId = student.studentId;

  if (!studentId) {
    throw new Error('Isabella Steinberg resolved without a canonical Student ID.');
  }

  // Verify the imported records directly from the authoritative Grade Records
  // sheet before asking the Grade Engine to calculate either term. Raw sheet
  // objects are used because that is the Grade Engine's native input schema.
  const gradeSheet = gradeRecords_getSheet_(ss);
  const allGradeRows = gradeRecords_getSheetObjects_(gradeSheet);
  const records = allGradeRows.filter(row =>
    historicalGradeImport_normalize_(row['Student ID']) ===
      historicalGradeImport_normalize_(studentId) &&
    historicalGradeImport_normalize_(row.Source) ===
      historicalGradeImport_normalize_(HISTORICAL_GRADE_IMPORT_CONFIG.source) &&
    gradeRecords_isTrue_(row.Active)
  );

  if (records.length !== 11) {
    throw new Error(
      'Expected 11 active Historical Grade Import records for ' +
      student.student + ' (' + studentId + '), but found ' + records.length +
      '. No grade summary was calculated.'
    );
  }

  // The public Grade Engine API owns student matching, effective-term
  // attribution, category inclusion, conversion, weighting, and deductions.
  // This inspector must never reconstruct any of those inputs independently.
  const canonicalGradeEngineSummary = gradeEngine_getStudentSummary(studentId);
  const build = term => (canonicalGradeEngineSummary.terms || []).find(item =>
    historicalGradeImport_normalize_(item.clinicalTerm) ===
      historicalGradeImport_normalize_(term)
  ) || null;

  const number = value => {
    if (value === null || value === undefined || value === '') return null;
    const parsed = Number(value);
    return isFinite(parsed) ? parsed : null;
  };
  const firstNumber = function() {
    for (let i = 0; i < arguments.length; i += 1) {
      const parsed = number(arguments[i]);
      if (parsed !== null) return parsed;
    }
    return null;
  };
  const categoryAverage = category =>
    category && category.available ? number(category.percentAverage) : null;
  const summarize = summary => {
    const categories = summary && summary.categories ? summary.categories : {};
    const deductions = summary && summary.deductions ? summary.deductions : {};
    const deductionTotal = firstNumber(deductions.total, summary.deductionTotal, 0);
    const finalAverage = firstNumber(
      summary.recordedAverage,
      summary.finalAverage,
      summary.final,
      summary.raw,
      summary.value,
      summary.semesterAverage && summary.semesterAverage.raw,
      summary.semesterAverage && summary.semesterAverage.value
    );
    const preDeduction = firstNumber(
      summary.averageBeforeDeductions,
      summary.preDeductionAverage,
      summary.courseAverageBeforeDeduction,
      summary.weightedAverageBeforeDeductions,
      finalAverage === null ? null : finalAverage - deductionTotal
    );
    return {
      competency: categoryAverage(categories.competency),
      iso: categoryAverage(categories.iso),
      professional: categoryAverage(categories.professional),
      quizzes: categoryAverage(categories.quizzes),
      clinicalTest: categoryAverage(categories.clinicalTest),
      casePresentation: categoryAverage(categories.casePresentation),
      deductions: deductionTotal,
      preDeductionAverage: preDeduction,
      finalAverage: finalAverage
    };
  };
  const compare = (actual, expected) => {
    const discrepancies = {};
    Object.keys(expected).forEach(key => {
      const actualNumber = number(actual[key]);
      const expectedNumber = number(expected[key]);
      if (actualNumber === null || Math.abs(actualNumber - expectedNumber) > 0.01) {
        discrepancies[key] = { actual: actual[key], expected: expected[key] };
      }
    });
    return { actual: actual, expected: expected, discrepancies: discrepancies };
  };

  const clinical1 = build('Clinical 1');
  const clinical2 = build('Clinical 2');
  const clinical1Actual = summarize(clinical1);
  const clinical2Actual = summarize(clinical2);
  const report = {
    success: true,
    readOnly: true,
    canonicalStudent: {
      studentIdSentToGradeEngine: studentId,
      student: student.student,
      gradeEngineLookupCompleted: Boolean(canonicalGradeEngineSummary),
      authoritativeReadApi: 'gradeEngine_getStudentSummary'
    },
    existingGradeRecordsVerifiedBeforeCalculation: {
      expected: 11,
      found: records.length,
      gradeRecordIds: records.map(row => row['Grade Record ID'])
    },
    clinicalTrackingWrites: 0,
    recordsWrittenOrChanged: 0,
    clinical1: compare(clinical1Actual, {
      professional: 89.38, quizzes: 91.6, clinicalTest: 89.7,
      deductions: -2, preDeductionAverage: 88.44, finalAverage: 86.44
    }),
    clinical2: compare(clinical2Actual, {
      competency: 88.00, iso: 88.13, professional: 90.42,
      quizzes: 91.8, clinicalTest: 92.5,
      casePresentation: 95.9, finalAverage: 91.08
    }),
    note: 'Read-only inspection. Category and semester values come directly from gradeEngine_getStudentSummary(). No records, weights, Semester Rules, or calculations were changed.'
  };
  Logger.log(JSON.stringify(report, null, 2));
  return report;
}


/**
 * Compact read-only diagnostic for the remaining Isabella Clinical 2
 * Competency difference. This reports source rows, but the comparison value
 * is read directly from gradeEngine_getStudentSummary().
 */
function diagnoseIsabellaClinical2CompetencyInputs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const student = gradeRecords_resolveStudent_(ss, '', 'Isabella Steinberg');
  const studentId = student.studentId;
  const summary = gradeEngine_getStudentSummary(studentId);
  const term = (summary.terms || []).find(item =>
    appNormalize_(item.clinicalTerm) === 'clinical 2'
  ) || null;

  const trackingSheet = appGetSheet_(ss, APP_CONFIG.SHEETS.CLINICAL_TRACKING);
  const rows = appSheetToObjects_(trackingSheet).filter(row =>
    gradeEngine_sameStudent_(row, studentId, student.student) &&
    appNormalize_(row['Evaluation Type']) === 'competency'
  );

  const records = rows.map(row => {
    const effectiveClinicalTerm = gradeEngine_effectiveClinicalTerm_(row);
    const finalScore = gradeEngine_numeric_(row['Final Score']);
    const historicalPercent = gradeEngine_numeric_(row['Historical Percent Grade']);
    let convertedPercentage = finalScore === null
      ? null
      : gradeEngine_scoreToPercent(finalScore);
    let source = finalScore === null ? '' : 'Final Score';

    if (
      convertedPercentage === null &&
      gradeEngine_isHistoricalTrackingRecord_(row) &&
      historicalPercent !== null &&
      historicalPercent >= 0 &&
      historicalPercent <= 100
    ) {
      convertedPercentage = historicalPercent <= 1
        ? historicalPercent * 100
        : historicalPercent;
      source = 'Historical Percent Grade';
    }

    let exclusionReason = '';
    if (appNormalize_(effectiveClinicalTerm) !== 'clinical 2') {
      exclusionReason = 'Effective Clinical Term is ' +
        (effectiveClinicalTerm || 'blank') + ', not Clinical 2.';
    } else if (convertedPercentage === null) {
      exclusionReason = 'No numeric Final Score or eligible historical percentage.';
    }

    return {
      submissionId: String(row['Submission ID'] || '').trim(),
      examination: String(row.Examination || '').trim(),
      storedClinicalTerm: String(row['Clinical Term'] || '').trim(),
      effectiveClinicalTerm: effectiveClinicalTerm,
      finalScore: finalScore,
      historicalPercentGrade: historicalPercent,
      convertedPercentage: convertedPercentage,
      gradeValueSource: source,
      included: exclusionReason === '',
      exclusionReason: exclusionReason
    };
  });

  const included = records.filter(record => record.included);
  const includedValues = included.map(record => record.convertedPercentage);
  const arithmeticAverage = includedValues.length
    ? gradeEngine_round2_(includedValues.reduce((sum, value) => sum + value, 0) /
        includedValues.length)
    : null;
  const engineAverage = term && term.categories && term.categories.competency
    ? term.categories.competency.percentAverage
    : null;

  const report = {
    success: true,
    readOnly: true,
    recordsWrittenOrChanged: 0,
    studentId: studentId,
    student: student.student,
    clinicalTerm: 'Clinical 2',
    authoritativeReadApi: 'gradeEngine_getStudentSummary',
    records: records,
    includedSubmissionIds: included.map(record => record.submissionId),
    includedValues: includedValues,
    arithmetic: includedValues.length
      ? '(' + includedValues.join(' + ') + ') / ' + includedValues.length
      : 'No included values',
    arithmeticAverage: arithmeticAverage,
    gradeEngineCompetencyAverage: engineAverage,
    matchesGradeEngine: arithmeticAverage === engineAverage
  };

  Logger.log(JSON.stringify(report, null, 2));
  return report;
}


function historicalGradeImport_summaryComparison_(summary, expected) {
  const categories = summary && summary.categories ? summary.categories : {};
  const actual = {
    professional: historicalGradeImport_categoryAverage_(categories.professional),
    quizzes: historicalGradeImport_categoryAverage_(categories.quizzes),
    clinicalTest: historicalGradeImport_categoryAverage_(categories.clinicalTest),
    casePresentation: historicalGradeImport_categoryAverage_(categories.casePresentation),
    deductions: summary ? summary.deductions : null,
    finalAverage: summary && summary.semesterAverage
      ? summary.semesterAverage.value : null
  };
  const discrepancies = {};
  Object.keys(expected).forEach(key => {
    const left = historicalGradeImport_number_(actual[key]);
    const right = historicalGradeImport_number_(expected[key]);
    if (left === null || Math.abs(left - right) > 0.01) {
      discrepancies[key] = { actual: actual[key], expected: expected[key] };
    }
  });
  return { actual: actual, expected: expected, discrepancies: discrepancies };
}


function historicalGradeImport_categoryAverage_(category) {
  if (!category || !category.available) return null;
  return category.percentAverage;
}


/**
 * READ-ONLY diagnostic for Isabella Steinberg's Clinical 1 and Clinical 2
 * Grade Engine inputs and output. This function performs no sheet writes and
 * calls no importer, setup, append, update, void, or recalculation mutation.
 */
function diagnoseIsabellaGradeEngineInputs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const student = gradeRecords_resolveStudent_(ss, '', 'Isabella Steinberg');
  const studentId = student.studentId;
  const trackingSheet = appGetSheet_(ss, APP_CONFIG.SHEETS.CLINICAL_TRACKING);
  const semesterRulesSheet = appGetSheet_(ss, APP_CONFIG.SHEETS.SEMESTER_RULES);
  const allTracking = appSheetToObjects_(trackingSheet);
  const semesterRules = appSheetToObjects_(semesterRulesSheet);
  const publicSummary = gradeEngine_getStudentSummary(studentId);
  const targetSubmissionIds = [
    'HIST-HI-E61AA8BF16',
    'HIST-HI-82C9D62DC3',
    'HIST-EVAL-970'
  ];

  const sanitize = value => {
    if (value === null || value === undefined) return value;
    if (Object.prototype.toString.call(value) === '[object Date]') {
      return isNaN(value.getTime()) ? String(value) : value.toISOString();
    }
    if (Array.isArray(value)) return value.map(sanitize);
    if (typeof value === 'object') {
      const result = {};
      Object.keys(value).forEach(key => { result[key] = sanitize(value[key]); });
      return result;
    }
    return value;
  };
  const clean = value => value === null || value === undefined
    ? '' : String(value).trim();
  const normalized = value => appNormalize_(value);
  const numeric = value => gradeEngine_numeric_(value);
  const effectiveTerm = row => gradeEngine_effectiveClinicalTerm_(row);
  const sameStudent = row => gradeEngine_sameStudent_(
    row, studentId, student.student
  );
  const engineTerm = term => (publicSummary.terms || []).find(item =>
    normalized(item.clinicalTerm) === normalized(term)
  ) || null;
  const engineWeights = term => gradeEngine_getCourseWeights_(term);
  const categoryKey = row => {
    const type = normalized(row['Evaluation Type']);
    if (type === 'competency') return 'competency';
    if (type === 'iso') return 'iso';
    return '';
  };
  const inclusion = (row, term) => {
    if (!sameStudent(row)) {
      return { included: false, category: '', reason: 'Student ID/name does not match Isabella.' };
    }
    const rowTerm = effectiveTerm(row);
    if (normalized(rowTerm) !== normalized(term)) {
      return {
        included: false,
        category: categoryKey(row),
        reason: 'Effective Clinical Term is ' + (rowTerm || 'blank') + ', not ' + term + '.'
      };
    }
    const category = categoryKey(row);
    if (!category) {
      return {
        included: false,
        category: '',
        reason: 'Evaluation Type is not exactly Competency or ISO.'
      };
    }
    if (Number(engineWeights(term)[category] || 0) <= 0) {
      return {
        included: false,
        category: category,
        reason: 'The Grade Engine weight for ' + category + ' is zero in ' + term + '.'
      };
    }
    const rawScore = numeric(row['Final Score']);
    let percent = rawScore === null ? null : gradeEngine_scoreToPercent(rawScore);
    let scoreSource = rawScore === null ? '' : 'Final Score';
    if (percent === null && gradeEngine_isHistoricalTrackingRecord_(row)) {
      const historicalPercent = numeric(row['Historical Percent Grade']);
      if (historicalPercent !== null && historicalPercent >= 0 &&
          historicalPercent <= 100) {
        percent = historicalPercent <= 1 ? historicalPercent * 100 : historicalPercent;
        scoreSource = 'Historical Percent Grade';
      }
    }
    if (percent === null) {
      return {
        included: false,
        category: category,
        reason: 'Neither numeric Final Score nor eligible Historical Percent Grade is available; the Grade Engine ignores this row.'
      };
    }
    if (percent === null) {
      return {
        included: false,
        category: category,
        reason: 'Final Score could not be converted to a percentage.'
      };
    }
    return {
      included: true,
      category: category,
      rawScore: rawScore,
      calculatedPercentage: percent,
      scoreSource: scoreSource,
      reason: 'Included using ' + scoreSource + '.'
    };
  };
  const describeRow = (row, term) => {
    const decision = inclusion(row, term);
    return {
      submissionId: clean(row['Submission ID']),
      storedClinicalTerm: clean(row['Clinical Term']),
      attributedClinicalTerm: clean(row['Attributed Clinical Term']),
      effectiveClinicalTermUsedByGradeEngine: effectiveTerm(row),
      evaluationType: clean(row['Evaluation Type']),
      examination: clean(row.Examination),
      result: clean(row.Result),
      performanceScore: row['Performance Score'],
      imageScore: row['Image Score'],
      finalScore: row['Final Score'],
      historicalPercentGrade: row['Historical Percent Grade'],
      gradeValueSourceUsedByGradeEngine:
        decision.included ? decision.scoreSource : null,
      calculatedPercentageUsedByGradeEngine:
        decision.included ? decision.calculatedPercentage : null,
      includedInSemesterGradeCalculation: decision.included,
      includedCategory: decision.category,
      inclusionOrExclusionReason: decision.reason
    };
  };
  const arithmetic = (rows, term, category) => {
    const included = rows.map(row => ({ row: row, decision: inclusion(row, term) }))
      .filter(item => item.decision.included && item.decision.category === category);
    const values = included.map(item => item.decision.calculatedPercentage);
    const average = values.length
      ? gradeEngine_round2_(values.reduce((sum, value) => sum + value, 0) / values.length)
      : null;
    return {
      submissionIds: included.map(item => clean(item.row['Submission ID'])),
      values: values,
      formula: values.length ? '(' + values.join(' + ') + ') / ' + values.length : 'No included values',
      arithmeticAverage: average
    };
  };
  const studentTracking = allTracking.filter(sameStudent);
  const termReport = term => {
    const rows = studentTracking.filter(row =>
      normalized(effectiveTerm(row)) === normalized(term)
    );
    const rawSummary = engineTerm(term);
    const competencyArithmetic = arithmetic(rows, term, 'competency');
    const isoArithmetic = arithmetic(rows, term, 'iso');
    return {
      clinicalTerm: term,
      allClinicalTrackingRowsConsideredForTerm: rows.map(row => describeRow(row, term)),
      competencyAverageInputs: competencyArithmetic,
      isoAverageInputs: isoArithmetic,
      gradeEngineReportedCategoryAverages: {
        competency: rawSummary && rawSummary.categories
          ? rawSummary.categories.competency : null,
        iso: rawSummary && rawSummary.categories
          ? rawSummary.categories.iso : null
      },
      rawGradeEngineSemesterSummary: sanitize(rawSummary)
    };
  };
  const targetRecords = targetSubmissionIds.map(submissionId => {
    const physicalMatches = allTracking.filter(row =>
      normalized(row['Submission ID']) === normalized(submissionId)
    );
    return {
      submissionId: submissionId,
      physicallyExistsInClinicalTracking: physicalMatches.length > 0,
      matchCount: physicalMatches.length,
      matches: physicalMatches.map(row => {
        const term = effectiveTerm(row);
        const decision = inclusion(row, term);
        return {
          studentId: clean(row['Student ID']),
          student: clean(row.Student),
          storedClinicalTerm: clean(row['Clinical Term']),
          attributedClinicalTerm: clean(row['Attributed Clinical Term']),
          effectiveClinicalTermUsedByGradeEngine: term,
          evaluationType: clean(row['Evaluation Type']),
          examination: clean(row.Examination),
          result: clean(row.Result),
          performanceScore: row['Performance Score'],
          imageScore: row['Image Score'],
          finalScore: row['Final Score'],
          historicalPercentGrade: row['Historical Percent Grade'],
          gradeValueSourceUsedByGradeEngine:
            decision.included ? decision.scoreSource : null,
          calculatedPercentageUsedByGradeEngine:
            decision.included ? decision.calculatedPercentage : null,
          includedByGradeEngine: decision.included,
          inclusionOrExclusionReason: decision.reason
        };
      })
    };
  });
  const firstPresent = (row, names) => {
    for (let i = 0; i < names.length; i += 1) {
      if (Object.prototype.hasOwnProperty.call(row, names[i])) return row[names[i]];
    }
    return null;
  };
  const requestedRuleFields = row => ({
    competencyWeight: firstPresent(row, ['Competency Weight', 'Competency %', 'Competency']),
    isoWeight: firstPresent(row, ['ISO Weight', 'ISO %', 'ISO']),
    professionalCompetencyWeight: firstPresent(row,
      ['Professional Competency Weight', 'Professional Competency %', 'Professional Competency']),
    clinicalTestWeight: firstPresent(row, ['Clinical Test Weight', 'Clinical Test %', 'Clinical Test']),
    casePresentationWeight: firstPresent(row,
      ['Case Presentation Weight', 'Case Presentation %', 'Case Presentation']),
    clinicalQuizWeight: firstPresent(row,
      ['Clinical Quiz Weight', 'Clinical Quizzes Weight', 'Clinical Quiz %', 'Clinical Quizzes'])
  });
  const ruleReport = term => {
    const rows = semesterRules.filter(row =>
      normalized(row['Clinical Term']) === normalized(term)
    );
    return {
      clinicalTerm: term,
      matchingSemesterRulesRowCount: rows.length,
      matchingRows: rows.map(row => ({
        requestedWeightFields: requestedRuleFields(row),
        rawRow: sanitize(row)
      })),
      gradeEngineActuallyUsesSemesterRulesForWeights: false,
      gradeEngineActualWeights: engineWeights(term),
      explanation: 'The current Grade Engine uses GRADE_ENGINE_COURSE_WEIGHTS; it does not load Semester Rules for category weights.'
    };
  };

  const clinical1 = termReport('Clinical 1');
  const clinical2 = termReport('Clinical 2');
  const classifications = [];
  targetRecords.forEach(target => {
    if (!target.physicallyExistsInClinicalTracking) {
      classifications.push({
        item: target.submissionId,
        classification: 'A. Missing Clinical Tracking record'
      });
      return;
    }
    target.matches.forEach(match => {
      if (match.includedByGradeEngine) return;
      const reason = match.inclusionOrExclusionReason;
      let classification = 'B. Record exists but Grade Engine excludes it';
      if (reason.indexOf('Final Score') >= 0) {
        classification = 'C. Historical record lacks a numeric score required by current Grade Engine';
      } else if (reason.indexOf('Clinical Term') >= 0 ||
                 reason.indexOf('Evaluation Type') >= 0 ||
                 reason.indexOf('Student') >= 0) {
        classification = 'D. Wrong term/type/filter';
      }
      classifications.push({
        item: target.submissionId,
        classification: classification,
        reason: reason
      });
    });
  });
  classifications.push({
    item: 'inspectIsabellaHistoricalGradeImport() null averages',
    classification: 'E. Inspector is reading wrong Grade Engine property',
    reason: 'Grade Engine properties are averageBeforeDeductions and recordedAverage.'
  });
  const verifyArithmetic = (term, key, arithmeticResult) => {
    const summary = engineTerm(term);
    const reported = summary && summary.categories && summary.categories[key]
      ? summary.categories[key].percentAverage : null;
    if (numeric(reported) !== numeric(arithmeticResult.arithmeticAverage)) {
      classifications.push({
        item: term + ' ' + key + ' arithmetic',
        classification: 'F. Grade Engine calculation defect',
        reason: 'Diagnostic arithmetic=' + arithmeticResult.arithmeticAverage +
          '; Grade Engine=' + reported + '.'
      });
    }
  };
  verifyArithmetic('Clinical 1', 'competency', clinical1.competencyAverageInputs);
  verifyArithmetic('Clinical 1', 'iso', clinical1.isoAverageInputs);
  verifyArithmetic('Clinical 2', 'competency', clinical2.competencyAverageInputs);
  verifyArithmetic('Clinical 2', 'iso', clinical2.isoAverageInputs);
  classifications.push({
    item: 'Legacy Clinical 2 Professional Competency value 90.13',
    classification: 'G. Historical workbook formula/data discrepancy',
    reason: 'Legacy workbook duplicated one 89.25 record; Grade Records correctly contains only the three valid evaluations.'
  });

  const report = {
    success: true,
    readOnly: true,
    student: { studentId: studentId, student: student.student },
    recordsWrittenOrChanged: 0,
    clinical1: clinical1,
    clinical2: clinical2,
    specificallyRequestedClinicalTrackingRecords: targetRecords,
    semesterRulesAndActualEngineWeights: {
      clinical1: ruleReport('Clinical 1'),
      clinical2: ruleReport('Clinical 2')
    },
    gradeEnginePropertyDiagnosis: {
      preDeductionProperty: 'averageBeforeDeductions',
      finalRecordedProperty: 'recordedAverage',
      inspectorPreviouslyRead: ['preDeductionAverage', 'finalAverage', 'raw', 'value'],
      conclusion: 'The Grade Engine values are present; the historical inspector read the wrong property names.'
    },
    discrepancyClassifications: classifications
  };
  Logger.log(JSON.stringify(report, null, 2));
  return report;
}


/**
 * Preview the exact Isabella-only controlled repair. READ ONLY.
 *
 * The returned approvalToken must be passed unchanged to
 * applyIsabellaHistoricalGradeRepairs(). If the underlying candidate set
 * changes, the token changes and the apply is refused.
 */
function previewIsabellaHistoricalGradeRepairs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const student = gradeRecords_resolveStudent_(ss, '', 'Isabella Steinberg');
  const trackingSheet = appGetSheet_(ss, APP_CONFIG.SHEETS.CLINICAL_TRACKING);
  const tracking = appSheetToObjects_(trackingSheet);
  const gradeSheet = gradeRecords_getSheet_(ss);
  const gradeRows = gradeRecords_getSheetObjects_(gradeSheet);
  const sameStudent = row => gradeEngine_sameStudent_(
    row, student.studentId, student.student
  );
  const historical = row => gradeEngine_isHistoricalTrackingRecord_(row);
  const expectedBackfills = [
    { submissionId: 'HIST-HI-1AF8BC64A0', clinicalTerm: 'Clinical 1', evaluationType: 'Competency', examination: 'Wrist', validateExamination: true, percent: 94 },
    { submissionId: 'HIST-HI-27849A60B5', clinicalTerm: 'Clinical 1', evaluationType: 'Competency', examination: 'Hand', validateExamination: true, percent: 85.25 },
    { submissionId: 'HIST-HI-546C81FCF1', clinicalTerm: 'Clinical 1', evaluationType: 'Competency', examination: 'Shoulder', validateExamination: true, percent: 83.25 },
    { submissionId: 'HIST-HI-E61AA8BF16', clinicalTerm: 'Clinical 1', evaluationType: 'ISO', examination: 'Wrist', validateExamination: true, percent: 85.5 },
    { submissionId: 'HIST-HI-82C9D62DC3', clinicalTerm: 'Clinical 2', evaluationType: 'ISO', examination: 'Lower Leg ISO', validateExamination: false, percent: 88.75 }
  ];
  const backfills = [];
  const errors = [];

  expectedBackfills.forEach(expected => {
    const matches = tracking.filter(row =>
      historicalGradeImport_normalize_(row['Submission ID']) ===
        historicalGradeImport_normalize_(expected.submissionId)
    );
    if (matches.length !== 1) {
      errors.push(
        expected.submissionId + ': expected exactly one Clinical Tracking row; found ' +
        matches.length + '.'
      );
      return;
    }
    const row = matches[0];
    const finalScore = gradeEngine_numeric_(row['Final Score']);
    const existingHistoricalPercent = gradeEngine_numeric_(
      row['Historical Percent Grade']
    );
    const stored = {
      submissionId: String(row['Submission ID'] || '').trim(),
      studentId: String(row['Student ID'] || '').trim(),
      student: String(row.Student || '').trim(),
      clinicalTerm: String(row['Clinical Term'] || '').trim(),
      attributedClinicalTerm: String(row['Attributed Clinical Term'] || '').trim(),
      evaluationType: String(row['Evaluation Type'] || '').trim(),
      examination: String(row.Examination || '').trim(),
      finalScore: row['Final Score'],
      historicalPercentGrade: row['Historical Percent Grade'] || ''
    };
    const validationFailures = [];
    if (!sameStudent(row)) validationFailures.push('Student ID/name does not match Isabella.');
    if (!historical(row)) validationFailures.push('Record is not identified as historical.');
    if (historicalGradeImport_normalize_(row['Clinical Term']) !==
        historicalGradeImport_normalize_(expected.clinicalTerm)) {
      validationFailures.push('Stored Clinical Term does not equal ' + expected.clinicalTerm + '.');
    }
    if (historicalGradeImport_normalize_(row['Evaluation Type']) !==
        historicalGradeImport_normalize_(expected.evaluationType)) {
      validationFailures.push('Evaluation Type does not equal ' + expected.evaluationType + '.');
    }
    if (expected.validateExamination &&
        historicalGradeImport_normalize_(row.Examination) !==
          historicalGradeImport_normalize_(expected.examination)) {
      validationFailures.push('Examination does not equal ' + expected.examination + '.');
    }
    if (finalScore !== null) {
      validationFailures.push(
        'Final Score is numeric; Historical Percent Grade fallback must not be written.'
      );
    }
    if (existingHistoricalPercent !== null &&
        Math.abs(existingHistoricalPercent - expected.percent) > 0.001) {
      validationFailures.push(
        'Existing Historical Percent Grade ' +
        existingHistoricalPercent + ' conflicts with approved value ' +
        expected.percent + '.'
      );
    }
    if (validationFailures.length) {
      errors.push({
        submissionId: expected.submissionId,
        validationFailures: validationFailures,
        exactStoredRowValues: stored
      });
      return;
    }
    backfills.push({
      submissionId: expected.submissionId,
      studentId: student.studentId,
      clinicalTerm: String(row['Clinical Term'] || '').trim(),
      attributedClinicalTerm: String(row['Attributed Clinical Term'] || '').trim(),
      evaluationType: String(row['Evaluation Type'] || '').trim(),
      examination: String(row.Examination || '').trim(),
      finalScore: row['Final Score'],
      currentHistoricalPercentGrade: row['Historical Percent Grade'] || '',
      approvedHistoricalPercentGrade: expected.percent,
      action: existingHistoricalPercent === null
        ? 'Backfill Historical Percent Grade'
        : 'Already matches approved Historical Percent Grade'
    });
  });

  const noRepairRowMatches = tracking.filter(row =>
    historicalGradeImport_normalize_(row['Submission ID']) ===
      'hist-eval-970'
  );
  let noRepairNeeded = {
    submissionId: 'HIST-EVAL-970',
    noRepairNeeded: false,
    reason: 'Expected exactly one Clinical Tracking row; found ' +
      noRepairRowMatches.length + '.',
    exactStoredRowValues: null
  };
  if (noRepairRowMatches.length === 1) {
    const row = noRepairRowMatches[0];
    const finalScore = gradeEngine_numeric_(row['Final Score']);
    const stored = {
      studentId: String(row['Student ID'] || '').trim(),
      student: String(row.Student || '').trim(),
      clinicalTerm: String(row['Clinical Term'] || '').trim(),
      attributedClinicalTerm: String(row['Attributed Clinical Term'] || '').trim(),
      evaluationType: String(row['Evaluation Type'] || '').trim(),
      examination: String(row.Examination || '').trim(),
      finalScore: row['Final Score'],
      historicalPercentGrade: row['Historical Percent Grade'] || ''
    };
    const valid = sameStudent(row) && historical(row) &&
      historicalGradeImport_normalize_(row['Clinical Term']) === 'clinical 2' &&
      historicalGradeImport_normalize_(row['Evaluation Type']) === 'iso' &&
      finalScore !== null;
    noRepairNeeded = {
      submissionId: 'HIST-EVAL-970',
      noRepairNeeded: valid,
      reason: valid
        ? 'Numeric Final Score already present; normal Grade Engine conversion applies.'
        : 'Stored row failed Isabella/Clinical 2/ISO/numeric Final Score validation.',
      finalScore: finalScore,
      calculatedPercentageFromNormalGradeEngineConversion:
        finalScore === null ? null : gradeEngine_scoreToPercent(finalScore),
      exactStoredRowValues: stored
    };
    if (!valid) {
      errors.push({
        submissionId: 'HIST-EVAL-970',
        validationFailures: [noRepairNeeded.reason],
        exactStoredRowValues: stored
      });
    }
  } else {
    errors.push({
      submissionId: 'HIST-EVAL-970',
      validationFailures: [noRepairNeeded.reason],
      exactStoredRowValues: null
    });
  }

  const duplicateSpecifications = [
    { gradeRecordId: 'GR-67C72790688341F2', categoryKey: 'professional', category: 'Professional Competency', percent: 89.38 },
    { gradeRecordId: 'GR-0181FDADC9454A05', categoryKey: 'clinicalTest', category: 'Clinical Test', percent: 89.70 }
  ];
  const duplicates = [];
  duplicateSpecifications.forEach(expected => {
    const matches = gradeRows.filter(row =>
      historicalGradeImport_normalize_(row['Grade Record ID']) ===
        historicalGradeImport_normalize_(expected.gradeRecordId)
    );
    if (matches.length !== 1) {
      errors.push(
        expected.gradeRecordId + ': expected exactly one Grade Record; found ' +
        matches.length + '.'
      );
      return;
    }
    const row = matches[0];
    const valid =
      historicalGradeImport_normalize_(row['Student ID']) ===
        historicalGradeImport_normalize_(student.studentId) &&
      historicalGradeImport_normalize_(row['Clinical Term']) === 'clinical 1' &&
      gradeRecords_normalizeGradeCategoryForRepair_(row['Grade Category']) ===
        expected.categoryKey &&
      historicalGradeImport_normalize_(row.Source) === 'dashboard manual entry' &&
      gradeRecords_isTrue_(row.Active) &&
      Math.abs(Number(row['Percent Grade']) - expected.percent) < 0.001;
    if (!valid) {
      errors.push({
        gradeRecordId: expected.gradeRecordId,
        validationFailures: ['Confirmed duplicate no longer matches the approved student/term/category/source/grade/active criteria.'],
        exactStoredRowValues: row
      });
      return;
    }
    duplicates.push({
      gradeRecordId: expected.gradeRecordId,
      category: String(row['Grade Category'] || '').trim(),
      grade: row['Percent Grade'],
      recordTerm: String(row['Clinical Term'] || '').trim(),
      source: String(row.Source || '').trim(),
      sourceId: String(row['Source ID'] || '').trim(),
      createdDate: historicalGradeImport_formatMaybeDate_(row['Entered At']),
      action: 'Deactivate through gradeRecords_void()',
      reason: 'Prior Isabella Clinical 1 Dashboard Manual Entry development artifact; authoritative Historical Grade Import source records now exist.'
    });
  });

  if (backfills.length !== 5) {
    errors.push('Expected exactly 5 Historical Percent Grade backfills; found ' +
      backfills.length + '.');
  }
  if (duplicates.length !== 2) {
    errors.push('Expected exactly 2 Dashboard Manual Entry deactivations; found ' +
      duplicates.length + '.');
  }
  if (!noRepairNeeded.noRepairNeeded) {
    errors.push('HIST-EVAL-970 was not safely confirmed as requiring no repair.');
  }

  const tokenMaterial = backfills.map(item =>
    item.submissionId + ':' + item.approvedHistoricalPercentGrade
  ).concat(duplicates.map(item => item.gradeRecordId)).sort().join('|');
  const approvalToken = 'ISABELLA-' + Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    tokenMaterial,
    Utilities.Charset.UTF_8
  ).slice(0, 8).map(byte => ('0' + ((byte + 256) % 256).toString(16)).slice(-2))
    .join('').toUpperCase();

  const result = {
    success: errors.length === 0,
    readOnly: true,
    student: student,
    schemaPreview: {
      clinicalTrackingHistoricalPercentGradeExists:
        historicalGradeImport_sheetHasHeader_(trackingSheet, 'Historical Percent Grade'),
      historicalImportHistoricalPercentGradeExists: (() => {
        const sheet = ss.getSheetByName('Historical Import');
        return sheet ? historicalGradeImport_sheetHasHeader_(
          sheet, 'Historical Percent Grade'
        ) : false;
      })(),
      proposedField: 'Historical Percent Grade'
    },
    historicalPercentageBackfills: backfills,
    historicalRecordsRequiringNoRepair: [noRepairNeeded],
    dashboardDevelopmentDuplicatesToDeactivate: duplicates,
    approvedLiveRecordChangeCount: backfills.length + duplicates.length,
    expectedApprovedLiveRecordChangeCount: 7,
    errors: errors,
    approvalToken: errors.length ? null : approvalToken,
    recordsWrittenOrChanged: 0
  };
  Logger.log(JSON.stringify(result, null, 2));
  return result;
}


/**
 * Apply only the exact Isabella repairs shown by the immediately matching
 * preview token. No other student can be changed by this function.
 */
function applyIsabellaHistoricalGradeRepairs(approvalToken) {
  const preview = previewIsabellaHistoricalGradeRepairs();
  if (!preview.success) {
    throw new Error('Repair preview contains errors. Nothing was changed.');
  }
  if (!approvalToken || String(approvalToken).trim() !== preview.approvalToken) {
    throw new Error(
      'Approval token does not match the current preview. Nothing was changed.'
    );
  }
  if (preview.student.studentId !== 'DEPT # 22' ||
      preview.student.student !== 'Isabella Steinberg') {
    throw new Error('Canonical Isabella identity check failed. Nothing was changed.');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const trackingSheet = appGetSheet_(ss, APP_CONFIG.SHEETS.CLINICAL_TRACKING);
  const stagingSheet = ss.getSheetByName('Historical Import');
  historicalGradeImport_ensureHeaderAfter_(
    trackingSheet, 'Historical Percent Grade', 'Final Score'
  );
  if (stagingSheet) {
    historicalGradeImport_ensureHeaderAfter_(
      stagingSheet, 'Historical Percent Grade', 'Final Score'
    );
  }

  const trackingRows = appSheetToObjects_(trackingSheet);
  const trackingHeaders = historicalGradeImport_headerMap_(trackingSheet);
  const backfilled = [];
  preview.historicalPercentageBackfills.forEach(candidate => {
    const matches = trackingRows.filter(row =>
      historicalGradeImport_normalize_(row['Submission ID']) ===
        historicalGradeImport_normalize_(candidate.submissionId) &&
      historicalGradeImport_normalize_(row['Student ID']) ===
        historicalGradeImport_normalize_(preview.student.studentId)
    );
    if (matches.length !== 1) {
      throw new Error(
        candidate.submissionId + ': record changed after preview. Repair stopped.'
      );
    }
    const rowNumber = matches[0]._rowNumber || historicalGradeImport_findSheetRow_(
      trackingSheet, 'Submission ID', candidate.submissionId
    );
    trackingSheet.getRange(
      rowNumber, trackingHeaders['Historical Percent Grade']
    ).setValue(candidate.approvedHistoricalPercentGrade);
    backfilled.push({
      submissionId: candidate.submissionId,
      historicalPercentGrade: candidate.approvedHistoricalPercentGrade
    });
  });

  const deactivated = [];
  preview.dashboardDevelopmentDuplicatesToDeactivate.forEach(candidate => {
    const result = gradeRecords_void(
      candidate.gradeRecordId,
      'Removed controlled Isabella development duplicate after Historical Grade Import.'
    );
    deactivated.push({
      gradeRecordId: candidate.gradeRecordId,
      category: candidate.category,
      result: result
    });
  });
  SpreadsheetApp.flush();

  const result = {
    success: true,
    student: preview.student,
    schema: {
      clinicalTracking: 'Historical Percent Grade',
      historicalImport: stagingSheet ? 'Historical Percent Grade' : 'Historical Import sheet not found'
    },
    backfilled: backfilled,
    deactivatedDevelopmentDuplicates: deactivated,
    unrelatedStudentRecordsChanged: 0,
    nextFunctions: [
      'diagnoseIsabellaGradeEngineInputs',
      'inspectIsabellaHistoricalGradeImport'
    ]
  };
  Logger.log(JSON.stringify(result, null, 2));
  return result;
}


function applyApprovedIsabellaHistoricalGradeRepairs() {
  if (!ISABELLA_HISTORICAL_REPAIR_APPROVAL_TOKEN) {
    throw new Error(
      'Run previewIsabellaHistoricalGradeRepairs(), review it, and paste its ' +
      'approvalToken into ISABELLA_HISTORICAL_REPAIR_APPROVAL_TOKEN first.'
    );
  }
  return applyIsabellaHistoricalGradeRepairs(
    ISABELLA_HISTORICAL_REPAIR_APPROVAL_TOKEN
  );
}


function gradeRecords_normalizeGradeCategoryForRepair_(value) {
  return gradeRecords_normalizeCategory_(value);
}


function historicalGradeImport_sheetHasHeader_(sheet, header) {
  return Object.prototype.hasOwnProperty.call(
    historicalGradeImport_headerMap_(sheet), header
  );
}


function historicalGradeImport_ensureHeaderAfter_(sheet, header, afterHeader) {
  const headers = historicalGradeImport_headerMap_(sheet);
  if (Object.prototype.hasOwnProperty.call(headers, header)) return;
  if (!Object.prototype.hasOwnProperty.call(headers, afterHeader)) {
    throw new Error(sheet.getName() + ' is missing header: ' + afterHeader + '.');
  }
  const afterColumn = headers[afterHeader];
  sheet.insertColumnAfter(afterColumn);
  sheet.getRange(1, afterColumn + 1).setValue(header);
}


function historicalGradeImport_findSheetRow_(sheet, idHeader, idValue) {
  const headers = historicalGradeImport_headerMap_(sheet);
  const idColumn = headers[idHeader];
  if (!idColumn) throw new Error(sheet.getName() + ' is missing header: ' + idHeader + '.');
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) throw new Error('No data rows found in ' + sheet.getName() + '.');
  const values = sheet.getRange(2, idColumn, lastRow - 1, 1).getValues();
  const matches = [];
  values.forEach((row, index) => {
    if (historicalGradeImport_normalize_(row[0]) ===
        historicalGradeImport_normalize_(idValue)) {
      matches.push(index + 2);
    }
  });
  if (matches.length !== 1) {
    throw new Error(idHeader + ' ' + idValue + ': expected one row; found ' + matches.length + '.');
  }
  return matches[0];
}


function historicalGradeImport_formatMaybeDate_(value) {
  if (!value) return '';
  const date = historicalGradeImport_date_(value);
  return date ? historicalGradeImport_formatDate_(date) : String(value);
}


/**
 * READ-ONLY Isabella Clinical 2 historical Competency repair preview.
 *
 * This function does not alter Clinical Tracking or the Grade Engine. To
 * simulate the proposed historical-value precedence with today's engine, it
 * clones records in memory and blanks Final Score only on the in-memory clone
 * when an eligible historical percentage exists. That makes the existing
 * historical fallback execute without changing any live value.
 */
function previewIsabellaClinical2HistoricalCompetencyRepair(suppressLog) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const student = gradeRecords_resolveStudent_(ss, '', 'Isabella Steinberg');
  const studentId = student.studentId;
  const clinicalTerm = 'Clinical 2';
  const trackingSheet = appGetSheet_(ss, APP_CONFIG.SHEETS.CLINICAL_TRACKING);
  const allTracking = appSheetToObjects_(trackingSheet);

  const documented = [
    { examination: 'Foot', aliases: ['foot'], percent: 92.75 },
    { examination: 'Knee', aliases: ['knee'], percent: 87.00 },
    { examination: 'Lower Leg (Tibia-Fibula)', aliases: ['lower leg tibia fibula', 'lower leg', 'tibia fibula'], percent: 86.25 },
    { examination: 'Cervical Spine', aliases: ['cervical spine'], percent: 86.75 },
    { examination: 'Lumbar Spine', aliases: ['lumbar spine'], percent: 86.25 },
    { examination: 'Chest (PA & Lateral)', aliases: ['chest pa and lateral', 'chest pa lateral'], percent: 90.50 },
    { examination: 'Ankle', aliases: ['ankle'], percent: 93.25 },
    { examination: 'Pelvis', aliases: ['pelvis'], percent: 93.25 },
    { examination: 'Forearm', aliases: ['forearm'], percent: 88.50 },
    { examination: 'Abdomen Supine (KUB)', aliases: ['abdomen supine kub', 'abdomen kub', 'kub'], percent: 87.75 },
    { examination: 'Clavicle', aliases: ['clavicle'], percent: 94.25 },
    { examination: 'Portable Chest', aliases: ['portable chest'], percent: 80.25 },
    { examination: 'Ribs', aliases: ['ribs'], percent: 88.50 },
    { examination: 'Sacrum and/or Coccyx', aliases: ['sacrum and or coccyx', 'sacrum coccyx'], percent: 76.75 }
  ];

  const normalizeExamination = value => String(value || '')
    .trim()
    .toLowerCase()
    .replace(/&/g, ' and ')
    .replace(/[^a-z0-9]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();

  const studentTracking = allTracking.filter(row =>
    gradeEngine_sameStudent_(row, studentId, student.student)
  );
  const eligiblePool = studentTracking.filter(row =>
    gradeEngine_isHistoricalTrackingRecord_(row) &&
    appNormalize_(row['Clinical Term']) === 'clinical 2' &&
    appNormalize_(row['Evaluation Type']) === 'competency'
  );
  const errors = [];
  const usedSubmissionIds = {};

  const candidates = documented.map(expected => {
    const aliases = expected.aliases.map(normalizeExamination);
    const matches = eligiblePool.filter(row =>
      aliases.indexOf(normalizeExamination(row.Examination)) !== -1
    );

    if (matches.length !== 1) {
      errors.push(
        expected.examination + ': expected exactly one historical Isabella ' +
        'Clinical 2 Competency match; found ' + matches.length + '.'
      );
      return {
        documentedExamination: expected.examination,
        proposedHistoricalPercentGrade: expected.percent,
        matchCount: matches.length,
        storedMatches: matches.map(row => ({
          submissionId: String(row['Submission ID'] || '').trim(),
          examination: String(row.Examination || '').trim(),
          storedClinicalTerm: String(row['Clinical Term'] || '').trim(),
          evaluationType: String(row['Evaluation Type'] || '').trim(),
          finalScore: row['Final Score'],
          historicalPercentGrade: row['Historical Percent Grade']
        }))
      };
    }

    const row = matches[0];
    const submissionId = String(row['Submission ID'] || '').trim();
    if (!submissionId) {
      errors.push(expected.examination + ': matched row has no Submission ID.');
    } else if (usedSubmissionIds[appNormalize_(submissionId)]) {
      errors.push(submissionId + ': one source row matched more than one documented examination.');
    }
    usedSubmissionIds[appNormalize_(submissionId)] = true;

    const finalScore = gradeEngine_numeric_(row['Final Score']);
    const existingHistoricalPercent = gradeEngine_numeric_(row['Historical Percent Grade']);
    const currentSource = finalScore !== null
      ? 'Final Score → gradeEngine_scoreToPercent()'
      : existingHistoricalPercent !== null
        ? 'Historical Percent Grade fallback'
        : 'Excluded — no numeric grade source';

    return {
      submissionId: submissionId,
      examination: String(row.Examination || '').trim(),
      documentedExamination: expected.examination,
      storedClinicalTerm: String(row['Clinical Term'] || '').trim(),
      effectiveClinicalTerm: gradeEngine_effectiveClinicalTerm_(row),
      evaluationType: String(row['Evaluation Type'] || '').trim(),
      existingFinalScore: finalScore,
      existingHistoricalPercentGrade: existingHistoricalPercent,
      proposedHistoricalPercentGrade: expected.percent,
      sourceCurrentlyUsed: currentSource,
      sourceAfterApprovedRepair: 'Historical Percent Grade',
      proposedAction: existingHistoricalPercent === expected.percent
        ? 'No value change required'
        : existingHistoricalPercent === null
          ? 'Set Historical Percent Grade'
          : 'Update Historical Percent Grade'
    };
  });

  const resolved = candidates.filter(candidate => candidate.submissionId);
  if (resolved.length !== documented.length) {
    errors.push('Expected 14 uniquely resolved repair rows; found ' + resolved.length + '.');
  }

  const proposedBySubmissionId = {};
  resolved.forEach(candidate => {
    proposedBySubmissionId[appNormalize_(candidate.submissionId)] =
      candidate.proposedHistoricalPercentGrade;
  });

  const simulatedTracking = studentTracking
    .filter(row =>
      appNormalize_(gradeEngine_effectiveClinicalTerm_(row)) === 'clinical 2'
    )
    .map(row => {
      const clone = Object.assign({}, row);
      const proposed = proposedBySubmissionId[
        appNormalize_(row['Submission ID'])
      ];
      if (proposed !== undefined) {
        clone['Historical Percent Grade'] = proposed;
      }

      // Preview the requested historical precedence for every eligible
      // historical row without changing either the live row or engine code.
      const historicalPercent = gradeEngine_numeric_(clone['Historical Percent Grade']);
      if (
        gradeEngine_isHistoricalTrackingRecord_(clone) &&
        historicalPercent !== null
      ) {
        clone['Final Score'] = '';
      }
      return clone;
    });

  const gradeRecords = gradeEngine_getSupplementalGradeRecords_(ss).filter(row =>
    gradeEngine_sameStudent_(row, studentId, student.student) &&
    appNormalize_(row['Clinical Term']) === 'clinical 2'
  );
  const simulatedSummary = errors.length === 0
    ? gradeEngine_buildTermSummary_({
        clinicalTerm: clinicalTerm,
        tracking: simulatedTracking,
        gradeRecords: gradeRecords
      })
    : null;
  const simulatedCompetency = simulatedSummary && simulatedSummary.categories
    ? simulatedSummary.categories.competency.percentAverage
    : null;
  const documentedAverage = gradeEngine_round2_(
    documented.reduce((sum, item) => sum + item.percent, 0) / documented.length
  );

  if (errors.length === 0 && simulatedCompetency !== documentedAverage) {
    errors.push(
      'Simulated Grade Engine Competency average is ' + simulatedCompetency +
      '; documented control is ' + documentedAverage + '.'
    );
  }

  const result = {
    success: errors.length === 0,
    readOnly: true,
    approvalRequiredBeforeWrite: true,
    recordsWrittenOrChanged: 0,
    studentId: studentId,
    student: student.student,
    clinicalTerm: clinicalTerm,
    proposedPrecedence: [
      'Historical + numeric Historical Percent Grade → Historical Percent Grade',
      'Otherwise numeric Final Score → existing conversion',
      'Otherwise excluded',
      'Current/live records never use Historical Percent Grade'
    ],
    expectedCandidateCount: 14,
    resolvedCandidateCount: resolved.length,
    candidates: candidates,
    documentedIncludedValues: documented.map(item => item.percent),
    documentedArithmetic: '(' + documented.map(item => item.percent).join(' + ') + ') / 14',
    documentedCompetencyAverage: documentedAverage,
    simulatedGradeEngineCompetencyAverage: simulatedCompetency,
    simulatedClinical2Semester: simulatedSummary,
    errors: errors,
    note: 'Preview only. No repair or schema/data/calculation change was applied.'
  };

  if (!suppressLog) {
    Logger.log(JSON.stringify(result, null, 2));
  }
  return result;
}


/**
 * Compact read-only verification of the already-prepared Isabella Clinical 2
 * historical repair preview. It intentionally returns only the approved
 * verification contract and performs no writes.
 */
function verifyIsabellaClinical2HistoricalRepairPreview() {
  const preview = previewIsabellaClinical2HistoricalCompetencyRepair(true);
  const candidates = (preview.candidates || []).filter(candidate =>
    candidate && candidate.submissionId
  );
  const finalCandidate = candidates[candidates.length - 1] || null;

  if (
    !finalCandidate ||
    finalCandidate.documentedExamination !== 'Sacrum and/or Coccyx' ||
    Number(finalCandidate.proposedHistoricalPercentGrade) !== 76.75
  ) {
    throw new Error(
      'Final candidate verification failed. Expected Sacrum and/or Coccyx = 76.75.'
    );
  }

  const summary = preview.simulatedClinical2Semester || {};
  const categories = summary.categories || {};
  const categoryAverage = key => {
    const category = categories[key];
    return category && category.available
      ? category.percentAverage
      : null;
  };
  const expected = Number(preview.documentedCompetencyAverage);
  const simulated = Number(preview.simulatedGradeEngineCompetencyAverage);

  const result = {
    student: preview.student,
    studentId: preview.studentId,
    candidateCount: candidates.length,
    allCandidatesResolved:
      candidates.length === 14 &&
      preview.resolvedCandidateCount === 14 &&
      (!preview.errors || preview.errors.length === 0),
    submissionIds: candidates.map(candidate => candidate.submissionId),
    examinations: candidates.map(candidate => candidate.documentedExamination),
    proposedHistoricalPercentGrades: candidates.map(candidate =>
      candidate.proposedHistoricalPercentGrade
    ),
    arithmeticFormula:
      preview.documentedArithmetic + ' = ' +
      expected.toFixed(2),
    simulatedCompetencyAverage: simulated,
    expectedCompetencyAverage: expected,
    competencyAverageMatchesExpected:
      Number.isFinite(simulated) &&
      Number.isFinite(expected) &&
      Math.abs(simulated - expected) < 0.005,
    simulatedClinical2CategoryAverages: {
      competency: categoryAverage('competency'),
      iso: categoryAverage('iso'),
      professional: categoryAverage('professional'),
      clinicalQuiz: categoryAverage('quizzes'),
      clinicalTest: categoryAverage('clinicalTest'),
      casePresentation: categoryAverage('casePresentation')
    },
    simulatedClinical2AverageBeforeDeductions:
      summary.averageBeforeDeductions === undefined
        ? null
        : summary.averageBeforeDeductions,
    simulatedClinical2FinalAverage:
      summary.recordedAverage === undefined
        ? null
        : summary.recordedAverage,
    approvalToken: preview.approvalToken || null,
    recordsWrittenOrChanged: 0
  };

  Logger.log(JSON.stringify(result, null, 2));
  return result;
}
