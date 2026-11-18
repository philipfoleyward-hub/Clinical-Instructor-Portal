/**
 * =====================================================================
 * CLINICAL COMPETENCY TRACKER
 * CLINICAL TRACKING SERVICE
 * 04_ClinicalTrackingService.gs
 * =====================================================================
 *
 * PURPOSE
 * -------
 * Single authoritative writer for the Clinical Tracking sheet.
 *
 * Clinical Tracking is the permanent factual evaluation ledger.
 *
 * REFERENCE MODEL
 * ---------------
 * Clinical Site ID  -> stable site identity
 * Evaluator ID      -> stable evaluator identity
 * Examination ID    -> stable examination identity
 *
 * Human-readable values are retained as recorded and are not rewritten
 * merely because an ID resolves.
 *
 * PRODUCTION RESOLUTION:
 *   - supplied valid ID
 *   - exact normalized name
 *
 * HISTORICAL BACKFILL ONLY:
 *   - exact normalized name
 *   - approved legacy mapping
 *   - unique exact evaluator last-name match
 *
 * NO FUZZY MATCHING.
 * =====================================================================
 */


/**
 * =====================================================================
 * MAIN ENTRY POINT
 * =====================================================================
 */
function clinicalTracking_addRecord(record) {

  if (!record) {
    throw new Error(
      'Clinical Tracking requires an evaluation record.'
    );
  }

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const trackingSheet =
    appGetSheet_(
      ss,
      APP_CONFIG.SHEETS.CLINICAL_TRACKING
    );


  // ===================================================================
  // SUBMISSION ID
  // ===================================================================

  const submissionId =
    clinicalTracking_firstValue_(
      record,
      [
        'Submission ID',
        'Record ID',
        'Record/ID'
      ]
    );

  if (!submissionId) {
    throw new Error(
      'Cannot create Clinical Tracking record without a Submission ID.'
    );
  }

  if (
    clinicalTracking_submissionExists_(
      trackingSheet,
      submissionId
    )
  ) {

    console.log(
      'Clinical Tracking already contains Submission ID: ' +
      submissionId
    );

    return {
      success: true,
      duplicate: true,
      submissionId: submissionId
    };
  }


  // ===================================================================
  // STUDENT
  // ===================================================================

  const suppliedStudentId =
    clinicalTracking_firstValue_(
      record,
      ['Student ID']
    );

  const suppliedStudentName =
    clinicalTracking_firstValue_(
      record,
      ['Student', 'Student Name']
    );

  const student =
    clinicalTracking_lookupStudent_(
      ss,
      suppliedStudentId,
      suppliedStudentName
    );

  const studentId =
    student
      ? student.studentId
      : suppliedStudentId;

  const studentName =
    student
      ? student.student
      : suppliedStudentName;

  if (!studentName) {
    throw new Error(
      'Cannot create Clinical Tracking record without a Student.'
    );
  }


  // ===================================================================
  // CLINICAL TERM
  // ===================================================================

  const clinicalTerm =
    clinicalTracking_firstValue_(
      record,
      ['Clinical Term']
    ) ||
    (
      student
        ? student.clinicalTerm
        : ''
    );


  // ===================================================================
  // CLINICAL SITE
  // ===================================================================

  const suppliedClinicalSiteId =
    clinicalTracking_firstValue_(
      record,
      ['Clinical Site ID']
    );

  const suppliedClinicalSite =
    clinicalTracking_firstValue_(
      record,
      ['Clinical Site']
    );

  const clinicalSiteRef =
    clinicalTracking_resolveReference_(
      ss,
      clinicalTracking_sheetName_(
        'CLINICAL_SITES',
        'Clinical Sites'
      ),
      'Clinical Site ID',
      'Clinical Site',
      suppliedClinicalSiteId,
      suppliedClinicalSite
    );

  const clinicalSiteId =
    clinicalSiteRef
      ? clinicalSiteRef.id
      : '';

  const clinicalSite =
    suppliedClinicalSite ||
    (
      clinicalSiteRef
        ? clinicalSiteRef.name
        : ''
    );


  // ===================================================================
  // EVALUATOR
  // ===================================================================

  const suppliedEvaluatorId =
    clinicalTracking_firstValue_(
      record,
      ['Evaluator ID']
    );

  const suppliedEvaluator =
    clinicalTracking_firstValue_(
      record,
      ['Evaluator']
    );

  const evaluatorRef =
    clinicalTracking_resolveReference_(
      ss,
      clinicalTracking_sheetName_(
        'EVALUATORS',
        'Evaluators'
      ),
      'Evaluator ID',
      'Evaluator',
      suppliedEvaluatorId,
      suppliedEvaluator
    );

  const evaluatorId =
    evaluatorRef
      ? evaluatorRef.id
      : '';

  const evaluator =
    suppliedEvaluator ||
    (
      evaluatorRef
        ? evaluatorRef.name
        : ''
    );

  const suppliedEvaluatorType =
    clinicalTracking_firstValue_(
      record,
      ['Evaluator Type']
    );

  const evaluatorType =
    evaluatorRef && evaluatorRef.row
      ? String(
          evaluatorRef.row['Evaluator Type'] || ''
        ).trim()
      : (
          suppliedEvaluatorType ||
          clinicalTracking_lookupEvaluatorType_(
            ss,
            evaluator
          )
        );


  // ===================================================================
  // EXAMINATION
  // ===================================================================

  const suppliedExaminationId =
    clinicalTracking_firstValue_(
      record,
      ['Examination ID']
    );

  const suppliedExamination =
    clinicalTracking_firstValue_(
      record,
      ['Examination']
    );

  const examinationRef =
    clinicalTracking_resolveReference_(
      ss,
      clinicalTracking_sheetName_(
        'EXAMINATIONS',
        'Examinations'
      ),
      'Examination ID',
      'Examination',
      suppliedExaminationId,
      suppliedExamination
    );

  const examinationId =
    examinationRef
      ? examinationRef.id
      : '';

  const examination =
    suppliedExamination ||
    (
      examinationRef
        ? examinationRef.name
        : ''
    );


  // ===================================================================
  // SOURCE / SUBMISSION METHOD
  // ===================================================================

  const source =
    clinicalTracking_firstValue_(
      record,
      ['Source']
    );

  const submissionMethod =
    clinicalTracking_firstValue_(
      record,
      ['Submission Method']
    ) ||
    clinicalTracking_submissionMethod_(
      source
    );


  // ===================================================================
  // RESULT / REVIEW
  // ===================================================================

  // Store approved ledger labels consistently with historical rows.
  // The grading engine may continue to use PASS/FAIL as internal codes;
  // the official Clinical Tracking record uses Pass/Fail from 01_Config.
  const rawResult =
    clinicalTracking_firstValue_(
      record,
      ['Result']
    );

  const normalizedResult = appNormalize_(rawResult);
  const result = normalizedResult === 'pass'
    ? APP_CONFIG.TERMINOLOGY.RESULTS.PASS
    : normalizedResult === 'fail'
      ? APP_CONFIG.TERMINOLOGY.RESULTS.FAIL
      : rawResult;

  let reviewStatus =
    clinicalTracking_firstValue_(
      record,
      ['Coordinator Review Status']
    );

  if (!reviewStatus) {

    if (
      appNormalize_(source) ===
        appNormalize_(
          APP_CONFIG.SOURCES.HISTORICAL_IMPORT
        ) &&
      APP_CONFIG.SETTINGS
        .HISTORICAL_IMPORT_REQUIRES_LIVE_REVIEW ===
        false
    ) {

      reviewStatus =
        APP_CONFIG.TERMINOLOGY
          .REVIEW_STATUS
          .APPROVED;

    } else {

      reviewStatus =
        APP_CONFIG.SETTINGS
          .DEFAULT_REVIEW_STATUS;
    }
  }

  let coordinatorFlag =
    clinicalTracking_firstValue_(
      record,
      ['Coordinator Flag']
    );

  if (
    !coordinatorFlag &&
    appNormalize_(result) === 'fail'
  ) {
    coordinatorFlag =
      'FAIL - Review Required';
  }


  // ===================================================================
  // BUILD RECORD
  // ===================================================================

  const trackingRecord = {

    'Submission ID': submissionId,
    'Student ID': studentId,
    'Student': studentName,

    'Clinical Term': clinicalTerm,

    'Attributed Clinical Term':
      clinicalTracking_firstValue_(
        record,
        ['Attributed Clinical Term']
      ),

    'Clinical Site ID': clinicalSiteId,
    'Clinical Site': clinicalSite,

    'Evaluator ID': evaluatorId,
    'Evaluator': evaluator,
    'Evaluator Type': evaluatorType,

    'Evaluation Type':
      clinicalTracking_firstValue_(
        record,
        ['Evaluation Type']
      ),

    'Examination ID': examinationId,
    'Examination': examination,

    'Anatomical Area':
      clinicalTracking_firstValue_(
        record,
        ['Anatomical Area']
      ),

    'Requirement Type':
      clinicalTracking_firstValue_(
        record,
        ['Requirement Type']
      ),

    'Date':
      clinicalTracking_firstValue_(
        record,
        ['Evaluation Date', 'Date']
      ),

    'Performance Score':
      clinicalTracking_firstValuePreserveZero_(
        record,
        ['Performance Score']
      ),

    'Image Score':
      clinicalTracking_firstValuePreserveZero_(
        record,
        ['Image Score']
      ),

    'Final Score':
      clinicalTracking_firstValuePreserveZero_(
        record,
        ['Final Score']
      ),

    'Historical Percent Grade':
      clinicalTracking_firstValuePreserveZero_(
        record,
        ['Historical Percent Grade']
      ),

    'Result': result,

    'Automatic Failure':
      clinicalTracking_firstValue_(
        record,
        ['Automatic Failure']
      ),

    'Counts Towards HCI Limit': '',
    'HCI Competency #': '',
    'Remediation Required': '',

    'Remediation Complete':
      clinicalTracking_firstValue_(
        record,
        ['Remediation Complete']
      ),

    'Submission Method': submissionMethod,

    'Source Document':
      clinicalTracking_firstValue_(
        record,
        ['Source Document', 'File URL']
      ),

    'Entered By':
      clinicalTracking_firstValue_(
        record,
        ['Entered By']
      ) ||
      'Automated Submission System',

    'Coordinator Review Status': reviewStatus,
    'Coordinator Flag': coordinatorFlag,

    'Coordinator Notes':
      clinicalTracking_firstValue_(
        record,
        ['Coordinator Notes']
      ),

    'Reviewed By':
      clinicalTracking_firstValue_(
        record,
        ['Reviewed By']
      ),

    'Review Date':
      clinicalTracking_firstValue_(
        record,
        ['Review Date']
      )
  };


// ===================================================================
// WRITE — ATOMIC DUPLICATE CHECK + APPEND
// ===================================================================

const lock =
  LockService.getDocumentLock();

lock.waitLock(30000);

try {

  // Re-check while holding the lock.
  // Another execution may have written this Submission ID while this
  // record was being resolved and prepared above.
  if (
    clinicalTracking_submissionExists_(
      trackingSheet,
      submissionId
    )
  ) {

    console.log(
      'Clinical Tracking already contains Submission ID: ' +
      submissionId
    );

    return {
      success: true,
      duplicate: true,
      submissionId: submissionId
    };
  }

  appAppendByHeaders_(
    trackingSheet,
    trackingRecord
  );

  // appAppendByHeaders_ retains header-based, atomic append ownership.
  // Google Sheets dates are stored as true dates; only the NEWLY appended
  // record gets the existing M/D/YYYY visual format. Historical facts and
  // prior row formatting are never altered. Presentation must never turn
  // a successful ledger append into an apparent failed submission.
  try {
    const headers = trackingSheet
      .getRange(1, 1, 1, trackingSheet.getLastColumn())
      .getValues()[0]
      .map(h => String(h || '').trim());
    const dateColumn = headers.indexOf('Date') + 1;
    if (dateColumn > 0) {
      trackingSheet
        .getRange(trackingSheet.getLastRow(), dateColumn)
        .setNumberFormat('m/d/yyyy');
    }
  } catch (formatError) {
    console.log('Nonfatal Clinical Tracking date display-format warning: ' +
      String(formatError && formatError.message || formatError));
  }

} finally {

  lock.releaseLock();
}

console.log(
  'Clinical Tracking record created: ' +
  submissionId
);

return {
  success: true,
  duplicate: false,
  submissionId: submissionId,
  record: trackingRecord
};

}

/**
 * =====================================================================
 * STUDENT LOOKUP
 * =====================================================================
 */
function clinicalTracking_lookupStudent_(
  ss,
  studentId,
  studentName
) {

  const sheet =
    appFindSheet_(
      ss,
      APP_CONFIG.SHEETS.STUDENTS
    );

  if (!sheet) {
    return null;
  }

  const rows =
    appSheetToObjects_(sheet);

  const targetId =
    appNormalize_(studentId);

  const targetName =
    appNormalize_(studentName);

  if (targetId) {

    const byId =
      rows.find(row =>
        appNormalize_(
          row['Student ID']
        ) === targetId
      );

    if (byId) {
      return {
        studentId: byId['Student ID'] || '',
        student: byId['Student'] || '',
        clinicalTerm: byId['Clinical Term'] || '',
        cohort: byId['Cohort'] || '',
        clinicalSite: byId['Clinical Site'] || ''
      };
    }
  }

  if (targetName) {

    const byName =
      rows.find(row =>
        appNormalize_(
          row['Student']
        ) === targetName
      );

    if (byName) {
      return {
        studentId: byName['Student ID'] || '',
        student: byName['Student'] || '',
        clinicalTerm: byName['Clinical Term'] || '',
        cohort: byName['Cohort'] || '',
        clinicalSite: byName['Clinical Site'] || ''
      };
    }
  }

  return null;
}


/**
 * =====================================================================
 * PRODUCTION EXACT REFERENCE RESOLUTION
 * =====================================================================
 *
 * No last-name or legacy matching here.
 */
function clinicalTracking_resolveReference_(
  ss,
  sheetName,
  idHeader,
  nameHeader,
  suppliedId,
  suppliedName
) {

  const idText =
    String(suppliedId || '').trim();

  const nameText =
    String(suppliedName || '').trim();

  if (!idText && !nameText) {
    return null;
  }

  const sheet =
    appFindSheet_(
      ss,
      sheetName
    );

  if (!sheet) {

    if (idText) {
      throw new Error(
        'Cannot validate ' +
        idHeader +
        ' because sheet "' +
        sheetName +
        '" was not found.'
      );
    }

    return null;
  }

  const rows =
    appSheetToObjects_(sheet);

  const idTarget =
    appNormalize_(idText);

  const nameTarget =
    appNormalize_(nameText);

  if (idTarget) {

    const matches =
      rows.filter(row =>
        appNormalize_(
          row[idHeader]
        ) === idTarget
      );

    if (matches.length === 0) {
      throw new Error(
        'Unknown ' +
        idHeader +
        ': ' +
        idText
      );
    }

    if (matches.length > 1) {
      throw new Error(
        'Duplicate ' +
        idHeader +
        ': ' +
        idText
      );
    }

    const match =
      matches[0];

    if (
      nameTarget &&
      appNormalize_(
        match[nameHeader]
      ) !== nameTarget
    ) {
      throw new Error(
        idHeader +
        ' / ' +
        nameHeader +
        ' mismatch: ' +
        idText +
        ' does not belong to "' +
        nameText +
        '".'
      );
    }

    return {
      id:
        String(
          match[idHeader] || ''
        ).trim(),
      name:
        String(
          match[nameHeader] || ''
        ).trim(),
      row: match
    };
  }

  const matches =
    rows.filter(row =>
      appNormalize_(
        row[nameHeader]
      ) === nameTarget
    );

  if (matches.length !== 1) {
    return null;
  }

  return {
    id:
      String(
        matches[0][idHeader] || ''
      ).trim(),
    name:
      String(
        matches[0][nameHeader] || ''
      ).trim(),
    row: matches[0]
  };
}


/**
 * =====================================================================
 * EVALUATOR TYPE LOOKUP
 * =====================================================================
 */
function clinicalTracking_lookupEvaluatorType_(
  ss,
  evaluatorName
) {

  if (!evaluatorName) {
    return '';
  }

  const sheet =
    appFindSheet_(
      ss,
      clinicalTracking_sheetName_(
        'EVALUATORS',
        'Evaluators'
      )
    );

  if (!sheet) {
    return '';
  }

  const rows =
    appSheetToObjects_(sheet);

  const target =
    appNormalize_(evaluatorName);

  const matches =
    rows.filter(row =>
      appNormalize_(
        row['Evaluator']
      ) === target
    );

  if (matches.length !== 1) {
    return '';
  }

  return String(
    matches[0]['Evaluator Type'] || ''
  ).trim();
}


/**
 * =====================================================================
 * HISTORICAL LEGACY SITE MAPPINGS
 * =====================================================================
 *
 * Explicitly approved mappings only.
 */
function clinicalTracking_legacySiteMappings_() {

  return {
    'bs pj': 'SITE-012',
    'bsh': 'SITE-012',
    'lhm': 'SITE-008'
  };
}


/**
 * =====================================================================
 * HISTORICAL LEGACY EVALUATOR MAPPINGS
 * =====================================================================
 *
 * Explicit non-person generic alias.
 *
 * Named evaluator variants are handled separately by unique exact
 * last-name matching.
 */
function clinicalTracking_legacyEvaluatorMappings_() {

  return {
    'ci hospital': 'EV-HCI'
  };
}


/**
 * =====================================================================
 * EXTRACT EVALUATOR LAST NAME
 * =====================================================================
 *
 * Examples:
 *   Amanda Rhinehart -> rhinehart
 *   A. Rhinehart     -> rhinehart
 *   A, Rider         -> rider
 *   E.Stebbins       -> stebbins
 *
 * This is NOT fuzzy matching.
 */
function clinicalTracking_evaluatorLastName_(
  value
) {

  let text =
    String(value || '')
      .trim()
      .toLowerCase();

  if (!text) {
    return '';
  }

  text =
    text.replace(
      /[,]+/g,
      ' '
    );

  text =
    text.replace(
      /\./g,
      ' '
    );

  text =
    text.replace(
      /\s+/g,
      ' '
    ).trim();

  if (!text) {
    return '';
  }

  const parts =
    text.split(' ');

  return parts[
    parts.length - 1
  ] || '';
}


/**
 * =====================================================================
 * BUILD REFERENCE INDEX
 * =====================================================================
 */
function clinicalTracking_buildReferenceIndex_(
  ss,
  sheetName,
  idHeader,
  nameHeader
) {

  const sheet =
    appFindSheet_(
      ss,
      sheetName
    );

  if (!sheet) {
    throw new Error(
      'Required reference sheet not found: ' +
      sheetName
    );
  }

  const rows =
    appSheetToObjects_(sheet);

  const byId = {};
  const byName = {};
  const byLastName = {};

  rows.forEach(row => {

    const id =
      appNormalize_(
        row[idHeader]
      );

    const name =
      appNormalize_(
        row[nameHeader]
      );

    if (id) {

      if (byId[id]) {
        throw new Error(
          'Duplicate ' +
          idHeader +
          ' in ' +
          sheetName +
          ': ' +
          row[idHeader]
        );
      }

      byId[id] = row;
    }

    if (name) {

      if (!byName[name]) {
        byName[name] = [];
      }

      byName[name].push(row);
    }

    if (
      nameHeader === 'Evaluator'
    ) {

      const lastName =
        clinicalTracking_evaluatorLastName_(
          row[nameHeader]
        );

      if (lastName) {

        if (!byLastName[lastName]) {
          byLastName[lastName] = [];
        }

        byLastName[lastName].push(row);
      }
    }
  });

  return {
    byId: byId,
    byName: byName,
    byLastName: byLastName
  };
}


/**
 * =====================================================================
 * HISTORICAL BACKFILL RESOLUTION
 * =====================================================================
 */
function clinicalTracking_resolveBackfillReference_(
  value,
  index,
  idHeader,
  nameHeader,
  legacyMappings,
  allowUniqueLastName
) {

  const displayValue =
    String(value || '').trim();

  if (!displayValue) {
    return {
      match: null,
      method: 'BLANK'
    };
  }

  const normalized =
    appNormalize_(
      displayValue
    );


  // -------------------------------------------------------------------
  // 1. EXACT NAME
  // -------------------------------------------------------------------

  const exactMatches =
    index.byName[normalized] || [];

  if (exactMatches.length === 1) {
    return {
      match: exactMatches[0],
      method: 'EXACT NAME'
    };
  }

  if (exactMatches.length > 1) {
    return {
      match: null,
      method: 'AMBIGUOUS EXACT NAME'
    };
  }


  // -------------------------------------------------------------------
  // 2. EXPLICIT APPROVED LEGACY MAPPING
  // -------------------------------------------------------------------

  const mappedId =
    legacyMappings
      ? legacyMappings[normalized]
      : '';

  if (mappedId) {

    const mapped =
      index.byId[
        appNormalize_(
          mappedId
        )
      ];

    if (!mapped) {
      throw new Error(
        'Approved legacy mapping points to unknown ' +
        idHeader +
        ': ' +
        mappedId
      );
    }

    return {
      match: mapped,
      method: 'APPROVED LEGACY MAPPING'
    };
  }


  // -------------------------------------------------------------------
  // 3. UNIQUE EXACT LAST NAME — EVALUATORS ONLY
  // -------------------------------------------------------------------

  if (
    allowUniqueLastName &&
    nameHeader === 'Evaluator'
  ) {

    const lastName =
      clinicalTracking_evaluatorLastName_(
        displayValue
      );

    if (lastName) {

      const lastNameMatches =
        index.byLastName[
          lastName
        ] || [];

      if (
        lastNameMatches.length === 1
      ) {
        return {
          match:
            lastNameMatches[0],
          method:
            'UNIQUE LAST NAME'
        };
      }

      if (
        lastNameMatches.length > 1
      ) {
        return {
          match: null,
          method:
            'AMBIGUOUS LAST NAME'
        };
      }
    }
  }

  return {
    match: null,
    method: 'UNRESOLVED'
  };
}


/**
 * =====================================================================
 * PREVIEW REFERENCE BACKFILL
 * =====================================================================
 *
 * SAFE. DOES NOT WRITE.
 */
function clinicalTracking_previewReferenceIdBackfill() {

  return clinicalTracking_backfillReferenceIds_(
    false
  );
}


/**
 * =====================================================================
 * APPLY REFERENCE BACKFILL
 * =====================================================================
 *
 * Run preview first.
 *
 * Fills blank:
 *   - Clinical Site ID
 *   - Evaluator ID
 *   - Evaluator Type when evaluator identity resolves
 *   - Examination ID
 *
 * Existing nonblank values are never overwritten.
 * Human-readable historical names are never rewritten.
 */
function clinicalTracking_applyReferenceIdBackfill() {

  return clinicalTracking_backfillReferenceIds_(
    true
  );
}


/**
 * =====================================================================
 * BACKFILL IMPLEMENTATION
 * =====================================================================
 */
function clinicalTracking_backfillReferenceIds_(
  writeChanges
) {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheet =
    appGetSheet_(
      ss,
      APP_CONFIG.SHEETS.CLINICAL_TRACKING
    );

  const lastRow =
    sheet.getLastRow();

  const lastColumn =
    sheet.getLastColumn();

  if (lastRow < 2) {

    const emptyReport = {
      mode:
        writeChanges
          ? 'APPLY'
          : 'PREVIEW',
      rowsReviewed: 0
    };

    console.log(
      JSON.stringify(
        emptyReport,
        null,
        2
      )
    );

    return emptyReport;
  }

  const headers =
    sheet
      .getRange(
        1,
        1,
        1,
        lastColumn
      )
      .getValues()[0];

  const headerMap =
    appHeaderMap_(headers);

  const requiredHeaders = [
    'Clinical Site ID',
    'Clinical Site',
    'Evaluator ID',
    'Evaluator',
    'Evaluator Type',
    'Examination ID',
    'Examination'
  ];

  requiredHeaders.forEach(header => {

    if (
      headerMap[
        appNormalize_(header)
      ] === undefined
    ) {
      throw new Error(
        'Clinical Tracking is missing required header: ' +
        header
      );
    }
  });

  const dataRange =
    sheet.getRange(
      2,
      1,
      lastRow - 1,
      lastColumn
    );

  const values =
    dataRange.getValues();

  const siteIndex =
    clinicalTracking_buildReferenceIndex_(
      ss,
      clinicalTracking_sheetName_(
        'CLINICAL_SITES',
        'Clinical Sites'
      ),
      'Clinical Site ID',
      'Clinical Site'
    );

  const evaluatorIndex =
    clinicalTracking_buildReferenceIndex_(
      ss,
      clinicalTracking_sheetName_(
        'EVALUATORS',
        'Evaluators'
      ),
      'Evaluator ID',
      'Evaluator'
    );

  const examinationIndex =
    clinicalTracking_buildReferenceIndex_(
      ss,
      clinicalTracking_sheetName_(
        'EXAMINATIONS',
        'Examinations'
      ),
      'Examination ID',
      'Examination'
    );

  const siteMappings =
    clinicalTracking_legacySiteMappings_();

  const evaluatorMappings =
    clinicalTracking_legacyEvaluatorMappings_();

  const c = {};

  requiredHeaders.forEach(header => {
    c[header] =
      headerMap[
        appNormalize_(header)
      ];
  });

  const report = {

    mode:
      writeChanges
        ? 'APPLY'
        : 'PREVIEW',

    rowsReviewed:
      values.length,

    fillable: {
      clinicalSiteId: 0,
      evaluatorId: 0,
      evaluatorType: 0,
      examinationId: 0
    },

    resolutionMethods: {
      clinicalSites: {},
      evaluators: {},
      examinations: {}
    },

    unresolved: {
      clinicalSites: {},
      evaluators: {},
      examinations: {}
    },

    existingValuesPreserved: {
      clinicalSiteId: 0,
      evaluatorId: 0,
      evaluatorType: 0,
      examinationId: 0
    },

    cellsWritten: 0
  };


  // ===================================================================
  // PROCESS ROWS
  // ===================================================================

  values.forEach(row => {

    // -----------------------------------------------------------------
    // CLINICAL SITE
    // -----------------------------------------------------------------

    const existingSiteId =
      String(
        row[c['Clinical Site ID']] || ''
      ).trim();

    const siteName =
      String(
        row[c['Clinical Site']] || ''
      ).trim();

    if (existingSiteId) {

      report
        .existingValuesPreserved
        .clinicalSiteId++;

    } else if (siteName) {

      const resolution =
        clinicalTracking_resolveBackfillReference_(
          siteName,
          siteIndex,
          'Clinical Site ID',
          'Clinical Site',
          siteMappings,
          false
        );

      clinicalTracking_countResolution_(
        report.resolutionMethods.clinicalSites,
        resolution.method
      );

      if (resolution.match) {

        const id =
          String(
            resolution.match[
              'Clinical Site ID'
            ] || ''
          ).trim();

        if (id) {
          row[c['Clinical Site ID']] =
            id;

          report
            .fillable
            .clinicalSiteId++;
        }

      } else {

        clinicalTracking_countResolution_(
          report.unresolved.clinicalSites,
          siteName
        );
      }
    }


    // -----------------------------------------------------------------
    // EVALUATOR
    // -----------------------------------------------------------------

    const existingEvaluatorId =
      String(
        row[c['Evaluator ID']] || ''
      ).trim();

    const evaluatorName =
      String(
        row[c['Evaluator']] || ''
      ).trim();

    const existingEvaluatorType =
      String(
        row[c['Evaluator Type']] || ''
      ).trim();

    let evaluatorMatch = null;

    if (existingEvaluatorId) {

      report
        .existingValuesPreserved
        .evaluatorId++;

      evaluatorMatch =
        evaluatorIndex.byId[
          appNormalize_(
            existingEvaluatorId
          )
        ] || null;

    } else if (evaluatorName) {

      const resolution =
        clinicalTracking_resolveBackfillReference_(
          evaluatorName,
          evaluatorIndex,
          'Evaluator ID',
          'Evaluator',
          evaluatorMappings,
          true
        );

      clinicalTracking_countResolution_(
        report.resolutionMethods.evaluators,
        resolution.method
      );

      evaluatorMatch =
        resolution.match;

      if (evaluatorMatch) {

        const id =
          String(
            evaluatorMatch[
              'Evaluator ID'
            ] || ''
          ).trim();

        if (id) {
          row[c['Evaluator ID']] =
            id;

          report
            .fillable
            .evaluatorId++;
        }

      } else {

        clinicalTracking_countResolution_(
          report.unresolved.evaluators,
          evaluatorName
        );
      }
    }

        if (existingEvaluatorType) {

      report
        .existingValuesPreserved
        .evaluatorType++;

    } else if (evaluatorMatch) {

      const evaluatorType =
        String(
          evaluatorMatch[
            'Evaluator Type'
          ] || ''
        ).trim();

      if (evaluatorType) {

        row[c['Evaluator Type']] =
          evaluatorType;

        report
          .fillable
          .evaluatorType++;
      }

    } else if (evaluatorName) {

      // Historical rule:
      // A named evaluator that does not resolve to a known College CI
      // is classified as Hospital CI for HCI purposes.
      //
      // Do NOT assign EV-HCI here. EV-HCI represents the generic
      // Hospital CI identity, not this individual person's identity.

      row[c['Evaluator Type']] =
        'Hospital CI';

      report
        .fillable
        .evaluatorType++;
    }

    // -----------------------------------------------------------------
    // EXAMINATION
    // -----------------------------------------------------------------

    const existingExaminationId =
      String(
        row[c['Examination ID']] || ''
      ).trim();

    const examinationName =
      String(
        row[c['Examination']] || ''
      ).trim();

    if (existingExaminationId) {

      report
        .existingValuesPreserved
        .examinationId++;

    } else if (examinationName) {

      const resolution =
        clinicalTracking_resolveBackfillReference_(
          examinationName,
          examinationIndex,
          'Examination ID',
          'Examination',
          null,
          false
        );

      clinicalTracking_countResolution_(
        report.resolutionMethods.examinations,
        resolution.method
      );

      if (resolution.match) {

        const id =
          String(
            resolution.match[
              'Examination ID'
            ] || ''
          ).trim();

        if (id) {

          row[c['Examination ID']] =
            id;

          report
            .fillable
            .examinationId++;
        }

      } else {

        clinicalTracking_countResolution_(
          report.unresolved.examinations,
          examinationName
        );
      }
    }
  });


  // ===================================================================
  // WRITE ONLY WHEN EXPLICITLY APPLYING
  // ===================================================================

  const totalFillable =
    report.fillable.clinicalSiteId +
    report.fillable.evaluatorId +
    report.fillable.evaluatorType +
    report.fillable.examinationId;

  if (
    writeChanges &&
    totalFillable > 0
  ) {

    dataRange.setValues(
      values
    );

    report.cellsWritten =
      totalFillable;
  }


  console.log(
    JSON.stringify(
      report,
      null,
      2
    )
  );

  return report;
}


/**
 * =====================================================================
 * REPORT COUNTER
 * =====================================================================
 */
function clinicalTracking_countResolution_(
  object,
  key
) {

  const value =
    String(key || 'UNKNOWN');

  if (!object[value]) {
    object[value] = 0;
  }

  object[value]++;
}


/**
 * =====================================================================
 * CONFIGURED SHEET NAME WITH FALLBACK
 * =====================================================================
 */
function clinicalTracking_sheetName_(
  configKey,
  fallbackName
) {

  if (
    APP_CONFIG &&
    APP_CONFIG.SHEETS &&
    APP_CONFIG.SHEETS[configKey]
  ) {
    return APP_CONFIG.SHEETS[configKey];
  }

  return fallbackName;
}


/**
 * =====================================================================
 * SOURCE -> SUBMISSION METHOD
 * =====================================================================
 */
function clinicalTracking_submissionMethod_(
  source
) {

  const text =
    appNormalize_(source);

  if (!text) {
    return 'Other';
  }

  if (text.includes('historical')) {
    return 'Historical Import';
  }

  if (
    text.includes('qr') ||
    text.includes('submission form') ||
    text.includes('uploaded')
  ) {
    return 'Uploaded Evaluation';
  }

  if (text.includes('digital')) {
    return 'Digital Evaluation';
  }

  if (text.includes('legacy')) {
    return 'Legacy Evaluation';
  }

  return String(
    source || 'Other'
  ).trim();
}


/**
 * =====================================================================
 * DUPLICATE PROTECTION
 * =====================================================================
 */
function clinicalTracking_submissionExists_(
  sheet,
  submissionId
) {

  const lastRow =
    sheet.getLastRow();

  const lastColumn =
    sheet.getLastColumn();

  if (
    lastRow < 2 ||
    lastColumn < 1
  ) {
    return false;
  }

  const headers =
    sheet
      .getRange(
        1,
        1,
        1,
        lastColumn
      )
      .getValues()[0];

  const headerMap =
    appHeaderMap_(headers);

  const idCol =
    headerMap['submission id'];

  if (idCol === undefined) {
    throw new Error(
      'Clinical Tracking is missing the Submission ID header.'
    );
  }

  const ids =
    sheet
      .getRange(
        2,
        idCol + 1,
        lastRow - 1,
        1
      )
      .getValues();

  const target =
    String(
      submissionId || ''
    ).trim();

  for (
    let i = 0;
    i < ids.length;
    i++
  ) {

    if (
      String(
        ids[i][0] || ''
      ).trim() === target
    ) {
      return true;
    }
  }

  return false;
}


/**
 * =====================================================================
 * FIRST NON-BLANK VALUE
 * =====================================================================
 */
function clinicalTracking_firstValue_(
  record,
  keys
) {

  for (
    let i = 0;
    i < keys.length;
    i++
  ) {

    const key =
      keys[i];

    if (
      !Object.prototype
        .hasOwnProperty
        .call(
          record,
          key
        )
    ) {
      continue;
    }

    const value =
      record[key];

    if (
      value !== null &&
      value !== undefined &&
      String(value).trim() !== ''
    ) {
      return value;
    }
  }

  return '';
}


/**
 * =====================================================================
 * FIRST VALUE - PRESERVE ZERO
 * =====================================================================
 */
function clinicalTracking_firstValuePreserveZero_(
  record,
  keys
) {

  for (
    let i = 0;
    i < keys.length;
    i++
  ) {

    const key =
      keys[i];

    if (
      !Object.prototype
        .hasOwnProperty
        .call(
          record,
          key
        )
    ) {
      continue;
    }

    const value =
      record[key];

    if (value === 0) {
      return 0;
    }

    if (
      value !== null &&
      value !== undefined &&
      String(value).trim() !== ''
    ) {
      return value;
    }
  }

  return '';
}