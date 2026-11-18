/**
 * ================================================================
 * HISTORICAL CLINICAL RECORD IMPORT SYSTEM
 * Version 2.0
 * ================================================================
 *
 * STAGING SHEET:
 *   Historical Import
 *
 * READS:
 *   Students
 *   Examinations
 *   Clinical Tracking
 *
 * WRITES:
 *   Historical Import
 *   Clinical Tracking
 *
 * WORKFLOW:
 *
 *   Historical source
 *        ↓
 *   Historical Import staging
 *        ↓
 *   prepareHistoricalImportMatches()
 *        ↓
 *   Human review
 *        ↓
 *   Ready to Import = TRUE
 *        ↓
 *   importHistoricalRecords()
 *        ↓
 *   Clinical Tracking
 *
 * IMPORTANT:
 *
 *   - Historical Import remains the audit/review table.
 *   - Clinical Tracking remains the permanent evaluation database.
 *   - The script does NOT silently guess uncertain examination matches.
 * ================================================================
 */


const HIST_IMPORT_CONFIG = {

  sheets: {
    staging: 'Historical Import',
    students: 'Students',
    examinations: 'Examinations',
    tracking: 'Clinical Tracking'
  },

  stagingHeaders: [
    'Import ID',
    'Student ID',
    'Student',
    'Clinical Term',
    'Evaluation Type',
    'Source Examination',
    'Matched Examination ID',
    'Matched Examination',
    'Match Status',
    'Date',
    'Result',
    'Performance Score',
    'Image Score',
    'Final Score',
    'Evaluator',
    'Evaluator Type',
    'Clinical Site',
    'Source Type',
    'Source Document',
    'Source Evaluation ID',
    'Notes',
    'Ready to Import',
    'Imported',
    'Imported Submission ID',
    'Import Error'
  ],

  trackingHeaders: [
    'Submission ID',
    'Student ID',
    'Student',
    'Clinical Term',
    'Attributed Clinical Term',
    'Clinical Site',
    'Evaluator',
    'Evaluator Type',
    'Evaluation Type',
    'Examination',
    'Anatomical Area',
    'Requirement Type',
    'Date',
    'Performance Score',
    'Image Score',
    'Final Score',
    'Historical Percent Grade',
    'Result',
    'Automatic Failure',
    'Counts Towards HCI Limit',
    'HCI Competency #',
    'Remediation Required',
    'Remediation Complete',
    'Submission Method',
    'Source Document',
    'Entered By',
    'Coordinator Review Status',
    'Coordinator Flag',
    'Coordinator Notes',
    'Reviewed By',
    'Review Date'
  ]
};


/**
 * ================================================================
 * STEP 1
 * PREPARE / MATCH HISTORICAL RECORDS
 * ================================================================
 *
 * Safe behavior:
 *
 * If Source Examination EXACTLY matches an active examination,
 * the script fills:
 *
 *   Matched Examination ID
 *   Matched Examination
 *   Match Status = Matched
 *
 * If it cannot safely determine the examination:
 *
 *   Match Status = Review Needed
 *
 * It does NOT guess.
 * ================================================================
 */

function prepareHistoricalImportMatches() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cfg = HIST_IMPORT_CONFIG;

  const stagingSheet =
    histGetSheet_(ss, cfg.sheets.staging);

  const examinationsSheet =
    histGetSheet_(ss, cfg.sheets.examinations);

  histValidateExactHeaders_(
    stagingSheet,
    cfg.stagingHeaders
  );

  const staging =
    histSheetData_(stagingSheet);

  const examinations =
    histSheetToObjects_(examinationsSheet);


  // ------------------------------------------------------------
  // Examination lookups
  // ------------------------------------------------------------

  const examsByName = {};
  const examsById = {};


  examinations.forEach(exam => {

    if (!histTrue_(exam['Active'])) {
      return;
    }

    const nameKey =
      histNormalize_(exam['Examination']);

    const idKey =
      histNormalize_(exam['Examination ID']);


    if (nameKey) {
      examsByName[nameKey] = exam;
    }

    if (idKey) {
      examsById[idKey] = exam;
    }

  });


  let matched = 0;
  let reviewNeeded = 0;


  for (
    let i = 1;
    i < staging.values.length;
    i++
  ) {

    const row = staging.values[i];
    const sheetRow = i + 1;


    if (!histRowHasData_(row)) {
      continue;
    }


    // ----------------------------------------------------------
    // Assign permanent Import ID
    // ----------------------------------------------------------

    histEnsureImportId_(
      stagingSheet,
      row,
      sheetRow,
      staging.col
    );


    // ----------------------------------------------------------
    // Don't change rows already imported
    // ----------------------------------------------------------

    if (
      histTrue_(
        row[
          staging.col['Imported']
        ]
      )
    ) {
      continue;
    }


    const sourceExam =
      String(
        row[
          staging.col['Source Examination']
        ] || ''
      ).trim();


    const existingMatchedId =
      String(
        row[
          staging.col['Matched Examination ID']
        ] || ''
      ).trim();


    const existingMatchedName =
      String(
        row[
          staging.col['Matched Examination']
        ] || ''
      ).trim();


    // ----------------------------------------------------------
    // If user already supplied a valid Examination ID,
    // respect it.
    // ----------------------------------------------------------

    if (existingMatchedId) {

      const exam =
        examsById[
          histNormalize_(existingMatchedId)
        ];


      if (exam) {

        histSetStagingValue_(
          stagingSheet,
          sheetRow,
          staging.col,
          'Matched Examination ID',
          exam['Examination ID']
        );

        histSetStagingValue_(
          stagingSheet,
          sheetRow,
          staging.col,
          'Matched Examination',
          exam['Examination']
        );

        histSetStagingValue_(
          stagingSheet,
          sheetRow,
          staging.col,
          'Match Status',
          'Matched'
        );

        matched++;

        continue;
      }
    }


    // ----------------------------------------------------------
    // If user supplied a valid standardized examination name,
    // accept it.
    // ----------------------------------------------------------

    if (existingMatchedName) {

      const exam =
        examsByName[
          histNormalize_(existingMatchedName)
        ];


      if (exam) {

        histSetStagingValue_(
          stagingSheet,
          sheetRow,
          staging.col,
          'Matched Examination ID',
          exam['Examination ID']
        );

        histSetStagingValue_(
          stagingSheet,
          sheetRow,
          staging.col,
          'Matched Examination',
          exam['Examination']
        );

        histSetStagingValue_(
          stagingSheet,
          sheetRow,
          staging.col,
          'Match Status',
          'Matched'
        );

        matched++;

        continue;
      }
    }


    // ----------------------------------------------------------
    // Exact Source Examination match
    // ----------------------------------------------------------

    const exactExam =
      examsByName[
        histNormalize_(sourceExam)
      ];


    if (exactExam) {

      histSetStagingValue_(
        stagingSheet,
        sheetRow,
        staging.col,
        'Matched Examination ID',
        exactExam['Examination ID']
      );

      histSetStagingValue_(
        stagingSheet,
        sheetRow,
        staging.col,
        'Matched Examination',
        exactExam['Examination']
      );

      histSetStagingValue_(
        stagingSheet,
        sheetRow,
        staging.col,
        'Match Status',
        'Matched'
      );

      matched++;

    } else {

      // --------------------------------------------------------
      // No guessing.
      // --------------------------------------------------------

      histSetStagingValue_(
        stagingSheet,
        sheetRow,
        staging.col,
        'Match Status',
        sourceExam
          ? 'Review Needed'
          : 'Missing Source Examination'
      );

      reviewNeeded++;
    }

  }


  SpreadsheetApp.flush();


  SpreadsheetApp
    .getActive()
    .toast(
      matched +
      ' matched; ' +
      reviewNeeded +
      ' need review.',
      'Historical Matching',
      8
    );
}


/**
 * ================================================================
 * STEP 2
 * IMPORT APPROVED HISTORICAL RECORDS
 * ================================================================
 *
 * Only imports rows where:
 *
 *   Ready to Import = TRUE
 *   Imported != TRUE
 *   Match Status = Matched
 *
 * ================================================================
 */

function importHistoricalRecords() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const cfg =
    HIST_IMPORT_CONFIG;


  const stagingSheet =
    histGetSheet_(
      ss,
      cfg.sheets.staging
    );

  const studentsSheet =
    histGetSheet_(
      ss,
      cfg.sheets.students
    );

  const examinationsSheet =
    histGetSheet_(
      ss,
      cfg.sheets.examinations
    );

  const trackingSheet =
    histGetSheet_(
      ss,
      cfg.sheets.tracking
    );


  // ------------------------------------------------------------
  // Validate structure BEFORE writing
  // ------------------------------------------------------------

  histValidateExactHeaders_(
    stagingSheet,
    cfg.stagingHeaders
  );

  histValidateExactHeaders_(
    trackingSheet,
    cfg.trackingHeaders
  );


  const staging =
    histSheetData_(stagingSheet);

  const students =
    histSheetToObjects_(studentsSheet);

  const examinations =
    histSheetToObjects_(examinationsSheet);

  const tracking =
    histSheetToObjects_(trackingSheet);


  // ------------------------------------------------------------
  // Student lookup
  // ------------------------------------------------------------

  const studentsById = {};


  students.forEach(student => {

    const key =
      histNormalize_(
        student['Student ID']
      );


    if (
      key &&
      !studentsById[key]
    ) {

      studentsById[key] =
        student;

    }

  });


  // ------------------------------------------------------------
  // Examination lookup
  // ------------------------------------------------------------

  const examsById = {};


  examinations.forEach(exam => {

    if (
      !histTrue_(
        exam['Active']
      )
    ) {
      return;
    }


    const key =
      histNormalize_(
        exam['Examination ID']
      );


    if (key) {
      examsById[key] = exam;
    }

  });


  // ------------------------------------------------------------
  // Existing Clinical Tracking IDs
  // ------------------------------------------------------------

  const existingSubmissionIds =
    new Set(
      tracking
        .map(record =>
          histNormalize_(
            record['Submission ID']
          )
        )
        .filter(Boolean)
    );


  const rowsToAppend = [];

  let importedCount = 0;
  let skippedCount = 0;
  let errorCount = 0;


  // ------------------------------------------------------------
  // Process staging rows
  // ------------------------------------------------------------

  for (
    let i = 1;
    i < staging.values.length;
    i++
  ) {

    const row =
      staging.values[i];

    const sheetRow =
      i + 1;


    if (!histRowHasData_(row)) {
      continue;
    }


    // ----------------------------------------------------------
    // Ensure permanent Import ID
    // ----------------------------------------------------------

    const importId =
      histEnsureImportId_(
        stagingSheet,
        row,
        sheetRow,
        staging.col
      );


    // ----------------------------------------------------------
    // Only approved rows
    // ----------------------------------------------------------

    const ready =
      histTrue_(
        row[
          staging.col['Ready to Import']
        ]
      );


    if (!ready) {

      skippedCount++;

      continue;
    }


    // ----------------------------------------------------------
    // Already imported
    // ----------------------------------------------------------

    if (
      histTrue_(
        row[
          staging.col['Imported']
        ]
      )
    ) {

      skippedCount++;

      continue;
    }


    // ----------------------------------------------------------
    // Clear previous error
    // ----------------------------------------------------------

    histSetStagingValue_(
      stagingSheet,
      sheetRow,
      staging.col,
      'Import Error',
      ''
    );


    try {

      // ========================================================
      // MATCH STATUS
      // ========================================================

      const matchStatus =
        String(
          row[
            staging.col['Match Status']
          ] || ''
        ).trim();


      if (
        histNormalize_(matchStatus) !==
        'matched'
      ) {

        throw new Error(
          'Match Status must be Matched before import.'
        );

      }


      // ========================================================
      // STUDENT
      // ========================================================

      const studentId =
        String(
          row[
            staging.col['Student ID']
          ] || ''
        ).trim();


      if (!studentId) {

        throw new Error(
          'Student ID is required.'
        );

      }


      const student =
        studentsById[
          histNormalize_(studentId)
        ];


      if (!student) {

        throw new Error(
          'Student ID "' +
          studentId +
          '" was not found in Students.'
        );

      }


      const officialStudentName =
        String(
          student['Student'] || ''
        ).trim();


      // ========================================================
      // TERM
      // ========================================================

      const clinicalTerm =
        String(
          row[
            staging.col['Clinical Term']
          ] || ''
        ).trim();


      if (
        !/^clinical\s+[1-5]$/i
          .test(clinicalTerm)
      ) {

        throw new Error(
          'Clinical Term must be Clinical 1 through Clinical 5.'
        );

      }


      // ========================================================
      // EVALUATION TYPE
      // ========================================================

      const evaluationType =
        String(
          row[
            staging.col['Evaluation Type']
          ] || ''
        ).trim();


      const allowedTypes = [
        'competency',
        'iso',
        'sso',
        'simulation'
      ];


      if (
        !allowedTypes.includes(
          histNormalize_(
            evaluationType
          )
        )
      ) {

        throw new Error(
          'Evaluation Type must be Competency, ISO, SSO, or Simulation.'
        );

      }


      // ========================================================
      // RESULT
      // ========================================================

      const result =
        String(
          row[
            staging.col['Result']
          ] || ''
        ).trim();


      if (
        ![
          'pass',
          'fail'
        ].includes(
          histNormalize_(result)
        )
      ) {

        throw new Error(
          'Result must be Pass or Fail.'
        );

      }


      // ========================================================
      // EXAMINATION
      // ========================================================

      const matchedExamId =
        String(
          row[
            staging.col['Matched Examination ID']
          ] || ''
        ).trim();


      if (!matchedExamId) {

        throw new Error(
          'Matched Examination ID is required.'
        );

      }


      const exam =
        examsById[
          histNormalize_(
            matchedExamId
          )
        ];


      if (!exam) {

        throw new Error(
          'Matched Examination ID "' +
          matchedExamId +
          '" was not found in Examinations.'
        );

      }


      const matchedExam =
        String(
          exam['Examination'] || ''
        ).trim();


      const requirementType =
        String(
          exam['Requirement Type'] || ''
        ).trim();


      // ========================================================
      // SOURCE / DUPLICATE ID
      // ========================================================

      const sourceEvaluationId =
        String(
          row[
            staging.col['Source Evaluation ID']
          ] || ''
        ).trim();


      let submissionId;


      if (sourceEvaluationId) {

        submissionId =
          'HIST-EVAL-' +
          histSafeId_(
            sourceEvaluationId
          );

      } else {

        submissionId =
          'HIST-' +
          histSafeId_(
            importId
          );

      }


      if (
        existingSubmissionIds.has(
          histNormalize_(
            submissionId
          )
        )
      ) {

        throw new Error(
          'Historical record already exists in Clinical Tracking as ' +
          submissionId +
          '.'
        );

      }


      // ========================================================
      // SOURCE VALUES
      // ========================================================

      const sourceExam =
        String(
          row[
            staging.col['Source Examination']
          ] || ''
        ).trim();


      const sourceType =
        String(
          row[
            staging.col['Source Type']
          ] || ''
        ).trim();


      const sourceDocument =
        String(
          row[
            staging.col['Source Document']
          ] || ''
        ).trim();


      const notes =
        String(
          row[
            staging.col['Notes']
          ] || ''
        ).trim();


      // ========================================================
      // AUDIT NOTE
      // ========================================================

      let auditNote =
        'Historical import';


      if (
        sourceExam &&
        histNormalize_(sourceExam) !==
        histNormalize_(matchedExam)
      ) {

        auditNote +=
          '. Source examination: "' +
          sourceExam +
          '". Matched to: "' +
          matchedExam +
          '".';

      }


      if (sourceEvaluationId) {

        auditNote +=
          ' Source Evaluation ID: ' +
          sourceEvaluationId +
          '.';

      }


      if (sourceType) {

        auditNote +=
          ' Source Type: ' +
          sourceType +
          '.';

      }


      if (notes) {

        auditNote +=
          ' Notes: ' +
          notes;

      }


      // ========================================================
      // BUILD CLINICAL TRACKING ROW
      // ========================================================

      const trackingRow = [

        submissionId,

        studentId,

        officialStudentName,

        clinicalTerm,

        clinicalTerm, // Attributed Clinical Term - historical stored term is authoritative

        String(
          row[
            staging.col['Clinical Site']
          ] || ''
        ).trim(),

        String(
          row[
            staging.col['Evaluator']
          ] || ''
        ).trim(),

        String(
          row[
            staging.col['Evaluator Type']
          ] || ''
        ).trim(),

        evaluationType,

        matchedExam,

        '', // Anatomical Area - existing importer has no source mapping for this field

        requirementType,

        row[
          staging.col['Date']
        ] || '',

        histBlankSafe_(
          row[
            staging.col['Performance Score']
          ]
        ),

        histBlankSafe_(
          row[
            staging.col['Image Score']
          ]
        ),

        histBlankSafe_(
          row[
            staging.col['Final Score']
          ]
        ),

        '', // Historical Percent Grade - not present in current Historical Import staging schema

        result,

        '', // Automatic Failure

        '', // Counts Towards HCI Limit

        '', // HCI Competency #

        '', // Remediation Required

        '', // Remediation Complete

        'Historical Import',

        sourceDocument,

        histCurrentUser_(),

        '', // Coordinator Review Status

        '', // Coordinator Flag

        auditNote,

        '', // Reviewed By

        ''  // Review Date
      ];


      rowsToAppend.push({

        sheetRow,

        submissionId,

        trackingRow
      });


      existingSubmissionIds.add(
        histNormalize_(
          submissionId
        )
      );


    } catch (error) {

      errorCount++;


      histSetStagingValue_(
        stagingSheet,
        sheetRow,
        staging.col,
        'Import Error',
        error.message
      );

    }

  }


  // ------------------------------------------------------------
  // WRITE VALID RECORDS
  // ------------------------------------------------------------

  if (
    rowsToAppend.length > 0
  ) {

    const startRow =
      trackingSheet.getLastRow() + 1;


    const values =
      rowsToAppend.map(item =>
        item.trackingRow
      );


    trackingSheet
      .getRange(
        startRow,
        1,
        values.length,
        cfg.trackingHeaders.length
      )
      .setValues(values);


    // ----------------------------------------------------------
    // Mark staging rows imported
    // ----------------------------------------------------------

    rowsToAppend.forEach(item => {

      histSetStagingValue_(
        stagingSheet,
        item.sheetRow,
        staging.col,
        'Imported',
        true
      );


      histSetStagingValue_(
        stagingSheet,
        item.sheetRow,
        staging.col,
        'Imported Submission ID',
        item.submissionId
      );


      histSetStagingValue_(
        stagingSheet,
        item.sheetRow,
        staging.col,
        'Import Error',
        ''
      );


      importedCount++;
    });

  }


  SpreadsheetApp.flush();


  SpreadsheetApp
    .getActive()
    .toast(
      importedCount +
      ' imported; ' +
      errorCount +
      ' error(s).',
      'Historical Import',
      8
    );


  Logger.log(
    'Historical Import complete. ' +
    importedCount +
    ' imported, ' +
    skippedCount +
    ' skipped, ' +
    errorCount +
    ' errors.'
  );
}


/**
 * ================================================================
 * IMPORT ID
 * ================================================================
 */

function histEnsureImportId_(
  sheet,
  row,
  sheetRow,
  col
) {

  let importId =
    String(
      row[
        col['Import ID']
      ] || ''
    ).trim();


  if (!importId) {

    importId =
      'HI-' +
      Utilities
        .getUuid()
        .replace(/-/g, '')
        .substring(0, 10)
        .toUpperCase();


    histSetStagingValue_(
      sheet,
      sheetRow,
      col,
      'Import ID',
      importId
    );


    row[
      col['Import ID']
    ] = importId;

  }


  return importId;
}


/**
 * ================================================================
 * SHEET DATA
 * ================================================================
 */

function histSheetData_(sheet) {

  const values =
    sheet
      .getDataRange()
      .getValues();


  const headers =
    values[0]
      .map(header =>
        String(header).trim()
      );


  const col = {};


  headers.forEach(
    (header, index) => {

      col[header] =
        index;

    }
  );


  return {
    values,
    headers,
    col
  };
}


/**
 * ================================================================
 * SHEET TO OBJECTS
 * ================================================================
 */

function histSheetToObjects_(sheet) {

  const values =
    sheet
      .getDataRange()
      .getValues();


  if (
    values.length < 2
  ) {
    return [];
  }


  const headers =
    values[0]
      .map(header =>
        String(header).trim()
      );


  return values
    .slice(1)
    .filter(row =>
      histRowHasData_(row)
    )
    .map(row => {

      const object = {};


      headers.forEach(
        (header, index) => {

          object[header] =
            row[index];

        }
      );


      return object;
    });
}


/**
 * ================================================================
 * HEADER VALIDATION
 * ================================================================
 */

function histValidateExactHeaders_(
  sheet,
  expected
) {

  const actual =
    sheet
      .getRange(
        1,
        1,
        1,
        expected.length
      )
      .getValues()[0]
      .map(header =>
        String(header).trim()
      );


  expected.forEach(
    (header, index) => {

      if (
        actual[index] !==
        header
      ) {

        throw new Error(
          'Header mismatch in "' +
          sheet.getName() +
          '" column ' +
          (index + 1) +
          '. Expected "' +
          header +
          '" but found "' +
          actual[index] +
          '".'
        );

      }

    }
  );
}


/**
 * ================================================================
 * SET STAGING VALUE
 * ================================================================
 */

function histSetStagingValue_(
  sheet,
  row,
  col,
  header,
  value
) {

  sheet
    .getRange(
      row,
      col[header] + 1
    )
    .setValue(value);
}


/**
 * ================================================================
 * HELPERS
 * ================================================================
 */

function histGetSheet_(
  ss,
  name
) {

  const sheet =
    ss.getSheetByName(
      name
    );


  if (!sheet) {

    throw new Error(
      'Required sheet not found: "' +
      name +
      '".'
    );

  }


  return sheet;
}


function histNormalize_(value) {

  return String(
    value || ''
  )
    .trim()
    .toLowerCase()
    .replace(
      /\s+/g,
      ' '
    );
}


function histTrue_(value) {

  if (
    value === true
  ) {
    return true;
  }


  const normalized =
    histNormalize_(value);


  return (
    normalized === 'true' ||
    normalized === 'yes' ||
    normalized === '1'
  );
}


function histBlankSafe_(value) {

  if (
    value === null ||
    typeof value === 'undefined'
  ) {

    return '';

  }


  return value;
}


function histRowHasData_(row) {

  return row.some(cell =>
    cell !== '' &&
    cell !== null
  );
}


function histSafeId_(value) {

  return String(
    value || ''
  )
    .trim()
    .replace(
      /[^A-Za-z0-9_-]/g,
      ''
    )
    .toUpperCase();
}


function histCurrentUser_() {

  try {

    return (
      Session
        .getActiveUser()
        .getEmail() ||
      ''
    );

  } catch (error) {

    return '';

  }
}