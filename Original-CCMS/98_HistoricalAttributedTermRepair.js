/**
 * =====================================================================
 * HISTORICAL ATTRIBUTED CLINICAL TERM REPAIR
 * 98_HistoricalAttributedTermRepair.gs
 * =====================================================================
 *
 * PURPOSE
 * -------
 * Repairs Attributed Clinical Term for historical Clinical Tracking rows.
 *
 * RULE
 * ----
 * Clinical Term = academic term the competency was credited toward.
 * Attributed Clinical Term = term in which the evaluation actually occurred,
 * determined strictly from Date + Semester Rules date ranges.
 *
 * SAFETY
 * ------
 * - Only HIST-* Submission IDs are considered.
 * - Clinical Term is NEVER changed.
 * - No examination, evaluator, grade, result, remediation, or other field
 *   is changed.
 * - Preview function performs NO writes.
 * - Apply function writes ONLY Attributed Clinical Term.
 * =====================================================================
 */


/**
 * =====================================================================
 * PREVIEW ONLY
 * =====================================================================
 *
 * Run this FIRST.
 *
 * Returns a report showing:
 * - rows reviewed
 * - rows that would change
 * - rows already correct
 * - rows with no usable date
 * - rows whose date does not fall inside exactly one active Semester Rule
 *
 * MAKES NO CHANGES.
 */
function previewHistoricalAttributedClinicalTermRepair() {

  return historicalAttributedTermRepair_(false);

}


/**
 * =====================================================================
 * APPLY REPAIR
 * =====================================================================
 *
 * Run only after reviewing preview output.
 *
 * Writes ONLY the Attributed Clinical Term column for historical rows
 * whose Date resolves to exactly one Semester Rules term.
 */
function applyHistoricalAttributedClinicalTermRepair() {

  return historicalAttributedTermRepair_(true);

}


/**
 * =====================================================================
 * CORE
 * =====================================================================
 */
function historicalAttributedTermRepair_(writeChanges) {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();


  const trackingSheet =
    appGetSheet_(
      ss,
      APP_CONFIG.SHEETS.CLINICAL_TRACKING
    );


  const semesterRulesSheet =
    appGetSheet_(
      ss,
      APP_CONFIG.SHEETS.SEMESTER_RULES
    );


  // ===================================================================
  // VALIDATE HEADERS
  // ===================================================================

  const trackingHeaders =
    historicalAttributedTerm_getHeaders_(
      trackingSheet
    );


  const requiredTrackingHeaders = [
    'Submission ID',
    'Student ID',
    'Student',
    'Clinical Term',
    'Attributed Clinical Term',
    'Date'
  ];


  requiredTrackingHeaders.forEach(header => {

    if (!trackingHeaders.includes(header)) {

      throw new Error(
        'Clinical Tracking is missing required header: "' +
        header +
        '".'
      );

    }

  });


  const semesterHeaders =
    historicalAttributedTerm_getHeaders_(
      semesterRulesSheet
    );


  const requiredSemesterHeaders = [
    'Clinical Term',
    'Start Date',
    'End Date',
    'Active'
  ];


  requiredSemesterHeaders.forEach(header => {

    if (!semesterHeaders.includes(header)) {

      throw new Error(
        'Semester Rules is missing required header: "' +
        header +
        '".'
      );

    }

  });


  // ===================================================================
  // READ SEMESTER RULES
  // ===================================================================

  const semesterRules =
    appSheetToObjects_(
      semesterRulesSheet
    )
      .filter(rule =>
        appIsTrue_(
          rule['Active']
        )
      )
      .map(rule => {

        const term =
          String(
            rule['Clinical Term'] || ''
          ).trim();


        const start =
          historicalAttributedTerm_dateOnly_(
            rule['Start Date']
          );


        const end =
          historicalAttributedTerm_dateOnly_(
            rule['End Date']
          );


        if (!term || !start || !end) {

          throw new Error(
            'Active Semester Rules row has missing/invalid term dates: ' +
            JSON.stringify(rule)
          );

        }


        return {
          term,
          start,
          end
        };

      });


  if (semesterRules.length === 0) {

    throw new Error(
      'No active Semester Rules were found.'
    );

  }


  // ===================================================================
  // READ CLINICAL TRACKING
  // ===================================================================

  const values =
    trackingSheet
      .getDataRange()
      .getValues();


  if (values.length <= 1) {

    return {
      success: true,
      writeChanges,
      rowsReviewed: 0,
      historicalRows: 0,
      rowsChanged: 0,
      alreadyCorrect: 0,
      missingDate: 0,
      unresolvedDate: 0,
      ambiguousDate: 0,
      changes: [],
      unresolved: []
    };

  }


  const headerMap = {};


  trackingHeaders.forEach(
    (header, index) => {

      headerMap[header] = index;

    }
  );


  const attributedColumn =
    headerMap['Attributed Clinical Term'] + 1;


  let rowsReviewed = 0;
  let historicalRows = 0;
  let rowsChanged = 0;
  let alreadyCorrect = 0;
  let missingDate = 0;
  let unresolvedDate = 0;
  let ambiguousDate = 0;


  const changes = [];
  const unresolved = [];


  // ===================================================================
  // REVIEW HISTORICAL ROWS
  // ===================================================================

  for (
    let r = 1;
    r < values.length;
    r++
  ) {

    rowsReviewed++;


    const row =
      values[r];


    const submissionId =
      String(
        row[
          headerMap['Submission ID']
        ] || ''
      ).trim();


    // Only historical imported records.
    if (
      !/^HIST-/i.test(
        submissionId
      )
    ) {
      continue;
    }


    historicalRows++;


    const studentId =
      String(
        row[
          headerMap['Student ID']
        ] || ''
      ).trim();


    const student =
      String(
        row[
          headerMap['Student']
        ] || ''
      ).trim();


    const clinicalTerm =
      String(
        row[
          headerMap['Clinical Term']
        ] || ''
      ).trim();


    const currentAttributedTerm =
      String(
        row[
          headerMap['Attributed Clinical Term']
        ] || ''
      ).trim();


    const rawDate =
      row[
        headerMap['Date']
      ];


    const eventDate =
      historicalAttributedTerm_dateOnly_(
        rawDate
      );


    if (!eventDate) {

      missingDate++;


      unresolved.push({
        row: r + 1,
        submissionId,
        studentId,
        student,
        clinicalTerm,
        currentAttributedTerm,
        date: rawDate,
        reason: 'Missing or invalid Date'
      });


      continue;
    }


    const matchingTerms =
      semesterRules.filter(rule =>

        eventDate.getTime() >=
          rule.start.getTime() &&

        eventDate.getTime() <=
          rule.end.getTime()

      );


    if (matchingTerms.length === 0) {

      unresolvedDate++;


      unresolved.push({
        row: r + 1,
        submissionId,
        studentId,
        student,
        clinicalTerm,
        currentAttributedTerm,
        date:
          historicalAttributedTerm_formatDate_(
            eventDate
          ),
        reason:
          'Date does not fall within any active Semester Rules range'
      });


      continue;
    }


    if (matchingTerms.length > 1) {

      ambiguousDate++;


      unresolved.push({
        row: r + 1,
        submissionId,
        studentId,
        student,
        clinicalTerm,
        currentAttributedTerm,
        date:
          historicalAttributedTerm_formatDate_(
            eventDate
          ),
        reason:
          'Date falls within more than one active Semester Rules range',
        matchingTerms:
          matchingTerms.map(
            item => item.term
          )
      });


      continue;
    }


    const correctAttributedTerm =
      matchingTerms[0].term;


    if (
      historicalAttributedTerm_normalize_(
        currentAttributedTerm
      ) ===
      historicalAttributedTerm_normalize_(
        correctAttributedTerm
      )
    ) {

      alreadyCorrect++;

      continue;
    }


    changes.push({
      row: r + 1,
      submissionId,
      studentId,
      student,
      date:
        historicalAttributedTerm_formatDate_(
          eventDate
        ),
      clinicalTerm,
      fromAttributedTerm:
        currentAttributedTerm,
      toAttributedTerm:
        correctAttributedTerm
    });


    if (writeChanges) {

      trackingSheet
        .getRange(
          r + 1,
          attributedColumn
        )
        .setValue(
          correctAttributedTerm
        );


      rowsChanged++;

    }

  }


  if (writeChanges) {

    SpreadsheetApp.flush();

  } else {

    rowsChanged =
      changes.length;

  }


  // ===================================================================
  // REPORT
  // ===================================================================

  const report = {

    success: true,

    mode:
      writeChanges
        ? 'APPLY'
        : 'PREVIEW',

    rowsReviewed,

    historicalRows,

    rowsThatNeedChange:
      changes.length,

    rowsChanged:
      writeChanges
        ? rowsChanged
        : 0,

    alreadyCorrect,

    missingDate,

    unresolvedDate,

    ambiguousDate,

    changes,

    unresolved

  };


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
 * HEADER HELPER
 * =====================================================================
 */
function historicalAttributedTerm_getHeaders_(
  sheet
) {

  const lastColumn =
    sheet.getLastColumn();


  if (lastColumn < 1) {
    return [];
  }


  return sheet
    .getRange(
      1,
      1,
      1,
      lastColumn
    )
    .getValues()[0]
    .map(header =>
      String(header).trim()
    );

}


/**
 * =====================================================================
 * DATE HELPERS
 * =====================================================================
 */
function historicalAttributedTerm_dateOnly_(
  value
) {

  if (!value) {
    return null;
  }


  let date;


  if (value instanceof Date) {

    date =
      new Date(
        value.getTime()
      );

  } else {

    date =
      new Date(value);

  }


  if (
    isNaN(
      date.getTime()
    )
  ) {
    return null;
  }


  return new Date(
    date.getFullYear(),
    date.getMonth(),
    date.getDate()
  );

}


function historicalAttributedTerm_formatDate_(
  value
) {

  return Utilities.formatDate(
    value,
    Session.getScriptTimeZone(),
    'M/d/yyyy'
  );

}


/**
 * =====================================================================
 * NORMALIZE
 * =====================================================================
 */
function historicalAttributedTerm_normalize_(
  value
) {

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
