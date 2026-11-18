/**
 * =====================================================================
 * CLINICAL COMPETENCY TRACKER
 * CLINICAL TERM ATTRIBUTION
 * 04_ClinicalTermAttribution.gs
 * =====================================================================
 *
 * PURPOSE
 * -------
 * Determines which clinical semester an evaluation ACTUALLY occurred in
 * based on:
 *
 *   Clinical Tracking -> Date
 *   Semester Rules    -> Start Date / End Date
 *
 * Writes:
 *
 *   Clinical Tracking -> Attributed Clinical Term
 *
 * IMPORTANT
 * ---------
 * This does NOT overwrite Clinical Term.
 *
 * Clinical Term:
 *   Snapshot / term associated with the record when it entered the system.
 *
 * Attributed Clinical Term:
 *   Term determined from the actual evaluation date.
 *
 * THIS FILE DOES NOT:
 *
 *   - grade evaluations
 *   - calculate requirements
 *   - calculate HCI
 *   - determine remediation
 *   - generate documents
 *   - send emails
 * =====================================================================
 */


/**
 * =====================================================================
 * REBUILD ALL ATTRIBUTED CLINICAL TERMS
 * =====================================================================
 *
 * Use this when:
 *
 *   - historical records are imported
 *   - Semester Rules dates change
 *   - Clinical Tracking data is corrected
 *   - a full reconciliation is desired
 */
function rebuildAttributedClinicalTerms() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();


  const trackingSheet =
    appGetSheet_(
      ss,
      APP_CONFIG.SHEETS
        .CLINICAL_TRACKING
    );


  const semesterSheet =
    appGetSheet_(
      ss,
      APP_CONFIG.SHEETS
        .SEMESTER_RULES
    );


  // ===================================================================
  // VALIDATE REQUIRED HEADERS
  // ===================================================================

  clinicalTerm_validateHeaders_(
    trackingSheet,
    [
      'Submission ID',
      'Date',
      'Clinical Term',
      'Attributed Clinical Term'
    ]
  );


  clinicalTerm_validateHeaders_(
    semesterSheet,
    [
      'Clinical Term',
      'Start Date',
      'End Date',
      'Active'
    ]
  );


  // ===================================================================
  // LOAD ACTIVE SEMESTER WINDOWS
  // ===================================================================

  const semesterRules =
    clinicalTerm_getSemesterWindows_(
      semesterSheet
    );


  if (
    semesterRules.length === 0
  ) {

    throw new Error(
      'No active Semester Rules were found.'
    );
  }


  // ===================================================================
  // READ CLINICAL TRACKING
  // ===================================================================

  const data =
    trackingSheet
      .getDataRange()
      .getValues();


  if (
    data.length < 2
  ) {

    console.log(
      'Clinical Tracking has no data rows.'
    );


    return {

      success:
        true,

      updated:
        0,

      unattributed:
        0

    };
  }


  const headers =
    data[0];


  const headerMap =
    appHeaderMap_(
      headers
    );


  const submissionIdCol =
    headerMap[
      'submission id'
    ];


  const dateCol =
    headerMap[
      'date'
    ];


  const attributedTermCol =
    headerMap[
      'attributed clinical term'
    ];


  // ===================================================================
  // CALCULATE ATTRIBUTED TERM FOR EACH ROW
  // ===================================================================

  const output = [];


  let attributedCount =
    0;


  let unattributedCount =
    0;


  for (
    let r = 1;
    r < data.length;
    r++
  ) {

    const row =
      data[r];


    const submissionId =
      String(
        row[
          submissionIdCol
        ] || ''
      ).trim();


    const evaluationDate =
      row[
        dateCol
      ];


    /*
     * Empty rows remain blank.
     */
    if (
      !submissionId &&
      !evaluationDate
    ) {

      output.push([
        ''
      ]);

      continue;
    }


    const attributedTerm =
      clinicalTerm_getForDateFromWindows_(
        evaluationDate,
        semesterRules
      );


    if (
      attributedTerm ===
      'UNATTRIBUTED'
    ) {

      unattributedCount++;

    } else if (
      attributedTerm
    ) {

      attributedCount++;
    }


    output.push([
      attributedTerm
    ]);
  }


  // ===================================================================
  // WRITE ENTIRE COLUMN IN ONE OPERATION
  // ===================================================================

  if (
    output.length > 0
  ) {

    trackingSheet
      .getRange(
        2,
        attributedTermCol + 1,
        output.length,
        1
      )
      .setValues(
        output
      );
  }


  console.log(
    '================================'
  );

  console.log(
    'CLINICAL TERM ATTRIBUTION COMPLETE'
  );

  console.log(
    'Attributed: ' +
    attributedCount
  );

  console.log(
    'Unattributed: ' +
    unattributedCount
  );

  console.log(
    '================================'
  );


  return {

    success:
      true,

    updated:
      attributedCount,

    unattributed:
      unattributedCount

  };
}


/**
 * =====================================================================
 * GET ATTRIBUTED TERM FOR ONE DATE
 * =====================================================================
 *
 * This is the reusable production function.
 *
 * New evaluation intake will eventually be able to call:
 *
 *   clinicalTerm_getForDate(evaluationDate)
 *
 * instead of rebuilding the entire column every time.
 */
function clinicalTerm_getForDate(
  evaluationDate
) {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();


  const semesterSheet =
    appGetSheet_(
      ss,
      APP_CONFIG.SHEETS
        .SEMESTER_RULES
    );


  const semesterRules =
    clinicalTerm_getSemesterWindows_(
      semesterSheet
    );


  return clinicalTerm_getForDateFromWindows_(
    evaluationDate,
    semesterRules
  );
}


/**
 * =====================================================================
 * LOAD ACTIVE SEMESTER WINDOWS
 * =====================================================================
 *
 * Returns:
 *
 * [
 *   {
 *     clinicalTerm: 'Clinical 1',
 *     startDate: Date,
 *     endDate: Date
 *   },
 *   ...
 * ]
 */
function clinicalTerm_getSemesterWindows_(
  semesterSheet
) {

  const rows =
    appSheetToObjects_(
      semesterSheet
    );


  const windows =
    rows

      .filter(row =>

        appIsTrue_(
          row['Active']
        )

      )

      .map(row => {

        const clinicalTerm =
          String(
            row[
              'Clinical Term'
            ] || ''
          ).trim();


        const startDate =
          appDateOnly_(
            row[
              'Start Date'
            ]
          );


        const endDate =
          appDateOnly_(
            row[
              'End Date'
            ]
          );


        if (
          !clinicalTerm
        ) {

          throw new Error(
            'Semester Rules contains an active row without a Clinical Term.'
          );
        }


        if (
          !startDate ||
          !endDate
        ) {

          throw new Error(
            'Invalid Start Date or End Date in Semester Rules for "' +
            clinicalTerm +
            '".'
          );
        }


        if (
          endDate.getTime() <
          startDate.getTime()
        ) {

          throw new Error(
            'Semester Rules has an End Date before Start Date for "' +
            clinicalTerm +
            '".'
          );
        }


        return {

          clinicalTerm:
            clinicalTerm,

          startDate:
            startDate,

          endDate:
            endDate

        };
      });


  windows.sort(
    (a, b) =>

      a.startDate.getTime() -
      b.startDate.getTime()
  );


  return windows;
}


/**
 * =====================================================================
 * DATE → CLINICAL TERM
 * =====================================================================
 *
 * Semester windows are inclusive:
 *
 *   Start Date <= Evaluation Date <= End Date
 *
 * If no semester contains the date:
 *
 *   UNATTRIBUTED
 */
function clinicalTerm_getForDateFromWindows_(
  evaluationDate,
  semesterRules
) {

  const date =
    appDateOnly_(
      evaluationDate
    );


  if (!date) {

    return (
      'UNATTRIBUTED'
    );
  }


  const matches =
    semesterRules.filter(rule =>

      appDateWithin_(
        date,
        rule.startDate,
        rule.endDate
      )

    );


  if (
    matches.length === 0
  ) {

    return (
      'UNATTRIBUTED'
    );
  }


  /*
   * Overlapping semester windows are a configuration problem.
   *
   * Do not silently guess which semester should win.
   */
  if (
    matches.length > 1
  ) {

    throw new Error(
      'Evaluation date ' +
      Utilities.formatDate(
        date,
        Session.getScriptTimeZone(),
        'yyyy-MM-dd'
      ) +
      ' matches more than one active Semester Rules window: ' +
      matches
        .map(
          rule =>
            rule.clinicalTerm
        )
        .join(', ')
    );
  }


  return (
    matches[0]
      .clinicalTerm
  );
}


/**
 * =====================================================================
 * REQUIRED HEADER VALIDATION
 * =====================================================================
 */
function clinicalTerm_validateHeaders_(
  sheet,
  requiredHeaders
) {

  const lastColumn =
    sheet.getLastColumn();


  if (
    lastColumn < 1
  ) {

    throw new Error(
      'Sheet "' +
      sheet.getName() +
      '" has no headers.'
    );
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
    appHeaderMap_(
      headers
    );


  const missing =
    requiredHeaders.filter(
      header =>

        headerMap[
          appNormalize_(
            header
          )
        ] ===
        undefined
    );


  if (
    missing.length > 0
  ) {

    throw new Error(
      'Sheet "' +
      sheet.getName() +
      '" is missing required header(s): ' +
      missing.join(', ')
    );
  }
}