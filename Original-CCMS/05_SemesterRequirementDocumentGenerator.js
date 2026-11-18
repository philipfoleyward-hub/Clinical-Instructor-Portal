/**
 * =====================================================================
 * CLINICAL COMPETENCY TRACKER
 * SEMESTER REQUIREMENT DOCUMENT GENERATOR
 * 05_SemesterRequirementDocumentGenerator.gs
 * =====================================================================
 *
 * PURPOSE
 * -------
 * Generates the official Semester Requirement Sheet for a student.
 *
 * CURRENTLY VALIDATED
 * -------------------
 * Clinical 2
 *
 * FUTURE
 * ------
 * Clinical 1, 3, 4 and 5 templates can be added to APP_CONFIG and
 * their layouts added to semesterDocument_getTemplateLayout_().
 *
 * THIS FILE OWNS
 * --------------
 * - Semester Requirement Sheet generation
 * - Selecting factual records needed by the document
 * - Populating semester requirement template tokens
 * - HCI numbering/display on the document
 * - Creating the student's generated Google Doc
 *
 * THIS FILE DOES NOT
 * ------------------
 * - Grade individual evaluations
 * - Write Clinical Tracking records
 * - Modify Forms
 * - Calculate the master Student Requirement Status sheet
 * - Send email
 * - Decide Drive folder structure
 *
 * Drive location is delegated to:
 *
 *   05_DriveService.gs
 *
 * =====================================================================
 */


/**
 * =====================================================================
 * MAIN PUBLIC FUNCTION
 * =====================================================================
 *
 * This name is intentionally preserved because Dashboard already calls:
 *
 *   generateSemesterRequirementSheet(studentId, clinicalTerm)
 *
 * Returns:
 *
 * {
 *   success: true,
 *   name: "...",
 *   url: "...",
 *   documentId: "...",
 *   counts: {...}
 * }
 */
function generateSemesterRequirementSheet(
  studentId,
  clinicalTerm
) {

  if (!studentId) {

    throw new Error(
      'Student ID is required.'
    );
  }


  if (!clinicalTerm) {

    throw new Error(
      'Clinical term is required.'
    );
  }


  const layout =
    semesterDocument_getTemplateLayout_(
      clinicalTerm
    );


  if (!layout) {

    throw new Error(
      'Semester Requirement generation is not configured for ' +
      clinicalTerm +
      '.'
    );
  }


  const templateId =
    layout.templateId;


  if (!templateId) {

    throw new Error(
      clinicalTerm +
      ' Semester Requirement template ID has not been configured.'
    );
  }


  // ===================================================================
  // WORKBOOK DATA
  // ===================================================================

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();


  const students =
    appSheetToObjects_(

      appGetSheet_(
        ss,
        APP_CONFIG.SHEETS.STUDENTS
      )

    );


  const tracking =
    appSheetToObjects_(

      appGetSheet_(
        ss,
        APP_CONFIG.SHEETS.CLINICAL_TRACKING
      )

    );


  const semesterRules =
    appSheetToObjects_(

      appGetSheet_(
        ss,
        APP_CONFIG.SHEETS.SEMESTER_RULES
      )

    );


  const evaluators =
    appSheetToObjects_(

      appGetSheet_(
        ss,
        APP_CONFIG.SHEETS.EVALUATORS
      )

    );


  // ===================================================================
  // STUDENT
  // ===================================================================

  const studentRows =
    students.filter(row =>

      appNormalize_(
        row['Student ID']
      ) ===
      appNormalize_(
        studentId
      )

    );


  if (
    studentRows.length === 0
  ) {

    throw new Error(
      'Student not found: ' +
      studentId
    );
  }


  /*
   * A student may have more than one Students row because of
   * site assignments.
   *
   * Prefer an active row for the requested clinical term.
   */
  const student =
    semesterDocument_selectStudentRow_(
      studentRows,
      clinicalTerm
    );


  const studentName =
    String(
      student['Student'] || ''
    ).trim();


  const cohort =
    String(
      student['Cohort'] || ''
    ).trim();


  if (!studentName) {

    throw new Error(
      'Student name is missing for Student ID: ' +
      studentId
    );
  }


  if (!cohort) {

    throw new Error(
      'Student has no Cohort assigned: ' +
      studentName
    );
  }


  const displayStudentId =
    semesterDocument_formatStudentId_(
      studentId
    );


  // ===================================================================
  // SEMESTER RULE
  // ===================================================================

  const rule =
    semesterRules.find(row =>

      appNormalize_(
        row['Clinical Term']
      ) ===
      appNormalize_(
        clinicalTerm
      ) &&

      appIsTrue_(
        row['Active']
      )

    );


  if (!rule) {

    throw new Error(
      'No active Semester Rules row found for ' +
      clinicalTerm +
      '.'
    );
  }


  const termStart =
    appDateOnly_(
      rule['Start Date']
    );


  const termEnd =
    appDateOnly_(
      rule['End Date']
    );


  if (
    !termStart ||
    !termEnd
  ) {

    throw new Error(
      clinicalTerm +
      ' has an invalid Start Date or End Date in Semester Rules.'
    );
  }


  // ===================================================================
  // NEXT CLINICAL TERM
  // ===================================================================

  const currentClinicalNumber =
    semesterDocument_clinicalNumber_(
      clinicalTerm
    );


  const nextClinicalTerm =
    currentClinicalNumber > 0

      ? 'Clinical ' +
        (
          currentClinicalNumber + 1
        )

      : '';


  // ===================================================================
  // STUDENT CLINICAL TRACKING RECORDS
  // ===================================================================

  const records =
    tracking.filter(row =>

      semesterDocument_sameStudent_(
        row,
        studentId,
        studentName
      )

    );


  // ===================================================================
  // ACTIVE COLLEGE CI LIST
  // ===================================================================

  const collegeCIs =
    evaluators.filter(row =>

      appIsTrue_(
        row['Active']
      ) &&

      appNormalize_(
        row['Evaluator Type']
      ) ===
      'college ci'

    );


  // ===================================================================
  // SEMESTER COMPETENCIES
  //
  // Passing competencies CREDITED to the selected term.
  //
  // This intentionally includes carryover credited to this term.
  // ===================================================================

  const semesterCompetencies =
    records

      .filter(record =>

        appNormalize_(
          record['Clinical Term']
        ) ===
        appNormalize_(
          clinicalTerm
        ) &&

        semesterDocument_isPassingCompetency_(
          record
        )

      )

      .sort(
        semesterDocument_recordDateSort_
      );


  // ===================================================================
  // ADDITIONAL TOWARD NEXT SEMESTER
  //
  // Example for Clinical 2:
  //
  //   - credited to Clinical 3
  //   - physically completed during Clinical 2 date window
  //
  // This preserves the logic of the known-good Clinical 2 generator.
  // ===================================================================

  const additionalCompetencies =
    nextClinicalTerm

      ? records

          .filter(record =>

            appNormalize_(
              record['Clinical Term']
            ) ===
            appNormalize_(
              nextClinicalTerm
            ) &&

            semesterDocument_isPassingCompetency_(
              record
            ) &&

            appDateWithin_(
              record['Date'],
              termStart,
              termEnd
            )

          )

          .sort(
            semesterDocument_recordDateSort_
          )

      : [];


  // ===================================================================
  // ISO
  //
  // ISO belongs to the semester in which it was physically performed.
  // ===================================================================

  const isoRecords =
    records

      .filter(record =>

        appNormalize_(
          record['Evaluation Type']
        ) ===
        'iso' &&

        appDateWithin_(
          record['Date'],
          termStart,
          termEnd
        )

      )

      .sort(
        semesterDocument_recordDateSort_
      );


  // ===================================================================
  // FAILED COMPETENCIES
  //
  // Failure belongs to the semester in which it occurred.
  // ===================================================================

  const failedRecords =
    records

      .filter(record =>

        semesterDocument_isFailedCompetency_(
          record
        ) &&

        appDateWithin_(
          record['Date'],
          termStart,
          termEnd
        )

      )

      .sort(
        semesterDocument_recordDateSort_
      );


  // ===================================================================
  // HCI NUMBERING
  //
  // Carryover from before this semester does NOT consume this
  // semester's HCI slots.
  //
  // Additional/future competencies physically completed during this
  // semester DO consume an HCI slot when completed by a non-College CI.
  // ===================================================================

  const hciNumberMap =
    semesterDocument_buildHCINumberMap_({

      records:
        records,

      termStart:
        termStart,

      termEnd:
        termEnd,

      collegeCIs:
        collegeCIs

    });


  // ===================================================================
  // TEMPLATE REPLACEMENTS
  // ===================================================================

  const replacements = {};


  // -------------------------------------------------------------------
  // STUDENT HEADER
  // -------------------------------------------------------------------

  replacements[
    layout.studentNameToken
  ] =
    studentName;


  replacements[
    layout.studentIdToken
  ] =
    displayStudentId;


  // -------------------------------------------------------------------
  // ISO ROWS
  // -------------------------------------------------------------------

  for (
    let i = 1;
    i <= layout.isoRows;
    i++
  ) {

    const record =
      isoRecords[
        i - 1
      ] || null;


    replacements[
      '{{I' + i + 'E}}'
    ] =
      record
        ? semesterDocument_displayExam_(
            record
          )
        : '';


    replacements[
      '{{I' + i + layout.evaluatorTokenSuffix + '}}'
    ] =
      record
        ? semesterDocument_displayEvaluator_(
            record
          )
        : '';


    replacements[
      '{{I' + i + 'G}}'
    ] =
      record
        ? semesterDocument_displayScore_(
            record
          )
        : '';


    replacements[
      '{{I' + i + 'D}}'
    ] =
      record
        ? semesterDocument_displayDate_(
            record['Date']
          )
        : '';
  }


  // -------------------------------------------------------------------
  // FAILED COMPETENCY ROWS
  // -------------------------------------------------------------------

  for (
    let i = 1;
    i <= layout.failureRows;
    i++
  ) {

    const record =
      failedRecords[
        i - 1
      ] || null;


    replacements[
      '{{F' + i + 'E}}'
    ] =
      record
        ? semesterDocument_displayExam_(
            record
          )
        : '';


    replacements[
      '{{F' + i + layout.evaluatorTokenSuffix + '}}'
    ] =
      record
        ? semesterDocument_displayEvaluator_(
            record
          )
        : '';


    replacements[
      '{{F' + i + 'G}}'
    ] =
      record
        ? semesterDocument_displayScore_(
            record
          )
        : '';


    replacements[
      '{{F' + i + 'D}}'
    ] =
      record
        ? semesterDocument_displayDate_(
            record['Date']
          )
        : '';
  }


  // -------------------------------------------------------------------
  // SEMESTER COMPETENCY ROWS
  // -------------------------------------------------------------------

  for (
    let i = 1;
    i <= layout.competencyRows;
    i++
  ) {

    const record =
      semesterCompetencies[
        i - 1
      ] || null;


    semesterDocument_populateCompetencyTokens_({

      replacements:
        replacements,

      prefix:
        'C' + i,

      record:
        record,

      evaluatorTokenSuffix:
        layout.evaluatorTokenSuffix,

      hciNumberMap:
        hciNumberMap,

      termStart:
        termStart,

      termEnd:
        termEnd

    });
  }


  // -------------------------------------------------------------------
  // ADDITIONAL / FUTURE COMPETENCY ROWS
  // -------------------------------------------------------------------

  for (
    let i = 1;
    i <= layout.additionalRows;
    i++
  ) {

    const record =
      additionalCompetencies[
        i - 1
      ] || null;


    semesterDocument_populateCompetencyTokens_({

      replacements:
        replacements,

      prefix:
        'A' + i,

      record:
        record,

      evaluatorTokenSuffix:
        layout.evaluatorTokenSuffix,

      hciNumberMap:
        hciNumberMap,

      termStart:
        termStart,

      termEnd:
        termEnd

    });
  }


  // ===================================================================
  // VERIFY TEMPLATE
  // ===================================================================

  const templateFile =
    DriveApp.getFileById(
      templateId
    );


  if (
    templateFile.getMimeType() !==
    'application/vnd.google-apps.document'
  ) {

    throw new Error(
      clinicalTerm +
      ' Semester Requirement template is not a native Google Doc.'
    );
  }


  // ===================================================================
  // OUTPUT NAME
  // ===================================================================

  const outputName =
    studentName +
    ' - ' +
    clinicalTerm +
    ' Semester Requirements';


  // ===================================================================
  // COPY TEMPLATE TO STUDENT GENERATED DOCUMENTS FOLDER
  // ===================================================================

  const copyResult =
    driveService_copyTemplateToStudentFolder({

      templateId:
        templateId,

      fileName:
        outputName,

      student:
        studentName,

      cohort:
        cohort,

      destination:
        'documents'

    });


  // ===================================================================
  // OPEN COPY
  // ===================================================================

  const document =
    semesterDocument_openWithRetry_(
      copyResult.fileId
    );


  const body =
    document.getBody();


  // ===================================================================
  // REPLACE TOKENS
  // ===================================================================

  Object
    .keys(
      replacements
    )
    .forEach(token => {

      const isStudentName =
        token ===
        layout.studentNameToken;

      const isStudentId =
        token ===
        layout.studentIdToken;

      const isGrade =
        /\{\{(?:C|A|I|F)\d+G\}\}/
          .test(
            token
          );

      semesterDocument_replaceFormattedToken_(
        body,
        token,
        replacements[token],
        {
          fontSize:
            isStudentName
              ? (layout.studentHeaderFontSize || 16)
              : isStudentId
                ? (layout.studentHeaderFontSize || 16)
                : (layout.detailFontSize || 10),

          bold:
            isStudentName ||
            isStudentId ||
            isGrade
        }
      );

    });


  document.saveAndClose();


  // ===================================================================
  // RESULT
  // ===================================================================

  console.log(
    '================================'
  );

  console.log(
    'SEMESTER REQUIREMENT SHEET CREATED'
  );

  console.log(
    'Student: ' +
    studentName
  );

  console.log(
    'Clinical Term: ' +
    clinicalTerm
  );

  console.log(
    'Document: ' +
    copyResult.fileUrl
  );

  console.log(
    '================================'
  );


  return {

    success:
      true,

    name:
      outputName,

    url:
      copyResult.fileUrl,

    documentId:
      copyResult.fileId,

    counts: {

      semesterCompetencies:
        semesterCompetencies.length,

      additionalCompetencies:
        additionalCompetencies.length,

      iso:
        isoRecords.length,

      failures:
        failedRecords.length,

      hci:
        Object
          .keys(
            hciNumberMap
          )
          .length

    }

  };
}


/**
 * =====================================================================
 * TEMPLATE / LAYOUT CONFIGURATION
 * =====================================================================
 *
 * APP_CONFIG stores the actual template IDs.
 *
 * This function defines the STRUCTURE expected inside each template.
 *
 * Only Clinical 2 has been validated so far.
 */
function semesterDocument_getTemplateLayout_(
  clinicalTerm
) {

  const normalized =
    appNormalize_(
      clinicalTerm
    );


  const layouts = {

    'clinical 1': {
      templateId: '1nNQ0Ythz5atDNMlHZ5yn3w08OSHlfGjZrAKQSfiyhLA',
      competencyRows: 3,
      additionalRows: 3,
      isoRows: 4,
      failureRows: 5,
      studentNameToken: '{{STUDENT_NAME}}',
      studentIdToken: '{{STUDENT_NO}}',
      studentHeaderFontSize: 24,
      evaluatorTokenSuffix: 'I'
    },

    'clinical 2': {
      templateId: '1KrZL7kH_LtdqLmSe9BP9SFF-XRfWRe0D6o8eN3v9Gcc',
      competencyRows: 13,
      additionalRows: 10,
      isoRows: 9,
      failureRows: 7,
      studentNameToken: '{{NAME}}',
      studentIdToken: '{{ID}}',
      studentHeaderFontSize: 14,
      detailFontSize: 9,
      evaluatorTokenSuffix: 'N'
    },

    'clinical 3': {
      templateId: '1JMq0xF3hav_nD8vY7MVQS-e_OYGHPa69YA7j1WXytV4',
      competencyRows: 10,
      additionalRows: 11,
      isoRows: 9,
      failureRows: 7,
      studentNameToken: '{{STUDENT_NAME}}',
      studentIdToken: '{{STUDENT_NO}}',
      evaluatorTokenSuffix: 'I'
    },

    'clinical 4': {
      templateId: '1-hpHhMD6rM43OKKsOCqCnVKeAzWaZdydG5ypS2hPJxU',
      competencyRows: 13,
      additionalRows: 15,
      isoRows: 9,
      failureRows: 7,
      studentNameToken: '{{STUDENT_NAME}}',
      studentIdToken: '{{STUDENT_NO}}',
      evaluatorTokenSuffix: 'I'
    },

    'clinical 5': {
      templateId: '1ab_UwXewYAoyqIts3gifvBukkVUMWbPaitGBoBZK-_I',
      competencyRows: 20,
      additionalRows: 0,
      isoRows: 9,
      failureRows: 7,
      studentNameToken: '{{STUDENT_NAME}}',
      studentIdToken: '{{STUDENT_NO}}',
      evaluatorTokenSuffix: 'I'
    }

  };


  return (
    layouts[normalized] ||
    null
  );
}

/**
 * =====================================================================
 * SELECT STUDENT ROW
 * =====================================================================
 *
 * Students can appear more than once because clinical-site assignments
 * may create multiple student rows.
 *
 * Preferred order:
 *
 *   1. Active + requested Clinical Term
 *   2. Requested Clinical Term
 *   3. Active
 *   4. First matching Student ID row
 */
function semesterDocument_selectStudentRow_(
  rows,
  clinicalTerm
) {

  const requestedTerm =
    appNormalize_(
      clinicalTerm
    );


  let match =
    rows.find(row =>

      appIsTrue_(
        row['Active']
      ) &&

      appNormalize_(
        row['Clinical Term']
      ) ===
      requestedTerm

    );


  if (match) {
    return match;
  }


  match =
    rows.find(row =>

      appNormalize_(
        row['Clinical Term']
      ) ===
      requestedTerm

    );


  if (match) {
    return match;
  }


  match =
    rows.find(row =>

      appIsTrue_(
        row['Active']
      )

    );


  return (
    match ||
    rows[0]
  );
}


/**
 * =====================================================================
 * POPULATE ONE COMPETENCY ROW
 * =====================================================================
 *
 * Expected tokens:
 *
 *   {{C1E}}  Examination
 *   {{C1N}}  Evaluator
 *   {{C1H}}  Hospital CI / HCI number
 *   {{C1G}}  Grade
 *   {{C1D}}  Date
 *
 * Additional rows use A instead of C:
 *
 *   {{A1E}}
 *   {{A1N}}
 *   {{A1H}}
 *   {{A1G}}
 *   {{A1D}}
 */
function semesterDocument_populateCompetencyTokens_(
  args
) {

  const replacements =
    args.replacements;


  const prefix =
    args.prefix;


  const record =
    args.record;


  const evaluatorTokenSuffix =
    args.evaluatorTokenSuffix || 'N';


  const hciNumberMap =
    args.hciNumberMap;


  const termStart =
    args.termStart;


  const termEnd =
    args.termEnd;


  replacements[
    '{{' + prefix + 'E}}'
  ] =
    record

      ? semesterDocument_displayExam_(
          record
        )

      : '';


  replacements[
    '{{' + prefix + evaluatorTokenSuffix + '}}'
  ] =
    record

      ? semesterDocument_displayEvaluator_(
          record
        )

      : '';


  replacements[
    '{{' + prefix + 'G}}'
  ] =
    record

      ? semesterDocument_displayScore_(
          record
        )

      : '';


  replacements[
    '{{' + prefix + 'D}}'
  ] =
    record

      ? semesterDocument_displayDate_(
          record['Date']
        )

      : '';


  // -------------------------------------------------------------------
  // EMPTY ROW
  // -------------------------------------------------------------------

  if (!record) {

    replacements[
      '{{' + prefix + 'H}}'
    ] =
      '';


    return;
  }


  // -------------------------------------------------------------------
  // CARRYOVER
  //
  // A competency physically performed before the selected semester
  // does not consume an HCI slot for the selected semester.
  // -------------------------------------------------------------------

  if (
    !appDateWithin_(
      record['Date'],
      termStart,
      termEnd
    )
  ) {

    replacements[
      '{{' + prefix + 'H}}'
    ] =
      'No';


    return;
  }


  // -------------------------------------------------------------------
  // CURRENT-TERM HCI
  // -------------------------------------------------------------------

  const key =
    semesterDocument_recordKey_(
      record
    );


  if (
    Object.prototype
      .hasOwnProperty
      .call(
        hciNumberMap,
        key
      )
  ) {

    replacements[
      '{{' + prefix + 'H}}'
    ] =
      'Yes #' +
      hciNumberMap[key];

  } else {

    replacements[
      '{{' + prefix + 'H}}'
    ] =
      'No';
  }
}


/**
 * =====================================================================
 * BUILD HCI NUMBER MAP
 * =====================================================================
 *
 * Qualifying record:
 *
 *   - passing competency
 *   - physically completed inside selected semester
 *   - evaluator is NOT a College CI
 *
 * Failures do not consume HCI slots.
 */
function semesterDocument_buildHCINumberMap_(
  args
) {

  const records =
    args.records || [];


  const termStart =
    args.termStart;


  const termEnd =
    args.termEnd;


  const collegeCIs =
    args.collegeCIs || [];


  const qualifying =
    records

      .filter(record =>

        semesterDocument_isPassingCompetency_(
          record
        ) &&

        appDateWithin_(
          record['Date'],
          termStart,
          termEnd
        ) &&

        !semesterDocument_isCollegeCI_(
          record,
          collegeCIs
        )

      )

      .sort(
        semesterDocument_recordDateSort_
      );


  const map = {};


  qualifying.forEach(
    (
      record,
      index
    ) => {

      map[
        semesterDocument_recordKey_(
          record
        )
      ] =
        index + 1;

    }
  );


  return map;
}


/**
 * =====================================================================
 * COLLEGE CI DETERMINATION
 * =====================================================================
 *
 * Priority:
 *
 *   1. Explicit Evaluator Type = College CI
 *   2. Explicit Evaluator Type = Hospital CI
 *   3. Match evaluator name against authoritative College CI list
 *
 * Anything else is treated as non-College CI for HCI purposes,
 * preserving the behavior of the validated Clinical 2 generator.
 */
function semesterDocument_isCollegeCI_(
  record,
  collegeCIs
) {

  const explicitType =
    appNormalize_(
      record[
        'Evaluator Type'
      ]
    );


  if (
    explicitType ===
    'college ci'
  ) {

    return true;
  }


  if (
    explicitType ===
    'hospital ci'
  ) {

    return false;
  }


  const evaluator =
    String(
      record['Evaluator'] || ''
    ).trim();


  if (!evaluator) {

    return false;
  }


  return collegeCIs.some(ci =>

    semesterDocument_namesMatch_(
      evaluator,
      ci['Evaluator']
    )

  );
}


/**
 * =====================================================================
 * HISTORICAL EVALUATOR NAME MATCHING
 * =====================================================================
 *
 * Examples:
 *
 *   Amanda Rhinehart  <->  A. Rhinehart
 *   Nicole Murray     <->  N. Murray
 */
function semesterDocument_namesMatch_(
  nameA,
  nameB
) {

  const a =
    semesterDocument_nameParts_(
      nameA
    );


  const b =
    semesterDocument_nameParts_(
      nameB
    );


  if (
    !a.last ||
    !b.last
  ) {

    return false;
  }


  if (
    a.last !==
    b.last
  ) {

    return false;
  }


  if (
    !a.firstInitial ||
    !b.firstInitial
  ) {

    return (
      appNormalize_(
        nameA
      ) ===
      appNormalize_(
        nameB
      )
    );
  }


  return (
    a.firstInitial ===
    b.firstInitial
  );
}


/**
 * =====================================================================
 * NAME PARTS
 * =====================================================================
 */
function semesterDocument_nameParts_(
  value
) {

  const cleaned =
    String(
      value || ''
    )
      .toLowerCase()
      .replace(
        /[^a-z\s]/g,
        ' '
      )
      .replace(
        /\s+/g,
        ' '
      )
      .trim();


  if (!cleaned) {

    return {

      firstInitial:
        '',

      last:
        ''

    };
  }


  const parts =
    cleaned.split(
      ' '
    );


  return {

    firstInitial:
      parts[0]
        ? parts[0][0]
        : '',

    last:
      parts[
        parts.length - 1
      ] || ''

  };
}


/**
 * =====================================================================
 * PASSING COMPETENCY
 * =====================================================================
 */
function semesterDocument_isPassingCompetency_(
  record
) {

  return (

    appNormalize_(
      record[
        'Evaluation Type'
      ]
    ) ===
      'competency' &&

    appNormalize_(
      record[
        'Result'
      ]
    ) ===
      'pass' &&

    !appIsTrue_(
      record[
        'Automatic Failure'
      ]
    )

  );
}


/**
 * =====================================================================
 * FAILED COMPETENCY
 * =====================================================================
 */
function semesterDocument_isFailedCompetency_(
  record
) {

  return (

    appNormalize_(
      record[
        'Evaluation Type'
      ]
    ) ===
      'competency' &&

    (

      appNormalize_(
        record[
          'Result'
        ]
      ) ===
        'fail' ||

      appIsTrue_(
        record[
          'Automatic Failure'
        ]
      )

    )

  );
}


/**
 * =====================================================================
 * SAME STUDENT
 * =====================================================================
 *
 * Student ID is authoritative.
 *
 * Student name is only a fallback for older imported records that may
 * not have a Student ID.
 */
function semesterDocument_sameStudent_(
  record,
  studentId,
  studentName
) {

  const targetId =
    appNormalize_(
      studentId
    );


  const recordId =
    appNormalize_(
      record[
        'Student ID'
      ]
    );


  if (
    targetId &&
    recordId
  ) {

    return (
      targetId ===
      recordId
    );
  }


  return (

    appNormalize_(
      record[
        'Student'
      ]
    ) ===

    appNormalize_(
      studentName
    )

  );
}


/**
 * =====================================================================
 * RECORD KEY
 * =====================================================================
 *
 * Submission ID is preferred.
 *
 * A composite key is used only as fallback for historical records.
 */
function semesterDocument_recordKey_(
  record
) {

  const submissionId =
    String(
      record[
        'Submission ID'
      ] || ''
    ).trim();


  if (submissionId) {

    return submissionId;
  }


  return [

    record[
      'Student ID'
    ],

    record[
      'Examination'
    ],

    record[
      'Date'
    ],

    record[
      'Evaluator'
    ]

  ]

    .map(
      appNormalize_
    )

    .join(
      '|'
    );
}


/**
 * =====================================================================
 * DATE SORT
 * =====================================================================
 */
function semesterDocument_recordDateSort_(
  a,
  b
) {

  return (

    appDateValue_(
      a['Date']
    ) -

    appDateValue_(
      b['Date']
    )

  );
}


/**
 * =====================================================================
 * EXAMINATION DISPLAY
 * =====================================================================
 */
function semesterDocument_displayExam_(
  record
) {

  return String(
    record[
      'Examination'
    ] || ''
  ).trim();
}


/**
 * =====================================================================
 * EVALUATOR DISPLAY
 * =====================================================================
 */
function semesterDocument_displayEvaluator_(
  record
) {

  return String(
    record[
      'Evaluator'
    ] || ''
  ).trim();
}


/**
 * =====================================================================
 * SCORE DISPLAY
 * =====================================================================
 *
 * Historical records may legitimately have no numeric Final Score.
 *
 * Do not invent one.
 */
function semesterDocument_displayScore_(
  record
) {

  record = record || {};

  const finalScore =
    semesterDocument_numericScore_(
      record['Final Score']
    );

  if (finalScore !== null) {
    return finalScore.toFixed(2);
  }

  const historicalRawFields = [
    'Historical Raw Score',
    'Original Raw Score',
    'Reconstructed Raw Score'
  ];

  for (let i = 0; i < historicalRawFields.length; i++) {
    const field = historicalRawFields[i];

    if (!Object.prototype.hasOwnProperty.call(record, field)) {
      continue;
    }

    const rawScore =
      semesterDocument_numericScore_(
        record[field]
      );

    if (rawScore !== null) {
      return rawScore.toFixed(2);
    }
  }

  const evaluationType =
    appNormalize_(
      record['Evaluation Type']
    );

  const result =
    String(
      record['Result'] || ''
    ).trim();

  const normalizedResult =
    appNormalize_(
      result
    );

  if (
    evaluationType === 'simulation' ||
    normalizedResult === 'simulation' ||
    normalizedResult === 'sim'
  ) {
    return 'Simulation';
  }

  if (normalizedResult === 'pass') {
    return 'Pass';
  }

  if (normalizedResult === 'fail') {
    return 'Fail';
  }

  return '';
}


function semesterDocument_numericScore_(
  value
) {

  if (
    value === '' ||
    value === null ||
    value === undefined
  ) {
    return null;
  }

  const numeric = Number(value);

  return Number.isFinite(numeric)
    ? numeric
    : null;
}



/**
 * =====================================================================
 * REPLACE AND FORMAT POPULATED TEMPLATE TOKEN
 * =====================================================================
 *
 * Replaces one template token and applies formatting only to the
 * inserted value. The institutional form around it is not resized.
 */
function semesterDocument_replaceFormattedToken_(
  body,
  token,
  value,
  options
) {

  options = options || {};

  const replacement =
    String(
      value === null ||
      value === undefined
        ? ''
        : value
    );

  const pattern =
    appEscapeRegex_(
      token
    );

  let replaced = false;
  let found = null;

  while (
    (
      found =
        body.findText(
          pattern
        )
    )
  ) {

    replaced = true;

    const text =
      found
        .getElement()
        .asText();

    const start =
      found.getStartOffset();

    const end =
      found.getEndOffsetInclusive();

    text.deleteText(
      start,
      end
    );

    if (!replacement) {
      continue;
    }

    text.insertText(
      start,
      replacement
    );

    const insertedEnd =
      start +
      replacement.length -
      1;

    let fontSize =
      Number(
        options.fontSize || 10
      );

    if (replacement.length > 28) {
      fontSize = Math.min(fontSize, 8);
    } else if (replacement.length > 20) {
      fontSize = Math.min(fontSize, 9);
    }

    text.setFontSize(
      start,
      insertedEnd,
      fontSize
    );

    if (options.bold === true) {
      text.setBold(
        start,
        insertedEnd,
        true
      );
    }
  }

  return replaced;
}

/**
 * =====================================================================
 * DATE DISPLAY
 * =====================================================================
 */
function semesterDocument_displayDate_(
  value
) {

  const date =
    appDateOnly_(
      value
    );


  if (!date) {

    return '';
  }


  return Utilities.formatDate(
    date,
    Session.getScriptTimeZone(),
    'M/d/yy'
  );
}


/**
 * =====================================================================
 * STUDENT ID DISPLAY
 * =====================================================================
 *
 * Example:
 *
 *   DEPT # 22
 *
 * becomes:
 *
 *   #22
 */
function semesterDocument_formatStudentId_(
  studentId
) {

  const value =
    String(
      studentId || ''
    ).trim();


  const match =
    value.match(
      /#\s*(\d+)/i
    );


  if (match) {

    return (
      '#' +
      match[1]
    );
  }


  return value;
}


/**
 * =====================================================================
 * CLINICAL NUMBER
 * =====================================================================
 *
 * Clinical 2 -> 2
 */
function semesterDocument_clinicalNumber_(
  value
) {

  const match =
    String(
      value || ''
    )
      .match(
        /clinical\s*(\d+)/i
      );


  return match

    ? Number(
        match[1]
      )

    : 0;
}


/**
 * =====================================================================
 * OPEN GENERATED DOCUMENT WITH RETRY
 * =====================================================================
 *
 * Drive may finish creating a Google Doc slightly before DocumentApp
 * can open it.
 */
function semesterDocument_openWithRetry_(
  documentId
) {

  const maxAttempts =
    6;


  let lastError =
    null;


  for (
    let attempt = 1;
    attempt <= maxAttempts;
    attempt++
  ) {

    try {

      return DocumentApp.openById(
        documentId
      );

    } catch (error) {

      lastError =
        error;


      Utilities.sleep(
        1000 * attempt
      );
    }
  }


  throw new Error(
    'Could not open the generated Semester Requirement document after ' +
    maxAttempts +
    ' attempts. ' +
    (
      lastError
        ? lastError.message
        : ''
    )
  );
}


/**
 * =====================================================================
 * MANUAL TEST
 * =====================================================================
 *
 * TEST ONLY.
 *
 * Enter a Student ID and choose Clinical 1 through Clinical 5.
 *
 * Running this WILL create a real generated document in that student's
 * Generated Documents folder.
 */
function testSemesterRequirementDocument() {

  const TEST_STUDENT_ID =
    'DEPT # 22';


  const TEST_CLINICAL_TERM =
    'Clinical 4';


  if (!TEST_STUDENT_ID) {

    throw new Error(
      'Enter a Student ID in testSemesterRequirementDocument() before running the test.'
    );
  }


  const result =
    generateSemesterRequirementSheet(
      TEST_STUDENT_ID,
      TEST_CLINICAL_TERM
    );


  console.log(
    JSON.stringify(
      result,
      null,
      2
    )
  );


  return result;
}