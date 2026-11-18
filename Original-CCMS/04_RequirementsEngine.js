/**
 * =====================================================================
 * CLINICAL COMPETENCY TRACKER
 * REQUIREMENTS ENGINE
 * =====================================================================
 *
 * PURPOSE
 * -------
 * Calculates student semester and program requirement status.
 *
 * SOURCES
 * -------
 * Students
 * Clinical Tracking
 * Examinations
 * Anatomical Areas
 * Program Rules
 * Semester Rules
 *
 * OUTPUT
 * ------
 * Student Requirement Status
 *
 * IMPORTANT DESIGN PRINCIPLES
 * ---------------------------
 * 1. Clinical Tracking is the permanent source record.
 * 2. Semester Rules defines the rules for Clinical 1-5.
 * 3. Evaluation DATE determines semester activity.
 * 4. Current-term counters do NOT inherit old activity.
 * 5. Carryover is shown separately.
 * 6. Program totals remain cumulative.
 * 7. HCI usage counts only NEW passing competencies during that term.
 * 8. Failed competencies do not consume HCI slots.
 * 9. Three failures on the same exam/body part trigger remediation.
 *
 * =====================================================================
 */

/**
 * =====================================================================
 * CLINICAL COMPETENCY TRACKER
 * REQUIREMENTS ENGINE
 * 04_RequirementsEngine.gs
 * =====================================================================
 *
 * PURPOSE
 * -------
 * Authoritative academic-rules engine.
 *
 * Reads factual evaluation history from Clinical Tracking and determines
 * what that history means for each student's semester and overall
 * program requirements.
 *
 * READS
 * -----
 * Students
 * Clinical Tracking
 * Examinations
 * Anatomical Areas
 * Program Rules
 * Semester Rules
 *
 * WRITES
 * ------
 * Student Requirement Status
 *
 * OWNS
 * ----
 * - Semester competency requirements
 * - Carryover
 * - New competencies this semester
 * - Additional competencies toward future semesters
 * - ISO requirements
 * - HCI usage and limits
 * - Failed competency summaries
 * - Three-failure remediation rule
 * - Clinical history
 * - Program requirements
 * - Mandatory / elective / group progress
 * - Examination eligibility and prerequisites
 *
 * DOES NOT OWN
 * ------------
 * - Evaluation grading
 * - Clinical Tracking writes
 * - Form synchronization
 * - Drive filing
 * - Document generation
 * - Email notifications
 *
 * DESIGN PRINCIPLES
 * -----------------
 * 1. Clinical Tracking is the permanent factual source.
 * 2. Semester Rules defines Clinical 1-5 rules.
 * 3. Evaluation DATE determines semester activity.
 * 4. Current-semester counters do not inherit old activity.
 * 5. Carryover is calculated separately.
 * 6. Program totals remain cumulative.
 * 7. HCI counts only qualifying NEW passing competencies that term.
 * 8. Failed competencies do not consume HCI slots.
 * 9. Three failures on the same examination/body part trigger remediation.
 * =====================================================================
 */


const REQUIREMENTS_STATUS_HEADERS = [

  'Student ID',
  'Student',
  'Clinical Term',
  'Status Type',
  'Requirement Scope',
  'Requirement ID',
  'Requirement Name',
  'Requirement Type',
  'Requirement Group',
  'Requirement Timing',
  'Required Count',
  'Completed Count',
  'Remaining Count',
  'Status',
  'Last Completed Date',
  'Source Record ID',
  'Notes'

];



/**
 * =====================================================================
 * MAIN ENGINE
 * =====================================================================
 */

function rebuildStudentRequirementStatus() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();


  // ===================================================================
  // AUTHORITATIVE SHEETS
  // ===================================================================

  const studentsSheet =
    appGetSheet_(
      ss,
      APP_CONFIG.SHEETS.STUDENTS
    );


  const trackingSheet =
    appGetSheet_(
      ss,
      APP_CONFIG.SHEETS.CLINICAL_TRACKING
    );


  const examinationsSheet =
    appGetSheet_(
      ss,
      APP_CONFIG.SHEETS.EXAMINATIONS
    );


  const anatomicalSheet =
    appGetSheet_(
      ss,
      APP_CONFIG.SHEETS.ANATOMICAL_AREAS
    );


  const programRulesSheet =
    appGetSheet_(
      ss,
      APP_CONFIG.SHEETS.PROGRAM_RULES
    );


  const semesterRulesSheet =
    appGetSheet_(
      ss,
      APP_CONFIG.SHEETS.SEMESTER_RULES
    );


  const statusSheet =
    appGetSheet_(
      ss,
      APP_CONFIG.SHEETS.STUDENT_REQUIREMENT_STATUS
    );


  // ===================================================================
  // VALIDATE IMPORTANT HEADERS
  // ===================================================================

  reqValidateRequiredHeaders_(
    studentsSheet,
    [
      'Student ID',
      'Student',
      'Clinical Term',
      'Active'
    ]
  );


  reqValidateRequiredHeaders_(
    trackingSheet,
    [
      'Submission ID',
      'Student ID',
      'Student',
      'Evaluation Type',
      'Examination ID',
      'Examination',
      'Date',
      'Result',
      'Evaluator',
      'Evaluator Type',
      'Automatic Failure'
    ]
  );


  reqValidateRequiredHeaders_(
    examinationsSheet,
    [
      'Examination ID',
      'Examination',
      'Requirement Type',
      'Requirement Group',
      'Earliest Clinical',
      'Prerequisite Type',
      'Prerequisite',
      'Active'
    ]
  );


  reqValidateRequiredHeaders_(
    semesterRulesSheet,
    [
      'Clinical Term',
      'Semester Competency Requirement',
      'Cumulative Minimum',
      'Cumulative Maximum',
      'Additional Future Maximum',
      'Minimum ISO',
      'Start Date',
      'End Date',
      'HCI Pass Limit',
      'Active'
    ]
  );


  reqValidateExactHeaders_(
    statusSheet,
    REQUIREMENTS_STATUS_HEADERS
  );


  // ===================================================================
  // READ DATA
  // ===================================================================

  const students =
    appSheetToObjects_(
      studentsSheet
    );


  const tracking =
    appSheetToObjects_(
      trackingSheet
    );


  const examinations =
    appSheetToObjects_(
      examinationsSheet
    );


  const anatomicalAreas =
    appSheetToObjects_(
      anatomicalSheet
    );


  const programRules =
    appSheetToObjects_(
      programRulesSheet
    );


  const semesterRules =
    appSheetToObjects_(
      semesterRulesSheet
    );


  // ===================================================================
  // LOOKUPS
  // ===================================================================

  const examByName =
    reqBuildLookup_(
      examinations,
      'Examination'
    );


  const examById =
    reqBuildLookup_(
      examinations,
      'Examination ID'
    );


  const semesterByTerm =
    reqBuildLookup_(
      semesterRules,
      'Clinical Term'
    );


  // ===================================================================
  // ACTIVE STUDENTS
  // ===================================================================

  const activeStudents =
    reqGetUniqueActiveStudents_(
      students
    );


  const output = [];


  activeStudents
    .forEach(student => {

      reqProcessStudent_({

        student,
        tracking,
        examinations,
        anatomicalAreas,
        programRules,
        semesterRules,

        examByName,
        examById,
        semesterByTerm,

        output

      });

    });


  // ===================================================================
  // WRITE OUTPUT
  // ===================================================================

  reqClearStatusData_(
    statusSheet,
    REQUIREMENTS_STATUS_HEADERS.length
  );


  if (
    output.length > 0
  ) {

    statusSheet
      .getRange(
        2,
        1,
        output.length,
        REQUIREMENTS_STATUS_HEADERS.length
      )
      .setValues(
        output
      );
  }


  SpreadsheetApp.flush();


  console.log(
    'Requirements Engine complete. ' +
    output.length +
    ' rows created.'
  );


  return {

    success:
      true,

    rowsCreated:
      output.length,

    studentsProcessed:
      activeStudents.length

  };
}


/**
 * =====================================================================
 * PROCESS ONE STUDENT
 * =====================================================================
 */

function reqProcessStudent_(args) {

  const {
    student,
    tracking,
    examinations,
    anatomicalAreas,
    programRules,
    semesterRules,
    examByName,
    examById,
    semesterByTerm,
    output
  } = args;


  const studentId =
    String(
      student['Student ID'] || ''
    ).trim();


  const studentName =
    String(
      student['Student'] || ''
    ).trim();


  const clinicalTerm =
    String(
      student['Clinical Term'] || ''
    ).trim();


  if (!studentId && !studentName) {
    return;
  }


  const records =
    tracking.filter(record =>
      reqSameStudent_(
        record,
        studentId,
        studentName
      )
    );


  // ---------------------------------------------------------------
  // NO CLINICAL TERM = PRE-CLINICAL
  // ---------------------------------------------------------------

  if (!clinicalTerm) {

    output.push(
      reqStatusRow_({
        studentId,
        studentName,
        clinicalTerm: '',
        statusType: 'Summary',
        scope: 'Semester',
        id: 'SEM-PRECLINICAL',
        name: 'Clinical Status',
        type: 'Term Status',
        timing: 'Pre-Clinical',
        completed: 0,
        status: 'Clinical Not Started',
        notes:
          'Student has not yet begun a clinical term.'
      })
    );

    return;
  }


  const semesterRule =
    semesterByTerm[
      reqNormalize_(clinicalTerm)
    ];


  if (!semesterRule) {

    throw new Error(
      'No Semester Rules row found for "' +
      clinicalTerm +
      '".'
    );
  }


  const termStart =
    reqDateOnly_(
      semesterRule['Start Date']
    );


  const termEnd =
    reqDateOnly_(
      semesterRule['End Date']
    );


  if (!termStart || !termEnd) {

    throw new Error(
      clinicalTerm +
      ' is missing a valid Start Date or End Date.'
    );
  }


  const termStatus =
    reqGetTermStatus_(
      termStart,
      termEnd
    );


  // ---------------------------------------------------------------
  // TERM STATUS
  // ---------------------------------------------------------------

  output.push(
    reqStatusRow_({
      studentId,
      studentName,
      clinicalTerm,
      statusType: 'Summary',
      scope: 'Semester',
      id:
        reqTermId_(clinicalTerm) +
        '-STATUS',
      name:
        clinicalTerm +
        ' Status',
      type: 'Term Status',
      timing: termStatus,
      completed: 0,
      status: termStatus,
      notes:
        reqFormatDateRange_(
          termStart,
          termEnd
        )
    })
  );


  // ---------------------------------------------------------------
  // RECORD BUCKETS
  // ---------------------------------------------------------------

  /*
   * Semester ownership is factual record data, not entry timing.
   *
   * Attributed Clinical Term is the event-term field produced by the
   * current intake architecture. Clinical Term is the authoritative
   * fallback used by historical imports and older records.
   *
   * Do not use Date, Created At, submission time, or today's date here.
   */
  const beforeTerm =
    records.filter(record =>
      reqClinicalNumber_(
        reqRecordClinicalTerm_(record)
      ) > 0 &&
      reqClinicalNumber_(
        reqRecordClinicalTerm_(record)
      ) <
      reqClinicalNumber_(clinicalTerm)
    );


  const currentTermRecords =
    records.filter(record =>
      reqNormalize_(
        reqRecordClinicalTerm_(record)
      ) ===
      reqNormalize_(clinicalTerm)
    );


  // ---------------------------------------------------------------
  // VALID UNIQUE PASSING COMPETENCIES
  // ---------------------------------------------------------------

  const passingBeforeTerm =
    reqGetUniquePassingCompetencies_(
      beforeTerm.filter(
        reqIsPassingCompetency_
      )
    );


  const passingThisTerm =
    reqGetUniquePassingCompetencies_(
      currentTermRecords.filter(
        reqIsPassingCompetency_
      )
    );


  // ---------------------------------------------------------------
  // SEMESTER REQUIREMENT
  // ---------------------------------------------------------------

  const semesterRequirementRaw =
    semesterRule[
      'Semester Competency Requirement'
    ];


  const semesterRequirement =
    reqParseRequirement_(
      semesterRequirementRaw
    );


  // ---------------------------------------------------------------
  // CARRYOVER
  // ---------------------------------------------------------------

  const carryoverAvailable =
    reqCalculateCarryoverAvailable_({

      clinicalTerm,
      passingBeforeTerm,
      semesterRules

    });


  let carryoverApplied = 0;
  let semesterProgress = 0;
  let stillNeeded = '';


  if (
    typeof semesterRequirement ===
    'number'
  ) {

    carryoverApplied =
      Math.min(
        carryoverAvailable,
        semesterRequirement
      );


    semesterProgress =
      Math.min(
        semesterRequirement,
        carryoverApplied +
        passingThisTerm.length
      );


    stillNeeded =
      Math.max(
        semesterRequirement -
        semesterProgress,
        0
      );
  }


  // ---------------------------------------------------------------
  // SEMESTER REQUIREMENT SUMMARY
  // ---------------------------------------------------------------

  output.push(
    reqStatusRow_({

      studentId,
      studentName,
      clinicalTerm,

      statusType: 'Summary',
      scope: 'Semester',

      id:
        reqTermId_(clinicalTerm) +
        '-COMP',

      name:
        clinicalTerm +
        ' Competency Requirement',

      type: 'Semester Competency',

      timing: termStatus,

      required:
        semesterRequirementRaw,

      completed:
        semesterProgress,

      remaining:
        stillNeeded,

      status:
        reqProgressStatus_(
          semesterRequirement,
          semesterProgress
        ),

      notes:
        'Carryover applied: ' +
        carryoverApplied +
        '. New passing competencies this term: ' +
        passingThisTerm.length +
        '.'
    })
  );


  // ---------------------------------------------------------------
  // CARRYOVER SUMMARY
  // ---------------------------------------------------------------

  output.push(
    reqStatusRow_({

      studentId,
      studentName,
      clinicalTerm,

      statusType: 'Summary',
      scope: 'Semester',

      id:
        reqTermId_(clinicalTerm) +
        '-CARRY',

      name:
        'Carryover Applied',

      type: 'Carryover',

      timing: termStatus,

      completed:
        carryoverApplied,

      status:
        carryoverApplied > 0
          ? 'Available'
          : 'None',

      notes:
        carryoverAvailable +
        ' qualifying carryover competencies available entering ' +
        clinicalTerm +
        '.'
    })
  );


  // ---------------------------------------------------------------
  // NEW COMPETENCIES THIS TERM
  // ---------------------------------------------------------------

  output.push(
    reqStatusRow_({

      studentId,
      studentName,
      clinicalTerm,

      statusType: 'Summary',
      scope: 'Semester',

      id:
        reqTermId_(clinicalTerm) +
        '-NEW',

      name:
        'New Competencies This Term',

      type: 'New Competencies',

      timing: termStatus,

      completed:
        passingThisTerm.length,

      status:
        passingThisTerm.length > 0
          ? 'Recorded'
          : 'None',

      lastDate:
        reqLatestDateFromRecords_(
          passingThisTerm
        ),

      sourceId:
        reqLatestSubmissionFromRecords_(
          passingThisTerm
        )
    })
  );


  // ---------------------------------------------------------------
  // ADDITIONAL / FUTURE COMPETENCIES
  // ---------------------------------------------------------------

  const additionalFuture =
    reqCalculateAdditionalFuture_({

      semesterRequirement,
      carryoverApplied,
      passingThisTerm,
      semesterRule

    });


  output.push(
    reqStatusRow_({

      studentId,
      studentName,
      clinicalTerm,

      statusType: 'Summary',
      scope: 'Semester',

      id:
        reqTermId_(clinicalTerm) +
        '-ADDITIONAL',

      name:
        'Additional Competencies for Future Terms',

      type: 'Additional Competencies',

      timing: termStatus,

      required:
        semesterRule[
          'Additional Future Maximum'
        ] || '',

      completed:
        additionalFuture.count,

      remaining:
        additionalFuture.remaining,

      status:
        additionalFuture.count > 0
          ? 'Banked'
          : 'None',

      notes:
        additionalFuture.notes
    })
  );


  // ---------------------------------------------------------------
  // ISO
  // ---------------------------------------------------------------

  const isoRecords =
    currentTermRecords.filter(record =>
      reqNormalize_(
        record['Evaluation Type']
      ) === 'iso'
    );


  const minimumISO =
    Number(
      semesterRule['Minimum ISO']
    ) || 0;


  output.push(
    reqStatusRow_({

      studentId,
      studentName,
      clinicalTerm,

      statusType: 'Summary',
      scope: 'Semester',

      id:
        reqTermId_(clinicalTerm) +
        '-ISO',

      name:
        'Instructor Selected Ongoing Evaluations',

      type: 'ISO',

      timing: termStatus,

      required:
        minimumISO,

      completed:
        isoRecords.length,

      remaining:
        Math.max(
          minimumISO -
          isoRecords.length,
          0
        ),

      status:
        isoRecords.length >=
        minimumISO
          ? 'Complete'
          : (
              isoRecords.length > 0
                ? 'In Progress'
                : 'Not Started'
            ),

      lastDate:
        reqLatestDateFromRecords_(
          isoRecords
        ),

      sourceId:
        reqLatestSubmissionFromRecords_(
          isoRecords
        )
    })
  );


  // ---------------------------------------------------------------
  // FAILURES THIS TERM
  // ---------------------------------------------------------------

  const failuresThisTerm =
    currentTermRecords.filter(
      reqIsFailedCompetency_
    );


  output.push(
    reqStatusRow_({

      studentId,
      studentName,
      clinicalTerm,

      statusType: 'Summary',
      scope: 'Semester',

      id:
        reqTermId_(clinicalTerm) +
        '-FAIL',

      name:
        'Failed Competencies This Term',

      type: 'Failure',

      timing: termStatus,

      completed:
        failuresThisTerm.length,

      status:
        failuresThisTerm.length > 0
          ? 'Needs Attention'
          : 'None',

      lastDate:
        reqLatestDateFromRecords_(
          failuresThisTerm
        ),

      sourceId:
        reqLatestSubmissionFromRecords_(
          failuresThisTerm
        )
    })
  );


  // ---------------------------------------------------------------
  // HCI USAGE
  // ---------------------------------------------------------------

  const hciLimit =
    Number(
      semesterRule[
        'HCI Pass Limit'
      ]
    ) || 0;


  const hciRecords =
    passingThisTerm.filter(
      reqConsumesHCISlot_
    );


  const hciUsed =
    hciRecords.length;


  const hciRemaining =
    Math.max(
      hciLimit - hciUsed,
      0
    );


  let hciStatus = 'Available';


  if (
    hciLimit > 0 &&
    hciUsed >= hciLimit
  ) {
    hciStatus = 'Limit Reached';
  }


  if (
    hciLimit > 0 &&
    hciUsed > hciLimit
  ) {
    hciStatus =
      'Coordinator Review Required';
  }


  output.push(
    reqStatusRow_({

      studentId,
      studentName,
      clinicalTerm,

      statusType: 'Summary',
      scope: 'Semester',

      id:
        reqTermId_(clinicalTerm) +
        '-HCI',

      name:
        'Hospital CI Passing Competencies',

      type: 'HCI Limit',

      timing: termStatus,

      required:
        hciLimit,

      completed:
        hciUsed,

      remaining:
        hciRemaining,

      status:
        hciStatus,

      lastDate:
        reqLatestDateFromRecords_(
          hciRecords
        ),

      sourceId:
        reqLatestSubmissionFromRecords_(
          hciRecords
        ),

      notes:
        hciUsed +
        ' of ' +
        hciLimit +
        ' HCI passing competency slots used this term.'
    })
  );


  // ---------------------------------------------------------------
  // REMEDIATION
  // ---------------------------------------------------------------

  reqBuildRemediationRows_({

    studentId,
    studentName,
    clinicalTerm,
    records,
    output

  });


  // ---------------------------------------------------------------
  // CLINICAL HISTORY
  // ---------------------------------------------------------------

  reqBuildClinicalHistory_({

    studentId,
    studentName,
    records,
    semesterRules,
    output

  });


  // ---------------------------------------------------------------
  // PROGRAM REQUIREMENTS
  // ---------------------------------------------------------------

  reqBuildProgramRequirements_({

    studentId,
    studentName,
    clinicalTerm,

    records,
    examinations,
    programRules,

    examByName,
    examById,

    output

  });

}


/**
 * =====================================================================
 * CARRYOVER
 * =====================================================================
 */

function reqCalculateCarryoverAvailable_(args) {

  const {
    clinicalTerm,
    passingBeforeTerm,
    semesterRules
  } = args;


  const clinicalNumber =
    reqClinicalNumber_(
      clinicalTerm
    );


  // Clinical 1 has no previous clinical semester.
  if (clinicalNumber <= 1) {
    return 0;
  }


  const previousTerm =
    'Clinical ' +
    (clinicalNumber - 1);


  const previousRule =
    semesterRules.find(rule =>
      reqNormalize_(
        rule['Clinical Term']
      ) ===
      reqNormalize_(
        previousTerm
      )
    );


  if (!previousRule) {
    return 0;
  }


  const previousMinimum =
    Number(
      previousRule[
        'Cumulative Minimum'
      ]
    ) || 0;


  return Math.max(
    passingBeforeTerm.length -
    previousMinimum,
    0
  );
}


/**
 * =====================================================================
 * ADDITIONAL FUTURE COMPETENCIES
 * =====================================================================
 */

function reqCalculateAdditionalFuture_(args) {

  const {
    semesterRequirement,
    carryoverApplied,
    passingThisTerm,
    semesterRule
  } = args;


  if (
    typeof semesterRequirement !==
    'number'
  ) {

    return {
      count: 0,
      remaining: '',
      notes:
        'Final clinical term requires completion of all remaining program requirements.'
    };
  }


  const newNeeded =
    Math.max(
      semesterRequirement -
      carryoverApplied,
      0
    );


  const additionalCount =
    Math.max(
      passingThisTerm.length -
      newNeeded,
      0
    );


  const rawMaximum =
    semesterRule[
      'Additional Future Maximum'
    ];


  if (
    rawMaximum === '' ||
    rawMaximum === null ||
    typeof rawMaximum ===
      'undefined'
  ) {

    return {
      count: additionalCount,
      remaining: '',
      notes:
        'No maximum additional future competency limit configured.'
    };
  }


  const maximum =
    Number(rawMaximum);


  if (isNaN(maximum)) {

    return {
      count: additionalCount,
      remaining: '',
      notes:
        'Additional future competency maximum could not be interpreted.'
    };
  }


  return {
    count: additionalCount,
    remaining:
      Math.max(
        maximum -
        additionalCount,
        0
      ),
    notes:
      additionalCount +
      ' of ' +
      maximum +
      ' additional competencies completed toward future semester requirements.'
  };
}


/**
 * =====================================================================
 * HCI LOGIC
 * =====================================================================
 */

function reqConsumesHCISlot_(record) {

  if (
    !reqIsPassingCompetency_(
      record
    )
  ) {
    return false;
  }


  const evaluatorType =
    reqNormalize_(
      record['Evaluator Type']
    );


  // Program rule:
  // anything NOT College CI is considered HCI.
  return (
    evaluatorType !==
    'college ci'
  );
}


/**
 * =====================================================================
 * REMEDIATION
 * =====================================================================
 */

function reqBuildRemediationRows_(args) {

  const {
    studentId,
    studentName,
    clinicalTerm,
    records,
    output
  } = args;


  const failureGroups = {};


  records
    .filter(
      reqIsFailedCompetency_
    )
    .forEach(record => {

      const key =
        reqRemediationKey_(
          record
        );


      if (!key) {
        return;
      }


      if (!failureGroups[key]) {
        failureGroups[key] = [];
      }


      failureGroups[key].push(
        record
      );
    });


  Object.keys(
    failureGroups
  ).forEach(key => {

    const failures =
      failureGroups[key];


    if (
      failures.length < 3
    ) {
      return;
    }


    const latest =
      reqGetLatestRecord_(
        failures
      );


    const remediationComplete =
      reqHasCompletedRemediation_(
        records,
        key
      );


    const examination =
      String(
        latest['Examination'] || ''
      ).trim();


    const anatomicalArea =
      String(
        latest['Anatomical Area'] || ''
      ).trim();


    const displayName =
      anatomicalArea
        ? examination +
          ' - ' +
          anatomicalArea
        : examination;


    output.push(
      reqStatusRow_({

        studentId,
        studentName,
        clinicalTerm,

        statusType: 'Alert',
        scope: 'Program',

        id:
          'REMED-' +
          reqSafeId_(key),

        name:
          displayName,

        type: 'Remediation',

        timing:
          remediationComplete
            ? 'Completed'
            : 'Required',

        required: 3,

        completed:
          failures.length,

        status:
          remediationComplete
            ? 'Remediation Complete'
            : 'Remediation Required',

        lastDate:
          latest['Date'] || '',

        sourceId:
          latest[
            'Submission ID'
          ] || '',

        notes:
          failures.length +
          ' failed competency evaluations on the same examination/body part.'
      })
    );

  });
}


function reqRemediationKey_(record) {

  const examination =
    reqNormalize_(
      record['Examination']
    );


  const anatomicalArea =
    reqNormalize_(
      record['Anatomical Area']
    );


  if (!examination) {
    return '';
  }


  return anatomicalArea
    ? examination +
      '|' +
      anatomicalArea
    : examination;
}


function reqHasCompletedRemediation_(
  records,
  remediationKey
) {

  return records.some(record => {

    const key =
      reqRemediationKey_(
        record
      );


    if (
      key !== remediationKey
    ) {
      return false;
    }


    return reqTrue_(
      record[
        'Remediation Complete'
      ]
    );

  });
}


/**
 * =====================================================================
 * CLINICAL HISTORY
 * =====================================================================
 */

function reqBuildClinicalHistory_(args) {

  const {
    studentId,
    studentName,
    records,
    semesterRules,
    output
  } = args;


  semesterRules
    .filter(rule =>
      reqTrue_(
        rule['Active']
      )
    )
    .sort(
      (a, b) =>
        reqClinicalNumber_(
          a['Clinical Term']
        ) -
        reqClinicalNumber_(
          b['Clinical Term']
        )
    )
    .forEach(rule => {

      const term =
        String(
          rule['Clinical Term'] || ''
        ).trim();


      const start =
        reqDateOnly_(
          rule['Start Date']
        );


      const end =
        reqDateOnly_(
          rule['End Date']
        );


      if (!start || !end) {
        return;
      }


      const termRecords =
        records.filter(record =>
          reqDateFallsWithin_(
            record['Date'],
            start,
            end
          )
        );


      const passed =
        reqGetUniquePassingCompetencies_(
          termRecords.filter(
            reqIsPassingCompetency_
          )
        );


      const failed =
        termRecords.filter(
          reqIsFailedCompetency_
        );


      const iso =
        termRecords.filter(record =>
          reqNormalize_(
            record[
              'Evaluation Type'
            ]
          ) === 'iso'
        );


      const names =
        passed
          .map(record =>
            String(
              record[
                'Examination'
              ] || ''
            ).trim()
          )
          .filter(Boolean);


      output.push(
        reqStatusRow_({

          studentId,
          studentName,
          clinicalTerm: term,

          statusType: 'History',
          scope: 'Semester',

          id:
            reqTermId_(term) +
            '-HISTORY',

          name:
            term +
            ' Clinical History',

          type:
            'Clinical History',

          timing:
            reqGetTermStatus_(
              start,
              end
            ),

          completed:
            passed.length,

          status:
            passed.length > 0
              ? 'Recorded'
              : 'No Competencies',

          lastDate:
            reqLatestDateFromRecords_(
              termRecords
            ),

          sourceId:
            reqLatestSubmissionFromRecords_(
              termRecords
            ),

          notes:
            'Passed competencies: ' +
            (
              names.length
                ? names.join(', ')
                : 'None'
            ) +
            '. ISO evaluations: ' +
            iso.length +
            '. Failed competencies: ' +
            failed.length +
            '.'
        })
      );

    });

}


/**
 * =====================================================================
 * PROGRAM REQUIREMENTS
 * =====================================================================
 */

function reqBuildProgramRequirements_(args) {

  const {
    studentId,
    studentName,
    clinicalTerm,
    records,
    examinations,
    programRules,
    examByName,
    examById,
    output
  } = args;


  const passed =
    reqGetUniquePassingCompetencies_(
      records.filter(
        reqIsPassingCompetency_
      )
    );


  const completedExamNames =
    new Set();


  const completedExamIds =
    new Set();


  const enriched = [];


  passed.forEach(record => {

    const recordExamId =
      reqNormalize_(
        record['Examination ID']
      );


    const recordExamName =
      reqNormalize_(
        record['Examination']
      );


    if (!recordExamId && !recordExamName) {
      return;
    }


    // Examination ID is authoritative when present.
    // Exact examination-name lookup is retained only as a legacy fallback
    // for established records that do not yet have an Examination ID.
    const exam =
      (
        recordExamId &&
        examById[recordExamId]
      ) ||
      (
        recordExamName &&
        examByName[recordExamName]
      );


    if (!exam) {
      return;
    }


    const canonicalExamId =
      reqNormalize_(
        exam['Examination ID']
      );


    const canonicalExamName =
      reqNormalize_(
        exam['Examination']
      );


    if (canonicalExamId) {
      completedExamIds.add(
        canonicalExamId
      );
    }


    if (canonicalExamName) {
      completedExamNames.add(
        canonicalExamName
      );
    }


    enriched.push({
      record,
      exam
    });

  });


  // ---------------------------------------------------------------
  // PROGRAM SUMMARY RULES
  // ---------------------------------------------------------------

  programRules
    .filter(rule =>
      !Object.prototype
        .hasOwnProperty
        .call(rule, 'Active') ||
      reqTrue_(
        rule['Active']
      )
    )
    .forEach(rule => {

      const ruleId =
        String(
          rule['Rule ID'] || ''
        ).trim();


      const ruleName =
        String(
          rule['Rule Name'] || ''
        ).trim();


      const ruleType =
        reqNormalize_(
          rule['Rule Type']
        );


      const requirementType =
        reqNormalize_(
          rule[
            'Requirement Type'
          ]
        );


      const requirementGroup =
        reqNormalize_(
          rule[
            'Requirement Group'
          ]
        );


      const required =
        Number(
          rule[
            'Required Count'
          ]
        ) || 0;


      if (!ruleId || !required) {
        return;
      }


      let matches = [];


      if (
        ruleType ===
        'minimum count'
      ) {

        matches =
          enriched.filter(item =>
            reqNormalize_(
              item.exam[
                'Requirement Type'
              ]
            ) ===
            requirementType
          );

      } else if (
        ruleType ===
        'minimum group count'
      ) {

        matches =
          enriched.filter(item =>
            reqNormalize_(
              item.exam[
                'Requirement Type'
              ]
            ) ===
              requirementType &&
            reqNormalize_(
              item.exam[
                'Requirement Group'
              ]
            ) ===
              requirementGroup
          );

      } else {

        return;
      }


      const completed =
        matches.length;


      output.push(
        reqStatusRow_({

          studentId,
          studentName,
          clinicalTerm,

          statusType: 'Summary',
          scope: 'Program',

          id: ruleId,
          name: ruleName,

          type:
            String(
              rule[
                'Requirement Type'
              ] || ''
            ),

          group:
            String(
              rule[
                'Requirement Group'
              ] || ''
            ),

          required,
          completed,

          remaining:
            Math.max(
              required -
              completed,
              0
            ),

          status:
            completed >= required
              ? 'Complete'
              : (
                  completed > 0
                    ? 'In Progress'
                    : 'Not Started'
                ),

          lastDate:
            reqLatestDateFromRecords_(
              matches.map(
                item =>
                  item.record
              )
            ),

          sourceId:
            reqLatestSubmissionFromRecords_(
              matches.map(
                item =>
                  item.record
              )
            )
        })
      );

    });


  // ---------------------------------------------------------------
  // INDIVIDUAL REQUIREMENT OPTIONS
  // ---------------------------------------------------------------

  examinations
    .filter(exam =>
      reqTrue_(
        exam['Active']
      )
    )
    .forEach(exam => {

      const examName =
        reqNormalize_(
          exam['Examination']
        );


      const examId =
        reqNormalize_(
          exam['Examination ID']
        );


      if (
        completedExamNames.has(
          examName
        ) ||
        (
          examId &&
          completedExamIds.has(
            examId
          )
        )
      ) {
        return;
      }


      const eligibility =
        reqEvaluateEligibility_({

          exam,
          clinicalTerm,
          completedExamIds,
          examById

        });


      output.push(
        reqStatusRow_({

          studentId,
          studentName,
          clinicalTerm,

          statusType: 'Detail',
          scope: 'Program',

          id:
            String(
              exam[
                'Examination ID'
              ] || ''
            ).trim(),

          name:
            String(
              exam[
                'Examination'
              ] || ''
            ).trim(),

          type:
            String(
              exam[
                'Requirement Type'
              ] || ''
            ).trim(),

          group:
            String(
              exam[
                'Requirement Group'
              ] || ''
            ).trim(),

          timing:
            eligibility.timing,

          required: 1,
          completed: 0,
          remaining: 1,

          status:
            'Still Needed',

          notes:
            eligibility.note
        })
      );

    });

}


/**
 * =====================================================================
 * PREREQUISITE / ELIGIBILITY
 * =====================================================================
 */

function reqEvaluateEligibility_(args) {

  const {
    exam,
    clinicalTerm,
    completedExamIds,
    examById
  } = args;


  const currentNumber =
    reqClinicalNumber_(
      clinicalTerm
    );


  const earliestNumber =
    reqClinicalNumber_(
      exam['Earliest Clinical']
    );


  if (
    earliestNumber &&
    currentNumber &&
    earliestNumber >
      currentNumber
  ) {

    return {
      timing:
        'Future Requirement',

      note:
        'Eligible beginning ' +
        exam['Earliest Clinical'] +
        '.'
    };
  }


  const type =
    reqNormalize_(
      exam['Prerequisite Type']
    ) || 'none';


  if (type === 'none') {

    return {
      timing: 'Eligible Now',
      note: ''
    };
  }


  if (
    type ===
    'specific_exam'
  ) {

    const prerequisiteId =
      String(
        exam['Prerequisite'] || ''
      ).trim();


    if (!prerequisiteId) {

      return {
        timing:
          'Prerequisite Verification Needed',

        note:
          'SPECIFIC_EXAM prerequisite has no Examination ID configured.'
      };
    }


    const prerequisite =
      examById[
        reqNormalize_(
          prerequisiteId
        )
      ];


    if (!prerequisite) {

      return {
        timing:
          'Prerequisite Verification Needed',

        note:
          'Prerequisite ' +
          prerequisiteId +
          ' was not found.'
      };
    }


    if (
      !completedExamIds.has(
        reqNormalize_(
          prerequisiteId
        )
      )
    ) {

      return {
        timing:
          'Prerequisite Not Met',

        note:
          'Must first complete ' +
          prerequisite['Examination'] +
          '.'
      };
    }


    return {
      timing: 'Eligible Now',

      note:
        'Prerequisite completed: ' +
        prerequisite[
          'Examination'
        ] +
        '.'
    };
  }


  // Existing corresponding/same-area rules
  // remain reviewable rather than making
  // unsafe assumptions.

  return {
    timing:
      'Prerequisite Verification Needed',

    note:
      String(
        exam['Prerequisite'] || ''
      ).trim()
  };
}


/**
 * =====================================================================
 * PASS / FAIL
 * =====================================================================
 */

function reqIsPassingCompetency_(record) {

  if (
    reqNormalize_(
      record['Evaluation Type']
    ) !== 'competency'
  ) {
    return false;
  }


  if (
    reqNormalize_(
      record['Result']
    ) !== 'pass'
  ) {
    return false;
  }


  if (
    reqTrue_(
      record['Automatic Failure']
    )
  ) {
    return false;
  }


  return true;
}


function reqIsFailedCompetency_(record) {

  if (
    reqNormalize_(
      record['Evaluation Type']
    ) !== 'competency'
  ) {
    return false;
  }


  return (
    reqNormalize_(
      record['Result']
    ) === 'fail' ||
    reqTrue_(
      record['Automatic Failure']
    )
  );
}


/**
 * =====================================================================
 * UNIQUE COMPETENCIES
 * =====================================================================
 */

function reqGetUniquePassingCompetencies_(records) {

  const unique = {};


  records.forEach(record => {

    const examination =
      reqNormalize_(
        record['Examination']
      );


    const anatomicalArea =
      reqNormalize_(
        record['Anatomical Area']
      );


    if (!examination) {
      return;
    }


    const key =
      anatomicalArea
        ? examination +
          '|' +
          anatomicalArea
        : examination;


    if (!unique[key]) {

      unique[key] = record;

      return;
    }


    if (
      reqDateValue_(
        record['Date']
      ) >
      reqDateValue_(
        unique[key]['Date']
      )
    ) {

      unique[key] = record;
    }

  });


  return Object.values(
    unique
  );
}


/**
 * =====================================================================
 * TERM / DATE
 * =====================================================================
 */

function reqGetTermStatus_(
  start,
  end
) {

  const today =
    reqDateOnly_(
      new Date()
    );


  if (
    today.getTime() <
    start.getTime()
  ) {
    return 'Upcoming';
  }


  if (
    today.getTime() >
    end.getTime()
  ) {
    return 'Completed';
  }


  return 'Current';
}


function reqDateFallsWithin_(
  value,
  start,
  end
) {

  const date =
    reqDateOnly_(value);


  if (!date) {
    return false;
  }


  return (
    date.getTime() >=
      start.getTime() &&
    date.getTime() <=
      end.getTime()
  );
}


function reqDateOnly_(value) {

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


function reqDateValue_(value) {

  const date =
    reqDateOnly_(value);


  return date
    ? date.getTime()
    : 0;
}


function reqFormatDateRange_(
  start,
  end
) {

  const timezone =
    Session.getScriptTimeZone();


  return (
    Utilities.formatDate(
      start,
      timezone,
      'M/d/yyyy'
    ) +
    ' - ' +
    Utilities.formatDate(
      end,
      timezone,
      'M/d/yyyy'
    )
  );
}


/**
 * =====================================================================
 * REQUIREMENT HELPERS
 * =====================================================================
 */

function reqParseRequirement_(value) {

  if (
    value === '' ||
    value === null ||
    typeof value ===
      'undefined'
  ) {
    return '';
  }


  const number =
    Number(value);


  if (!isNaN(number)) {
    return number;
  }


  return String(value).trim();
}


function reqProgressStatus_(
  required,
  completed
) {

  if (
    typeof required !==
    'number'
  ) {
    return 'Program Completion Required';
  }


  if (completed >= required) {
    return 'Complete';
  }


  if (completed > 0) {
    return 'In Progress';
  }


  return 'Not Started';
}


/**
 * =====================================================================
 * OUTPUT ROW
 * =====================================================================
 */

function reqStatusRow_(data) {

  return [

    data.studentId || '',
    data.studentName || '',
    data.clinicalTerm || '',

    data.statusType || '',
    data.scope || '',

    data.id || '',
    data.name || '',
    data.type || '',
    data.group || '',
    data.timing || '',

    (
      typeof data.required !==
      'undefined'
        ? data.required
        : ''
    ),

    (
      typeof data.completed !==
      'undefined'
        ? data.completed
        : ''
    ),

    (
      typeof data.remaining !==
      'undefined'
        ? data.remaining
        : ''
    ),

    data.status || '',
    data.lastDate || '',
    data.sourceId || '',
    data.notes || ''

  ];
}


/**
 * =====================================================================
 * LATEST RECORD
 * =====================================================================
 */

function reqGetLatestRecord_(records) {

  if (
    !records ||
    records.length === 0
  ) {
    return null;
  }


  return records.reduce(
    (latest, current) => {

      if (!latest) {
        return current;
      }


      return (
        reqDateValue_(
          current['Date']
        ) >
        reqDateValue_(
          latest['Date']
        )
      )
        ? current
        : latest;

    },
    null
  );
}


function reqLatestDateFromRecords_(records) {

  const latest =
    reqGetLatestRecord_(records);


  return latest
    ? latest['Date'] || ''
    : '';
}


function reqLatestSubmissionFromRecords_(records) {

  const latest =
    reqGetLatestRecord_(records);


  return latest
    ? latest['Submission ID'] || ''
    : '';
}


/**
 * =====================================================================
 * STUDENTS
 * =====================================================================
 */

function reqGetUniqueActiveStudents_(students) {

  const unique = {};


  students
    .filter(student =>
      reqTrue_(
        student['Active']
      )
    )
    .forEach(student => {

      const id =
        String(
          student['Student ID'] || ''
        ).trim();


      const name =
        String(
          student['Student'] || ''
        ).trim();


      const key =
        id ||
        reqNormalize_(name);


      if (!key) {
        return;
      }


      if (!unique[key]) {

        unique[key] = student;

        return;
      }


      const existingTerm =
        reqClinicalNumber_(
          unique[key][
            'Clinical Term'
          ]
        );


      const candidateTerm =
        reqClinicalNumber_(
          student[
            'Clinical Term'
          ]
        );


      if (
        candidateTerm >
        existingTerm
      ) {

        unique[key] = student;
      }

    });


  return Object.values(
    unique
  );
}


function reqSameStudent_(
  record,
  studentId,
  studentName
) {

  const recordId =
    String(
      record['Student ID'] || ''
    ).trim();


  if (
    studentId &&
    recordId
  ) {

    return (
      recordId === studentId
    );
  }


  return (
    reqNormalize_(
      record['Student']
    ) ===
    reqNormalize_(
      studentName
    )
  );
}


/**
 * =====================================================================
 * GENERIC HELPERS
 * =====================================================================
 */

function reqClinicalNumber_(value) {

  const match =
    String(
      value || ''
    ).match(
      /clinical\s*(\d+)/i
    );


  return match
    ? Number(match[1])
    : 0;
}


function reqTermId_(clinicalTerm) {

  return (
    'SEM-' +
    String(
      clinicalTerm || ''
    )
      .trim()
      .replace(
        /\s+/g,
        '-'
      )
      .toUpperCase()
  );
}


function reqSafeId_(value) {

  return String(
    value || ''
  )
    .toUpperCase()
    .replace(
      /[^A-Z0-9]+/g,
      '-'
    )
    .replace(
      /^-+|-+$/g,
      ''
    );
}


function reqNormalize_(value) {

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


function reqTrue_(value) {

  if (value === true) {
    return true;
  }


  const normalized =
    reqNormalize_(value);


  return (
    normalized === 'true' ||
    normalized === 'yes' ||
    normalized === '1'
  );
}


/**
 * =====================================================================
 * LOOKUPS
 * =====================================================================
 */

function reqBuildLookup_(
  rows,
  field
) {

  const lookup = {};


  rows.forEach(row => {

    if (
      Object.prototype
        .hasOwnProperty
        .call(row, 'Active') &&
      !reqTrue_(
        row['Active']
      )
    ) {
      return;
    }


    const key =
      reqNormalize_(
        row[field]
      );


    if (key) {
      lookup[key] = row;
    }

  });


  return lookup;
}


/**
 * =====================================================================
 * SHEET HELPERS
 * =====================================================================
 */

function reqGetSheet_(ss, name) {

  const sheet =
    ss.getSheetByName(name);


  if (!sheet) {

    throw new Error(
      'Required sheet not found: "' +
      name +
      '".'
    );
  }


  return sheet;
}


function reqSheetToObjects_(sheet) {

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
    values[0].map(header =>
      String(header).trim()
    );


  return values
    .slice(1)
    .filter(row =>
      row.some(cell =>
        cell !== ''
      )
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
 * =====================================================================
 * HEADER VALIDATION
 * =====================================================================
 */

function reqValidateRequiredHeaders_(
  sheet,
  requiredHeaders
) {

  const lastColumn =
    Math.max(
      sheet.getLastColumn(),
      1
    );


  const headers =
    sheet
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


  const missing =
    requiredHeaders.filter(header =>
      !headers.includes(header)
    );


  if (missing.length > 0) {

    throw new Error(
      'Sheet "' +
      sheet.getName() +
      '" is missing required header(s): ' +
      missing.join(', ')
    );
  }
}


function reqValidateExactHeaders_(
  sheet,
  expectedHeaders
) {

  const actual =
    sheet
      .getRange(
        1,
        1,
        1,
        expectedHeaders.length
      )
      .getValues()[0]
      .map(header =>
        String(header).trim()
      );


  expectedHeaders.forEach(
    (expected, index) => {

      if (
        actual[index] !==
        expected
      ) {

        throw new Error(
          'Header mismatch in "' +
          sheet.getName() +
          '" column ' +
          (index + 1) +
          '. Expected "' +
          expected +
          '" but found "' +
          actual[index] +
          '".'
        );
      }

    }
  );
}


/**
 * =====================================================================
 * CLEAR OUTPUT
 * =====================================================================
 */

function reqClearStatusData_(
  sheet,
  columnCount
) {

  const lastRow =
    sheet.getLastRow();


  if (lastRow <= 1) {
    return;
  }


  sheet
    .getRange(
      2,
      1,
      lastRow - 1,
      columnCount
    )
    .clearContent();
}


/**
 * Returns the semester in which a Clinical Tracking event belongs.
 * Attributed Clinical Term is the current event-term field. Clinical Term
 * remains the fallback for HIST-* and other established records.
 */
function reqRecordClinicalTerm_(record) {

  return String(
    record &&
    (
      record['Attributed Clinical Term'] ||
      record['Clinical Term'] ||
      ''
    )
  ).trim();
}


/**
 * Non-writing regression test for semester ownership and competency type.
 * Uses fixtures only; it does not alter Isabella, TEST-001, Clinical
 * Tracking, Grade Records, or Historical Import staging.
 */
function testRequirementsSemesterAttributionRegression() {

  const fixtures = [
    {
      name: 'Current Clinical 3 competency',
      record: {
        'Submission ID': 'TEST-C3-COMP',
        'Clinical Term': 'Clinical 3',
        'Evaluation Type': 'Competency',
        'Examination': 'Test Examination A',
        'Result': 'Pass',
        'Automatic Failure': false,
        'Date': new Date(2020, 0, 1)
      },
      expected: true
    },
    {
      name: 'Historical Clinical 1 competency entered today',
      record: {
        'Submission ID': 'TEST-C1-COMP',
        'Clinical Term': 'Clinical 1',
        'Evaluation Type': 'Competency',
        'Examination': 'Test Examination B',
        'Result': 'Pass',
        'Automatic Failure': false,
        'Date': new Date()
      },
      expected: false
    },
    {
      name: 'Current Clinical 3 ISO',
      record: {
        'Submission ID': 'TEST-C3-ISO',
        'Clinical Term': 'Clinical 3',
        'Evaluation Type': 'ISO',
        'Examination': 'Test Examination C',
        'Result': 'Pass',
        'Automatic Failure': false,
        'Date': new Date()
      },
      expected: false
    },
    {
      name: 'Attributed term overrides snapshot term',
      record: {
        'Submission ID': 'TEST-ATTRIBUTED-C1',
        'Clinical Term': 'Clinical 3',
        'Attributed Clinical Term': 'Clinical 1',
        'Evaluation Type': 'Competency',
        'Examination': 'Test Examination D',
        'Result': 'Pass',
        'Automatic Failure': false,
        'Date': new Date()
      },
      expected: false
    }
  ];

  const selectedTerm = 'Clinical 3';

  const results =
    fixtures.map(fixture => {

      const belongsToSelectedTerm =
        reqNormalize_(
          reqRecordClinicalTerm_(fixture.record)
        ) ===
        reqNormalize_(selectedTerm);

      const qualifies =
        belongsToSelectedTerm &&
        reqIsPassingCompetency_(fixture.record);

      if (qualifies !== fixture.expected) {
        throw new Error(
          fixture.name +
          ' failed. Expected ' +
          fixture.expected +
          ', received ' +
          qualifies +
          '.'
        );
      }

      return {
        test: fixture.name,
        passed: true,
        storedTerm:
          reqRecordClinicalTerm_(fixture.record),
        countsForClinical3:
          qualifies
      };
    });

  const clinical3Count =
    fixtures
      .filter(fixture =>
        reqNormalize_(
          reqRecordClinicalTerm_(fixture.record)
        ) ===
        reqNormalize_(selectedTerm)
      )
      .filter(fixture =>
        reqIsPassingCompetency_(fixture.record)
      )
      .length;

  if (clinical3Count !== 1) {
    throw new Error(
      'Expected exactly one qualifying Clinical 3 competency; received ' +
      clinical3Count +
      '.'
    );
  }

  return {
    success: true,
    selectedTerm: selectedTerm,
    qualifyingCompetencies: clinical3Count,
    results: results
  };
}
