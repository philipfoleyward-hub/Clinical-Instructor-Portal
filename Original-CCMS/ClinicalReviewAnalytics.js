/**
 * =====================================================================
 * CLINICAL ASSESSMENT ANALYTICS
 * Complete replacement for the existing ClinicalReviewAnalytics.gs
 * =====================================================================
 *
 * PURPOSE
 * -------
 * Read-only analytics backend for Dashboard > Assessment.
 *
 * REVISIONS
 * ---------
 * 1. The former Student dimension is now COHORT-WIDE.
 *    - The public key remains "student" for backward compatibility with
 *      the existing Dashboard UI.
 *    - The displayed label is "Cohort".
 *    - Options are dynamically populated from Students > Cohort.
 *    - Selecting a cohort returns every Clinical Tracking evaluation
 *      belonging to students in that cohort.
 *
 * 2. Examination analytics use FIRST ATTEMPT ONLY.
 *    - One qualifying attempt per Student ID + selected Examination.
 *    - Qualifying attempt = Pass, Fail, or Automatic Failure.
 *    - Later fails/passes remain in Clinical Tracking but are excluded
 *      from Examination Assessment metrics and detail results.
 *    - Earliest Date wins; sheet row is the stable tie-breaker.
 *
 * 3. Evaluator analytics are ID-based.
 *    - Evaluator ID is authoritative for College CIs.
 *    - Tracking records without Evaluator ID resolve to the generic
 *      Hospital CI evaluator ID from the Evaluators sheet.
 *    - Typed hospital evaluator names remain visible in detail rows.
 *
 * 4. Automatic Failure always counts as a failure for metrics.
 *
 * 5. No writes are performed by this file.
 * =====================================================================
 */


function clinicalReview_getDimensions() {
  return [
    { key: 'evaluator', label: 'Evaluator' },
    { key: 'examination', label: 'Examination' },

    /*
     * KEEPING key:"student" prevents the current UI from breaking.
     * Semantically this dimension is now COHORT.
     */
    { key: 'student', label: 'Cohort' },

    { key: 'clinicalSite', label: 'Clinical Site' }
  ];
}


/**
 * =====================================================================
 * OPTIONS
 * =====================================================================
 */
function clinicalReview_getOptions(dimension) {

  const key =
    clinicalReview_normalizeDimension_(dimension);

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const tracking =
    clinicalReview_sheetObjects_(
      clinicalReview_getSheet_(
        ss,
        APP_CONFIG.SHEETS.CLINICAL_TRACKING
      )
    );


  // -------------------------------------------------------------------
  // COHORT
  // -------------------------------------------------------------------
  if (key === 'student') {

    const students =
      clinicalReview_sheetObjects_(
        clinicalReview_getSheet_(
          ss,
          APP_CONFIG.SHEETS.STUDENTS
        )
      );

    const seen = {};

    return students
      .map(r =>
        clinicalReview_text_(
          r['Cohort']
        )
      )
      .filter(Boolean)
      .filter(value => {

        const normalized =
          clinicalReview_normalizeCohort_(value);

        if (!normalized || seen[normalized]) {
          return false;
        }

        seen[normalized] = true;
        return true;
      })
      .sort(
        (a, b) =>
          clinicalReview_cohortSortValue_(a) -
          clinicalReview_cohortSortValue_(b)
      )
      .map(value => ({
        value:
          clinicalReview_normalizeCohort_(value),

        label:
          clinicalReview_cohortLabel_(value),

        meta: ''
      }));
  }


  // -------------------------------------------------------------------
  // EVALUATOR — AUTHORITATIVE IDs ONLY
  // -------------------------------------------------------------------
  if (key === 'evaluator') {

    const evaluators =
      clinicalReview_getEvaluatorMaster_(ss);

    return evaluators
      .filter(r =>
        clinicalReview_evaluatorId_(r)
      )
      .filter(r => {

        /*
         * If an Active column exists, respect it.
         * If it does not exist, the row is still valid.
         */
        if (
          Object.prototype.hasOwnProperty.call(
            r,
            'Active'
          ) &&
          !clinicalReview_true_(
            r['Active']
          )
        ) {
          return false;
        }

        return true;
      })
      .map(r => {

        const id =
          clinicalReview_evaluatorId_(r);

        const name =
          clinicalReview_evaluatorName_(r) ||
          id;

        const type =
          clinicalReview_text_(
            r['Evaluator Type']
          );

        return {
          value: id,
          label: name,
          meta: type
        };
      })
      .filter(
        (item, index, items) =>
          items.findIndex(other =>
            clinicalReview_normalize_(other.value) ===
            clinicalReview_normalize_(item.value)
          ) === index
      )
      .sort(
        (a, b) =>
          a.label.localeCompare(
            b.label
          )
      );
  }


  // -------------------------------------------------------------------
  // EXAMINATION
  // -------------------------------------------------------------------
  if (key === 'examination') {

    let values = [];

    const master =
      clinicalReview_trySheetObjects_(
        ss,
        APP_CONFIG.SHEETS.EXAMINATIONS ||
          'Examinations'
      );

    master.forEach(r => {

      if (
        Object.prototype.hasOwnProperty.call(
          r,
          'Active'
        ) &&
        !clinicalReview_true_(
          r['Active']
        )
      ) {
        return;
      }

      const value =
        clinicalReview_text_(
          r['Examination']
        );

      if (value) {
        values.push(value);
      }
    });

    tracking.forEach(r => {

      const value =
        clinicalReview_text_(
          r['Examination']
        );

      if (value) {
        values.push(value);
      }
    });

    return clinicalReview_textOptions_(
      values
    );
  }


  // -------------------------------------------------------------------
  // CLINICAL SITE
  // -------------------------------------------------------------------
  if (key === 'clinicalSite') {

    let values = [];

    const master =
      clinicalReview_trySheetObjects_(
        ss,
        APP_CONFIG.SHEETS.CLINICAL_SITES ||
          'Clinical Sites'
      );

    master.forEach(r => {

      if (
        Object.prototype.hasOwnProperty.call(
          r,
          'Active'
        ) &&
        !clinicalReview_true_(
          r['Active']
        )
      ) {
        return;
      }

      const value =
        clinicalReview_text_(
          r['Clinical Site'] ||
          r['Site'] ||
          r['Clinical Site Name']
        );

      if (value) {
        values.push(value);
      }
    });

    tracking.forEach(r => {

      const value =
        clinicalReview_text_(
          r['Clinical Site']
        );

      if (value) {
        values.push(value);
      }
    });

    return clinicalReview_textOptions_(
      values
    );
  }


  return [];
}


/**
 * =====================================================================
 * MAIN QUERY
 * =====================================================================
 */
function clinicalReview_query(
  dimension,
  selectedValue
) {

  const key =
    clinicalReview_normalizeDimension_(
      dimension
    );

  const value =
    clinicalReview_text_(
      selectedValue
    );

  if (!value) {
    throw new Error(
      'An Assessment value is required.'
    );
  }


  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const tracking =
    clinicalReview_sheetObjects_(
      clinicalReview_getSheet_(
        ss,
        APP_CONFIG.SHEETS.CLINICAL_TRACKING
      )
    );

  const students =
    clinicalReview_sheetObjects_(
      clinicalReview_getSheet_(
        ss,
        APP_CONFIG.SHEETS.STUDENTS
      )
    );

  const evaluatorContext =
    clinicalReview_buildEvaluatorContext_(
      ss
    );


  let selectedLabel = value;

  let matchedRecords = [];


  // -------------------------------------------------------------------
  // COHORT
  // -------------------------------------------------------------------
  if (key === 'student') {

    const cohort =
      clinicalReview_normalizeCohort_(
        value
      );

    if (!cohort) {
      throw new Error(
        'A valid cohort is required.'
      );
    }

    const cohortStudents =
      students.filter(r =>
        clinicalReview_cohortMatches_(
          r['Cohort'],
          cohort
        )
      );

    if (!cohortStudents.length) {
      throw new Error(
        'No students were found for ' +
        clinicalReview_cohortLabel_(
          cohort
        ) +
        '.'
      );
    }

    const studentIds = {};

    const studentNames = {};

    cohortStudents.forEach(r => {

      const id =
        clinicalReview_normalize_(
          r['Student ID']
        );

      const name =
        clinicalReview_normalize_(
          r['Student']
        );

      if (id) {
        studentIds[id] = true;
      }

      if (name) {
        studentNames[name] = true;
      }
    });

    matchedRecords =
      tracking.filter(record => {

        const id =
          clinicalReview_normalize_(
            record['Student ID']
          );

        if (id && studentIds[id]) {
          return true;
        }

        /*
         * Name fallback exists only for older historical rows whose
         * Student ID may be blank. ID remains preferred.
         */
        const name =
          clinicalReview_normalize_(
            record['Student']
          );

        return Boolean(
          name &&
          studentNames[name]
        );
      });

    selectedLabel =
      clinicalReview_cohortLabel_(
        cohort
      );
  }


  // -------------------------------------------------------------------
  // EVALUATOR
  // -------------------------------------------------------------------
  else if (key === 'evaluator') {

    const selectedEvaluatorId =
      clinicalReview_normalize_(
        value
      );

    const evaluator =
      evaluatorContext.byId[
        selectedEvaluatorId
      ];

    if (!evaluator) {
      throw new Error(
        'Evaluator ID not found in Evaluators: ' +
        value
      );
    }

    selectedLabel =
      evaluator.label;

    matchedRecords =
      tracking.filter(record =>
        clinicalReview_normalize_(
          clinicalReview_resolveTrackingEvaluatorId_(
            record,
            evaluatorContext
          )
        ) ===
        selectedEvaluatorId
      );
  }


  // -------------------------------------------------------------------
  // EXAMINATION
  // -------------------------------------------------------------------
  else if (key === 'examination') {

    matchedRecords =
      tracking.filter(record =>
        clinicalReview_normalize_(
          record['Examination']
        ) ===
        clinicalReview_normalize_(
          value
        )
      );

    /*
     * CRITICAL RULE:
     * For Examination Assessment, retain only the FIRST qualifying
     * attempt per student.
     */
    matchedRecords =
      clinicalReview_firstAttemptRecords_(
        matchedRecords
      );
  }


  // -------------------------------------------------------------------
  // CLINICAL SITE
  // -------------------------------------------------------------------
  else if (key === 'clinicalSite') {

    matchedRecords =
      tracking.filter(record =>
        clinicalReview_normalize_(
          record['Clinical Site']
        ) ===
        clinicalReview_normalize_(
          value
        )
      );
  }


  const evaluations =
    matchedRecords
      .map(record =>
        clinicalReview_evaluationItem_(
          record,
          evaluatorContext
        )
      )
      .sort(
        clinicalReview_evaluationSort_
      );


  const metrics =
    clinicalReview_metrics_(
      evaluations
    );

  const numeric =
    evaluations.filter(
      item =>
        item.percentScore !== null
    );


  return {

    dimension:
      key,

    dimensionLabel:
      clinicalReview_dimensionLabel_(
        key
      ),

    selectedValue:
      value,

    selectedLabel:
      selectedLabel,

    firstAttemptsOnly:
      key === 'examination',

    contextLabel:
      key === 'examination'
        ? selectedLabel +
          ' · First attempts only'
        : selectedLabel,

    description:
      clinicalReview_description_(
        key
      ),

    metrics:
      metrics,

    /*
     * evaluations is already sorted high -> low among numeric records.
     */
    highestEvaluation:
      numeric.length
        ? numeric[0]
        : null,

    lowestEvaluation:
      numeric.length
        ? numeric[
            numeric.length - 1
          ]
        : null,

    evaluations:
      evaluations
  };
}


/**
 * =====================================================================
 * METRICS
 * =====================================================================
 */
function clinicalReview_metrics_(
  evaluations
) {

  const numeric =
    evaluations.filter(
      x =>
        x.percentScore !== null
    );

  const numericScores =
    numeric.map(
      x =>
        x.percentScore
    );


  /*
   * Automatic Failure overrides any inconsistent Result text.
   */
  const passCount =
    evaluations.filter(
      x =>
        clinicalReview_evaluationOutcome_(
          x
        ) === 'pass'
    ).length;

  const failCount =
    evaluations.filter(
      x =>
        clinicalReview_evaluationOutcome_(
          x
        ) === 'fail'
    ).length;

  const decided =
    passCount +
    failCount;


  return {

    evaluationCount:
      evaluations.length,

    numericScoreCount:
      numeric.length,

    numericScoreCoverage:
      evaluations.length
        ? clinicalReview_round2_(
            (
              numeric.length /
              evaluations.length
            ) *
            100
          )
        : 0,

    averageScore:
      numericScores.length
        ? clinicalReview_round2_(
            numericScores.reduce(
              (a, b) =>
                a + b,
              0
            ) /
            numericScores.length
          )
        : null,

    highestScore:
      numericScores.length
        ? Math.max.apply(
            null,
            numericScores
          )
        : null,

    lowestScore:
      numericScores.length
        ? Math.min.apply(
            null,
            numericScores
          )
        : null,

    passRate:
      decided
        ? clinicalReview_round2_(
            (
              passCount /
              decided
            ) *
            100
          )
        : null,

    passingCount:
      passCount,

    failingCount:
      failCount,

    passFailOnlyCount:
      evaluations.length -
      numeric.length,

    automaticFailureCount:
      evaluations.filter(
        x =>
          x.automaticFailure
      ).length
  };
}


/**
 * =====================================================================
 * FIRST ATTEMPT — EXAMINATION ANALYTICS
 * =====================================================================
 *
 * One qualifying record per student.
 *
 * Qualifying:
 *   - PASS
 *   - FAIL
 *   - Automatic Failure
 *
 * Earliest date wins.
 * If two records have the same date, the original Clinical Tracking
 * row number is the stable tie-breaker.
 */
function clinicalReview_firstAttemptRecords_(
  records
) {

  const qualifying =
    (records || [])
      .filter(
        clinicalReview_isQualifyingAttempt_
      )
      .slice()
      .sort(
        clinicalReview_firstAttemptSort_
      );

  const seen = {};

  const firstAttempts = [];

  qualifying.forEach(record => {

    const key =
      clinicalReview_studentAttemptKey_(
        record
      );

    if (!key || seen[key]) {
      return;
    }

    seen[key] = true;

    firstAttempts.push(
      record
    );
  });

  return firstAttempts;
}


function clinicalReview_isQualifyingAttempt_(
  record
) {

  if (
    clinicalReview_true_(
      record[
        'Automatic Failure'
      ]
    )
  ) {
    return true;
  }

  const result =
    clinicalReview_normalize_(
      record['Result']
    );

  return (
    result === 'pass' ||
    result === 'fail' ||
    result === 'auto-fail' ||
    result === 'auto fail' ||
    result === 'automatic fail' ||
    result === 'automatic failure'
  );
}


function clinicalReview_firstAttemptSort_(
  a,
  b
) {

  const ad =
    clinicalReview_dateValue_(
      a['Date']
    );

  const bd =
    clinicalReview_dateValue_(
      b['Date']
    );

  /*
   * Valid dated attempts sort before undated records.
   * If both are dated, chronological ascending.
   */
  if (ad && bd && ad !== bd) {
    return ad - bd;
  }

  if (ad && !bd) {
    return -1;
  }

  if (!ad && bd) {
    return 1;
  }


  const ar =
    Number(
      a.__clinicalReviewRow ||
      0
    );

  const br =
    Number(
      b.__clinicalReviewRow ||
      0
    );

  if (ar !== br) {
    return ar - br;
  }


  return clinicalReview_text_(
    a['Submission ID']
  ).localeCompare(
    clinicalReview_text_(
      b['Submission ID']
    )
  );
}


function clinicalReview_studentAttemptKey_(
  record
) {

  const studentId =
    clinicalReview_normalize_(
      record['Student ID']
    );

  if (studentId) {
    return 'id:' + studentId;
  }

  const student =
    clinicalReview_normalize_(
      record['Student']
    );

  return student
    ? 'name:' + student
    : '';
}


/**
 * =====================================================================
 * EVALUATOR ID ARCHITECTURE
 * =====================================================================
 */
function clinicalReview_buildEvaluatorContext_(
  ss
) {

  const master =
    clinicalReview_getEvaluatorMaster_(
      ss
    );

  const byId = {};

  let hospitalCiId = '';

  master.forEach(record => {

    const id =
      clinicalReview_evaluatorId_(
        record
      );

    if (!id) {
      return;
    }

    const idKey =
      clinicalReview_normalize_(
        id
      );

    const name =
      clinicalReview_evaluatorName_(
        record
      ) ||
      id;

    const type =
      clinicalReview_text_(
        record['Evaluator Type']
      );

    byId[idKey] = {
      id: id,
      label: name,
      type: type
    };


    if (
      clinicalReview_isHospitalCiMaster_(
        record
      )
    ) {
      hospitalCiId = id;
    }
  });


  if (!hospitalCiId) {

    throw new Error(
      'The Evaluators sheet must contain a generic Hospital CI row with its own Evaluator ID.'
    );
  }


  return {
    byId: byId,
    hospitalCiId:
      hospitalCiId
  };
}


function clinicalReview_getEvaluatorMaster_(
  ss
) {

  return clinicalReview_sheetObjects_(
    clinicalReview_getSheet_(
      ss,
      APP_CONFIG.SHEETS.EVALUATORS ||
        'Evaluators'
    )
  );
}


function clinicalReview_evaluatorId_(
  record
) {

  return clinicalReview_text_(
    record['Evaluator ID'] ||
    record['Evaluator Id'] ||
    record['ID'] ||
    record['Id']
  );
}


function clinicalReview_evaluatorName_(
  record
) {

  return clinicalReview_text_(
    record['Evaluator'] ||
    record['Evaluator Name'] ||
    record['Name']
  );
}


function clinicalReview_isHospitalCiMaster_(
  record
) {

  const name =
    clinicalReview_normalize_(
      clinicalReview_evaluatorName_(
        record
      )
    );

  const type =
    clinicalReview_normalize_(
      record['Evaluator Type']
    );

  return (
    name === 'hospital ci' ||
    name === 'hospital clinical instructor' ||
    type === 'hospital ci' ||
    type === 'hospital clinical instructor'
  );
}


function clinicalReview_resolveTrackingEvaluatorId_(
  record,
  context
) {

  const id =
    clinicalReview_text_(
      record['Evaluator ID']
    );

  /*
   * Any explicit Evaluator ID wins.
   */
  if (id) {
    return id;
  }

  /*
   * Per current architecture:
   * No evaluator ID = Hospital CI.
   */
  return context.hospitalCiId;
}


/**
 * =====================================================================
 * EVALUATION ITEM
 * =====================================================================
 */
function clinicalReview_evaluationItem_(
  record,
  evaluatorContext
) {

  const resolved =
    clinicalReview_resolveGrade_(
      record
    );

  const evaluatorId =
    evaluatorContext
      ? clinicalReview_resolveTrackingEvaluatorId_(
          record,
          evaluatorContext
        )
      : clinicalReview_text_(
          record['Evaluator ID']
        );

  const evaluatorMaster =
    evaluatorContext &&
    evaluatorContext.byId[
      clinicalReview_normalize_(
        evaluatorId
      )
    ]
      ? evaluatorContext.byId[
          clinicalReview_normalize_(
            evaluatorId
          )
        ]
      : null;


  return {

    submissionId:
      clinicalReview_text_(
        record['Submission ID']
      ),

    studentId:
      clinicalReview_text_(
        record['Student ID']
      ),

    student:
      clinicalReview_text_(
        record['Student']
      ),

    clinicalTerm:
      clinicalReview_text_(
        record['Clinical Term']
      ),

    attributedClinicalTerm:
      clinicalReview_text_(
        record[
          'Attributed Clinical Term'
        ]
      ),

    evaluationType:
      clinicalReview_text_(
        record['Evaluation Type']
      ),

    examination:
      clinicalReview_text_(
        record['Examination']
      ),

    anatomicalArea:
      clinicalReview_text_(
        record['Anatomical Area']
      ),

    clinicalSite:
      clinicalReview_text_(
        record['Clinical Site']
      ),

    evaluatorId:
      evaluatorId,

    /*
     * Keep the evaluator actually typed on the evaluation visible.
     * This is useful for Hospital CI detail rows.
     */
    evaluator:
      clinicalReview_text_(
        record['Evaluator']
      ) ||
      (
        evaluatorMaster
          ? evaluatorMaster.label
          : ''
      ),

    evaluatorGroup:
      evaluatorMaster
        ? evaluatorMaster.label
        : '',

    evaluatorType:
      clinicalReview_text_(
        record['Evaluator Type']
      ) ||
      (
        evaluatorMaster
          ? evaluatorMaster.type
          : ''
      ),

    date:
      clinicalReview_displayDate_(
        record['Date']
      ),

    dateValue:
      clinicalReview_dateValue_(
        record['Date']
      ),

    rawScore:
      resolved.rawScore,

    percentScore:
      resolved.percent,

    gradeSource:
      resolved.source,

    gradeDisplay:
      clinicalReview_gradeDisplay_(
        resolved,
        record
      ),

    result:
      clinicalReview_text_(
        record['Result']
      ),

    automaticFailure:
      clinicalReview_true_(
        record[
          'Automatic Failure'
        ]
      ),

    sourceDocument:
      clinicalReview_text_(
        record['Source Document']
      ),

    historical:
      clinicalReview_isHistorical_(
        record
      ),

    sourceRow:
      Number(
        record.__clinicalReviewRow ||
        0
      )
  };
}


/**
 * =====================================================================
 * OUTCOME
 * =====================================================================
 */
function clinicalReview_evaluationOutcome_(
  evaluation
) {

  if (
    evaluation &&
    evaluation.automaticFailure
  ) {
    return 'fail';
  }

  const result =
    clinicalReview_normalize_(
      evaluation
        ? evaluation.result
        : ''
    );

  if (result === 'pass') {
    return 'pass';
  }

  if (
    result === 'fail' ||
    result === 'auto-fail' ||
    result === 'auto fail' ||
    result === 'automatic fail' ||
    result === 'automatic failure'
  ) {
    return 'fail';
  }

  return '';
}


/**
 * =====================================================================
 * GRADE RESOLUTION
 * =====================================================================
 */
function clinicalReview_resolveGrade_(
  record
) {

  if (
    typeof gradeEngine_resolveTrackingGrade_ ===
    'function'
  ) {

    const resolved =
      gradeEngine_resolveTrackingGrade_(
        record
      ) ||
      {};

    return {

      available:
        Boolean(
          resolved.available
        ),

      rawScore:
        clinicalReview_numberOrNull_(
          resolved.rawScore
        ),

      percent:
        clinicalReview_numberOrNull_(
          resolved.percent
        ),

      source:
        clinicalReview_text_(
          resolved.source
        )
    };
  }


  const finalScore =
    clinicalReview_numberOrNull_(
      record['Final Score']
    );

  const historicalPercent =
    clinicalReview_numberOrNull_(
      record[
        'Historical Percent Grade'
      ]
    );


  if (
    clinicalReview_isHistorical_(
      record
    ) &&
    historicalPercent !== null &&
    historicalPercent >= 0 &&
    historicalPercent <= 100
  ) {

    return {

      available: true,

      rawScore:
        finalScore,

      percent:
        historicalPercent <= 1
          ? historicalPercent * 100
          : historicalPercent,

      source:
        'Historical Percent Grade'
    };
  }


  if (
    finalScore !== null &&
    typeof gradeEngine_scoreToPercent ===
      'function'
  ) {

    return {

      available: true,

      rawScore:
        finalScore,

      percent:
        clinicalReview_numberOrNull_(
          gradeEngine_scoreToPercent(
            finalScore
          )
        ),

      source:
        'Final Score'
    };
  }


  return {
    available: false,
    rawScore: finalScore,
    percent: null,
    source: ''
  };
}


/**
 * =====================================================================
 * DESCRIPTION / LABELS
 * =====================================================================
 */
function clinicalReview_description_(
  key
) {

  if (key === 'student') {

    return (
      'Assessment summarizes all recorded clinical evaluations for every student in the selected cohort. ' +
      'Numeric averages use the same grade-resolution rules as the Grade Engine. ' +
      'Pass/Fail-only records remain visible but are excluded from numeric averages.'
    );
  }


  if (key === 'examination') {

    return (
      'Assessment uses only each student’s first qualifying attempt at the selected examination. ' +
      'The first Pass, Fail, or Automatic Failure is retained; all later attempts are excluded from Assessment metrics but remain unchanged in Clinical Tracking.'
    );
  }


  if (key === 'evaluator') {

    return (
      'Assessment groups College Clinical Instructors by Evaluator ID. ' +
      'Evaluations without an Evaluator ID are grouped under the generic Hospital CI Evaluator ID.'
    );
  }


  return (
    'Assessment summarizes recorded clinical evaluations for the selected parameter. ' +
    'Numeric averages use the same grade-resolution rules as the Grade Engine. ' +
    'Pass/Fail-only records remain visible but are excluded from numeric averages.'
  );
}


function clinicalReview_dimensionLabel_(
  key
) {

  return ({
    evaluator:
      'Evaluator',

    examination:
      'Examination',

    /*
     * Legacy internal key, new UI meaning.
     */
    student:
      'Cohort',

    clinicalSite:
      'Clinical Site'
  })[key] ||
    key;
}


/**
 * =====================================================================
 * DISPLAY SORT
 * =====================================================================
 */
function clinicalReview_evaluationSort_(
  a,
  b
) {

  const an =
    a.percentScore !== null;

  const bn =
    b.percentScore !== null;


  if (an && bn) {

    if (
      b.percentScore !==
      a.percentScore
    ) {
      return (
        b.percentScore -
        a.percentScore
      );
    }

    if (
      b.dateValue !==
      a.dateValue
    ) {
      return (
        b.dateValue -
        a.dateValue
      );
    }
  }


  if (an) {
    return -1;
  }

  if (bn) {
    return 1;
  }


  return (
    b.dateValue -
    a.dateValue
  );
}


/**
 * =====================================================================
 * GRADE DISPLAY
 * =====================================================================
 */
function clinicalReview_gradeDisplay_(
  resolved,
  record
) {

  if (
    resolved.percent !== null
  ) {

    const pct =
      clinicalReview_trimNumber_(
        resolved.percent
      ) +
      '%';


    if (
      resolved.rawScore !== null &&
      resolved.source ===
        'Final Score'
    ) {

      return (
        clinicalReview_trimNumber_(
          resolved.rawScore
        ) +
        ' • ' +
        pct
      );
    }


    return pct;
  }


  return (
    clinicalReview_text_(
      record['Result']
    ) ||
    '—'
  );
}


/**
 * =====================================================================
 * COHORT HELPERS
 * =====================================================================
 */
function clinicalReview_normalizeCohort_(
  value
) {

  const text =
    clinicalReview_normalize_(
      value
    );

  if (!text) {
    return '';
  }

  const match =
    text.match(
      /(?:class\s+of\s+)?(\d{4})/
    );

  return match
    ? match[1]
    : text;
}


function clinicalReview_cohortMatches_(
  value,
  cohort
) {

  return (
    clinicalReview_normalizeCohort_(
      value
    ) ===
    clinicalReview_normalizeCohort_(
      cohort
    )
  );
}


function clinicalReview_cohortLabel_(
  value
) {

  const cohort =
    clinicalReview_normalizeCohort_(
      value
    );

  return /^\d{4}$/.test(
    cohort
  )
    ? 'Class of ' + cohort
    : cohort;
}


function clinicalReview_cohortSortValue_(
  value
) {

  const cohort =
    clinicalReview_normalizeCohort_(
      value
    );

  const number =
    Number(
      cohort
    );

  return Number.isFinite(
    number
  )
    ? number
    : 999999;
}


/**
 * =====================================================================
 * DIMENSION NORMALIZATION
 * =====================================================================
 */
function clinicalReview_normalizeDimension_(
  dimension
) {

  const value =
    clinicalReview_normalize_(
      dimension
    )
      .replace(
        /\s+/g,
        ''
      );


  const map = {

    evaluator:
      'evaluator',

    examination:
      'examination',

    exam:
      'examination',

    /*
     * "student" remains accepted for the current Dashboard.
     * "cohort" is also accepted for future UI cleanup.
     */
    student:
      'student',

    cohort:
      'student',

    clinicalsite:
      'clinicalSite',

    site:
      'clinicalSite'
  };


  if (!map[value]) {

    throw new Error(
      'Unsupported Assessment dimension: ' +
      dimension
    );
  }


  return map[value];
}


/**
 * =====================================================================
 * TEXT OPTIONS
 * =====================================================================
 */
function clinicalReview_textOptions_(
  values
) {

  const seen = {};

  return (
    values ||
    []
  )
    .map(
      clinicalReview_text_
    )
    .filter(Boolean)
    .filter(value => {

      const key =
        clinicalReview_normalize_(
          value
        );

      if (seen[key]) {
        return false;
      }

      seen[key] = true;

      return true;
    })
    .sort(
      (a, b) =>
        a.localeCompare(
          b
        )
    )
    .map(value => ({
      value: value,
      label: value,
      meta: ''
    }));
}


/**
 * =====================================================================
 * SHEET HELPERS
 * =====================================================================
 */
function clinicalReview_getSheet_(
  ss,
  name
) {

  const sheet =
    ss.getSheetByName(
      name
    );

  if (!sheet) {

    throw new Error(
      'Required sheet not found: ' +
      name
    );
  }

  return sheet;
}


function clinicalReview_trySheetObjects_(
  ss,
  name
) {

  if (!name) {
    return [];
  }

  const sheet =
    ss.getSheetByName(
      name
    );

  return sheet
    ? clinicalReview_sheetObjects_(
        sheet
      )
    : [];
}


function clinicalReview_sheetObjects_(
  sheet
) {

  if (
    !sheet ||
    sheet.getLastRow() < 2 ||
    sheet.getLastColumn() < 1
  ) {
    return [];
  }


  const data =
    sheet
      .getDataRange()
      .getValues();


  const headers =
    data[0].map(
      clinicalReview_text_
    );


  return data
    .slice(1)
    .map(
      (row, index) => {

        const obj = {};

        headers.forEach(
          (header, columnIndex) => {

            if (header) {
              obj[header] =
                row[columnIndex];
            }
          }
        );

        /*
         * Internal stable source-order key.
         * Never written back to the spreadsheet.
         */
        obj.__clinicalReviewRow =
          index + 2;

        return obj;
      }
    );
}


/**
 * =====================================================================
 * HISTORICAL / BOOLEAN / TEXT / DATE HELPERS
 * =====================================================================
 */
function clinicalReview_isHistorical_(
  record
) {

  if (
    typeof gradeEngine_isHistoricalTrackingRecord_ ===
    'function'
  ) {

    return Boolean(
      gradeEngine_isHistoricalTrackingRecord_(
        record
      )
    );
  }


  return (

    clinicalReview_normalize_(
      record['Submission ID']
    ).indexOf(
      'hist-'
    ) === 0 ||

    clinicalReview_normalize_(
      record['Submission Method']
    ).indexOf(
      'historical'
    ) !== -1 ||

    clinicalReview_normalize_(
      record['Source']
    ).indexOf(
      'historical'
    ) !== -1
  );
}


function clinicalReview_true_(
  value
) {

  if (
    typeof appIsTrue_ ===
    'function'
  ) {
    return appIsTrue_(
      value
    );
  }


  if (
    value === true ||
    value === 1
  ) {
    return true;
  }


  const normalized =
    clinicalReview_normalize_(
      value
    );


  return (
    normalized === 'true' ||
    normalized === 'yes' ||
    normalized === 'y' ||
    normalized === '1'
  );
}


function clinicalReview_text_(
  value
) {

  return (
    value === null ||
    value === undefined
  )
    ? ''
    : String(
        value
      ).trim();
}


function clinicalReview_normalize_(
  value
) {

  return clinicalReview_text_(
    value
  )
    .toLowerCase()
    .replace(
      /\s+/g,
      ' '
    )
    .trim();
}


function clinicalReview_numberOrNull_(
  value
) {

  if (
    value === '' ||
    value === null ||
    value === undefined
  ) {
    return null;
  }


  const number =
    Number(
      value
    );


  return Number.isFinite(
    number
  )
    ? number
    : null;
}


function clinicalReview_round2_(
  value
) {

  return Math.round(
    (
      Number(
        value
      ) +
      Number.EPSILON
    ) *
    100
  ) /
  100;
}


function clinicalReview_trimNumber_(
  value
) {

  return clinicalReview_round2_(
    value
  )
    .toFixed(
      2
    )
    .replace(
      /\.?0+$/,
      ''
    );
}


function clinicalReview_dateValue_(
  value
) {

  if (!value) {
    return 0;
  }


  const date =
    value instanceof Date
      ? value
      : new Date(
          value
        );


  return Number.isNaN(
    date.getTime()
  )
    ? 0
    : date.getTime();
}


function clinicalReview_displayDate_(
  value
) {

  const time =
    clinicalReview_dateValue_(
      value
    );

  if (!time) {
    return '';
  }


  return Utilities.formatDate(
    new Date(
      time
    ),
    Session.getScriptTimeZone(),
    'MM/dd/yyyy'
  );
}


/**
 * =====================================================================
 * READ-ONLY QA
 * =====================================================================
 *
 * Run this after replacing the existing ClinicalReviewAnalytics.gs.
 * It performs NO WRITES.
 */
function testClinicalAssessmentArchitecture() {

  const result = {

    success: true,

    readOnly: true,

    dimensions: []
  };


  clinicalReview_getDimensions()
    .forEach(dimension => {

      const options =
        clinicalReview_getOptions(
          dimension.key
        );

      const sample =
        options.length
          ? clinicalReview_query(
              dimension.key,
              options[0].value
            )
          : null;


      result.dimensions.push({

        key:
          dimension.key,

        label:
          dimension.label,

        optionCount:
          options.length,

        firstOption:
          options.length
            ? options[0]
            : null,

        sample:
          sample
            ? {

                selectedLabel:
                  sample.selectedLabel,

                firstAttemptsOnly:
                  sample.firstAttemptsOnly,

                evaluationCount:
                  sample.metrics.evaluationCount,

                averageScore:
                  sample.metrics.averageScore,

                highestScore:
                  sample.metrics.highestScore,

                lowestScore:
                  sample.metrics.lowestScore,

                passRate:
                  sample.metrics.passRate,

                automaticFailureCount:
                  sample.metrics.automaticFailureCount
              }
            : null
      });
    });


  Logger.log(
    JSON.stringify(
      result,
      null,
      2
    )
  );


  return result;
}


/**
 * Backward-compatible old test name.
 */
function testClinicalReviewArchitecture() {
  return testClinicalAssessmentArchitecture();
}
