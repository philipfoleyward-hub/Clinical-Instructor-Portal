/**
 * =====================================================================
 * CLINICAL COMPETENCY TRACKER
 * GRADE ENGINE
 * 04_GradeEngine.gs
 * =====================================================================
 *
 * PURPOSE
 * -------
 * Authoritative student-level clinical grade engine.
 *
 * THIS ENGINE:
 *
 *   1. Converts individual 0-3 clinical grades to the official percent.
 *   2. Calculates category averages.
 *   3. Applies the correct semester-specific category weights.
 *   4. Calculates a PARTIAL weighted average when some grade categories
 *      have not yet been entered.
 *   5. Calculates the true weighted semester grade automatically once
 *      all required categories are available.
 *   6. Calculates an overall recorded program average from semester
 *      averages currently available.
 *
 * IMPORTANT DISPLAY RULE
 * ----------------------
 *
 * Individual clinical evaluation:
 *
 *      2.74  •  90.5%
 *
 * Semester / Program summary:
 *
 *      90.5%
 *
 * If categories are missing:
 *
 *      Partial weighted average
 *      Does not yet include Professional Competency,
 *      Clinical Test, Clinical Quizzes, etc.
 *
 * =====================================================================
 */


/**
 * =====================================================================
 * OFFICIAL COURSE WEIGHTS
 * =====================================================================
 *
 * These are the official category weights for Clinical Practicum 1-5.
 */
const GRADE_ENGINE_COURSE_WEIGHTS = {

  'clinical 1': {

    competency: 35,
    iso: 15,
    professional: 15,
    clinicalTest: 25,
    casePresentation: 0,
    quizzes: 10

  },

  'clinical 2': {

    competency: 15,
    iso: 20,
    professional: 15,
    clinicalTest: 25,
    casePresentation: 15,
    quizzes: 10

  },

  'clinical 3': {

    competency: 20,
    iso: 25,
    professional: 15,
    clinicalTest: 30,
    casePresentation: 0,
    quizzes: 10

  },

  'clinical 4': {

    competency: 20,
    iso: 25,
    professional: 15,
    clinicalTest: 30,
    casePresentation: 0,
    quizzes: 10

  },

  'clinical 5': {

    competency: 15,
    iso: 20,
    professional: 15,
    clinicalTest: 25,
    casePresentation: 15,
    quizzes: 10

  }

};

/**
 * =====================================================================
 * PUBLIC STUDENT GRADE SUMMARY
 * =====================================================================
 */
function gradeEngine_getStudentSummary(
  studentId
) {

  if (!studentId) {

    throw new Error(
      'Student ID is required.'
    );
  }


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
        APP_CONFIG.SHEETS
          .CLINICAL_TRACKING
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


  const student =
    gradeEngine_getMostAdvancedStudentRow_(
      studentRows
    );


  const studentName =
    String(
      student['Student'] || ''
    ).trim();


  const currentClinicalTerm =
    String(
      student['Clinical Term'] || ''
    ).trim();


  // ===================================================================
  // CLINICAL TRACKING
  // ===================================================================

  const studentTracking =
    tracking.filter(record =>

      gradeEngine_sameStudent_(
        record,
        studentId,
        studentName
      )

    );


  // ===================================================================
  // OPTIONAL SUPPLEMENTAL GRADE RECORDS
  //
  // This prepares the engine for:
  //
  //   Add New -> Professional Competency
  //   Add New -> Clinical Test
  //   Add New -> Clinical Quiz
  //   Add New -> Case Presentation
  //   Add New -> Point Deduction
  //
  // The sheet does NOT have to exist yet.
  // ===================================================================

  const allGradeRecords =
    gradeEngine_getSupplementalGradeRecords_(
      ss
    );


  const studentGradeRecords =
    allGradeRecords.filter(record =>

      gradeEngine_sameStudent_(
        record,
        studentId,
        studentName
      )

    );


  // ===================================================================
  // BUILD TERM SUMMARIES
  // ===================================================================

  const terms =
    [];


  for (
    let clinicalNumber = 1;
    clinicalNumber <= 5;
    clinicalNumber++
  ) {

    const clinicalTerm =
      'Clinical ' +
      clinicalNumber;


    const termTracking =
      studentTracking.filter(record =>

        appNormalize_(
          gradeEngine_effectiveClinicalTerm_(
            record
          )
        ) ===
        appNormalize_(
          clinicalTerm
        )

      );


    const termGradeRecords =
      studentGradeRecords.filter(record =>

        appNormalize_(
          record['Clinical Term']
        ) ===
        appNormalize_(
          clinicalTerm
        )

      );


    const summary =
      gradeEngine_buildTermSummary_({

        clinicalTerm:
          clinicalTerm,

        tracking:
          termTracking,

        gradeRecords:
          termGradeRecords

      });


    /*
     * Do not clutter program calculations with future semesters
     * containing absolutely no grade information.
     */
    if (
      summary.hasGradeData
    ) {

      terms.push(
        summary
      );
    }
  }


  // ===================================================================
  // CURRENT SEMESTER
  // ===================================================================

  let currentSemester =
    terms.find(term =>

      appNormalize_(
        term.clinicalTerm
      ) ===
      appNormalize_(
        currentClinicalTerm
      )

    );


  if (!currentSemester) {

    currentSemester =
      gradeEngine_buildTermSummary_({

        clinicalTerm:
          currentClinicalTerm,

        tracking:
          [],

        gradeRecords:
          []

      });
  }


  // ===================================================================
  // OVERALL PROGRAM RECORDED AVERAGE
  //
  // Each available semester contributes one semester average.
  //
  // Future semesters with no grades are not included.
  // ===================================================================

  const program =
    gradeEngine_buildProgramSummary_(
      terms
    );


  return {

    studentId:
      String(
        student['Student ID'] ||
        studentId
      ).trim(),

    student:
      studentName,

    clinicalTerm:
      currentClinicalTerm,

    semester:
      currentSemester,

    program:
      program,

    terms:
      terms

  };
}


/**
 * =====================================================================
 * BUILD ONE SEMESTER
 * =====================================================================
 */
function gradeEngine_buildTermSummary_(
  args
) {

  const clinicalTerm =
    args.clinicalTerm;


  const tracking =
    args.tracking || [];


  const gradeRecords =
    args.gradeRecords || [];


  const weights =
    gradeEngine_getCourseWeights_(
      clinicalTerm
    );


  // ===================================================================
  // COMPETENCY RECORDS
  // ===================================================================

  const competencyRecords =
    tracking.filter(record =>

      appNormalize_(
        record['Evaluation Type']
      ) ===
      'competency'

    );


  // ===================================================================
  // ISO RECORDS
  // ===================================================================

  const isoRecords =
    tracking.filter(record =>

      appNormalize_(
        record['Evaluation Type']
      ) ===
      'iso'

    );


  // ===================================================================
  // SUPPLEMENTAL CATEGORIES
  // ===================================================================

  const professionalRecords =
    gradeEngine_filterGradeRecordsByCategory_(

      gradeRecords,
      'professional'

    );


  const clinicalTestRecords =
    gradeEngine_filterGradeRecordsByCategory_(

      gradeRecords,
      'clinicalTest'

    );


  const casePresentationRecords =
    gradeEngine_filterGradeRecordsByCategory_(

      gradeRecords,
      'casePresentation'

    );


  const quizRecords =
    gradeEngine_filterGradeRecordsByCategory_(

      gradeRecords,
      'quizzes'

    );


  const deductionRecords =
    gradeEngine_filterGradeRecordsByCategory_(

      gradeRecords,
      'deduction'

    );


  // ===================================================================
  // CATEGORY AVERAGES
  // ===================================================================

  const competency =
    gradeEngine_calculateClinicalCategory_(
      competencyRecords
    );


  const iso =
    gradeEngine_calculateClinicalCategory_(
      isoRecords
    );


  const professional =
    gradeEngine_calculateSupplementalCategory_(

      professionalRecords,
      true

    );


  const clinicalTest =
    gradeEngine_calculateSupplementalCategory_(

      clinicalTestRecords,
      false

    );


  const casePresentation =
    gradeEngine_calculateSupplementalCategory_(

      casePresentationRecords,
      false

    );


  const quizzes =
    gradeEngine_calculateSupplementalCategory_(

      quizRecords,
      false

    );


  const categories = {

    competency:
      competency,

    iso:
      iso,

    professional:
      professional,

    clinicalTest:
      clinicalTest,

    casePresentation:
      casePresentation,

    quizzes:
      quizzes

  };


  // ===================================================================
  // DETERMINE INCLUDED / MISSING CATEGORIES
  // ===================================================================

  const includedCategories =
    [];


  const missingCategories =
    [];


  let availableWeight =
    0;


  let weightedPoints =
    0;


  Object
    .keys(
      weights
    )
    .forEach(categoryKey => {

      const weight =
        Number(
          weights[
            categoryKey
          ]
        ) || 0;


      /*
       * Example:
       * Case Presentation has weight zero in Clinical 1/3/4.
       * It should not appear as missing.
       */
      if (
        weight <= 0
      ) {

        return;
      }


      const category =
        categories[
          categoryKey
        ];


      if (
        category &&
        category.available
      ) {

        includedCategories.push(
          gradeEngine_categoryLabel_(
            categoryKey
          )
        );


        availableWeight +=
          weight;


        weightedPoints +=

          category.percentAverage *
          weight;


      } else {

        missingCategories.push(
          gradeEngine_categoryLabel_(
            categoryKey
          )
        );
      }

    });


  // ===================================================================
  // PARTIAL WEIGHTED AVERAGE
  //
  // Example:
  //
  // Clinical 2:
  //
  // Competency = 92%, weight 15
  // ISO        = 93%, weight 20
  //
  // available weight = 35
  //
  // partial average =
  //
  // (92*15 + 93*20) / 35
  //
  // NOT:
  //
  // (92*15 + 93*20) / 100
  //
  // This gives the best current estimate from actually-recorded data
  // without pretending missing categories are zeros.
  // ===================================================================

  let averageBeforeDeductions =
    null;


  if (
    availableWeight > 0
  ) {

    averageBeforeDeductions =

      weightedPoints /
      availableWeight;
  }


  // ===================================================================
  // POINT DEDUCTIONS
  // ===================================================================

  const deductions =
    gradeEngine_calculateDeductions_(
      deductionRecords
    );


  let recordedAverage =
    averageBeforeDeductions;


  if (
    recordedAverage !== null
  ) {

    recordedAverage +=
      deductions.total;


    /*
     * Prevent display outside normal percentage boundaries.
     */
    recordedAverage =
      Math.max(
        0,
        Math.min(
          100,
          recordedAverage
        )
      );


    recordedAverage =
      gradeEngine_round2_(
        recordedAverage
      );
  }


  const complete =
    (
      missingCategories.length === 0 &&
      availableWeight === 100
    );


  // ===================================================================
  // STATUS / DISCLAIMER
  // ===================================================================

  let note =
    '';


  if (
    recordedAverage === null
  ) {

    note =
      'No grade-bearing records are currently available for this semester.';

  } else if (
    complete
  ) {

    note =
      'Includes all required weighted grade categories for ' +
      clinicalTerm +
      '.';

  } else {

    note =
      'Partial weighted average based only on grades currently recorded in this system.';


    if (
      missingCategories.length > 0
    ) {

      note +=

        ' Does not yet include ' +
        gradeEngine_joinList_(
          missingCategories
        ) +
        '.';
    }
  }


  return {

    clinicalTerm:
      clinicalTerm,

    hasGradeData:
      recordedAverage !== null,

    complete:
      complete,

    averageType:
      complete
        ? 'semester grade'
        : 'partial weighted average',

    recordedAverage:
      recordedAverage,

    recordedAverageText:
      gradeEngine_formatPercent_(
        recordedAverage
      ),

    averageBeforeDeductions:
      averageBeforeDeductions === null
        ? null
        : gradeEngine_round2_(
            averageBeforeDeductions
          ),

    deductions:
      deductions,

    availableWeight:
      availableWeight,

    totalRequiredWeight:
      100,

    includedCategories:
      includedCategories,

    missingCategories:
      missingCategories,

    note:
      note,

    weights:
      weights,

    categories:
      categories

  };
}


/**
 * =====================================================================
 * PROGRAM SUMMARY
 * =====================================================================
 */
function gradeEngine_buildProgramSummary_(
  terms
) {

  const availableTerms =
    (terms || [])
      .filter(term =>

        term.recordedAverage !==
        null

      );


  if (
    availableTerms.length === 0
  ) {

    return {

      available:
        false,

      complete:
        false,

      recordedAverage:
        null,

      recordedAverageText:
        '',

      semesterCount:
        0,

      includedTerms:
        [],

      incompleteTerms:
        [],

      missingCategories:
        [],

      note:
        'No semester grade data is currently available.'

    };
  }


  const average =

    availableTerms
      .reduce(
        (sum, term) =>

          sum +
          term.recordedAverage,

        0
      ) /

    availableTerms.length;


  const incompleteTerms =
    availableTerms

      .filter(term =>
        !term.complete
      )

      .map(term =>
        term.clinicalTerm
      );


  const missingMap = {};


  availableTerms.forEach(term => {

    term
      .missingCategories
      .forEach(category => {

        missingMap[
          appNormalize_(
            category
          )
        ] =
          category;

      });

  });


  const missingCategories =
    Object.values(
      missingMap
    );


  const complete =
    incompleteTerms.length === 0;


  let note;


  if (complete) {

    note =
      'Average of all complete semester clinical grades currently recorded.';

  } else {

    note =
      'Overall recorded average is based on the semester grade data currently available.';


    if (
      missingCategories.length > 0
    ) {

      note +=

        ' Some semester averages do not yet include ' +
        gradeEngine_joinList_(
          missingCategories
        ) +
        '.';
    }
  }


  return {

    available:
      true,

    complete:
      complete,

    recordedAverage:
      gradeEngine_round2_(
        average
      ),

    recordedAverageText:
      gradeEngine_formatPercent_(
        average
      ),

    semesterCount:
      availableTerms.length,

    includedTerms:
      availableTerms.map(term =>
        term.clinicalTerm
      ),

    incompleteTerms:
      incompleteTerms,

    missingCategories:
      missingCategories,

    note:
      note

  };
}


/**
 * =====================================================================
 * CLINICAL TRACKING CATEGORY
 * =====================================================================
 *
 * Used for Competency and ISO.
 *
 * Each individual Final Score is first converted:
 *
 *   2.74 -> 90.5%
 *
 * Then the resulting percentages are averaged.
 *
 * PASSES AND FAILURES ARE BOTH INCLUDED.
 */
function gradeEngine_calculateClinicalCategory_(
  records
) {

  const grades = [];


  (records || [])
    .forEach(record => {

      const resolved =
        gradeEngine_resolveTrackingGrade_(
          record
        );


      if (
        !resolved.available ||
        resolved.percent === null
      ) {

        return;
      }


      grades.push({

        rawScore:
          resolved.rawScore,

        percent:
          resolved.percent,

        source:
          resolved.source

      });

    });


  if (
    grades.length === 0
  ) {

    return {

      available:
        false,

      count:
        0,

      rawAverage:
        null,

      percentAverage:
        null,

      percentAverageText:
        ''

    };
  }


  const rawValues =
    grades

      .map(item =>
        item.rawScore
      )

      .filter(value =>
        value !== null &&
        value <= 3
      );


  let rawAverage =
    null;


  if (
    rawValues.length > 0
  ) {

    rawAverage =

      rawValues.reduce(
        (sum, value) =>
          sum + value,
        0
      ) /
      rawValues.length;


    rawAverage =
      gradeEngine_round2_(
        rawAverage
      );
  }


  const percentAverage =

    grades.reduce(
      (sum, item) =>
        sum + item.percent,
      0
    ) /
    grades.length;


  return {

    available:
      true,

    count:
      grades.length,

    rawAverage:
      rawAverage,

    percentAverage:
      gradeEngine_round2_(
        percentAverage
      ),

    percentAverageText:
      gradeEngine_formatPercent_(
        percentAverage
      )

  };
}


/**
 * =====================================================================
 * RESOLVE ONE CLINICAL TRACKING GRADE
 * =====================================================================
 *
 * Central record-level rule used by the Grade Engine and document
 * generators.
 *
 * Historical records:
 * - Historical Percent Grade is authoritative when present.
 * - A true Final Score remains available as raw evidence but does not
 *   override the reviewed historical awarded percentage.
 *
 * Current/live records:
 * - Final Score is converted through the official 0-3 scale.
 * - Historical Percent Grade is never used.
 *
 * Pass/Fail-only records:
 * - No numeric grade is invented.
 * =====================================================================
 */
function gradeEngine_resolveTrackingGrade_(
  record
) {

  record =
    record || {};


  const rawScore =
    gradeEngine_numeric_(
      record['Final Score']
    );


  if (
    gradeEngine_isHistoricalTrackingRecord_(
      record
    )
  ) {

    const historicalPercent =
      gradeEngine_numeric_(
        record['Historical Percent Grade']
      );


    if (
      historicalPercent !== null &&
      historicalPercent >= 0 &&
      historicalPercent <= 100
    ) {

      const normalizedHistoricalPercent =
        historicalPercent >= 0 &&
        historicalPercent <= 1
          ? historicalPercent * 100
          : historicalPercent;


      return {

        available:
          true,

        rawScore:
          rawScore,

        percent:
          gradeEngine_round2_(
            normalizedHistoricalPercent
          ),

        source:
          'Historical Percent Grade'

      };
    }
  }


  if (
    rawScore !== null
  ) {

    const percent =
      gradeEngine_scoreToPercent(
        rawScore
      );


    if (
      percent !== null
    ) {

      return {

        available:
          true,

        rawScore:
          rawScore,

        percent:
          gradeEngine_round2_(
            percent
          ),

        source:
          'Final Score'

      };
    }
  }


  return {

    available:
      false,

    rawScore:
      rawScore,

    percent:
      null,

    source:
      ''

  };
}


/**
 * =====================================================================
 * SUPPLEMENTAL GRADE CATEGORY
 * =====================================================================
 *
 * For Professional Competency:
 *
 * raw score may be 0-3 and is converted using the official table.
 *
 * For tests/quizzes/case presentations:
 *
 * Percent Grade should normally be entered directly.
 */
function gradeEngine_calculateSupplementalCategory_(
  records,
  allowClinicalScale
) {

  const percentages = [];


  (records || [])
    .forEach(record => {

      let percent =
        gradeEngine_numeric_(
          record['Percent Grade']
        );


      if (
        percent !== null
      ) {

        /*
         * Spreadsheet-style percent:
         *
         * .905 -> 90.5
         */
        if (
          percent >= 0 &&
          percent <= 1
        ) {

          percent *=
            100;
        }


        percentages.push(
          percent
        );


        return;
      }


      if (
        !allowClinicalScale
      ) {

        return;
      }


      const raw =
        gradeEngine_numeric_(
          record['Raw Score']
        );


      if (
        raw === null
      ) {

        return;
      }


      percent =
        gradeEngine_scoreToPercent(
          raw
        );


      if (
        percent !== null
      ) {

        percentages.push(
          percent
        );
      }

    });


  if (
    percentages.length === 0
  ) {

    return {

      available:
        false,

      count:
        0,

      percentAverage:
        null,

      percentAverageText:
        ''

    };
  }


  const average =

    percentages.reduce(
      (sum, value) =>
        sum + value,
      0
    ) /
    percentages.length;


  return {

    available:
      true,

    count:
      percentages.length,

    percentAverage:
      gradeEngine_round2_(
        average
      ),

    percentAverageText:
      gradeEngine_formatPercent_(
        average
      )

  };
}


/**
 * =====================================================================
 * POINT DEDUCTIONS
 * =====================================================================
 *
 * Future Grade Records may contain:
 *
 * Grade Category = Point Deduction
 *
 * Percent Grade / Points / Raw Score may contain:
 *
 * -2
 * -5
 * -10
 *
 * A positive entered value is interpreted as points OFF and converted
 * to negative.
 */
function gradeEngine_calculateDeductions_(
  records
) {

  const items = [];


  let total =
    0;


  (records || [])
    .forEach(record => {

      let value =
        gradeEngine_numeric_(
          record['Points']
        );


      if (
        value === null
      ) {

        value =
          gradeEngine_numeric_(
            record['Percent Grade']
          );
      }


      if (
        value === null
      ) {

        value =
          gradeEngine_numeric_(
            record['Raw Score']
          );
      }


      if (
        value === null ||
        value === 0
      ) {

        return;
      }


      /*
       * User enters:
       *
       * 2 = two points OFF
       *
       * or:
       *
       * -2
       *
       * Both result in -2.
       */
      const deduction =
        value > 0
          ? -value
          : value;


      total +=
        deduction;


      items.push({

        name:
          String(
            record[
              'Assignment / Evaluation Name'
            ] ||
            record[
              'Assignment'
            ] ||
            'Point Deduction'
          ).trim(),

        points:
          deduction,

        notes:
          String(
            record['Notes'] || ''
          ).trim()

      });

    });


  return {

    total:
      gradeEngine_round2_(
        total
      ),

    items:
      items

  };
}


/**
 * =====================================================================
 * OFFICIAL 0-3 -> PERCENT CONVERSION
 * =====================================================================
 *
 * Reproduces Sheet 2 from the Professional Competency workbook.
 */
function gradeEngine_scoreToPercent(
  rawScore
) {

  const numeric =
    gradeEngine_numeric_(
      rawScore
    );


  if (
    numeric === null
  ) {

    return null;
  }


  /*
   * Historical imported value already expressed as percent.
   */
  if (
    numeric > 3
  ) {

    return gradeEngine_round2_(
      numeric
    );
  }


  const bounded =
    Math.max(
      0,
      Math.min(
        3,
        numeric
      )
    );


  const hundredths =
    Math.round(
      bounded *
      100
    );


  // ===================================================================
  // TOP OF SCALE
  // ===================================================================

  if (hundredths === 300) return 100;
  if (hundredths === 299) return 99;
  if (hundredths === 298) return 98;
  if (hundredths === 297) return 97;
  if (hundredths === 296) return 96;


  // ===================================================================
  // 2.95 THROUGH 2.17
  // ===================================================================

  if (
    hundredths >= 217 &&
    hundredths <= 295
  ) {

    return gradeEngine_round2_(

      95.75 -

      (
        295 -
        hundredths
      ) *
      0.25

    );
  }


  if (hundredths >= 210) return 76.25;
  if (hundredths >= 201) return 76;
  if (hundredths >= 199) return 75;
  if (hundredths >= 192) return 74;
  if (hundredths >= 189) return 73;
  if (hundredths >= 185) return 72;
  if (hundredths >= 181) return 71;
  if (hundredths >= 175) return 70;
  if (hundredths >= 170) return 69;
  if (hundredths >= 166) return 68;
  if (hundredths === 165) return 67;
  if (hundredths >= 161) return 66;
  if (hundredths >= 157) return 65;
  if (hundredths >= 153) return 64;
  if (hundredths >= 149) return 63;
  if (hundredths >= 145) return 62;
  if (hundredths >= 141) return 61;
  if (hundredths >= 137) return 60;
  if (hundredths >= 133) return 59;
  if (hundredths >= 129) return 58;
  if (hundredths >= 125) return 57;
  if (hundredths >= 121) return 56;
  if (hundredths >= 117) return 55;
  if (hundredths >= 113) return 54;
  if (hundredths >= 109) return 53;
  if (hundredths >= 105) return 52;
  if (hundredths >= 101) return 51;
  if (hundredths >= 97) return 50;
  if (hundredths >= 93) return 49;
  if (hundredths >= 89) return 48;
  if (hundredths >= 85) return 47;
  if (hundredths >= 81) return 46;
  if (hundredths >= 77) return 45;
  if (hundredths >= 73) return 44;
  if (hundredths >= 69) return 43;
  if (hundredths >= 65) return 42;

  /*
   * Professional Comp conversion table intentionally skips 41%.
   */
  if (hundredths >= 57) return 40;

  if (hundredths >= 53) return 39;
  if (hundredths >= 49) return 38;
  if (hundredths >= 45) return 37;
  if (hundredths >= 41) return 36;
  if (hundredths >= 37) return 35;
  if (hundredths >= 33) return 34;
  if (hundredths >= 29) return 33;
  if (hundredths >= 25) return 32;
  if (hundredths >= 21) return 31;
  if (hundredths >= 17) return 30;
  if (hundredths >= 13) return 29;
  if (hundredths >= 9) return 28;
  if (hundredths >= 5) return 27;
  if (hundredths >= 1) return 26;

  return 25;
}


/**
 * =====================================================================
 * INDIVIDUAL GRADE DISPLAY
 * =====================================================================
 *
 * Example:
 *
 * {
 *   rawText: "2.00",
 *   percentText: "75%",
 *   combinedText: "2.00 • 75%"
 * }
 */
function gradeEngine_formatEvaluationGrade(
  rawScore
) {

  const raw =
    gradeEngine_numeric_(
      rawScore
    );


  if (
    raw === null
  ) {

    return {

      available:
        false,

      raw:
        null,

      rawText:
        '',

      percent:
        null,

      percentText:
        '',

      combinedText:
        ''

    };
  }


  const percent =
    gradeEngine_scoreToPercent(
      raw
    );


  const rawText =
    raw <= 3
      ? raw.toFixed(2)
      : '';


  const percentText =
    gradeEngine_formatPercent_(
      percent
    );


  return {

    available:
      percent !== null,

    raw:
      raw,

    rawText:
      rawText,

    percent:
      percent,

    percentText:
      percentText,

    combinedText:
      rawText &&
      percentText

        ? rawText +
          ' • ' +
          percentText

        : (
            percentText ||
            rawText
          )

  };
}


/**
 * =====================================================================
 * OPTIONAL GRADE RECORDS
 * =====================================================================
 *
 * This creates the architectural landing place for future Add New.
 *
 * We are NOT creating the sheet yet.
 *
 * When it exists, expected columns are:
 *
 *   Student ID
 *   Student
 *   Clinical Term
 *   Grade Category
 *   Assignment / Evaluation Name
 *   Raw Score
 *   Percent Grade
 *   Points
 *   Date
 *   Entered By
 *   Source
 *   Source ID
 *   Notes
 *   Active
 *
 * One sheet can support ALL non-Clinical-Tracking grade components.
 */
function gradeEngine_getSupplementalGradeRecords_(
  ss
) {

  let configuredName =
    '';


  if (
    APP_CONFIG &&
    APP_CONFIG.SHEETS &&
    APP_CONFIG.SHEETS.GRADE_RECORDS
  ) {

    configuredName =
      APP_CONFIG.SHEETS
        .GRADE_RECORDS;
  }


  /*
   * Transitional future-ready fallback.
   *
   * This does not require the sheet to exist.
   */
  const sheetName =
    configuredName ||
    'Grade Records';


  const sheet =
    ss.getSheetByName(
      sheetName
    );


  if (!sheet) {

    return [];
  }


  return appSheetToObjects_(
    sheet
  )
    .filter(record => {

      if (
        Object.prototype
          .hasOwnProperty
          .call(
            record,
            'Active'
          )
      ) {

        return appIsTrue_(
          record['Active']
        );
      }


      return true;

    });
}


/**
 * =====================================================================
 * FILTER SUPPLEMENTAL CATEGORY
 * =====================================================================
 */
function gradeEngine_filterGradeRecordsByCategory_(
  records,
  categoryKey
) {

  return (records || [])
    .filter(record =>

      gradeEngine_normalizeGradeCategory_(
        record['Grade Category']
      ) ===
      categoryKey

    );
}


/**
 * =====================================================================
 * NORMALIZE GRADE CATEGORY
 * =====================================================================
 */
function gradeEngine_normalizeGradeCategory_(
  value
) {

  const normalized =
    appNormalize_(
      value
    );


  if (
    normalized.includes(
      'professional'
    )
  ) {

    return 'professional';
  }


  if (
    normalized.includes(
      'clinical test'
    ) ||
    normalized ===
      'test'
  ) {

    return 'clinicalTest';
  }


  if (
    normalized.includes(
      'case presentation'
    )
  ) {

    return 'casePresentation';
  }


  if (
    normalized.includes(
      'quiz'
    )
  ) {

    return 'quizzes';
  }


  if (
    normalized.includes(
      'deduction'
    ) ||
    normalized.includes(
      'points off'
    )
  ) {

    return 'deduction';
  }


  return '';
}


/**
 * =====================================================================
 * COURSE WEIGHTS
 * =====================================================================
 */
function gradeEngine_getCourseWeights_(
  clinicalTerm
) {

  const key =
    appNormalize_(
      clinicalTerm
    );


  const weights =
    GRADE_ENGINE_COURSE_WEIGHTS[
      key
    ];


  if (!weights) {

    return {

      competency:
        0,

      iso:
        0,

      professional:
        0,

      clinicalTest:
        0,

      casePresentation:
        0,

      quizzes:
        0

    };
  }


  return Object.assign(
    {},
    weights
  );
}


/**
 * =====================================================================
 * CATEGORY LABEL
 * =====================================================================
 */
function gradeEngine_categoryLabel_(
  key
) {

  const labels = {

    competency:
      'Competency Evaluations',

    iso:
      'Instructor Selected Ongoings',

    professional:
      'Professional Competency',

    clinicalTest:
      'Clinical Test',

    casePresentation:
      'Case Presentation',

    quizzes:
      'Clinical Quizzes'

  };


  return (
    labels[key] ||
    key
  );
}


/**
 * =====================================================================
 * EFFECTIVE TERM
 * =====================================================================
 */
function gradeEngine_effectiveClinicalTerm_(
  record
) {

  /*
   * Reviewed historical imports own their stored Clinical Term. Their source
   * dates may fall inside a modern Semester Rules window and must never cause
   * retroactive reassignment through Attributed Clinical Term.
   */
  if (
    gradeEngine_isHistoricalTrackingRecord_(
      record
    )
  ) {

    return String(
      record['Clinical Term'] || ''
    ).trim();
  }

  const attributed =
    String(
      record[
        'Attributed Clinical Term'
      ] || ''
    ).trim();


  if (
    attributed &&
    appNormalize_(
      attributed
    ) !==
    'unattributed'
  ) {

    return attributed;
  }


  return String(
    record['Clinical Term'] || ''
  ).trim();
}


/**
 * =====================================================================
 * SAME STUDENT
 * =====================================================================
 */
function gradeEngine_sameStudent_(
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
      record['Student ID']
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
      record['Student']
    ) ===

    appNormalize_(
      studentName
    )

  );
}


/**
 * =====================================================================
 * MOST ADVANCED STUDENT ROW
 * =====================================================================
 */
function gradeEngine_getMostAdvancedStudentRow_(
  rows
) {

  const active =
    (rows || [])
      .filter(row =>

        appIsTrue_(
          row['Active']
        )

      );


  const candidates =
    active.length
      ? active
      : rows;


  return candidates
    .slice()
    .sort(
      (a, b) =>

        gradeEngine_clinicalNumber_(
          b['Clinical Term']
        ) -

        gradeEngine_clinicalNumber_(
          a['Clinical Term']
        )
    )[0];
}


/**
 * =====================================================================
 * CLINICAL NUMBER
 * =====================================================================
 */
function gradeEngine_clinicalNumber_(
  value
) {

  const match =
    String(
      value || ''
    ).match(
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
 * NUMBER
 * =====================================================================
 */
function gradeEngine_numeric_(
  value
) {

  if (
    value === null ||
    value === undefined ||
    String(value).trim() === ''
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


/**
 * =====================================================================
 * FORMAT PERCENT
 * =====================================================================
 */
function gradeEngine_formatPercent_(
  value
) {

  if (
    value === null ||
    value === undefined ||
    !Number.isFinite(
      Number(value)
    )
  ) {

    return '';
  }


  const number =
    gradeEngine_round2_(
      Number(value)
    );


  if (
    Number.isInteger(
      number
    )
  ) {

    return (
      number.toFixed(0) +
      '%'
    );
  }


  return (
    Number(
      number.toFixed(2)
    ) +
    '%'
  );
}


/**
 * =====================================================================
 * JOIN HUMAN LIST
 * =====================================================================
 */
function gradeEngine_joinList_(
  items
) {

  const values =
    (items || [])
      .filter(Boolean);


  if (
    values.length === 0
  ) {

    return '';
  }


  if (
    values.length === 1
  ) {

    return values[0];
  }


  if (
    values.length === 2
  ) {

    return (
      values[0] +
      ' and ' +
      values[1]
    );
  }


  return (

    values
      .slice(
        0,
        values.length - 1
      )
      .join(', ') +

    ', and ' +

    values[
      values.length - 1
    ]

  );
}


/**
 * =====================================================================
 * ROUND TWO
 * =====================================================================
 */
function gradeEngine_round2_(
  value
) {

  return Math.round(
    (
      Number(value) +
      Number.EPSILON
    ) *
    100
  ) / 100;
}


/**
 * True only for records explicitly identifiable as historical imports.
 */
function gradeEngine_isHistoricalTrackingRecord_(
  record
) {

  const submissionId =
    appNormalize_(
      record['Submission ID']
    );


  const submissionMethod =
    appNormalize_(
      record['Submission Method'] ||
      record['Source'] ||
      record['Source Type']
    );


  return (
    submissionId.indexOf('hist-') === 0 ||
    submissionMethod.indexOf('historical import') >= 0
  );
}
