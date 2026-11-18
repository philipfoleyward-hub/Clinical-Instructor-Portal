/**
 * Portal-side authoritative scoring rules support — READ ONLY.
 * Extracted without formula changes from the supplied original CCMS HEAD:
 * 04_ClinicalEvaluationGradingEngine.js (helper methods enumerated below).
 * Prefixing prevents collisions with other Portal services.
 * Reads Scoring Weights + Image Scoring Rules from the same CCMS workbook
 * configured for Portal's existing ciDataBook_() service.
 * No form submission, spreadsheet write, Drive write or email.
 */
const CI_PORTAL_GRADE_CONFIG_ = Object.freeze({
  SHEETS: Object.freeze({ SCORING_WEIGHTS:'Scoring Weights', IMAGE_SCORING_RULES:'Image Scoring Rules' }),
  TERMINOLOGY: Object.freeze({ RESULTS: Object.freeze({PASS:'Pass', FAIL:'Fail'}) })
});
function ciPortalAcademicGetSheet_(ss,name) {
  const sheet=ss.getSheetByName(name);
  if(!sheet) throw new Error('Authoritative CCMS grading sheet unavailable: '+name);
  return sheet;
}
function ciPortalAcademicNormalize_(value) {
  // Exact original CCMS 01_SharedUtils.gs appNormalize_ normalization.
  return String(value===null||value===undefined?'':value).trim().toLowerCase().replace(/\s+/g,' ');
}

// Unmodified grading logic from original CCMS: grading_loadPerformanceWeights_
function ciPortalAcademic_loadPerformanceWeights_(
  ss
) {

  const sheet =
    ciPortalAcademicGetSheet_(
      ss,
      CI_PORTAL_GRADE_CONFIG_.SHEETS
        .SCORING_WEIGHTS
    );

  const values =
    sheet
      .getDataRange()
      .getValues();

  if (
    values.length < 2
  ) {
    throw new Error(
      'Scoring Weights contains no scoring rows.'
    );
  }

  const headers =
    values[0].map(
      ciPortalAcademicNormalize_
    );

  const sectionIndex =
    ciPortalAcademic_findHeader_(
      headers,
      [
        'section',
        'grading group',
        'performance section',
        'category'
      ]
    );

  const criterionIndex =
    ciPortalAcademic_findHeader_(
      headers,
      [
        'criterion',
        'criteria',
        'performance criterion',
        'question'
      ]
    );

  const weightIndex =
    ciPortalAcademic_findHeader_(
      headers,
      [
        'weight',
        'scoring weight'
      ]
    );

  if (
    sectionIndex === -1 ||
    criterionIndex === -1 ||
    weightIndex === -1
  ) {
    throw new Error(
      'Could not identify Section, Criterion and Weight columns in Scoring Weights.'
    );
  }

  const rows = [];

  values
    .slice(1)
    .forEach(row => {

      const criterion =
        ciPortalAcademic_clean_(
          row[
            criterionIndex
          ]
        );

      const weight =
        Number(
          row[
            weightIndex
          ]
        );

      if (
        !criterion ||
        !Number.isFinite(
          weight
        )
      ) {
        return;
      }

      rows.push({
        section:
          ciPortalAcademic_clean_(
            row[
              sectionIndex
            ]
          ),

        criterion:
          criterion,

        criterionKey:
          ciPortalAcademic_criterionKey_(
            criterion
          ),

        weight:
          weight
      });
    });

  if (
    rows.length === 0
  ) {
    throw new Error(
      'No usable rows were loaded from Scoring Weights.'
    );
  }

  return rows;
}

// Unmodified grading logic from original CCMS: grading_findWeight_
function ciPortalAcademic_findWeight_(
  weights,
  formSection,
  criterionKey
) {

  const sectionKey =
    ciPortalAcademic_normalizeSection_(
      formSection
    );

  let match =
    weights.find(row =>

      ciPortalAcademic_normalizeSection_(
        row.section
      ) === sectionKey &&

      row.criterionKey ===
        criterionKey
    );

  if (match) {
    return match;
  }

  match =
    weights.find(row =>
      row.criterionKey ===
        criterionKey
    );

  return match || null;
}

// Unmodified grading logic from original CCMS: grading_parseImageScoreAnswer_
function ciPortalAcademic_parseImageScoreAnswer_(
  answer
) {

  const text =
    ciPortalAcademic_clean_(
      answer
    );

  if (!text) {
    return {
      applicable:
        false,

      score:
        null,

      explanation:
        'No response'
    };
  }

  const normalized =
    ciPortalAcademicNormalize_(
      text
    );

  if (
    normalized === 'n/a' ||
    normalized.startsWith(
      'n/a '
    )
  ) {
    return {
      applicable:
        false,

      score:
        null,

      explanation:
        text
    };
  }

  const match =
    text.match(
      /^\s*([0-3])\s*[=:;-]?\s*(.*)$/
    );

  if (!match) {
    return {
      applicable:
        false,

      score:
        null,

      explanation:
        text
    };
  }

  const score =
    Number(
      match[1]
    );

  const explanation =
    ciPortalAcademic_clean_(
      match[2]
    ) ||
    (
      'Score ' +
      score
    );

  return {
    applicable:
      true,

    score:
      score,

    explanation:
      explanation
  };
}

// Unmodified grading logic from original CCMS: grading_loadImagePenalties_
function ciPortalAcademic_loadImagePenalties_(
  ss
) {

  const sheet =
    ciPortalAcademicGetSheet_(
      ss,
      CI_PORTAL_GRADE_CONFIG_.SHEETS
        .IMAGE_SCORING_RULES
    );

  const values =
    sheet
      .getDataRange()
      .getValues();

  let zeroPenalty =
    null;

  let missingMarkerPenalty =
    null;

  let wrongMarkerPenalty =
    null;

  let illegibleMarkerPenalty =
    null;


  for (
    let i = 0;
    i < values.length;
    i++
  ) {

    /*
     * Historically the penalty table has lived in G:H.
     * Scan every adjacent pair so moving the table does not break grading.
     *
     * Rule labels are matched by meaning rather than one exact phrase.
     * Examples accepted:
     *
     *   Zero In View Penalty
     *   Zero Rule Penalty
     *   Zero Penalty
     *
     *   Missing Marker Penalty
     *   Wrong Marker Penalty
     *   Illegible Marker Penalty
     *   Illegible or Poorly Placed Marker Penalty
     */
    for (
      let j = 0;
      j < values[i].length - 1;
      j++
    ) {

      const rule =
        ciPortalAcademicNormalize_(
          values[i][j]
        );

      const value =
        Number(
          values[i][
            j + 1
          ]
        );

      if (
        !rule ||
        !Number.isFinite(
          value
        )
      ) {
        continue;
      }


      const isPenalty =
        rule.includes(
          'penalt'
        );


      if (
        isPenalty &&
        rule.includes(
          'zero'
        )
      ) {
        zeroPenalty =
          value;
        continue;
      }


      if (
        isPenalty &&
        rule.includes(
          'marker'
        ) &&
        rule.includes(
          'missing'
        ) &&
        !rule.includes(
          'wrong/missing'
        )
      ) {
        missingMarkerPenalty =
          value;
        continue;
      }


      if (
        isPenalty &&
        rule.includes(
          'marker'
        ) &&
        rule.includes(
          'wrong'
        ) &&
        !rule.includes(
          'wrong/missing'
        )
      ) {
        wrongMarkerPenalty =
          value;
        continue;
      }


      if (
        isPenalty &&
        rule.includes(
          'marker'
        ) &&
        (
          rule.includes(
            'illegible'
          ) ||
          rule.includes(
            'poorly placed'
          )
        )
      ) {
        illegibleMarkerPenalty =
          value;
      }
    }
  }


  const missingRules = [];

  if (
    !Number.isFinite(
      zeroPenalty
    )
  ) {
    missingRules.push(
      'Zero penalty'
    );
  }

  if (
    !Number.isFinite(
      missingMarkerPenalty
    )
  ) {
    missingRules.push(
      'Missing Marker penalty'
    );
  }

  if (
    !Number.isFinite(
      wrongMarkerPenalty
    )
  ) {
    missingRules.push(
      'Wrong Marker penalty'
    );
  }

  if (
    !Number.isFinite(
      illegibleMarkerPenalty
    )
  ) {
    missingRules.push(
      'Illegible/Poorly Placed Marker penalty'
    );
  }


  if (
    missingRules.length > 0
  ) {

    throw new Error(
      'Could not identify these rules in Image Scoring Rules: ' +
      missingRules.join(', ') +
      '. Confirm each rule label is immediately followed by its numeric value.'
    );
  }


  return {
    zeroPenalty:
      zeroPenalty,

    missingMarkerPenalty:
      missingMarkerPenalty,

    wrongMarkerPenalty:
      wrongMarkerPenalty,

    illegibleMarkerPenalty:
      illegibleMarkerPenalty
  };
}

// Unmodified grading logic from original CCMS: grading_findHeader_
function ciPortalAcademic_findHeader_(
  normalizedHeaders,
  candidates
) {

  for (
    let i = 0;
    i < candidates.length;
    i++
  ) {

    const index =
      normalizedHeaders.indexOf(
        ciPortalAcademicNormalize_(
          candidates[i]
        )
      );

    if (
      index !== -1
    ) {
      return index;
    }
  }

  return -1;
}

// Unmodified grading logic from original CCMS: grading_normalizeSection_
function ciPortalAcademic_normalizeSection_(
  value
) {

  const text =
    ciPortalAcademicNormalize_(
      value
    );

  if (
    text.includes(
      'portable'
    ) ||
    text.includes(
      'contrast'
    )
  ) {
    return (
      'portable / contrast'
    );
  }

  return text;
}

// Unmodified grading logic from original CCMS: grading_criterionKey_
function ciPortalAcademic_criterionKey_(
  value
) {

  return ciPortalAcademic_clean_(
    value
  )
    .toLowerCase()

    .replace(
      /^\s*\d+\s*[\.\)\-:]?\s*/,
      ''
    )

    .replace(
      /[’‘]/g,
      "'"
    )

    .replace(
      /[^a-z0-9]+/g,
      ' '
    )

    .replace(
      /\s+/g,
      ' '
    )

    .trim();
}

// Unmodified grading logic from original CCMS: grading_clean_
function ciPortalAcademic_clean_(
  value
) {

  if (
    value === null ||
    value === undefined
  ) {
    return '';
  }

  if (
    Array.isArray(value)
  ) {
    return value
      .map(
        ciPortalAcademic_clean_
      )
      .join(', ');
  }

  return String(value)
    .trim();
}

// Unmodified grading logic from original CCMS: grading_round2_
function ciPortalAcademic_round2_(
  value
) {

  return Math.round(
    (
      Number(value) +
      Number.EPSILON
    ) * 100
  ) / 100;
}


/** Official 0–3 scale conversion from original CCMS 04_GradeEngine.js.
 * No percentages are computed in the browser. The original mapping is retained
 * for consistency with Clinical Grades and historical documents.
 */
function ciPortalAcademic_numeric_(value){
  if (value===null||value===undefined||value==='')return null;
  const numeric=Number(value);
  return Number.isFinite(numeric)?numeric:null;
}
function ciPortalAcademic_scoreToPercent_(
  rawScore
) {

  const numeric =
    ciPortalAcademic_numeric_(
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

    return ciPortalAcademic_round2_(
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

    return ciPortalAcademic_round2_(

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
