/**
 * =====================================================================
 * CLINICAL COMPETENCY TRACKER
 * CLINICAL EVALUATION GRADING ENGINE
 * 04_ClinicalEvaluationGradingEngine.gs
 * =====================================================================
 *
 * PURPOSE
 * -------
 * Authoritative grading engine for the DIGITAL Clinical Evaluation Form.
 *
 * THIS FILE OWNS:
 *
 *   - Reading one submitted Digital Evaluation
 *   - Performance scoring
 *   - Image scoring
 *   - Image penalties
 *   - Automatic Failure detection
 *   - Final competency score
 *   - Pass / Fail determination
 *
 * THIS FILE DOES NOT:
 *
 *   - Modify the Google Form
 *   - Write to Clinical Tracking
 *   - Generate documents
 *   - Move files
 *   - Send email
 *   - Calculate semester/program requirements
 *
 * IMPORTANT
 * ---------
 * This engine GRADES ONE FormResponse.
 *
 * Production intake will eventually call:
 *
 *   gradeClinicalEvaluation(formResponse)
 *
 * =====================================================================
 */


/**
 * =====================================================================
 * MAIN PRODUCTION GRADING FUNCTION
 * =====================================================================
 */
function gradeClinicalEvaluation(formResponse) {

  if (!formResponse) {
    throw new Error(
      'A Clinical Evaluation FormResponse is required.'
    );
  }

  const ss =
    ccmsPortalWorkbook_();

  const form =
    FormApp.openById(
      APP_CONFIG.FORMS
        .DIGITAL_CLINICAL_EVALUATION
    );

  const responseByItemId =
    grading_buildResponseMap_(
      formResponse
    );

  const evaluation =
    grading_readEvaluationInfo_(
      form,
      responseByItemId
    );

  const automaticFailure =
    grading_readAutomaticFailure_(
      form,
      responseByItemId
    );

  const weights =
    grading_loadPerformanceWeights_(
      ss
    );

  const performance =
    grading_gradePerformance_(
      form,
      responseByItemId,
      weights
    );

  const penalties =
    grading_loadImagePenalties_(
      ss
    );

  const image =
    grading_gradeImages_(
      form,
      responseByItemId,
      penalties
    );

    let finalScore = null;

  /*
   * -------------------------------------------------------------
   * FINAL COMPETENCY SCORE
   * -------------------------------------------------------------
   *
   * Program Handbook §13.7 — Automatic Failure:
   *
   * A student who meets an Automatic Failure condition earns
   * a competency grade of 0.0.
   *
   * Performance and Image scores remain intact in the grade
   * object for documentation/audit purposes. Automatic Failure
   * overrides only the final competency grade.
   */

  if (automaticFailure.occurred) {

    finalScore = 0;

  } else if (
    performance.score !== null &&
    image.score !== null
  ) {

    finalScore =
      (
        performance.score +
        image.score
      ) / 2;
  }

  const performancePass =
    performance.score !== null &&
    performance.score >= 2.0;

  const imagePass =
    image.score !== null &&
    image.score >= 2.0;

  let result =
    APP_CONFIG.TERMINOLOGY
      .RESULTS.FAIL
      .toUpperCase();

  if (
    !automaticFailure.occurred &&
    performancePass &&
    imagePass
  ) {
    result =
      APP_CONFIG.TERMINOLOGY
        .RESULTS.PASS
        .toUpperCase();
  }

  return {
    responseId:
      formResponse.getId
        ? formResponse.getId()
        : '',

    submittedAt:
      formResponse.getTimestamp
        ? formResponse.getTimestamp()
        : '',

    evaluation:
      evaluation,

    automaticFailure:
      automaticFailure,

    performance:
      performance,

    image:
      image,

    finalScore:
      finalScore === null
        ? null
        : grading_round2_(
            finalScore
          ),

    result:
      result
  };
}


/**
 * =====================================================================
 * TEST — GRADE MOST RECENT RESPONSE
 * =====================================================================
 */
function testGradeLatestClinicalEvaluation() {

  const form =
    FormApp.openById(
      APP_CONFIG.FORMS
        .DIGITAL_CLINICAL_EVALUATION
    );

  const responses =
    form.getResponses();

  if (responses.length === 0) {
    throw new Error(
      'The Digital Clinical Evaluation Form has no responses.'
    );
  }

  const response =
    responses[
      responses.length - 1
    ];

  const grade =
    gradeClinicalEvaluation(
      response
    );

  console.log(
    JSON.stringify(
      grade,
      null,
      2
    )
  );

  return grade;
}


/**
 * =====================================================================
 * TEST — IMAGE SCORING SUMMARY FOR MOST RECENT RESPONSE
 * =====================================================================
 *
 * Logs only the image-scoring portion so Apps Script does not truncate
 * the output with the full Performance Evaluation detail.
 */
function testLatestClinicalEvaluationImageScoring() {

  const form =
    FormApp.openById(
      APP_CONFIG.FORMS
        .DIGITAL_CLINICAL_EVALUATION
    );

  const responses =
    form.getResponses();

  if (
    responses.length === 0
  ) {
    throw new Error(
      'The Digital Clinical Evaluation Form has no responses.'
    );
  }

  const response =
    responses[
      responses.length - 1
    ];

  const grade =
    gradeClinicalEvaluation(
      response
    );

  const image =
    grade.image || {};

  const summary = {
    student:
      grade.evaluation
        ? grade.evaluation.student
        : '',

    examination:
      grade.evaluation
        ? grade.evaluation.examination
        : '',

    responseId:
      grade.responseId,

    penalties: {
      zero:
        image.zeroPenalty,

      missingMarker:
        image.missingMarkerPenalty,

      wrongMarker:
        image.wrongMarkerPenalty,

      illegibleMarker:
        image.illegibleMarkerPenalty
    },

    grossPoints:
      image.grossPoints,

    netPoints:
      image.netPoints,

    sectionsScored:
      image.sectionsScored,

    rawImageScore:
      image.rawScore,

    storedImageScore:
      image.score,

    finalScore:
      grade.finalScore,

    result:
      grade.result,

    views:
      (image.views || []).map(
        view => ({
          viewNumber:
            view.viewNumber,

          projection:
            view.projection,

          grossPoints:
            view.grossPoints,

          zeroPenalty:
            view.zeroPenalty,

          missingMarkerPenalty:
            view.missingMarkerPenalty,

          wrongMarkerPenalty:
            view.wrongMarkerPenalty,

          illegibleMarkerPenalty:
            view.illegibleMarkerPenalty,

          netPoints:
            view.netPoints,

          sectionsScored:
            view.sectionsScored,

          rawViewScore:
            view.imageScore,

          marker:
            view.marker
        })
      )
  };

  console.log(
    JSON.stringify(
      summary,
      null,
      2
    )
  );

  return summary;
}


/**
 * =====================================================================
 * RESPONSE MAP
 * =====================================================================
 */
function grading_buildResponseMap_(
  formResponse
) {

  const map = {};

  formResponse
    .getItemResponses()
    .forEach(itemResponse => {

      const item =
        itemResponse.getItem();

      map[
        String(
          item.getId()
        )
      ] =
        itemResponse.getResponse();
    });

  return map;
}


/**
 * =====================================================================
 * BASIC EVALUATION INFORMATION
 * =====================================================================
 */
function grading_readEvaluationInfo_(
  form,
  responseByItemId
) {

  const result = {
    student: '',
    clinicalSite: '',
    evaluator: '',
    examination: '',
    evaluationType: '',
    date: '',
    comments: ''
  };

  form
    .getItems()
    .forEach(item => {

      const title =
        appNormalize_(
          item.getTitle()
        );

      const answer =
        responseByItemId[
          String(
            item.getId()
          )
        ];

      if (
        answer === undefined ||
        answer === null ||
        answer === ''
      ) {
        return;
      }

      if (
        title === 'student'
      ) {
        result.student =
          grading_clean_(
            answer
          );
        return;
      }

      if (
        title ===
        'clinical site'
      ) {
        result.clinicalSite =
          grading_clean_(
            answer
          );
        return;
      }

      if (
        title ===
        'evaluator'
      ) {
        result.evaluator =
          grading_clean_(
            answer
          );
        return;
      }

      if (
        title === 'examination' ||
        (
          title.startsWith(
            'examination'
          ) &&
          title !==
            'examination category'
        )
      ) {
        result.examination =
          grading_clean_(
            answer
          );
        return;
      }

      if (
        title ===
        'evaluation type'
      ) {
        result.evaluationType =
          grading_clean_(
            answer
          );
        return;
      }

      if (
        title === 'date' ||
        title ===
          'evaluation date'
      ) {
        result.date =
          answer;
        return;
      }

      if (
        title ===
          'evaluator comments'
      ) {
        result.comments =
          grading_clean_(
            answer
          );
      }
    });

  return result;
}


/**
 * =====================================================================
 * AUTOMATIC FAILURE
 * =====================================================================
 */
function grading_readAutomaticFailure_(
  form,
  responseByItemId
) {

  const result = {
    occurred: false,
    reasons: [],
    comments: ''
  };

  form
    .getItems()
    .forEach(item => {

      const title =
        appNormalize_(
          item.getTitle()
        );

      const answer =
        responseByItemId[
          String(
            item.getId()
          )
        ];

      if (
        title ===
         'did an automatic failure occur?'||
        title ===
          'did an automatic-failure condition occur during this evaluation?'
      ) {
       if (
          appNormalize_(answer) === 'yes'
        ) {
          result.occurred = true;
        }
      }

      if (
        title ===
        'automatic failure reason(s)'
      ) {

        if (
          Array.isArray(answer)
        ) {
          result.reasons =
            answer.map(
              grading_clean_
            );
        } else if (
          grading_clean_(
            answer
          )
        ) {
          result.reasons = [
            grading_clean_(
              answer
            )
          ];
        }
      }

      if (
        title ===
        'automatic failure comments'
      ) {
        result.comments =
          grading_clean_(
            answer
          );
      }
    });

  return result;
}


/**
 * =====================================================================
 * LOAD PERFORMANCE WEIGHTS
 * =====================================================================
 */
function grading_loadPerformanceWeights_(
  ss
) {

  const sheet =
    appGetSheet_(
      ss,
      APP_CONFIG.SHEETS
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
      appNormalize_
    );

  const sectionIndex =
    grading_findHeader_(
      headers,
      [
        'section',
        'grading group',
        'performance section',
        'category'
      ]
    );

  const criterionIndex =
    grading_findHeader_(
      headers,
      [
        'criterion',
        'criteria',
        'performance criterion',
        'question'
      ]
    );

  const weightIndex =
    grading_findHeader_(
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
        grading_clean_(
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
          grading_clean_(
            row[
              sectionIndex
            ]
          ),

        criterion:
          criterion,

        criterionKey:
          grading_criterionKey_(
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


/**
 * =====================================================================
 * PERFORMANCE GRADING
 * =====================================================================
 */
function grading_gradePerformance_(
  form,
  responseByItemId,
  weights
) {

  const performanceGridTitles = [
    'exam preparation',
    'patient care',
    'radiation protection',
    'equipment manipulation',
    'positioning skills',
    'portable / contrast'
  ];

  const result = {
    sections: [],
    totalPoints: 0,
    totalWeight: 0,
    score: null,
    portableContrastApplicable: grading_readPortableContrastApplicability_(form, responseByItemId)
  };

  form
    .getItems(
      FormApp.ItemType.GRID
    )
    .forEach(rawItem => {

      const grid =
        rawItem.asGridItem();

      const gridTitle =
        appNormalize_(
          grid.getTitle()
        );

      if (
        !performanceGridTitles
          .includes(
            gridTitle
          )
      ) {
        return;
      }

      const answer =
        responseByItemId[
          String(
            grid.getId()
          )
        ];

      const criteria =
        grid.getRows();

      const section = {
        name:
          grid.getTitle(),

        points:
          0,

        weight:
          0,

        score:
          null,

        rows:
          [],

        comments:
          grading_readSectionComment_(
            form,
            grid.getId(),
            responseByItemId
          )
      };

      const ratings =
        Array.isArray(answer)
          ? answer
          : [];

      criteria.forEach(
        (criterion, index) => {

          const rawRating =
            grading_clean_(
              ratings[index]
            );

          const criterionKey =
            grading_criterionKey_(
              criterion
            );

          const weightRow =
            grading_findWeight_(
              weights,
              grid.getTitle(),
              criterionKey
            );

          if (!weightRow) {
            throw new Error(
              'Could not find scoring weight for:\n\n' +
              grid.getTitle() +
              '\n' +
              criterion
            );
          }

          const isNA =
            !rawRating ||
            appNormalize_(
              rawRating
            ) === 'n/a';

          if (isNA) {

            section.rows.push({
              criterion:
                criterion,

              ratingText:
                rawRating ||
                'N/A',

              applicable:
                false,

              weight:
                weightRow.weight,

              points:
                0
            });

            return;
          }

          const rating =
            Number(
              rawRating
            );

          if (
            !Number.isFinite(
              rating
            )
          ) {
            throw new Error(
              'Invalid rating "' +
              rawRating +
              '" for:\n' +
              criterion
            );
          }

          const points =
            rating *
            weightRow.weight;

          section.points +=
            points;

          section.weight +=
            weightRow.weight;

          section.rows.push({
            criterion:
              criterion,

            ratingText:
              rawRating,

            applicable:
              true,

            weight:
              weightRow.weight,

            points:
              points
          });
        }
      );

      if (
        section.weight > 0
      ) {
        section.score =
          section.points /
          section.weight;
      }

      result.sections.push(
        section
      );

      result.totalPoints +=
        section.points;

      result.totalWeight +=
        section.weight;
    });

  if (
    result.totalWeight > 0
  ) {
    result.score =
      result.totalPoints /
      result.totalWeight;
  }

  return result;
}


/**
 * Reads a free-text Comments/Notes item that belongs to the same Form
 * section as a performance grid. It stops at the next PageBreak and
 * deliberately ignores overall evaluator/automatic-failure/image fields.
 */
function grading_readSectionComment_(
  form,
  gridItemId,
  responseByItemId
) {
  const items = form.getItems();
  let gridIndex = -1;
  let gridTitle = '';

  for (let i = 0; i < items.length; i++) {
    if (String(items[i].getId()) === String(gridItemId)) {
      gridIndex = i;
      gridTitle = appNormalize_(items[i].getTitle());
      break;
    }
  }

  if (gridIndex === -1) return '';

  // The finalized Form uses explicit section-owned titles such as
  // "Patient Care Comments".  Require the title to agree with the
  // owning performance grid; do not assign a generic Comments item
  // from another section by title alone.
  const expectedTitle = gridTitle + ' comments';
  let legacyGeneric = '';

  for (let i = gridIndex + 1; i < items.length; i++) {
    const item = items[i];
    if (item.getType() === FormApp.ItemType.PAGE_BREAK) break;

    const title = appNormalize_(item.getTitle());
    const answer = responseByItemId[String(item.getId())];
    if (answer === undefined || answer === null || answer === '') continue;

    if (title === expectedTitle) {
      return grading_clean_(answer);
    }

    // Backward-compatible fallback for responses created before the
    // Form comment questions were explicitly renamed.  Section location
    // is still required, so a generic title cannot cross section bounds.
    if ((title === 'comments' || title === 'comment') && !legacyGeneric) {
      legacyGeneric = grading_clean_(answer);
    }
  }

  return legacyGeneric;
}


/**
 * Reads the finalized Form's explicit Portable/Contrast applicability gate.
 * true  = include the Portable / Contrast section in the generated record
 * false = omit it from the generated record
 * null  = gate response unavailable; downstream code must use a safe fallback
 */
function grading_readPortableContrastApplicability_(form, responseByItemId) {
  const targetTitle = 'was portable/contrast applicable during this examination?';

  const items = form.getItems();
  for (let i = 0; i < items.length; i++) {
    const item = items[i];
    if (appNormalize_(item.getTitle()) !== targetTitle) continue;

    const answer = grading_clean_(responseByItemId[String(item.getId())]);
    const normalized = appNormalize_(answer);
    if (normalized === 'yes') return true;
    if (normalized === 'no') return false;
    return null;
  }

  return null;
}

/**
 * =====================================================================
 * FIND PERFORMANCE WEIGHT
 * =====================================================================
 */
function grading_findWeight_(
  weights,
  formSection,
  criterionKey
) {

  const sectionKey =
    grading_normalizeSection_(
      formSection
    );

  let match =
    weights.find(row =>

      grading_normalizeSection_(
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


/**
 * =====================================================================
 * IMAGE GRADING
 * =====================================================================
 *
 * Raw document/audit math may be negative.
 * The grade-bearing Image Score returned as image.score is floored at 0.
 * =====================================================================
 */
function grading_gradeImages_(
  form,
  responseByItemId,
  penalties
) {

  const items =
    form.getItems();

  const views = [];

  for (
    let i = 0;
    i < items.length;
    i++
  ) {

    if (
      items[i].getType() !==
      FormApp.ItemType.PAGE_BREAK
    ) {
      continue;
    }

    const sectionTitle =
      grading_clean_(
        items[i].getTitle()
      );

    const match =
      sectionTitle.match(
        /^Image Evaluation\s*-\s*View\s+(\d+)$/i
      );

    if (!match) {
      continue;
    }

    const viewNumber =
      Number(
        match[1]
      );

    let sectionEnd =
      items.length;

    for (
      let j = i + 1;
      j < items.length;
      j++
    ) {

      if (
        items[j].getType() ===
        FormApp.ItemType.PAGE_BREAK
      ) {
        sectionEnd =
          j;
        break;
      }
    }

    const view =
      grading_gradeOneImageView_(
        items,
        i + 1,
        sectionEnd,
        viewNumber,
        responseByItemId,
        penalties
      );

    if (
      view.used
    ) {
      views.push(
        view
      );
    }
  }

  let grossPoints = 0;
  let zeroPenalty = 0;
  let missingMarkerPenalty = 0;
  let wrongMarkerPenalty = 0;
  let illegibleMarkerPenalty = 0;
  let netPoints = 0;
  let sectionsScored = 0;

  views.forEach(view => {

    grossPoints +=
      view.grossPoints;

    zeroPenalty +=
      view.zeroPenalty;

    missingMarkerPenalty +=
      view.missingMarkerPenalty;

    wrongMarkerPenalty +=
      view.wrongMarkerPenalty;

    illegibleMarkerPenalty +=
      view.illegibleMarkerPenalty;

    netPoints +=
      view.netPoints;

    sectionsScored +=
      view.sectionsScored;
  });

  const rawScore =
    sectionsScored > 0
      ? netPoints /
        sectionsScored
      : null;

  const score =
    rawScore === null
      ? null
      : Math.max(
          0,
          rawScore
        );

  return {
    views:
      views,

    grossPoints:
      grossPoints,

    zeroPenalty:
      zeroPenalty,

    missingMarkerPenalty:
      missingMarkerPenalty,

    wrongMarkerPenalty:
      wrongMarkerPenalty,

    illegibleMarkerPenalty:
      illegibleMarkerPenalty,

    netPoints:
      netPoints,

    sectionsScored:
      sectionsScored,

    rawScore:
      rawScore,

    score:
      score
  };
}


/**
 * =====================================================================
 * GRADE ONE IMAGE VIEW
 * =====================================================================
 */
function grading_gradeOneImageView_(
  items,
  startIndex,
  endIndex,
  viewNumber,
  responseByItemId,
  penalties
) {

  const result = {
    used:
      false,

    viewNumber:
      viewNumber,

    projection:
      '',

    components:
      [],

    exposureIndex:
      '',

    marker:
      '',

    comments:
      '',

    grossPoints:
      0,

    zeroPenalty:
      0,

    missingMarkerPenalty:
      0,

    wrongMarkerPenalty:
      0,

    illegibleMarkerPenalty:
      0,

    netPoints:
      0,

    sectionsScored:
      0,

    imageScore:
      null
  };

  const componentNames = {
    'positioning':
      'Positioning',

    'equipment':
      'Equipment / Central Ray',

    'radiation protection':
      'Radiation Protection / Artifacts',

    'exposure':
      'Exposure'
  };

  for (
    let i = startIndex;
    i < endIndex;
    i++
  ) {

    const item =
      items[i];

    const title =
      appNormalize_(
        item.getTitle()
      );

    const answer =
      responseByItemId[
        String(
          item.getId()
        )
      ];

    if (
      title ===
      'view/projection'
    ) {

      result.projection =
        grading_clean_(
          answer
        );

      if (
        result.projection
      ) {
        result.used =
          true;
      }

      continue;
    }

    if (
      Object.prototype
        .hasOwnProperty
        .call(
          componentNames,
          title
        )
    ) {

      const parsed =
        grading_parseImageScoreAnswer_(
          answer
        );

      const component = {
        name:
          componentNames[
            title
          ],

        rawAnswer:
          grading_clean_(
            answer
          ),

        applicable:
          parsed.applicable,

        score:
          parsed.score,

        explanation:
          parsed.explanation
      };

      result.components.push(
        component
      );

      if (
        !parsed.applicable ||
        parsed.score === null
      ) {
        continue;
      }

      result.grossPoints +=
        parsed.score;

      result.sectionsScored++;

      /*
       * One zero penalty per affected VIEW.
       * Additional zero-scored criteria in the same image do not create
       * additional -5 penalties. Separate views accumulate separately.
       */
      if (
        parsed.score === 0
      ) {
        result.zeroPenalty =
          penalties.zeroPenalty;
      }

      continue;
    }

    if (
      title ===
        'exposure index range:' ||

      title ===
        'exposure index range' ||

      title.includes(
        'exposure index'
      ) ||

      title.includes(
        's number'
      )
    ) {

      result.exposureIndex =
        grading_clean_(
          answer
        );

      continue;
    }

    if (
      title ===
        'marker placement' ||

      title ===
        'marker evaluation'
    ) {

      result.marker =
        grading_clean_(
          answer
        );

      const marker =
        appNormalize_(
          answer
        );

      if (
        marker ===
        'missing marker'
      ) {
        result.missingMarkerPenalty =
          penalties
            .missingMarkerPenalty;
      }

      if (
        marker ===
        'wrong marker'
      ) {
        result.wrongMarkerPenalty =
          penalties
            .wrongMarkerPenalty;
      }

      if (
        marker ===
          'illegible marker' ||

        marker ===
          'poorly placed marker' ||

        marker ===
          'illegible or poorly placed marker'
      ) {
        result.illegibleMarkerPenalty =
          penalties
            .illegibleMarkerPenalty;
      }

      continue;
    }

    /*
     * Image-view comments are explicitly titled in the finalized Form:
     *   View 1 Comments
     *   View 2 Comments
     *   ...
     *   View 8 Comments
     *
     * Historical responses may still contain a generic "Comments" item.
     * Because this function is already bounded to one Image Evaluation -
     * View N section, the generic title is a safe backward-compatible
     * fallback and cannot cross into another view.
     */
    const expectedViewCommentsTitle =
      'view ' + viewNumber + ' comments';

    if (
      title === expectedViewCommentsTitle ||
      title === 'comments' ||
      title === 'comment'
    ) {
      result.comments =
        grading_clean_(
          answer
        );
    }
  }

  /*
   * Penalties are cumulative across categories and across views.
   * Zero-in-view is capped at one -5 penalty per affected view.
   */
  result.netPoints =
    result.grossPoints +
    result.zeroPenalty +
    result.missingMarkerPenalty +
    result.wrongMarkerPenalty +
    result.illegibleMarkerPenalty;

  /*
   * Per-view detailed score is preserved raw and may be negative.
   */
  if (
    result.sectionsScored > 0
  ) {
    result.imageScore =
      result.netPoints /
      result.sectionsScored;
  }

  return result;
}


/**
 * =====================================================================
 * PARSE IMAGE SCORE RESPONSE
 * =====================================================================
 */
function grading_parseImageScoreAnswer_(
  answer
) {

  const text =
    grading_clean_(
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
    appNormalize_(
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
    grading_clean_(
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


/**
 * =====================================================================
 * LOAD IMAGE PENALTIES
 * =====================================================================
 *
 * Authoritative source:
 *
 *   Image Scoring Rules
 *
 * Required rules:
 *
 *   Zero In View Penalty
 *   Missing Marker Penalty
 *   Wrong Marker Penalty
 *   Illegible Marker Penalty
 *
 * Accepted aliases for the -3 rule:
 *   Illegible or Poorly Placed Marker Penalty
 *   Poorly Placed Marker Penalty
 * =====================================================================
 */
function grading_loadImagePenalties_(
  ss
) {

  const sheet =
    appGetSheet_(
      ss,
      APP_CONFIG.SHEETS
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
        appNormalize_(
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


/**
 * =====================================================================
 * HEADER MATCHER
 * =====================================================================
 */
function grading_findHeader_(
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
        appNormalize_(
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


/**
 * =====================================================================
 * NORMALIZE PERFORMANCE SECTION
 * =====================================================================
 */
function grading_normalizeSection_(
  value
) {

  const text =
    appNormalize_(
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


/**
 * =====================================================================
 * NORMALIZE PERFORMANCE CRITERION
 * =====================================================================
 */
function grading_criterionKey_(
  value
) {

  return grading_clean_(
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


/**
 * =====================================================================
 * CLEAN VALUE
 * =====================================================================
 */
function grading_clean_(
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
        grading_clean_
      )
      .join(', ');
  }

  return String(value)
    .trim();
}


/**
 * =====================================================================
 * ROUND TO TWO DECIMAL PLACES
 * =====================================================================
 */
function grading_round2_(
  value
) {

  return Math.round(
    (
      Number(value) +
      Number.EPSILON
    ) * 100
  ) / 100;
}
/**
 * =====================================================================
 * TEMPORARY QA — COMMENT MAPPING
 * READ ONLY
 * =====================================================================
 *
 * Compares:
 *   1. Every populated comment-like response in the latest Form response
 *   2. The comments captured by gradeClinicalEvaluation()
 *
 * DELETE after comment-mapping QA is complete.
 */
function testLatestClinicalEvaluationComments() {

  const form = FormApp.openById(
    APP_CONFIG.FORMS.DIGITAL_CLINICAL_EVALUATION
  );

  const responses = form.getResponses();

  if (!responses.length) {
    throw new Error(
      'The Digital Clinical Evaluation Form has no responses.'
    );
  }

  const response = responses[responses.length - 1];

  const responseByItemId =
    grading_buildResponseMap_(response);

  const grade =
    gradeClinicalEvaluation(response);


  // ------------------------------------------------------------
  // 1. RAW FORM COMMENT RESPONSES
  // ------------------------------------------------------------

  const rawComments = [];

  form.getItems().forEach(item => {

    const title =
      grading_clean_(item.getTitle());

    const normalizedTitle =
      appNormalize_(title);

    // Only inspect comment-like fields.
    if (
      !normalizedTitle.includes('comment')
    ) {
      return;
    }

    const answer =
      responseByItemId[
        String(item.getId())
      ];

    rawComments.push({
      itemId: String(item.getId()),
      title: title,
      response:
        answer === undefined ||
        answer === null
          ? ''
          : grading_clean_(answer)
    });
  });


  // ------------------------------------------------------------
  // 2. COMMENTS CAPTURED BY GRADING ENGINE
  // ------------------------------------------------------------

  const performanceComments =
    (grade.performance.sections || [])
      .map(section => ({
        section: section.name,
        comments: section.comments || ''
      }));


  const imageComments =
    (grade.image.views || [])
      .map(view => ({
        viewNumber: view.viewNumber,
        projection: view.projection,
        comments: view.comments || ''
      }));


  const summary = {

    student:
      grade.evaluation
        ? grade.evaluation.student
        : '',

    examination:
      grade.evaluation
        ? grade.evaluation.examination
        : '',

    responseId:
      grade.responseId,

    rawFormCommentFields:
      rawComments,

    capturedPerformanceComments:
      performanceComments,

    capturedImageComments:
      imageComments,

    capturedOverallEvaluatorComments:
      grade.evaluation
        ? grade.evaluation.comments || ''
        : '',

    capturedAutomaticFailureComments:
      grade.automaticFailure
        ? grade.automaticFailure.comments || ''
        : ''
  };


  console.log(
    JSON.stringify(
      summary,
      null,
      2
    )
  );

  return summary;
}
/**
 * =====================================================================
 * TEMPORARY QA — AUTOMATIC FAILURE CHECKPOINT MAPPING
 * READ ONLY
 * =====================================================================
 *
 * PURPOSE:
 *   Inspect every Automatic Failure-related field in the latest Digital
 *   Clinical Evaluation Form response and compare the raw Form values
 *   with the final automaticFailure object returned by
 *   gradeClinicalEvaluation().
 *
 * DOES NOT:
 *   - modify the Form
 *   - modify Clinical Tracking
 *   - modify grading rules
 *   - generate documents
 *   - write any data
 *
 * DELETE after Automatic Failure checkpoint QA is complete.
 */
function testLatestClinicalEvaluationAutomaticFailure() {

  const form = FormApp.openById(
    APP_CONFIG.FORMS.DIGITAL_CLINICAL_EVALUATION
  );

  const responses = form.getResponses();

  if (!responses.length) {
    throw new Error(
      'The Digital Clinical Evaluation Form has no responses.'
    );
  }

  const response =
    responses[responses.length - 1];

  const responseByItemId =
    grading_buildResponseMap_(response);

  const grade =
    gradeClinicalEvaluation(response);


  // ------------------------------------------------------------
  // RAW AUTOMATIC FAILURE-RELATED FORM FIELDS
  // ------------------------------------------------------------

  const rawAutomaticFailureFields = [];

  form.getItems().forEach(item => {

    const title =
      grading_clean_(item.getTitle());

    const normalizedTitle =
      appNormalize_(title);

    /*
     * Capture all fields that appear to belong to the
     * Automatic Failure workflow.
     *
     * This deliberately uses a broad match so that duplicated
     * checkpoints, reasons, comments, or differently titled
     * AF fields are visible during QA rather than silently missed.
     */
    if (
      !normalizedTitle.includes('automatic failure')
    ) {
      return;
    }

    const answer =
      responseByItemId[
        String(item.getId())
      ];

    rawAutomaticFailureFields.push({

      itemId:
        String(item.getId()),

      title:
        title,

      response:
        answer === undefined ||
        answer === null
          ? ''
          : grading_clean_(answer)

    });

  });


  // ------------------------------------------------------------
  // FINAL GRADED AUTOMATIC FAILURE RESULT
  // ------------------------------------------------------------

  const automaticFailure =
    grade.automaticFailure || {};


  // ------------------------------------------------------------
  // QA SUMMARY
  // ------------------------------------------------------------

  const summary = {

    student:
      grade.evaluation
        ? grade.evaluation.student
        : '',

    examination:
      grade.evaluation
        ? grade.evaluation.examination
        : '',

    responseId:
      grade.responseId,

    rawAutomaticFailureFields:
      rawAutomaticFailureFields,

    capturedAutomaticFailure: {

      occurred:
        automaticFailure.occurred === true,

      reasons:
        automaticFailure.reasons || [],

      comments:
        automaticFailure.comments || ''

    },

    finalEvaluationResult:
      grade.result || '',

    finalScore:
      grade.finalScore

  };


  console.log(
    JSON.stringify(
      summary,
      null,
      2
    )
  );

  return summary;
}

