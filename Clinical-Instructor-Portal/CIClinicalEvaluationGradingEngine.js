/**
 * =====================================================================
 * CCMS — CLINICAL INSTRUCTOR PORTAL
 * CLINICAL EVALUATION GRADING ENGINE
 * CIClinicalEvaluationGradingEngine.gs
 * =====================================================================
 *
 * STATUS
 * ------
 * CANDIDATE — READ-ONLY PORTAL GRADING ENGINE
 *
 * PURPOSE
 * -------
 * Portal-native sibling of ClinicalEvaluationGradingEngine.gs.
 *
 * The production engine grades a Google FormResponse.
 * This engine grades a Clinical Instructor Portal payload.
 *
 * ACADEMIC AUTHORITY
 * ------------------
 * This engine intentionally reproduces the grading behavior of:
 *
 *   ClinicalEvaluationGradingEngine.gs
 *
 * and reads the same authoritative workbook configuration:
 *
 *   Scoring Weights
 *   Image Scoring Rules
 *
 * THIS FILE OWNS
 * --------------
 * - Reading one Portal evaluation payload
 * - Performance scoring
 * - Image scoring
 * - Image penalties
 * - Automatic Failure detection
 * - Final competency score
 * - Pass / Fail determination
 *
 * THIS FILE DOES NOT
 * ------------------
 * - Modify the Google Form
 * - Modify ClinicalEvaluationGradingEngine.gs
 * - Write Clinical Tracking
 * - Generate documents
 * - Move files
 * - Send notifications
 * - Calculate semester/program requirements
 * - Save drafts
 * - Submit evaluations
 *
 * IMPORTANT
 * ---------
 * Portal UI structure may differ from the Google Form.
 * Academic grading behavior must not.
 *
 * =====================================================================
 */


/**
 * =====================================================================
 * MAIN PORTAL GRADING FUNCTION
 * =====================================================================
 *
 * Expected payload:
 *
 * {
 *   evaluation: {
 *     studentId: "...",
 *     student: "...",
 *     clinicalSiteId: "...",
 *     clinicalSite: "...",
 *     evaluatorId: "...",
 *     evaluator: "...",
 *     examinationId: "...",
 *     examination: "...",
 *     evaluationType: "Competency",
 *     date: "...",
 *     comments: "..."
 *   },
 *
 *   automaticFailure: {
 *     occurred: false,
 *     reasons: [],
 *     comments: ""
 *   },
 *
 *   performance: {
 *     sections: [
 *       {
 *         name: "Patient Care",
 *         rows: [
 *           {
 *             criterion: "...",
 *             rating: 3
 *           }
 *         ]
 *       }
 *     ]
 *   },
 *
 *   image: {
 *     views: [
 *       {
 *         projection: "AP",
 *         components: {
 *           positioning: 3,
 *           equipment: 3,
 *           radiationProtection: 3,
 *           exposure: 3
 *         },
 *         exposureIndex: "...",
 *         marker: "Wrong Marker",
 *         comments: "..."
 *       }
 *     ]
 *   }
 * }
 *
 * @param {Object} payload
 * @return {Object}
 */
function ciClinicalGradeEvaluation(payload) {

  if (
    !payload ||
    typeof payload !== 'object'
  ) {
    throw new Error(
      'A Clinical Instructor Portal evaluation payload is required.'
    );
  }

  const ss =
    ciDataBook_();


  /*
   * ---------------------------------------------------------------
   * READ PORTAL INPUT
   * ---------------------------------------------------------------
   */

  const evaluation =
    ciClinicalGrading_readEvaluationInfo_(
      payload.evaluation || {}
    );

  const automaticFailure =
    ciClinicalGrading_readAutomaticFailure_(
      payload.automaticFailure || {}
    );


  /*
   * ---------------------------------------------------------------
   * PERFORMANCE
   *
   * Uses the SAME authoritative Scoring Weights loader as the
   * production Google Form grading engine.
   * ---------------------------------------------------------------
   */

  const weights =
    ciPortalAcademic_loadPerformanceWeights_(ss);

  const performance =
    ciClinicalGrading_gradePerformance_(
      payload.performance || {},
      weights
    );


  /*
   * ---------------------------------------------------------------
   * IMAGE EVALUATION
   *
   * Uses the SAME authoritative Image Scoring Rules loader as the
   * production Google Form grading engine.
   * ---------------------------------------------------------------
   */

  const penalties =
    ciPortalAcademic_loadImagePenalties_(ss);

  const image =
    ciClinicalGrading_gradeImages_(
      payload.image || {},
      penalties
    );


  /*
   * ---------------------------------------------------------------
   * FINAL SCORE
   *
   * Program Handbook August 2026, section 13.7 and original CCMS
   * 04_ClinicalEvaluationGradingEngine.js:
   *   0.0 on Automatic Failure; otherwise (Performance + Image) / 2
   * ---------------------------------------------------------------
   */

  let finalScore = null;

  // The original CCMS engine (04_ClinicalEvaluationGradingEngine.js)
  // overrides the final competency score to 0.0 for Automatic Failure.
  // Retain performance/image details unchanged for academic audit.
  if (automaticFailure.occurred) {
    finalScore = 0;
  } else if (
    performance.score !== null &&
    image.score !== null
  ) {
    finalScore = (performance.score + image.score) / 2;
  }


  /*
   * ---------------------------------------------------------------
   * PASS / FAIL
   *
   * EXACT production thresholds:
   *
   * Performance >= 2.00
   * Image       >= 2.00
   * No Automatic Failure
   * ---------------------------------------------------------------
   */

  const performancePass =
    performance.score !== null &&
    performance.score >= 2.0;

  const imagePass =
    image.score !== null &&
    image.score >= 2.0;


  let result =
    CI_PORTAL_GRADE_CONFIG_
      .TERMINOLOGY
      .RESULTS
      .FAIL
      .toUpperCase();


  if (
    !automaticFailure.occurred &&
    performancePass &&
    imagePass
  ) {

    result =
      CI_PORTAL_GRADE_CONFIG_
        .TERMINOLOGY
        .RESULTS
        .PASS
        .toUpperCase();
  }


  /*
   * ---------------------------------------------------------------
   * RETURN
   *
   * Mirrors the production engine result shape as closely as possible.
   * ---------------------------------------------------------------
   */

  return {

    source:
      'CI_PORTAL',

    responseId:
      '',

    submittedAt:
      '',

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
        : ciPortalAcademic_round2_(
            finalScore
          ),

    result:
      result
  };
}


/**
 * =====================================================================
 * BASIC EVALUATION INFORMATION
 * =====================================================================
 */
function ciClinicalGrading_readEvaluationInfo_(source) {

  return {

    studentId:
      ciClinicalGrading_clean_(
        source.studentId
      ),

    student:
      ciClinicalGrading_clean_(
        source.student
      ),

    clinicalSiteId:
      ciClinicalGrading_clean_(
        source.clinicalSiteId
      ),

    clinicalSite:
      ciClinicalGrading_clean_(
        source.clinicalSite
      ),

    evaluatorId:
      ciClinicalGrading_clean_(
        source.evaluatorId
      ),

    evaluator:
      ciClinicalGrading_clean_(
        source.evaluator
      ),

    examinationId:
      ciClinicalGrading_clean_(
        source.examinationId
      ),

    examination:
      ciClinicalGrading_clean_(
        source.examination
      ),

    evaluationType:
      ciClinicalGrading_clean_(
        source.evaluationType
      ),

    date:
      source.date || '',

    comments:
      ciClinicalGrading_clean_(
        source.comments
      )
  };
}


/**
 * =====================================================================
 * AUTOMATIC FAILURE
 * =====================================================================
 */
function ciClinicalGrading_readAutomaticFailure_(source) {

  let reasons = [];

  if (
    Array.isArray(
      source.reasons
    )
  ) {

    reasons =
      source.reasons
        .map(
          ciClinicalGrading_clean_
        )
        .filter(Boolean);

  } else {

    const reason =
      ciClinicalGrading_clean_(
        source.reasons
      );

    if (reason) {
      reasons = [
        reason
      ];
    }
  }


  return {

    occurred:
      source.occurred === true,

    reasons:
      reasons,

    comments:
      ciClinicalGrading_clean_(
        source.comments
      )
  };
}


/**
 * =====================================================================
 * PERFORMANCE GRADING
 * =====================================================================
 *
 * Mirrors ciPortalAcademic_gradePerformance_() from the production engine.
 *
 * DIFFERENCE:
 * Production reads Google Form GRID items.
 * Portal reads payload.performance.sections.
 *
 * Scoring behavior remains the same.
 * =====================================================================
 */
function ciClinicalGrading_gradePerformance_(
  performancePayload,
  weights
) {

  const sections =
    Array.isArray(
      performancePayload.sections
    )
      ? performancePayload.sections
      : [];


  const result = {

    sections:
      [],

    totalPoints:
      0,

    totalWeight:
      0,

    score:
      null
  };


  sections.forEach(
    sectionPayload => {

      const sectionName =
        ciClinicalGrading_clean_(
          sectionPayload.name
        );


      if (!sectionName) {

        throw new Error(
          'Portal Performance section is missing its name.'
        );
      }


      const rows =
        Array.isArray(
          sectionPayload.rows
        )
          ? sectionPayload.rows
          : [];


      const section = {

        name:
          sectionName,

        points:
          0,

        weight:
          0,

        score:
          null,

        rows:
          []
      };


      rows.forEach(
        rowPayload => {

          const criterion =
            ciClinicalGrading_clean_(
              rowPayload.criterion
            );


          if (!criterion) {

            throw new Error(
              'Portal Performance criterion is missing in section "' +
              sectionName +
              '".'
            );
          }


          /*
           * SAME production criterion normalization.
           */

          const criterionKey =
            ciPortalAcademic_criterionKey_(
              criterion
            );


          /*
           * SAME production authoritative weight lookup.
           */

          const weightRow =
            ciPortalAcademic_findWeight_(
              weights,
              sectionName,
              criterionKey
            );


          if (!weightRow) {

            throw new Error(
              'Could not find scoring weight for:\n\n' +
              sectionName +
              '\n' +
              criterion
            );
          }


          const rawRating =
            ciClinicalGrading_clean_(
              rowPayload.rating
            );


          /*
           * SAME grading semantics as production:
           *
           * Blank or N/A is non-applicable.
           *
           * Workflow validation preventing blank submission belongs
           * outside this grading engine.
           */

          const isNA =
            !rawRating ||
            ciPortalAcademicNormalize_(
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


          /*
           * SAME production calculation.
           */

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
    }
  );


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
 * =====================================================================
 * IMAGE GRADING
 * =====================================================================
 *
 * Mirrors ciPortalAcademic_gradeImages_() from the production engine.
 *
 * DIFFERENCE:
 * Production discovers Image Evaluation page sections in Google Forms.
 * Portal receives an explicit array of image views.
 * =====================================================================
 */
function ciClinicalGrading_gradeImages_(
  imagePayload,
  penalties
) {

  const sourceViews =
    Array.isArray(
      imagePayload.views
    )
      ? imagePayload.views
      : [];


  const views = [];


  sourceViews.forEach(
    (
      sourceView,
      index
    ) => {

      const view =
        ciClinicalGrading_gradeOneImageView_(
          sourceView || {},
          index + 1,
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
  );


  /*
   * SAME production aggregation fields.
   */

  let grossPoints = 0;

  let zeroPenalty = 0;

  let missingMarkerPenalty = 0;

  let wrongMarkerPenalty = 0;

  let illegibleMarkerPenalty = 0;

  let netPoints = 0;

  let sectionsScored = 0;


  views.forEach(
    view => {

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
    }
  );


  /*
   * SAME production image-score calculation.
   */

  const rawScore =
    sectionsScored > 0
      ? netPoints /
        sectionsScored
      : null;


  /*
   * SAME production floor.
   *
   * Detailed raw math may be negative.
   * Grade-bearing Image Score cannot be below zero.
   */

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
 * GRADE ONE PORTAL IMAGE VIEW
 * =====================================================================
 *
 * Mirrors ciPortalAcademic_gradeOneImageView_() from production.
 *
 * Portal image contract:
 *
 * {
 *   projection: "AP",
 *
 *   components: {
 *     positioning: 3,
 *     equipment: 3,
 *     radiationProtection: 3,
 *     exposure: 3
 *   },
 *
 *   exposureIndex: "...",
 *
 *   marker: "Wrong Marker",
 *
 *   comments: "..."
 * }
 *
 * IMPORTANT:
 * marker uses the SAME academic values recognized by production:
 *
 *   Missing Marker
 *   Wrong Marker
 *   Illegible Marker
 *   Poorly Placed Marker
 *   Illegible or Poorly Placed Marker
 *
 * =====================================================================
 */
function ciClinicalGrading_gradeOneImageView_(
  source,
  viewNumber,
  penalties
) {

  const result = {

    used:
      false,

    viewNumber:
      viewNumber,

    projection:
      ciClinicalGrading_clean_(
        source.projection
      ),

    components:
      [],

    exposureIndex:
      ciClinicalGrading_clean_(
        source.exposureIndex
      ),

    marker:
      ciClinicalGrading_clean_(
        source.marker
      ),

    comments:
      ciClinicalGrading_clean_(
        source.comments
      ),

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


  /*
   * SAME concept as production:
   * an image view is used when View/Projection exists.
   */

  if (
    result.projection
  ) {

    result.used =
      true;
  }


  if (
    !result.used
  ) {

    return result;
  }


  const components =
    source.components || {};


  /*
   * SAME four production image components.
   */

  const componentDefinitions = [

    {
      key:
        'positioning',

      name:
        'Positioning'
    },

    {
      key:
        'equipment',

      name:
        'Equipment / Central Ray'
    },

    {
      key:
        'radiationProtection',

      name:
        'Radiation Protection / Artifacts'
    },

    {
      key:
        'exposure',

      name:
        'Exposure'
    }
  ];


  componentDefinitions.forEach(
    definition => {

      const answer =
        components[
          definition.key
        ];


      /*
       * SAME production parser.
       */

      const parsed =
        ciPortalAcademic_parseImageScoreAnswer_(
          answer
        );


      const component = {

        name:
          definition.name,

        rawAnswer:
          ciClinicalGrading_clean_(
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

        return;
      }


      result.grossPoints +=
        parsed.score;


      result.sectionsScored++;


      /*
       * EXACT production Zero-in-View behavior:
       *
       * ONE zero penalty per affected view.
       *
       * Multiple zero-scored components in the same view do not
       * multiply the penalty.
       */

      if (
        parsed.score === 0
      ) {

        result.zeroPenalty =
          penalties.zeroPenalty;
      }
    }
  );


  /*
   * ---------------------------------------------------------------
   * MARKER PENALTIES
   *
   * EXACT production interpretation.
   * ---------------------------------------------------------------
   */

  const marker =
    ciPortalAcademicNormalize_(
      result.marker
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


  /*
   * SAME production cumulative calculation.
   */

  result.netPoints =
    result.grossPoints +
    result.zeroPenalty +
    result.missingMarkerPenalty +
    result.wrongMarkerPenalty +
    result.illegibleMarkerPenalty;


  /*
   * SAME production raw per-view calculation.
   *
   * This value may be negative.
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
 * PORTAL VALUE CLEANER
 * =====================================================================
 *
 * Namespaced so it cannot collide with the production grading engine.
 */
function ciClinicalGrading_clean_(value) {

  if (
    value === null ||
    value === undefined
  ) {

    return '';
  }


  if (
    Array.isArray(
      value
    )
  ) {

    return value
      .map(
        ciClinicalGrading_clean_
      )
      .join(', ');
  }


  return String(
    value
  )
    .trim();
}
