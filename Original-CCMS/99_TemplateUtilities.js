/**
 * =====================================================================
 * CLINICAL COMPETENCY TRACKER
 * TEMPLATE UTILITIES
 * 99_TemplateUtilities.gs
 * =====================================================================
 *
 * PURPOSE
 * -------
 * Development utilities for inspecting and validating Google Doc
 * templates used by the Clinical Competency Tracker.
 *
 * THIS FILE MAY:
 *
 *   - Verify template accessibility
 *   - Verify that a template is a native Google Doc
 *   - List template replacement tokens
 *   - Check expected replacement tokens
 *   - Inspect table structure
 *   - Inspect document paragraphs/headings
 *
 * THIS FILE DOES NOT:
 *
 *   - Generate production documents
 *   - Grade evaluations
 *   - Write Clinical Tracking
 *   - Modify requirements
 *   - Send notifications
 *   - Run automatically
 *
 * IMPORTANT
 * ---------
 * Utilities in this file are intended to be run manually by an
 * administrator/developer while building or troubleshooting templates.
 *
 * =====================================================================
 */


/**
 * =====================================================================
 * TEMPLATE UTILITY SETTINGS
 * =====================================================================
 *
 * OPTIONAL:
 *
 * Paste any Google Doc template ID into MANUAL_TEMPLATE_ID when you
 * want to inspect a template that is not yet configured in APP_CONFIG.
 */
const TEMPLATE_UTILITY_CONFIG = {

  MANUAL_TEMPLATE_ID:
    ''

};


/**
 * =====================================================================
 * INSPECT ALL CONFIGURED TEMPLATES
 * =====================================================================
 *
 * READ ONLY.
 *
 * Gives us one quick overview of every document template currently
 * configured in APP_CONFIG.
 */
function templateUtility_inspectConfiguredTemplates() {

  const templates = [];


  // ===================================================================
  // CLINICAL COMPETENCY EVALUATION
  // ===================================================================

  if (
    APP_CONFIG.TEMPLATES
      .CLINICAL_COMPETENCY_EVALUATION
  ) {

    templates.push({

      label:
        'Clinical Competency Evaluation',

      id:
        APP_CONFIG.TEMPLATES
          .CLINICAL_COMPETENCY_EVALUATION

    });
  }


  // ===================================================================
  // SEMESTER REQUIREMENT TEMPLATES
  // ===================================================================

  const semesterTemplates =
    APP_CONFIG.TEMPLATES
      .SEMESTER_REQUIREMENTS ||
    {};


  Object
    .keys(
      semesterTemplates
    )
    .forEach(key => {

      const id =
        semesterTemplates[key];


      if (!id) {
        return;
      }


      templates.push({

        label:
          'Semester Requirements - ' +
          key,

        id:
          id

      });

    });


  // ===================================================================
  // PROGRAM REQUIREMENT TEMPLATE
  // ===================================================================

  if (
    APP_CONFIG.TEMPLATES
      .PROGRAM_REQUIREMENTS
  ) {

    templates.push({

      label:
        'Program Requirements',

      id:
        APP_CONFIG.TEMPLATES
          .PROGRAM_REQUIREMENTS

    });
  }


  // ===================================================================
  // INSPECT
  // ===================================================================

  const results =
    templates.map(template => {

      try {

        const inspection =
          templateUtility_getTemplateInfo_(
            template.id
          );


        return {

          label:
            template.label,

          success:
            true,

          id:
            template.id,

          name:
            inspection.name,

          mimeType:
            inspection.mimeType,

          nativeGoogleDoc:
            inspection.nativeGoogleDoc,

          tableCount:
            inspection.tableCount,

          paragraphCount:
            inspection.paragraphCount,

          tokenCount:
            inspection.tokens.length,

          tokens:
            inspection.tokens

        };


      } catch (error) {

        return {

          label:
            template.label,

          success:
            false,

          id:
            template.id,

          error:
            error.message

        };
      }

    });


  console.log(
    JSON.stringify(
      results,
      null,
      2
    )
  );


  return results;
}


/**
 * =====================================================================
 * INSPECT MANUAL TEMPLATE
 * =====================================================================
 *
 * READ ONLY.
 *
 * Put a Google Doc ID into:
 *
 *   TEMPLATE_UTILITY_CONFIG.MANUAL_TEMPLATE_ID
 *
 * Then run this function.
 */
function templateUtility_inspectManualTemplate() {

  const templateId =
    String(
      TEMPLATE_UTILITY_CONFIG
        .MANUAL_TEMPLATE_ID ||
      ''
    ).trim();


  if (!templateId) {

    throw new Error(
      'Enter a Google Doc template ID in ' +
      'TEMPLATE_UTILITY_CONFIG.MANUAL_TEMPLATE_ID first.'
    );
  }


  const result =
    templateUtility_getTemplateInfo_(
      templateId
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


/**
 * =====================================================================
 * VERIFY CLINICAL COMPETENCY EVALUATION TEMPLATE
 * =====================================================================
 *
 * READ ONLY.
 *
 * This replaces the old standalone:
 *
 *   verifyCompetencyTemplate()
 *
 * The current Evaluation Document Generator does NOT use marker tokens.
 * It depends on the existing formatted table structure.
 *
 * Therefore this utility verifies the structural pieces that generator
 * expects to find.
 */
function templateUtility_verifyEvaluationTemplate() {

  const templateId =
    APP_CONFIG.TEMPLATES
      .CLINICAL_COMPETENCY_EVALUATION;


  if (!templateId) {

    throw new Error(
      'Clinical Competency Evaluation template is not configured.'
    );
  }


  const file =
    DriveApp.getFileById(
      templateId
    );


  console.log(
    '=========================================='
  );

  console.log(
    'CLINICAL COMPETENCY EVALUATION TEMPLATE'
  );

  console.log(
    '=========================================='
  );


  console.log(
    'Name: ' +
    file.getName()
  );


  console.log(
    'ID: ' +
    file.getId()
  );


  console.log(
    'MIME Type: ' +
    file.getMimeType()
  );


  if (
    file.getMimeType() !==
    'application/vnd.google-apps.document'
  ) {

    throw new Error(
      'Evaluation template is not a native Google Doc. MIME type: ' +
      file.getMimeType()
    );
  }


  const doc =
    DocumentApp.openById(
      templateId
    );


  const body =
    doc.getBody();


  const tables =
    body.getTables();


  const checks = {


    nativeGoogleDoc:
      true,


    minimumTableCount: {

      expected:
        13,

      actual:
        tables.length,

      pass:
        tables.length >= 13

    },


    view1Heading:
      Boolean(
        templateUtility_findParagraphStarting_(
          body,
          'View 1 -'
        )
      ),


    view2Heading:
      Boolean(
        templateUtility_findParagraphStarting_(
          body,
          'View 2 -'
        )
      ),


    view3Heading:
      Boolean(
        templateUtility_findParagraphStarting_(
          body,
          'View 3 -'
        )
      ),


    performanceTotal:
      Boolean(
        templateUtility_findParagraphStarting_(
          body,
          'Performance Total:'
        )
      ),


    finalCalculation:
      Boolean(
        templateUtility_findParagraphStarting_(
          body,
          'Performance:'
        )
      ),


    evaluatorComments:
      Boolean(
        templateUtility_findExactParagraph_(
          body,
          'Evaluator Comments'
        )
      ),


    finalEvaluationSummary:
      Boolean(
        templateUtility_findParagraphStarting_(
          body,
          'Final Evaluation Summary'
        )
      )

  };


  // ===================================================================
  // OVERALL IMAGE SUMMARY
  // ===================================================================

  const imageSummaryTable =
    templateUtility_findTableByFirstCell_(
      body,
      'Views'
    );


  checks.overallImageSummary =
    Boolean(
      imageSummaryTable
    );


  // ===================================================================
  // FINAL TABLE
  // ===================================================================

  const finalTable =
    templateUtility_findFinalEvaluationTable_(
      body
    );


  checks.finalEvaluationTable =
    Boolean(
      finalTable
    );


  // ===================================================================
  // OVERALL PASS / FAIL
  // ===================================================================

  const failures = [];


  if (
    !checks.minimumTableCount.pass
  ) {

    failures.push(
      'Template has fewer than 13 tables.'
    );
  }


  if (!checks.view1Heading) {

    failures.push(
      'View 1 heading is missing.'
    );
  }


  if (!checks.view2Heading) {

    failures.push(
      'View 2 heading is missing.'
    );
  }


  if (!checks.view3Heading) {

    failures.push(
      'View 3 heading is missing.'
    );
  }


  if (!checks.performanceTotal) {

    failures.push(
      'Performance Total paragraph is missing.'
    );
  }


  if (!checks.finalCalculation) {

    failures.push(
      'Final calculation paragraph is missing.'
    );
  }


  if (!checks.overallImageSummary) {

    failures.push(
      'Overall Image Summary table is missing.'
    );
  }


  if (!checks.finalEvaluationSummary) {

    failures.push(
      'Final Evaluation Summary heading is missing.'
    );
  }


  if (!checks.finalEvaluationTable) {

    failures.push(
      'Final Evaluation Summary table is missing.'
    );
  }


  if (!checks.evaluatorComments) {

    failures.push(
      'Evaluator Comments heading is missing.'
    );
  }


  const result = {

    success:
      failures.length === 0,

    templateName:
      file.getName(),

    templateId:
      templateId,

    tableCount:
      tables.length,

    checks:
      checks,

    failures:
      failures

  };


  console.log(
    JSON.stringify(
      result,
      null,
      2
    )
  );


  if (
    failures.length > 0
  ) {

    throw new Error(
      'Evaluation template verification failed:\n\n' +
      failures.join('\n')
    );
  }


  console.log(
    'PASS - Evaluation template structure is valid.'
  );


  return result;
}


/**
 * =====================================================================
 * VERIFY CLINICAL 2 SEMESTER REQUIREMENT TEMPLATE
 * =====================================================================
 *
 * READ ONLY.
 *
 * The Clinical 2 generator currently expects:
 *
 *   Student header:
 *     {{NAME}}
 *     {{ID}}
 *
 *   13 competency rows:
 *     C1-C13
 *
 *   10 additional rows:
 *     A1-A10
 *
 *   9 ISO rows:
 *     I1-I9
 *
 *   7 failed competency rows:
 *     F1-F7
 */
function templateUtility_verifyClinical2Template() {

  const templateId =
    APP_CONFIG
      .TEMPLATES
      .SEMESTER_REQUIREMENTS
      .CLINICAL_2;


  if (!templateId) {

    throw new Error(
      'Clinical 2 Semester Requirement template is not configured.'
    );
  }


  const expectedTokens = [

    '{{NAME}}',
    '{{ID}}'

  ];


  // ===================================================================
  // 13 SEMESTER COMPETENCIES
  // ===================================================================

  for (
    let i = 1;
    i <= 13;
    i++
  ) {

    expectedTokens.push(
      '{{C' + i + 'E}}'
    );

    expectedTokens.push(
      '{{C' + i + 'N}}'
    );

    expectedTokens.push(
      '{{C' + i + 'H}}'
    );

    expectedTokens.push(
      '{{C' + i + 'G}}'
    );

    expectedTokens.push(
      '{{C' + i + 'D}}'
    );
  }


  // ===================================================================
  // 10 ADDITIONAL COMPETENCIES
  // ===================================================================

  for (
    let i = 1;
    i <= 10;
    i++
  ) {

    expectedTokens.push(
      '{{A' + i + 'E}}'
    );

    expectedTokens.push(
      '{{A' + i + 'N}}'
    );

    expectedTokens.push(
      '{{A' + i + 'H}}'
    );

    expectedTokens.push(
      '{{A' + i + 'G}}'
    );

    expectedTokens.push(
      '{{A' + i + 'D}}'
    );
  }


  // ===================================================================
  // 9 ISO ROWS
  // ===================================================================

  for (
    let i = 1;
    i <= 9;
    i++
  ) {

    expectedTokens.push(
      '{{I' + i + 'E}}'
    );

    expectedTokens.push(
      '{{I' + i + 'N}}'
    );

    expectedTokens.push(
      '{{I' + i + 'G}}'
    );

    expectedTokens.push(
      '{{I' + i + 'D}}'
    );
  }


  // ===================================================================
  // 7 FAILURE ROWS
  // ===================================================================

  for (
    let i = 1;
    i <= 7;
    i++
  ) {

    expectedTokens.push(
      '{{F' + i + 'E}}'
    );

    expectedTokens.push(
      '{{F' + i + 'N}}'
    );

    expectedTokens.push(
      '{{F' + i + 'G}}'
    );

    expectedTokens.push(
      '{{F' + i + 'D}}'
    );
  }


  return templateUtility_verifyTokens_(

    templateId,

    expectedTokens,

    'Clinical 2 Semester Requirement'

  );
}


/**
 * =====================================================================
 * LIST TOKENS IN MANUAL TEMPLATE
 * =====================================================================
 *
 * READ ONLY.
 *
 * Useful when preparing a new template.
 *
 * Finds anything formatted like:
 *
 *   {{TOKEN}}
 */
function templateUtility_listManualTemplateTokens() {

  const templateId =
    String(
      TEMPLATE_UTILITY_CONFIG
        .MANUAL_TEMPLATE_ID ||
      ''
    ).trim();


  if (!templateId) {

    throw new Error(
      'Enter a Google Doc ID in ' +
      'TEMPLATE_UTILITY_CONFIG.MANUAL_TEMPLATE_ID first.'
    );
  }


  const tokens =
    templateUtility_getTokens_(
      templateId
    );


  console.log(
    '=========================================='
  );

  console.log(
    'TEMPLATE TOKENS'
  );

  console.log(
    '=========================================='
  );


  tokens.forEach(
    token => {

      console.log(
        token
      );

    }
  );


  console.log(
    '=========================================='
  );

  console.log(
    'TOTAL TOKENS: ' +
    tokens.length
  );

  console.log(
    '=========================================='
  );


  return tokens;
}


/**
 * =====================================================================
 * INSPECT MANUAL TEMPLATE TABLES
 * =====================================================================
 *
 * READ ONLY.
 *
 * Very useful when a document generator relies on existing formatted
 * tables instead of replacement tokens.
 *
 * For each table this reports:
 *
 *   - Table number
 *   - Row count
 *   - Cell count per row
 *   - First-cell text
 */
function templateUtility_inspectManualTemplateTables() {

  const templateId =
    String(
      TEMPLATE_UTILITY_CONFIG
        .MANUAL_TEMPLATE_ID ||
      ''
    ).trim();


  if (!templateId) {

    throw new Error(
      'Enter a Google Doc ID in ' +
      'TEMPLATE_UTILITY_CONFIG.MANUAL_TEMPLATE_ID first.'
    );
  }


  templateUtility_requireGoogleDoc_(
    templateId
  );


  const doc =
    DocumentApp.openById(
      templateId
    );


  const tables =
    doc
      .getBody()
      .getTables();


  const result =
    tables.map(
      (
        table,
        tableIndex
      ) => {

        const rows = [];


        for (
          let rowIndex = 0;
          rowIndex < table.getNumRows();
          rowIndex++
        ) {

          rows.push({

            row:
              rowIndex + 1,

            cells:
              table
                .getRow(
                  rowIndex
                )
                .getNumCells()

          });
        }


        let firstCellText =
          '';


        if (
          table.getNumRows() > 0 &&
          table
            .getRow(0)
            .getNumCells() > 0
        ) {

          firstCellText =
            table
              .getRow(0)
              .getCell(0)
              .getText()
              .trim();
        }


        return {

          table:
            tableIndex + 1,

          rowCount:
            table.getNumRows(),

          rows:
            rows,

          firstCellText:
            firstCellText

        };

      }
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


/**
 * =====================================================================
 * INSPECT MANUAL TEMPLATE PARAGRAPHS
 * =====================================================================
 *
 * READ ONLY.
 *
 * Prints all non-empty top-level paragraphs in document order.
 *
 * Useful when a generator needs to find headings such as:
 *
 *   Performance Total:
 *   View 1 -
 *   Final Evaluation Summary
 */
function templateUtility_inspectManualTemplateParagraphs() {

  const templateId =
    String(
      TEMPLATE_UTILITY_CONFIG
        .MANUAL_TEMPLATE_ID ||
      ''
    ).trim();


  if (!templateId) {

    throw new Error(
      'Enter a Google Doc ID in ' +
      'TEMPLATE_UTILITY_CONFIG.MANUAL_TEMPLATE_ID first.'
    );
  }


  templateUtility_requireGoogleDoc_(
    templateId
  );


  const doc =
    DocumentApp.openById(
      templateId
    );


  const body =
    doc.getBody();


  const paragraphs = [];


  for (
    let i = 0;
    i < body.getNumChildren();
    i++
  ) {

    const child =
      body.getChild(i);


    if (
      child.getType() !==
      DocumentApp.ElementType.PARAGRAPH
    ) {

      continue;
    }


    const paragraph =
      child.asParagraph();


    const text =
      paragraph
        .getText()
        .trim();


    if (!text) {
      continue;
    }


    paragraphs.push({

      childIndex:
        i,

      text:
        text,

      heading:
        String(
          paragraph.getHeading()
        )

    });
  }


  console.log(
    JSON.stringify(
      paragraphs,
      null,
      2
    )
  );


  return paragraphs;
}


/**
 * =====================================================================
 * VERIFY EXPECTED TOKENS
 * =====================================================================
 *
 * READ ONLY.
 */
function templateUtility_verifyTokens_(
  templateId,
  expectedTokens,
  label
) {

  templateUtility_requireGoogleDoc_(
    templateId
  );


  const actualTokens =
    templateUtility_getTokens_(
      templateId
    );


  const actualSet =
    new Set(
      actualTokens
    );


  const missing =
    expectedTokens.filter(
      token =>

        !actualSet.has(
          token
        )

    );


  const expectedSet =
    new Set(
      expectedTokens
    );


  const unexpected =
    actualTokens.filter(
      token =>

        !expectedSet.has(
          token
        )

    );


  const result = {

    success:
      missing.length === 0,

    template:
      label || '',

    expectedTokenCount:
      expectedTokens.length,

    actualTokenCount:
      actualTokens.length,

    missing:
      missing,

    unexpected:
      unexpected

  };


  console.log(
    JSON.stringify(
      result,
      null,
      2
    )
  );


  if (
    missing.length > 0
  ) {

    throw new Error(
      (label || 'Template') +
      ' is missing ' +
      missing.length +
      ' expected token(s):\n\n' +
      missing.join('\n')
    );
  }


  console.log(
    'PASS - All expected template tokens were found.'
  );


  return result;
}


/**
 * =====================================================================
 * GET TEMPLATE INFORMATION
 * =====================================================================
 *
 * READ ONLY.
 */
function templateUtility_getTemplateInfo_(
  templateId
) {

  const file =
    DriveApp.getFileById(
      templateId
    );


  const result = {

    id:
      file.getId(),

    name:
      file.getName(),

    mimeType:
      file.getMimeType(),

    nativeGoogleDoc:

      file.getMimeType() ===
      'application/vnd.google-apps.document',

    tableCount:
      null,

    paragraphCount:
      null,

    tokens:
      []

  };


  if (
    !result.nativeGoogleDoc
  ) {

    return result;
  }


  const doc =
    DocumentApp.openById(
      templateId
    );


  const body =
    doc.getBody();


  result.tableCount =
    body
      .getTables()
      .length;


  let paragraphCount =
    0;


  for (
    let i = 0;
    i < body.getNumChildren();
    i++
  ) {

    if (
      body
        .getChild(i)
        .getType() ===
      DocumentApp.ElementType.PARAGRAPH
    ) {

      paragraphCount++;
    }
  }


  result.paragraphCount =
    paragraphCount;


  result.tokens =
    templateUtility_extractTokensFromText_(
      body.getText()
    );


  return result;
}


/**
 * =====================================================================
 * GET TOKENS FROM GOOGLE DOC
 * =====================================================================
 */
function templateUtility_getTokens_(
  templateId
) {

  templateUtility_requireGoogleDoc_(
    templateId
  );


  const doc =
    DocumentApp.openById(
      templateId
    );


  return templateUtility_extractTokensFromText_(

    doc
      .getBody()
      .getText()

  );
}


/**
 * =====================================================================
 * EXTRACT TOKENS FROM TEXT
 * =====================================================================
 *
 * Recognizes:
 *
 *   {{NAME}}
 *   {{C1E}}
 *   {{PROGRAM_REQUIREMENT_1}}
 *
 * etc.
 */
function templateUtility_extractTokensFromText_(
  text
) {

  const matches =
    String(
      text || ''
    )
      .match(
        /\{\{[^{}]+\}\}/g
      ) ||
    [];


  return Array
    .from(
      new Set(
        matches
      )
    )
    .sort();
}


/**
 * =====================================================================
 * REQUIRE NATIVE GOOGLE DOC
 * =====================================================================
 */
function templateUtility_requireGoogleDoc_(
  templateId
) {

  if (!templateId) {

    throw new Error(
      'Template ID is required.'
    );
  }


  const file =
    DriveApp.getFileById(
      templateId
    );


  if (
    file.getMimeType() !==
    'application/vnd.google-apps.document'
  ) {

    throw new Error(
      '"' +
      file.getName() +
      '" is not a native Google Doc. MIME type: ' +
      file.getMimeType()
    );
  }


  return file;
}


/**
 * =====================================================================
 * FIND PARAGRAPH STARTING WITH TEXT
 * =====================================================================
 */
function templateUtility_findParagraphStarting_(
  body,
  startText
) {

  for (
    let i = 0;
    i < body.getNumChildren();
    i++
  ) {

    const child =
      body.getChild(i);


    if (
      child.getType() !==
      DocumentApp.ElementType.PARAGRAPH
    ) {

      continue;
    }


    const text =
      child
        .asParagraph()
        .getText()
        .trim();


    if (
      text.indexOf(
        startText
      ) === 0
    ) {

      return child.asParagraph();
    }
  }


  return null;
}


/**
 * =====================================================================
 * FIND EXACT PARAGRAPH
 * =====================================================================
 */
function templateUtility_findExactParagraph_(
  body,
  exactText
) {

  for (
    let i = 0;
    i < body.getNumChildren();
    i++
  ) {

    const child =
      body.getChild(i);


    if (
      child.getType() !==
      DocumentApp.ElementType.PARAGRAPH
    ) {

      continue;
    }


    const text =
      child
        .asParagraph()
        .getText()
        .trim();


    if (
      text ===
      exactText
    ) {

      return child.asParagraph();
    }
  }


  return null;
}


/**
 * =====================================================================
 * FIND TABLE BY FIRST CELL
 * =====================================================================
 */
function templateUtility_findTableByFirstCell_(
  body,
  firstCellText
) {

  const tables =
    body.getTables();


  for (
    let i = 0;
    i < tables.length;
    i++
  ) {

    const table =
      tables[i];


    if (
      table.getNumRows() < 1
    ) {

      continue;
    }


    if (
      table
        .getRow(0)
        .getNumCells() < 1
    ) {

      continue;
    }


    const text =
      table
        .getRow(0)
        .getCell(0)
        .getText()
        .trim();


    if (
      text ===
      firstCellText
    ) {

      return table;
    }
  }


  return null;
}


/**
 * =====================================================================
 * FIND FINAL EVALUATION TABLE
 * =====================================================================
 */
function templateUtility_findFinalEvaluationTable_(
  body
) {

  const tables =
    body.getTables();


  for (
    let i = 0;
    i < tables.length;
    i++
  ) {

    const table =
      tables[i];


    if (
      table.getNumRows() !== 4
    ) {

      continue;
    }


    if (
      table
        .getRow(0)
        .getNumCells() !== 3
    ) {

      continue;
    }


    const firstCell =
      table
        .getRow(0)
        .getCell(0)
        .getText()
        .trim();


    if (
      firstCell ===
      'Performance Score'
    ) {

      return table;
    }
  }


  return null;
}
