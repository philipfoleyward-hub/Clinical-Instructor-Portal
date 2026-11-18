/**
 * STUDENT CLINICAL HISTORY SERVICE
 * Read-only backend for Students > Clinical History.
 */

function clinicalHistory_getStudent(studentId) {
  if (!studentId) throw new Error('Student ID is required.');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const students = appSheetToObjects_(
    appGetSheet_(ss, APP_CONFIG.SHEETS.STUDENTS)
  );
  const tracking = appSheetToObjects_(
    appGetSheet_(ss, APP_CONFIG.SHEETS.CLINICAL_TRACKING)
  );

  const studentRows = students.filter(
    r => appNormalize_(r['Student ID']) === appNormalize_(studentId)
  );
  if (!studentRows.length) throw new Error('Student not found: ' + studentId);

  const student = clinicalHistory_mostAdvancedStudentRow_(studentRows);
  const studentName = clinicalHistory_text_(student['Student']);
  const resolvedStudentId = clinicalHistory_text_(student['Student ID']);
  const currentClinicalTerm = clinicalHistory_text_(student['Clinical Term']);

  const studentTracking = tracking.filter(
    r => clinicalHistory_sameStudent_(r, resolvedStudentId, studentName)
  );

  let gradeSummary = { terms: [] };
  if (typeof gradeEngine_getStudentSummary === 'function') {
    gradeSummary = gradeEngine_getStudentSummary(resolvedStudentId) || { terms: [] };
  }

  const gradeRecords = clinicalHistory_getGradeRecords_(
    ss, resolvedStudentId, studentName
  );

  const cards = [];
  for (let n = 1; n <= 5; n++) {
    cards.push(clinicalHistory_buildTermCard_({
      clinicalTerm: 'Clinical ' + n,
      currentClinicalTerm: currentClinicalTerm,
      tracking: studentTracking,
      gradeRecords: gradeRecords,
      gradeTerms: gradeSummary.terms || []
    }));
  }

  return {
    student: {
      studentId: resolvedStudentId,
      student: studentName,
      cohort: clinicalHistory_text_(student['Cohort']),
      currentClinicalTerm: currentClinicalTerm
    },
    description:
      'Clinical History is a semester-by-semester academic record showing grades, ' +
      'clinical evaluations, supplemental grade items, and documented carryover/banked activity.',
    terms: cards
  };
}

function clinicalHistory_buildTermCard_(args) {
  const clinicalTerm = args.clinicalTerm;
  const termNumber = clinicalHistory_termNumber_(clinicalTerm);
  const currentNumber = clinicalHistory_termNumber_(args.currentClinicalTerm);

  const ownedTracking = (args.tracking || []).filter(
    r => appNormalize_(clinicalHistory_ownedTerm_(r)) === appNormalize_(clinicalTerm)
  );

  const occurredTracking = (args.tracking || []).filter(
    r => appNormalize_(clinicalHistory_occurredTerm_(r)) === appNormalize_(clinicalTerm)
  );

  const bankedForFuture = occurredTracking
    .filter(r => {
      const owned = clinicalHistory_termNumber_(clinicalHistory_ownedTerm_(r));
      return owned > termNumber;
    })
    .map(clinicalHistory_trackingItem_);

  const carryoverIntoTerm = ownedTracking
    .filter(r => {
      const occurred = clinicalHistory_termNumber_(clinicalHistory_occurredTerm_(r));
      return occurred > 0 && occurred < termNumber;
    })
    .map(clinicalHistory_trackingItem_);

  const termGradeRecords = (args.gradeRecords || [])
    .filter(r => appNormalize_(r['Clinical Term']) === appNormalize_(clinicalTerm))
    .map(clinicalHistory_gradeRecordItem_)
    .sort((a,b) => b.dateValue - a.dateValue);

  const gradeTerm = (args.gradeTerms || []).find(
    t => appNormalize_(t.clinicalTerm) === appNormalize_(clinicalTerm)
  ) || null;

  const evaluations = ownedTracking
    .map(clinicalHistory_trackingItem_)
    .sort((a,b) => b.dateValue - a.dateValue);

  const competency = evaluations.filter(
    x => appNormalize_(x.evaluationType) === 'competency'
  );
  const iso = evaluations.filter(
    x => appNormalize_(x.evaluationType) === 'iso'
  );
  const sso = evaluations.filter(
    x => appNormalize_(x.evaluationType) === 'sso'
  );
  const simulation = evaluations.filter(
    x => appNormalize_(x.evaluationType) === 'simulation'
  );

  let status = 'Upcoming';
  if (currentNumber > 0 && termNumber < currentNumber) status = 'Completed';
  if (currentNumber > 0 && termNumber === currentNumber) status = 'Current';

  return {
    clinicalTerm: clinicalTerm,
    status: status,
    semesterAverage: clinicalHistory_gradeSummary_(gradeTerm),
    counts: {
      evaluations: evaluations.length,
      competencies: competency.length,
      passedCompetencies: competency.filter(
        x => appNormalize_(x.result) === 'pass'
      ).length,
      failedCompetencies: competency.filter(
        x => appNormalize_(x.result) === 'fail'
      ).length,
      iso: iso.length,
      sso: sso.length,
      simulation: simulation.length,
      supplementalGradeItems: termGradeRecords.length,
      carryoverIntoTerm: carryoverIntoTerm.length,
      bankedForFuture: bankedForFuture.length
    },
    evaluations: evaluations,
    supplementalGrades: termGradeRecords,
    carryoverIntoTerm: carryoverIntoTerm,
    bankedForFuture: bankedForFuture
  };
}

function clinicalHistory_gradeSummary_(term) {
  if (!term) {
    return {
      available: false,
      value: null,
      display: '—',
      complete: false,
      partial: false,
      note: 'No grade-bearing records are currently available for this semester.',
      includedCategories: [],
      missingCategories: []
    };
  }

  const value = clinicalHistory_numberOrNull_(term.recordedAverage);

  return {
    available: value !== null,
    value: value,
    display: value === null
      ? '—'
      : clinicalHistory_trimNumber_(value) + '%',
    complete: Boolean(term.complete),
    partial: value !== null && !term.complete,
    availableWeight: Number(term.availableWeight || 0),
    note: clinicalHistory_text_(term.note),
    includedCategories: term.includedCategories || [],
    missingCategories: term.missingCategories || [],
    categories: term.categories || {},
    deductions: term.deductions || {}
  };
}

function clinicalHistory_trackingItem_(record) {
  const resolved = clinicalHistory_resolveTrackingGrade_(record);

  return {
    submissionId: clinicalHistory_text_(record['Submission ID']),
    academicTerm: clinicalHistory_ownedTerm_(record),
    occurredTerm: clinicalHistory_occurredTerm_(record),
    evaluationType: clinicalHistory_text_(record['Evaluation Type']),
    examination: clinicalHistory_text_(record['Examination']),
    anatomicalArea: clinicalHistory_text_(record['Anatomical Area']),
    date: clinicalHistory_displayDate_(record['Date']),
    dateValue: clinicalHistory_dateValue_(record['Date']),
    evaluator: clinicalHistory_text_(record['Evaluator']),
    evaluatorType: clinicalHistory_text_(record['Evaluator Type']),
    clinicalSite: clinicalHistory_text_(record['Clinical Site']),
    result: clinicalHistory_text_(record['Result']),
    rawScore: resolved.rawScore,
    percentScore: resolved.percent,
    gradeSource: resolved.source,
    gradeDisplay: clinicalHistory_trackingGradeDisplay_(resolved, record),
    automaticFailure: clinicalHistory_true_(record['Automatic Failure']),
    sourceDocument: clinicalHistory_text_(record['Source Document'])
  };
}

function clinicalHistory_gradeRecordItem_(record) {
  return {
    gradeRecordId: clinicalHistory_text_(record['Grade Record ID']),
    clinicalTerm: clinicalHistory_text_(record['Clinical Term']),
    gradeCategory: clinicalHistory_text_(record['Grade Category']),
    assignmentName: clinicalHistory_text_(record['Assignment Name']),
    date: clinicalHistory_displayDate_(record['Date']),
    dateValue: clinicalHistory_dateValue_(record['Date']),
    instructor: clinicalHistory_text_(record['Instructor']),
    percentGrade: clinicalHistory_numberOrNull_(record['Percent Grade']),
    points: clinicalHistory_numberOrNull_(record['Points']),
    pointsPossible: clinicalHistory_numberOrNull_(record['Points Possible']),
    notes: clinicalHistory_text_(record['Notes'])
  };
}

function clinicalHistory_getGradeRecords_(ss, studentId, studentName) {
  let rows = [];

  if (typeof gradeEngine_getSupplementalGradeRecords_ === 'function') {
    rows = gradeEngine_getSupplementalGradeRecords_(ss) || [];
  } else {
    const sheet = ss.getSheetByName('Grade Records');
    rows = sheet ? clinicalHistory_sheetObjects_(sheet) : [];
  }

  return rows.filter(r => {
    if (
      Object.prototype.hasOwnProperty.call(r, 'Active') &&
      !clinicalHistory_true_(r['Active'])
    ) return false;

    return clinicalHistory_sameStudent_(r, studentId, studentName);
  });
}

/**
 * Clinical Term = academic ownership.
 * Attributed Clinical Term = physical occurrence.
 */
function clinicalHistory_ownedTerm_(record) {
  return clinicalHistory_text_(record['Clinical Term']);
}

function clinicalHistory_occurredTerm_(record) {
  return clinicalHistory_text_(
    record['Attributed Clinical Term'] || record['Clinical Term']
  );
}

function clinicalHistory_resolveTrackingGrade_(record) {
  if (typeof gradeEngine_resolveTrackingGrade_ === 'function') {
    const r = gradeEngine_resolveTrackingGrade_(record) || {};
    return {
      rawScore: clinicalHistory_numberOrNull_(r.rawScore),
      percent: clinicalHistory_numberOrNull_(r.percent),
      source: clinicalHistory_text_(r.source)
    };
  }

  const raw = clinicalHistory_numberOrNull_(record['Final Score']);
  if (raw !== null && typeof gradeEngine_scoreToPercent === 'function') {
    return {
      rawScore: raw,
      percent: clinicalHistory_numberOrNull_(gradeEngine_scoreToPercent(raw)),
      source: 'Final Score'
    };
  }

  return { rawScore: raw, percent: null, source: '' };
}

function clinicalHistory_trackingGradeDisplay_(resolved, record) {
  if (resolved.percent !== null) {
    const pct = clinicalHistory_trimNumber_(resolved.percent) + '%';
    if (resolved.rawScore !== null && resolved.source === 'Final Score') {
      return clinicalHistory_trimNumber_(resolved.rawScore) + ' • ' + pct;
    }
    return pct;
  }
  return clinicalHistory_text_(record['Result']) || '—';
}

function clinicalHistory_mostAdvancedStudentRow_(rows) {
  return rows.slice().sort(
    (a,b) =>
      clinicalHistory_termNumber_(b['Clinical Term']) -
      clinicalHistory_termNumber_(a['Clinical Term'])
  )[0];
}

function clinicalHistory_termNumber_(value) {
  const m = clinicalHistory_text_(value).match(/clinical\s*(\d+)/i);
  return m ? Number(m[1]) : 0;
}

function clinicalHistory_sameStudent_(record, studentId, studentName) {
  const recordId = appNormalize_(record['Student ID']);
  const targetId = appNormalize_(studentId);
  if (recordId && targetId) return recordId === targetId;
  return appNormalize_(record['Student']) === appNormalize_(studentName);
}

function clinicalHistory_sheetObjects_(sheet) {
  if (!sheet || sheet.getLastRow() < 2) return [];
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(clinicalHistory_text_);
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h,i) => { if (h) obj[h] = row[i]; });
    return obj;
  });
}

function clinicalHistory_true_(value) {
  if (typeof appIsTrue_ === 'function') return appIsTrue_(value);
  if (value === true || value === 1) return true;
  const v = appNormalize_(value);
  return v === 'true' || v === 'yes' || v === '1';
}

function clinicalHistory_text_(value) {
  return value === null || value === undefined ? '' : String(value).trim();
}

function clinicalHistory_numberOrNull_(value) {
  if (value === '' || value === null || value === undefined) return null;
  const n = Number(value);
  return Number.isFinite(n) ? n : null;
}

function clinicalHistory_round2_(value) {
  return Math.round((Number(value) + Number.EPSILON) * 100) / 100;
}

function clinicalHistory_trimNumber_(value) {
  return clinicalHistory_round2_(value).toFixed(2).replace(/\.?0+$/,'');
}

function clinicalHistory_dateValue_(value) {
  if (!value) return 0;
  const d = value instanceof Date ? value : new Date(value);
  return Number.isNaN(d.getTime()) ? 0 : d.getTime();
}

function clinicalHistory_displayDate_(value) {
  const time = clinicalHistory_dateValue_(value);
  if (!time) return '';
  return Utilities.formatDate(
    new Date(time),
    Session.getScriptTimeZone(),
    'MM/dd/yyyy'
  );
}

function testStudentClinicalHistoryArchitecture() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const students = appSheetToObjects_(
    appGetSheet_(ss, APP_CONFIG.SHEETS.STUDENTS)
  );
  const testStudent = students.find(
    r => appNormalize_(r['Student ID']) === 'test-001'
  );
  if (!testStudent) throw new Error('TEST-001 was not found.');

  const result = clinicalHistory_getStudent(testStudent['Student ID']);
  const compact = {
    success: true,
    readOnly: true,
    student: result.student,
    terms: result.terms.map(t => ({
      clinicalTerm: t.clinicalTerm,
      status: t.status,
      semesterAverage: t.semesterAverage.display,
      evaluations: t.counts.evaluations,
      competencies: t.counts.competencies,
      iso: t.counts.iso,
      carryoverIntoTerm: t.counts.carryoverIntoTerm,
      bankedForFuture: t.counts.bankedForFuture,
      supplementalGradeItems: t.counts.supplementalGradeItems
    }))
  };

  Logger.log(JSON.stringify(compact, null, 2));
  return compact;
}
