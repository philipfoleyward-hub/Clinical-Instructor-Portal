/**
 * =====================================================================
 * CLINICAL COMPETENCY TRACKER
 * ATTENDANCE SERVICE
 * AttendanceService.gs
 * =====================================================================
 *
 * PHASE 2B PURPOSE
 * ----------------
 * Authoritative attendance calculation service and controlled Attendance
 * Events writer.
 *
 * Authority:
 *   - Attendance Rules = effective-dated attendance policy configuration
 *   - Attendance Events = factual attendance exceptions / credits
 *   - SchedulingService = expected clinical schedule
 *   - Students = stable student identity
 *
 * This service DOES NOT:
 *   - silently alter or delete historical Attendance Events
 *   - award Earned Time outside an explicit coordinator event
 *   - modify grades / Grade Records
 *   - infer absence from a missing Attendance Event
 *   - modify Scheduling data
 *   - automatically dismiss or progress a student
 *
 * Phase 2A calculations:
 *   - base ETO allocation
 *   - earned time
 *   - ETO-eligible deductions
 *   - current available attendance balance
 *   - ordinary make-up required/completed/outstanding
 *   - evening-rotation replacement required/completed/outstanding
 *   - maximum-absence review flag
 *
 * IMPORTANT
 * ---------
 * 22.5 hours is an INITIAL ALLOCATION, not a hard ceiling. Approved
 * Earned Time may increase the calculated available balance above it.
 * =====================================================================
 */


/**
 * Public: returns the authoritative attendance rule for a clinical term/date.
 */
function attendance_getRule(clinicalTerm, asOfDate) {
  const term = attendance_normalizeTerm_(clinicalTerm);
  if (!term) {
    throw new Error('Clinical Term is required.');
  }

  const date = asOfDate ? attendance_toDate_(asOfDate) : new Date();
  const rules = attendance_readRules_();

  const matches = rules.filter(rule => {
    if (!rule.active) return false;
    if (attendance_normalizeTerm_(rule.clinicalTerm) !== term) return false;
    if (rule.effectiveStartDate && date < rule.effectiveStartDate) return false;
    if (rule.effectiveEndDate && date > rule.effectiveEndDate) return false;
    return true;
  });

  if (matches.length === 0) {
    throw new Error(
      'No active Attendance Rule was found for ' + clinicalTerm +
      ' on ' + attendance_formatDate_(date) + '.'
    );
  }

  if (matches.length > 1) {
    throw new Error(
      'Multiple active Attendance Rules apply to ' + clinicalTerm +
      ' on ' + attendance_formatDate_(date) + '. Resolve the rule overlap.'
    );
  }

  return attendance_publicRule_(matches[0]);
}


/**
 * Public: returns authoritative Attendance Events for one student/term.
 *
 * Rejected records and records superseded by another record are excluded
 * from authoritative calculations. Pending/unapproved records are returned
 * separately by attendance_getStudentTermSummary and do not affect balances.
 */
function attendance_getStudentEvents(studentId, clinicalTerm) {
  const id = attendance_text_(studentId);
  const term = attendance_normalizeTerm_(clinicalTerm);

  if (!id) throw new Error('Student ID is required.');
  if (!term) throw new Error('Clinical Term is required.');

  const all = attendance_readEvents_()
    .filter(event =>
      attendance_text_(event.studentId) === id &&
      attendance_normalizeTerm_(event.clinicalTerm) === term
    );

  const supersededIds = new Set(
    all
      .map(event => attendance_text_(event.supersedesEventId))
      .filter(Boolean)
  );

  return all
    .filter(event => !supersededIds.has(attendance_text_(event.attendanceEventId)))
    .map(attendance_publicEvent_)
    .sort((a, b) => {
      const ad = a.date ? new Date(a.date).getTime() : 0;
      const bd = b.date ? new Date(b.date).getTime() : 0;
      return ad - bd;
    });
}


/**
 * Public: authoritative student/term attendance summary.
 *
 * Approval rule for Phase 2A:
 *   - APPROVED events affect authoritative balances.
 *   - REJECTED events do not.
 *   - blank / PENDING / other statuses are returned as pending and do not
 *     affect authoritative balances until approved.
 *
 * This deliberately prevents unapproved event data from silently changing
 * progression-relevant attendance calculations.
 */
function attendance_getStudentTermSummary(studentId, clinicalTerm, asOfDate) {
  const id = attendance_text_(studentId);
  const term = attendance_normalizeTerm_(clinicalTerm);

  if (!id) throw new Error('Student ID is required.');
  if (!term) throw new Error('Clinical Term is required.');

  const student = attendance_getStudent_(id);
  const events = attendance_getStudentEvents(id, term);

  const approved = events.filter(event =>
    attendance_normalizeStatus_(event.approvalStatus) === 'APPROVED'
  );

  const rejected = events.filter(event =>
    attendance_normalizeStatus_(event.approvalStatus) === 'REJECTED'
  );

  const pending = events.filter(event => {
    const status = attendance_normalizeStatus_(event.approvalStatus);
    return status !== 'APPROVED' && status !== 'REJECTED';
  });

  const ruleDate = asOfDate
    ? attendance_toDate_(asOfDate)
    : attendance_ruleReferenceDate_(approved, new Date());

  const rule = attendance_getRule(term, ruleDate);
  const calculation = attendance_calculateFromEvents_(rule, approved);

  const scheduled = attendance_getScheduledHours_(student, term);
  const maximumAbsence = attendance_calculateMaximumAbsence_(
    calculation.totalMissedHours,
    scheduled.scheduledHours,
    rule.maximumAbsencePercentage
  );

  return {
    success: true,
    studentId: id,
    studentName: student.studentName,
    clinicalTerm: rule.clinicalTerm,
    rule: rule,
    balances: calculation.balances,
    makeUp: calculation.makeUp,
    eveningRotation: calculation.eveningRotation,
    attendance: {
      totalMissedHours: calculation.totalMissedHours,
      etoEligibleMissedHours: calculation.etoEligibleMissedHours,
      etoIneligibleMissedHours: calculation.etoIneligibleMissedHours,
      earnedTimeHours: calculation.earnedTimeHours,
      scheduledHours: scheduled.scheduledHours,
      scheduledHoursSource: scheduled.source,
      maximumAbsencePercentage: rule.maximumAbsencePercentage,
      absencePercentage: maximumAbsence.absencePercentage,
      maximumAbsenceReviewRequired: maximumAbsence.reviewRequired,
      maximumAbsenceReviewReason: maximumAbsence.reason
    },
    eventCounts: {
      approved: approved.length,
      pending: pending.length,
      rejected: rejected.length,
      totalCurrent: events.length
    },
    approvedEvents: approved,
    pendingEvents: pending,
    rejectedEvents: rejected,
    warnings: calculation.warnings.concat(scheduled.warnings)
  };
}


/**
 * Public: cohort summaries for a term. Read only.
 */
function attendance_getCohortTermSummary(cohort, clinicalTerm, asOfDate) {
  const students = attendance_getActiveStudentsForCohort_(cohort);
  return {
    success: true,
    cohort: cohort,
    clinicalTerm: clinicalTerm,
    studentCount: students.length,
    students: students.map(student =>
      attendance_getStudentTermSummary(
        student.studentId,
        clinicalTerm,
        asOfDate
      )
    )
  };
}


/**
 * Public: validates Attendance Rules and Attendance Events without writing.
 */

/**
 * Public: controlled writer for one Attendance Event.
 *
 * Coordinator-entered records default to APPROVED so that an explicit
 * coordinator decision is reflected immediately. Historical rows are never
 * overwritten. Corrections use attendance_amendEvent().
 */
function attendance_createEvent(input) {
  input = input || {};

  const studentId = attendance_text_(input.studentId);
  const clinicalTerm = attendance_normalizeTerm_(input.clinicalTerm);
  const eventType = attendance_normalizeEventType_(input.eventType);
  const eventDate = attendance_optionalDate_(input.date);
  const hours = attendance_number_(input.hours, null);
  const etoImpact = attendance_number_(input.etoImpact, null);
  const makeUpRequired = attendance_bool_(input.makeUpRequired);
  const relatedEventId = attendance_text_(input.relatedEventId);
  const notes = attendance_text_(input.notes);
  const reason = attendance_text_(input.reason);
  const approvalStatus = attendance_normalizeStatus_(
    input.approvalStatus || 'APPROVED'
  );
  const source = attendance_text_(input.source || 'CCMS Dashboard');

  if (!studentId) throw new Error('Student ID is required.');
  if (!clinicalTerm) throw new Error('Clinical Term is required.');
  if (!eventDate) throw new Error('A valid attendance date is required.');

  const allowedTypes = [
    'ABSENCE',
    'LATE_ARRIVAL',
    'EARLY_DEPARTURE',
    'LATE_RETURN',
    'DISCIPLINARY_REMOVAL',
    'EARNED_TIME',
    'MAKE_UP_TIME',
    'EVENING_ROTATION_REPLACEMENT',
    'ATTENDANCE_ADJUSTMENT'
  ];

  if (allowedTypes.indexOf(eventType) === -1) {
    throw new Error(
      'Unsupported Attendance Event Type: ' + (input.eventType || '') + '.'
    );
  }

  if (hours === null || hours < 0) {
    throw new Error('Hours must be a number greater than or equal to 0.');
  }

  if (['APPROVED','PENDING','REJECTED'].indexOf(approvalStatus) === -1) {
    throw new Error('Approval Status must be APPROVED, PENDING, or REJECTED.');
  }

  // Validate the student and the effective attendance rule before writing.
  attendance_getStudent_(studentId);
  attendance_getRule(clinicalTerm, eventDate);

  if (relatedEventId) {
    const related = attendance_readEvents_().find(event =>
      attendance_text_(event.attendanceEventId) === relatedEventId
    );
    if (!related) {
      throw new Error('Related Attendance Event ID was not found: ' + relatedEventId);
    }
    if (attendance_text_(related.studentId) !== studentId) {
      throw new Error('Related Attendance Event belongs to a different student.');
    }
  }

  const now = new Date();
  const actor = attendance_currentActor_();
  const eventId = attendance_newEventId_(eventDate);

  const record = {
    'Attendance Event ID': eventId,
    'Student ID': studentId,
    'Date': eventDate,
    'Clinical Term': attendance_publicTerm_(clinicalTerm),
    'Event Type': eventType,
    'Hours': hours,
    'ETO Impact': etoImpact === null ? '' : etoImpact,
    'Make-Up Required': makeUpRequired,
    'Related Event ID': relatedEventId,
    'Source': source,
    'Approval Status': approvalStatus,
    'Approved By': approvalStatus === 'APPROVED' ? actor : '',
    'Approved At': approvalStatus === 'APPROVED' ? now : '',
    'Notes': notes,
    'Created At': now,
    'Created By': actor,
    'Supersedes Event ID': '',
    'Amendment Reason': '',
    'Reason': reason
  };

  attendance_appendEventRecord_(record);

  return {
    success: true,
    attendanceEventId: eventId,
    event: attendance_getStudentEvents(studentId, clinicalTerm)
      .find(event => event.attendanceEventId === eventId) || null,
    summary: attendance_getStudentTermSummary(studentId, clinicalTerm, eventDate)
  };
}


/**
 * Public: correct an existing Attendance Event without overwriting history.
 * The new row supersedes the original row.
 */
function attendance_amendEvent(attendanceEventId, changes, amendmentReason) {
  const originalId = attendance_text_(attendanceEventId);
  const reasonText = attendance_text_(amendmentReason);
  const all = attendance_readEvents_();
  const original = all.find(event => event.attendanceEventId === originalId);

  if (!original) throw new Error('Attendance Event not found: ' + originalId);
  if (!reasonText) throw new Error('Amendment Reason is required.');

  const input = Object.assign({
    studentId: original.studentId,
    date: original.date,
    clinicalTerm: original.clinicalTerm,
    eventType: original.eventType,
    hours: original.hours,
    etoImpact: original.etoImpact,
    makeUpRequired: original.makeUpRequired,
    relatedEventId: original.relatedEventId,
    source: 'CCMS Dashboard Amendment',
    approvalStatus: original.approvalStatus || 'APPROVED',
    notes: original.notes,
    reason: original.reason
  }, changes || {});

  const studentId = attendance_text_(input.studentId);
  if (studentId !== attendance_text_(original.studentId)) {
    throw new Error('An amendment cannot move an event to another student.');
  }

  const clinicalTerm = attendance_normalizeTerm_(input.clinicalTerm);
  const eventType = attendance_normalizeEventType_(input.eventType);
  const eventDate = attendance_optionalDate_(input.date);
  const hours = attendance_number_(input.hours, null);
  const etoImpact = attendance_number_(input.etoImpact, null);
  const approvalStatus = attendance_normalizeStatus_(input.approvalStatus || 'APPROVED');

  if (!eventDate) throw new Error('A valid attendance date is required.');
  if (hours === null || hours < 0) throw new Error('Hours must be >= 0.');

  const allowedTypes = [
    'ABSENCE','LATE_ARRIVAL','EARLY_DEPARTURE','LATE_RETURN',
    'DISCIPLINARY_REMOVAL','EARNED_TIME','MAKE_UP_TIME',
    'EVENING_ROTATION_REPLACEMENT','ATTENDANCE_ADJUSTMENT'
  ];
  if (allowedTypes.indexOf(eventType) === -1) {
    throw new Error('Unsupported Attendance Event Type: ' + eventType + '.');
  }

  attendance_getRule(clinicalTerm, eventDate);

  const now = new Date();
  const actor = attendance_currentActor_();
  const newId = attendance_newEventId_(eventDate);

  attendance_appendEventRecord_({
    'Attendance Event ID': newId,
    'Student ID': studentId,
    'Date': eventDate,
    'Clinical Term': attendance_publicTerm_(clinicalTerm),
    'Event Type': eventType,
    'Hours': hours,
    'ETO Impact': etoImpact === null ? '' : etoImpact,
    'Make-Up Required': attendance_bool_(input.makeUpRequired),
    'Related Event ID': attendance_text_(input.relatedEventId),
    'Source': attendance_text_(input.source || 'CCMS Dashboard Amendment'),
    'Approval Status': approvalStatus,
    'Approved By': approvalStatus === 'APPROVED' ? actor : '',
    'Approved At': approvalStatus === 'APPROVED' ? now : '',
    'Notes': attendance_text_(input.notes),
    'Created At': now,
    'Created By': actor,
    'Supersedes Event ID': originalId,
    'Amendment Reason': reasonText,
    'Reason': attendance_text_(input.reason)
  });

  return {
    success: true,
    attendanceEventId: newId,
    supersedesEventId: originalId,
    summary: attendance_getStudentTermSummary(studentId, clinicalTerm, eventDate)
  };
}


function attendance_appendEventRecord_(record) {
  const sheet = attendance_requireSheet_(APP_CONFIG.SHEETS.ATTENDANCE_EVENTS);
  const required = [
    'Attendance Event ID','Student ID','Date','Clinical Term','Event Type',
    'Hours','ETO Impact','Make-Up Required','Related Event ID','Source',
    'Approval Status','Approved By','Approved At','Notes','Created At',
    'Created By','Supersedes Event ID','Amendment Reason','Reason'
  ];

  if (sheet.getLastColumn() < 1) {
    throw new Error('Attendance Events has no headers.');
  }

  const headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0]
    .map(value => attendance_text_(value));

  const missing = required.filter(header => headers.indexOf(header) === -1);
  if (missing.length) {
    throw new Error(
      'Attendance Events is missing required header(s): ' + missing.join(', ')
    );
  }

  const row = headers.map(header =>
    Object.prototype.hasOwnProperty.call(record, header) ? record[header] : ''
  );

  sheet.appendRow(row);
}


function attendance_newEventId_(date) {
  const stamp = Utilities.formatDate(
    attendance_toDate_(date),
    Session.getScriptTimeZone(),
    'yyyyMMdd'
  );
  return 'ATT-' + stamp + '-' +
    Utilities.getUuid().replace(/-/g,'').slice(0,12).toUpperCase();
}


function attendance_currentActor_() {
  try {
    return Session.getActiveUser().getEmail() ||
      Session.getEffectiveUser().getEmail() ||
      'CCMS Coordinator';
  } catch (error) {
    return 'CCMS Coordinator';
  }
}


function attendance_publicTerm_(normalizedTerm) {
  const match = attendance_text_(normalizedTerm).match(/CLINICAL\s*([1-5])/i);
  return match ? 'Clinical ' + match[1] : attendance_text_(normalizedTerm);
}


function attendance_validateData() {
  const issues = [];
  const warnings = [];
  const rules = attendance_readRules_();
  const events = attendance_readEvents_();

  const activeRules = rules.filter(rule => rule.active);

  ['Clinical 1', 'Clinical 2', 'Clinical 3', 'Clinical 4', 'Clinical 5']
    .forEach(term => {
      const termRules = activeRules.filter(rule =>
        attendance_normalizeTerm_(rule.clinicalTerm) ===
        attendance_normalizeTerm_(term)
      );
      if (termRules.length === 0) {
        issues.push('No active Attendance Rule exists for ' + term + '.');
      }
    });

  for (let i = 0; i < activeRules.length; i++) {
    const a = activeRules[i];

    if (!(a.initialEtoHours >= 0)) {
      issues.push('Attendance Rule row ' + a.rowNumber + ': Initial ETO Hours must be >= 0.');
    }
    if (!(a.clinicalDayHours > 0)) {
      issues.push('Attendance Rule row ' + a.rowNumber + ': Clinical Day Hours must be > 0.');
    }
    if (!(a.etoIncrementHours > 0)) {
      issues.push('Attendance Rule row ' + a.rowNumber + ': ETO Increment Hours must be > 0.');
    }
    if (a.maximumAbsencePercentage !== null &&
        (a.maximumAbsencePercentage <= 0 || a.maximumAbsencePercentage > 100)) {
      issues.push('Attendance Rule row ' + a.rowNumber + ': Maximum Absence Percentage must be > 0 and <= 100.');
    }

    for (let j = i + 1; j < activeRules.length; j++) {
      const b = activeRules[j];
      if (attendance_normalizeTerm_(a.clinicalTerm) !== attendance_normalizeTerm_(b.clinicalTerm)) continue;
      if (attendance_rangesOverlap_(a.effectiveStartDate, a.effectiveEndDate, b.effectiveStartDate, b.effectiveEndDate)) {
        issues.push(
          'Attendance Rules rows ' + a.rowNumber + ' and ' + b.rowNumber +
          ' overlap for ' + a.clinicalTerm + '.'
        );
      }
    }
  }

  const eventIds = new Set();
  events.forEach(event => {
    const id = attendance_text_(event.attendanceEventId);
    if (!id) {
      issues.push('Attendance Events row ' + event.rowNumber + ': Attendance Event ID is required.');
    } else if (eventIds.has(id)) {
      issues.push('Duplicate Attendance Event ID: ' + id + '.');
    } else {
      eventIds.add(id);
    }

    if (!attendance_text_(event.studentId)) {
      issues.push('Attendance Events row ' + event.rowNumber + ': Student ID is required.');
    }
    if (!attendance_normalizeTerm_(event.clinicalTerm)) {
      issues.push('Attendance Events row ' + event.rowNumber + ': Clinical Term is required.');
    }
    if (!attendance_text_(event.eventType)) {
      issues.push('Attendance Events row ' + event.rowNumber + ': Event Type is required.');
    }
    if (event.hours !== null && event.hours < 0) {
      issues.push('Attendance Events row ' + event.rowNumber + ': Hours must be stored as a positive magnitude.');
    }
    if (event.supersedesEventId && !eventIds.has(event.supersedesEventId)) {
      warnings.push(
        'Attendance Events row ' + event.rowNumber +
        ': Supersedes Event ID ' + event.supersedesEventId +
        ' was not found earlier in the sheet.'
      );
    }
  });

  return {
    success: issues.length === 0,
    issueCount: issues.length,
    warningCount: warnings.length,
    issues: issues,
    warnings: warnings,
    counts: {
      attendanceRules: rules.length,
      activeAttendanceRules: activeRules.length,
      attendanceEvents: events.length
    }
  };
}


/**
 * PURE CALCULATION CORE.
 * Exposed with underscore for QA only; production callers should use public
 * summary functions above.
 */
function attendance_calculateFromEvents_(rule, approvedEvents) {
  const warnings = [];

  let earnedTimeHours = 0;
  let etoEligibleMissedHours = 0;
  let etoIneligibleMissedHours = 0;
  let totalMissedHours = 0;
  let explicitMakeUpRequiredHours = 0;
  let makeUpCompletedHours = 0;
  let eveningRequiredCount = 0;
  let eveningCompletedCount = 0;

  approvedEvents.forEach(event => {
    const type = attendance_normalizeEventType_(event.eventType);
    const hours = attendance_number_(event.hours, 0);
    const etoImpact = event.etoImpact === null || event.etoImpact === ''
      ? null
      : Number(event.etoImpact);
    const makeUpRequired = attendance_bool_(event.makeUpRequired);

    switch (type) {
      case 'EARNED_TIME':
        earnedTimeHours += attendance_positiveCredit_(event, hours, etoImpact, warnings);
        break;

      case 'MAKE_UP_TIME':
        makeUpCompletedHours += hours;
        break;

      case 'EVENING_ROTATION_REPLACEMENT':
        eveningCompletedCount += 1;
        break;

      case 'ABSENCE':
      case 'LATE_ARRIVAL':
      case 'EARLY_DEPARTURE':
      case 'LATE_RETURN':
      case 'DISCIPLINARY_REMOVAL':
        totalMissedHours += hours;

        if (type === 'DISCIPLINARY_REMOVAL' && (etoImpact === null || etoImpact === 0)) {
          etoIneligibleMissedHours += hours;
        } else {
          const deduction = attendance_missedDeduction_(event, hours, etoImpact, warnings);
          etoEligibleMissedHours += deduction;
        }

        if (makeUpRequired) {
          explicitMakeUpRequiredHours += hours;
        }

        if (attendance_eventRequiresEveningReplacement_(event, rule)) {
          eveningRequiredCount += 1;
        }
        break;

      case 'ATTENDANCE_ADJUSTMENT':
        if (etoImpact === null) {
          warnings.push(
            'Attendance Adjustment ' + event.attendanceEventId +
            ' has no ETO Impact and therefore does not change the balance.'
          );
        } else if (etoImpact > 0) {
          earnedTimeHours += etoImpact;
        } else if (etoImpact < 0) {
          etoEligibleMissedHours += Math.abs(etoImpact);
        }
        if (makeUpRequired && hours > 0) {
          explicitMakeUpRequiredHours += hours;
        }
        break;

      default:
        warnings.push(
          'Unrecognized Attendance Event Type "' + event.eventType +
          '" for event ' + event.attendanceEventId +
          '; it was excluded from attendance calculations.'
        );
    }
  });

  const base = attendance_number_(rule.initialEtoHours, 0);
  const availableBeforeMissed = base + earnedTimeHours;
  const currentAvailableBalance = availableBeforeMissed - etoEligibleMissedHours;

  const excessEtoDeficit = rule.excessEtoRequiresMakeUp
    ? Math.max(0, etoEligibleMissedHours - availableBeforeMissed)
    : 0;

  // ETO-ineligible missed time is always an obligation. Explicit make-up
  // requirements may identify additional obligations. Use max rather than sum
  // to avoid double-counting the same missed hours when both conditions apply.
  const ordinaryRequiredHours = Math.max(
    excessEtoDeficit + etoIneligibleMissedHours,
    explicitMakeUpRequiredHours
  );

  const ordinaryOutstandingHours = Math.max(
    0,
    ordinaryRequiredHours - makeUpCompletedHours
  );

  const eveningOutstandingCount = Math.max(
    0,
    eveningRequiredCount - eveningCompletedCount
  );

  return {
    balances: {
      baseEtoAllocationHours: attendance_round_(base),
      earnedTimeHours: attendance_round_(earnedTimeHours),
      availableBeforeMissedHours: attendance_round_(availableBeforeMissed),
      etoEligibleMissedHours: attendance_round_(etoEligibleMissedHours),
      currentAvailableBalanceHours: attendance_round_(currentAvailableBalance)
    },
    makeUp: {
      requiredHours: attendance_round_(ordinaryRequiredHours),
      completedHours: attendance_round_(makeUpCompletedHours),
      outstandingHours: attendance_round_(ordinaryOutstandingHours)
    },
    eveningRotation: {
      requiredReplacementCount: eveningRequiredCount,
      completedReplacementCount: eveningCompletedCount,
      outstandingReplacementCount: eveningOutstandingCount,
      replacementRequired: eveningOutstandingCount > 0,
      startTime: rule.eveningRotationStart,
      endTime: rule.eveningRotationEnd
    },
    totalMissedHours: attendance_round_(totalMissedHours),
    etoEligibleMissedHours: attendance_round_(etoEligibleMissedHours),
    etoIneligibleMissedHours: attendance_round_(etoIneligibleMissedHours),
    earnedTimeHours: attendance_round_(earnedTimeHours),
    warnings: warnings
  };
}


function attendance_eventRequiresEveningReplacement_(event, rule) {
  if (!rule.eveningRotationRequired || !rule.eveningAbsenceRequiresFullMakeUp) {
    return false;
  }

  const type = attendance_normalizeEventType_(event.eventType);
  if (!['ABSENCE', 'LATE_ARRIVAL', 'EARLY_DEPARTURE'].includes(type)) {
    return false;
  }

  // Phase 2A does not infer an evening assignment from free text. The event
  // must explicitly identify the evening requirement through Reason or Notes.
  const context = (
    attendance_text_(event.reason) + ' ' + attendance_text_(event.notes)
  ).toUpperCase();

  return context.indexOf('EVENING') !== -1;
}


function attendance_positiveCredit_(event, hours, etoImpact, warnings) {
  if (etoImpact !== null) {
    if (etoImpact < 0) {
      warnings.push(
        'Earned Time event ' + event.attendanceEventId +
        ' has negative ETO Impact; Hours was used as the positive credit instead.'
      );
      return hours;
    }
    return etoImpact;
  }
  return hours;
}


function attendance_missedDeduction_(event, hours, etoImpact, warnings) {
  if (etoImpact === null) return hours;
  if (etoImpact > 0) {
    warnings.push(
      'Missed-time event ' + event.attendanceEventId +
      ' has positive ETO Impact; Hours was used as the deduction instead.'
    );
    return hours;
  }
  return Math.abs(etoImpact);
}


function attendance_calculateMaximumAbsence_(missedHours, scheduledHours, thresholdPercent) {
  if (!(scheduledHours > 0) || !(thresholdPercent > 0)) {
    return {
      absencePercentage: null,
      reviewRequired: false,
      reason: scheduledHours > 0
        ? 'Maximum absence percentage is not configured.'
        : 'Scheduled clinical hours could not be established; no automatic percentage decision was made.'
    };
  }

  const percentage = (missedHours / scheduledHours) * 100;
  return {
    absencePercentage: attendance_round_(percentage),
    reviewRequired: percentage >= thresholdPercent,
    reason: percentage >= thresholdPercent
      ? 'Recorded missed clinical hours meet or exceed the configured review threshold. Program review is required; CCMS does not make a dismissal decision.'
      : ''
  };
}


function attendance_getScheduledHours_(student, clinicalTerm) {
  const warnings = [];

  if (typeof scheduling_getClinicalCalendarForTerm !== 'function') {
    return {
      scheduledHours: null,
      source: 'Unavailable',
      warnings: ['SchedulingService is unavailable; maximum-absence percentage was not calculated.']
    };
  }

  try {
    const rows = scheduling_getClinicalCalendarForTerm(
      clinicalTerm,
      student.cohort
    ) || [];

    const hours = rows.reduce((sum, row) => {
      const value = Number(
        row.scheduledHours !== undefined
          ? row.scheduledHours
          : row['Scheduled Hours']
      );
      return sum + (Number.isFinite(value) ? value : 0);
    }, 0);

    if (!(hours > 0)) {
      warnings.push('Clinical Calendar returned no positive scheduled hours for this student cohort/term.');
      return { scheduledHours: null, source: 'Clinical Calendar', warnings: warnings };
    }

    return {
      scheduledHours: attendance_round_(hours),
      source: 'Clinical Calendar',
      warnings: warnings
    };
  } catch (error) {
    warnings.push('Scheduled-hours lookup failed: ' + error.message);
    return { scheduledHours: null, source: 'Clinical Calendar', warnings: warnings };
  }
}


function attendance_readRules_() {
  const sheet = attendance_requireSheet_(APP_CONFIG.SHEETS.ATTENDANCE_RULES);
  const rows = attendance_sheetObjects_(sheet);

  return rows.map(row => ({
    rowNumber: row.__rowNumber,
    clinicalTerm: attendance_text_(row['Clinical Term']),
    initialEtoHours: attendance_number_(row['Initial ETO Hours'], null),
    clinicalDayHours: attendance_number_(row['Clinical Day Hours'], null),
    etoIncrementHours: attendance_number_(row['ETO Increment Hours'], null),
    etoCarryoverAllowed: attendance_bool_(row['ETO Carryover Allowed']),
    excessEtoRequiresMakeUp: attendance_bool_(row['Excess ETO Requires Make-Up']),
    eveningRotationRequired: attendance_bool_(row['Evening Rotation Required']),
    eveningRotationStart: attendance_timeText_(row['Evening Rotation Start']),
    eveningRotationEnd: attendance_timeText_(row['Evening Rotation End']),
    eveningAbsenceRequiresFullMakeUp: attendance_bool_(row['Evening Absence Requires Full Make-Up']),
    maximumAbsencePercentage: attendance_number_(row['Maximum Absence Percentage'], null),
    effectiveStartDate: attendance_optionalDate_(row['Effective Start Date']),
    effectiveEndDate: attendance_optionalDate_(row['Effective End Date']),
    active: attendance_bool_(row['Active']),
    notes: attendance_text_(row['Notes'])
  }));
}


function attendance_readEvents_() {
  const sheet = attendance_requireSheet_(APP_CONFIG.SHEETS.ATTENDANCE_EVENTS);
  const rows = attendance_sheetObjects_(sheet);

  return rows.map(row => ({
    rowNumber: row.__rowNumber,
    attendanceEventId: attendance_text_(row['Attendance Event ID']),
    studentId: attendance_text_(row['Student ID']),
    date: attendance_optionalDate_(row['Date']),
    clinicalTerm: attendance_text_(row['Clinical Term']),
    eventType: attendance_text_(row['Event Type']),
    hours: attendance_number_(row['Hours'], null),
    etoImpact: attendance_number_(row['ETO Impact'], null),
    makeUpRequired: attendance_bool_(row['Make-Up Required']),
    relatedEventId: attendance_text_(row['Related Event ID']),
    source: attendance_text_(row['Source']),
    approvalStatus: attendance_text_(row['Approval Status']),
    approvedBy: attendance_text_(row['Approved By']),
    approvedAt: attendance_optionalDate_(row['Approved At']),
    notes: attendance_text_(row['Notes']),
    createdAt: attendance_optionalDate_(row['Created At']),
    createdBy: attendance_text_(row['Created By']),
    supersedesEventId: attendance_text_(row['Supersedes Event ID']),
    amendmentReason: attendance_text_(row['Amendment Reason']),
    reason: attendance_text_(row['Reason'])
  }));
}


function attendance_getStudent_(studentId) {
  const sheet = attendance_requireSheet_(APP_CONFIG.SHEETS.STUDENTS);
  const rows = attendance_sheetObjects_(sheet);

  const matches = rows.filter(item =>
    attendance_text_(item['Student ID']) === studentId
  );

  if (matches.length === 0) {
    throw new Error('Student ID not found: ' + studentId);
  }

  const row = attendance_selectCanonicalStudentRow_(matches);

  return {
    studentId: studentId,
    studentName: attendance_studentDisplayName_(row),
    cohort: row['Cohort'] !== undefined ? row['Cohort'] : '',
    active: attendance_bool_(row['Active'])
  };
}


function attendance_getActiveStudentsForCohort_(cohort) {
  const sheet = attendance_requireSheet_(APP_CONFIG.SHEETS.STUDENTS);
  const rows = attendance_sheetObjects_(sheet);
  const target = attendance_normalizeCohort_(cohort);

  const byStudentId = {};

  rows
    .filter(row =>
      attendance_bool_(row['Active']) &&
      attendance_normalizeCohort_(row['Cohort']) === target
    )
    .forEach(row => {
      const studentId = attendance_text_(row['Student ID']);
      if (!studentId) return;

      if (!byStudentId[studentId]) {
        byStudentId[studentId] = [];
      }

      byStudentId[studentId].push(row);
    });

  return Object.keys(byStudentId)
    .map(studentId => {
      const row = attendance_selectCanonicalStudentRow_(
        byStudentId[studentId]
      );

      return {
        studentId: studentId,
        studentName: attendance_studentDisplayName_(row),
        cohort: row['Cohort']
      };
    })
    .sort((a, b) =>
      String(a.studentName || a.studentId)
        .localeCompare(String(b.studentName || b.studentId))
    );
}


/**
 * Resolve one canonical Students row for a stable Student ID without changing
 * source data. Attendance must not treat repeated Students rows as separate
 * people.
 *
 * Preference order:
 *   1. Active rows over inactive rows.
 *   2. Highest Clinical Term number when repeated rows disagree.
 *   3. Latest source row as a deterministic tie-breaker.
 *
 * Student ID remains the relational key. Names are display fields only.
 */
function attendance_selectCanonicalStudentRow_(rows) {
  const candidates = (rows || []).slice();

  if (candidates.length === 0) {
    return {};
  }

  candidates.sort((a, b) => {
    const activeA = attendance_bool_(a['Active']) ? 1 : 0;
    const activeB = attendance_bool_(b['Active']) ? 1 : 0;

    if (activeA !== activeB) {
      return activeB - activeA;
    }

    const termA = attendance_clinicalNumber_(a['Clinical Term']);
    const termB = attendance_clinicalNumber_(b['Clinical Term']);

    if (termA !== termB) {
      return termB - termA;
    }

    return Number(b.__rowNumber || 0) - Number(a.__rowNumber || 0);
  });

  return candidates[0];
}


function attendance_clinicalNumber_(value) {
  const match = attendance_text_(value)
    .match(/CLINICAL\s*([1-5])/i);

  return match ? Number(match[1]) : 0;
}


function attendance_sheetObjects_(sheet) {
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];

  const headers = values[0].map(attendance_text_);
  return values.slice(1)
    .map((row, index) => {
      const obj = { __rowNumber: index + 2 };
      headers.forEach((header, col) => {
        if (header) obj[header] = row[col];
      });
      return obj;
    })
    .filter(obj => Object.keys(obj).some(key =>
      key !== '__rowNumber' && attendance_text_(obj[key]) !== ''
    ));
}


function attendance_requireSheet_(sheetName) {
  if (!sheetName) throw new Error('Required APP_CONFIG sheet alias is missing.');
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) throw new Error('Required sheet not found: ' + sheetName);
  return sheet;
}


function attendance_publicRule_(rule) {
  return {
    clinicalTerm: rule.clinicalTerm,
    initialEtoHours: rule.initialEtoHours,
    clinicalDayHours: rule.clinicalDayHours,
    etoIncrementHours: rule.etoIncrementHours,
    etoCarryoverAllowed: rule.etoCarryoverAllowed,
    excessEtoRequiresMakeUp: rule.excessEtoRequiresMakeUp,
    eveningRotationRequired: rule.eveningRotationRequired,
    eveningRotationStart: rule.eveningRotationStart,
    eveningRotationEnd: rule.eveningRotationEnd,
    eveningAbsenceRequiresFullMakeUp: rule.eveningAbsenceRequiresFullMakeUp,
    maximumAbsencePercentage: rule.maximumAbsencePercentage,
    effectiveStartDate: rule.effectiveStartDate,
    effectiveEndDate: rule.effectiveEndDate,
    active: rule.active,
    notes: rule.notes
  };
}


function attendance_publicEvent_(event) {
  return {
    attendanceEventId: event.attendanceEventId,
    studentId: event.studentId,
    date: event.date,
    clinicalTerm: event.clinicalTerm,
    eventType: event.eventType,
    hours: event.hours,
    etoImpact: event.etoImpact,
    makeUpRequired: event.makeUpRequired,
    relatedEventId: event.relatedEventId,
    source: event.source,
    approvalStatus: event.approvalStatus,
    approvedBy: event.approvedBy,
    approvedAt: event.approvedAt,
    notes: event.notes,
    createdAt: event.createdAt,
    createdBy: event.createdBy,
    supersedesEventId: event.supersedesEventId,
    amendmentReason: event.amendmentReason,
    reason: event.reason
  };
}


function attendance_ruleReferenceDate_(events, fallback) {
  const dates = events.map(e => e.date).filter(Boolean);
  if (!dates.length) return fallback;
  return new Date(Math.max.apply(null, dates.map(d => new Date(d).getTime())));
}


function attendance_rangesOverlap_(aStart, aEnd, bStart, bEnd) {
  const min = new Date(1900, 0, 1);
  const max = new Date(2999, 11, 31);
  const as = aStart || min;
  const ae = aEnd || max;
  const bs = bStart || min;
  const be = bEnd || max;
  return as <= be && bs <= ae;
}


function attendance_normalizeEventType_(value) {
  return attendance_text_(value)
    .toUpperCase()
    .replace(/[\s\-\/]+/g, '_')
    .replace(/[^A-Z0-9_]/g, '');
}


function attendance_normalizeStatus_(value) {
  return attendance_text_(value).toUpperCase();
}


function attendance_normalizeTerm_(value) {
  const text = attendance_text_(value).toUpperCase().replace(/\s+/g, ' ');
  const match = text.match(/CLINICAL\s*([1-5])/);
  return match ? 'CLINICAL ' + match[1] : text;
}


function attendance_normalizeCohort_(value) {
  const text = attendance_text_(value).toUpperCase();
  const year = text.match(/20\d{2}/);
  return year ? year[0] : text.replace(/\s+/g, ' ').trim();
}


function attendance_studentDisplayName_(row) {
  const candidates = [
    row['Student'],
    row['Student Name'],
    row['Name'],
    row['Last Name, First Name'],
    row['Full Name']
  ];

  for (let i = 0; i < candidates.length; i++) {
    if (attendance_text_(candidates[i])) {
      return attendance_text_(candidates[i]);
    }
  }

  const first = attendance_text_(row['First Name']);
  const last = attendance_text_(row['Last Name']);

  return (first + ' ' + last).trim() ||
    attendance_text_(row['Student ID']);
}


function attendance_bool_(value) {
  if (value === true) return true;
  if (value === false || value === null || value === undefined || value === '') return false;
  const text = attendance_text_(value).toUpperCase();
  return ['TRUE', 'YES', 'Y', '1', 'ACTIVE', 'REQUIRED'].includes(text);
}


function attendance_number_(value, fallback) {
  if (value === null || value === undefined || value === '') return fallback;
  const number = Number(value);
  return Number.isFinite(number) ? number : fallback;
}


function attendance_text_(value) {
  if (value === null || value === undefined) return '';
  if (value instanceof Date) return value;
  return String(value).trim();
}



/**
 * Read-only QA: confirms that cohort population is unique by Student ID and
 * exposes display names selected by AttendanceService.
 */
function attendance_testCohortIdentity(cohort) {
  const students = attendance_getActiveStudentsForCohort_(cohort);
  const ids = students.map(student => student.studentId);
  const uniqueIds = new Set(ids);

  const result = {
    success: ids.length === uniqueIds.size,
    cohort: cohort,
    studentCount: students.length,
    uniqueStudentIdCount: uniqueIds.size,
    duplicateStudentIds: ids.filter((id, index) =>
      ids.indexOf(id) !== index
    ),
    students: students
  };

  console.log(JSON.stringify(result, null, 2));
  return result;
}


function attendance_toDate_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) return new Date(value.getTime());
  const date = new Date(value);
  if (isNaN(date.getTime())) throw new Error('Invalid date: ' + value);
  return date;
}


function attendance_optionalDate_(value) {
  if (value === null || value === undefined || value === '') return null;
  try { return attendance_toDate_(value); } catch (error) { return null; }
}


function attendance_timeText_(value) {
  if (value === null || value === undefined || value === '') return '';
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'h:mm a');
  }
  return String(value).trim();
}


function attendance_formatDate_(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}


function attendance_round_(value) {
  return Math.round((Number(value) + Number.EPSILON) * 100) / 100;
}
