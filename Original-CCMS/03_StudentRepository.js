/**
 * CCMS Students retrieval only. Authoritative source: APP_CONFIG.SHEETS.STUDENTS.
 * Uses the existing execution-scoped DataContext; no CacheService for Students.
 * Services retain active/cohort/clinical-term selection and duplicate-row policy.
 * IDs and all original headers/values are returned unchanged. No ID generation,
 * name fallback, DEPT Number derivation, writes, or academic business rules.
 * Default repositories share appDataContext_(); pass the request context when supplied.
 */
function appStudentRepository_(context) {
  context = context || appDataContext_();
  function students() { return context.getObjects(APP_CONFIG.SHEETS.STUDENTS); }
  function byId(studentId) {
    const target = appNormalize_(studentId);
    if (!target) return [];
    return students().filter(row => appNormalize_(row['Student ID']) === target);
  }
  return Object.freeze({
    getStudents: students,
    // Preserve every matching source row; services decide which clinical-term row to use.
    getStudentsById: byId,
    getStudentById: function(studentId) {
      const matches = byId(studentId);
      if (matches.length > 1) throw new Error('Multiple Students rows match Student ID; use getStudentsById and resolve in the service.');
      return matches.length ? matches[0] : null;
    }
  });
}
