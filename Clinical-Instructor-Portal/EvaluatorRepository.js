/**
 * CCMS Evaluators retrieval only. Source: APP_CONFIG.SHEETS.EVALUATORS.
 * Uses existing execution-scoped DataContext. No cross-request cache is enabled
 * in this slice: evaluator rows include Active/Portal Access facts, and no
 * guaranteed invalidation path for script/manual updates has been established.
 * Authentication, authorization, password checks, and login eligibility remain
 * entirely in PortalAccessService. This repository neither filters access flags
 * nor synthesizes Hospital CI records. IDs/headers/source values stay unchanged.
 */
function appEvaluatorRepository_(context) {
  context = context || appDataContext_();
  function evaluators() { return context.getObjects(APP_CONFIG.SHEETS.EVALUATORS); }
  function byId(evaluatorId) {
    const target = appNormalize_(evaluatorId);
    if (!target) return [];
    return evaluators().filter(row => appNormalize_(row['Evaluator ID']) === target);
  }
  return Object.freeze({
    getEvaluators: evaluators,
    getEvaluatorsById: byId,
    getEvaluatorById: function(evaluatorId) {
      const matches = byId(evaluatorId);
      if (matches.length > 1) throw new Error('Multiple Evaluators rows match Evaluator ID; use getEvaluatorsById and resolve in the service.');
      return matches.length ? matches[0] : null;
    }
  });
}
