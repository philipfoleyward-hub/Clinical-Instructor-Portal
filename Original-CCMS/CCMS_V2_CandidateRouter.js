/**
 * =====================================================================
 * CCMS V2 CANDIDATE ROUTER
 * CandidateRouter.gs
 * =====================================================================
 *
 * PURPOSE
 * -------
 * Provides a safe side-by-side route for the Students V2 candidate without
 * changing the default CCMS web-app behavior.
 *
 * DEFAULT:
 *   .../exec
 *   -> DashboardrunV2
 *
 * STUDENTS V2 CANDIDATE:
 *   .../exec?ui=students-v2
 *   -> DashboardrunV2_Students_Candidate
 *
 * INSTALLATION
 * ------------
 * 1. Add DashboardrunV2_Students_Candidate.html to the SAME Apps Script
 *    project as the existing CCMS services.
 * 2. Add StudentsV2Service.gs to that project.
 * 3. Add this file.
 * 4. In Dashboard.gs, replace ONLY the existing doGet(e) function with the
 *    doGet(e) implementation below.
 *
 * IMPORTANT
 * ---------
 * Do not keep two doGet() definitions in the Apps Script project.
 */


/**
 * Candidate-aware CCMS web-app entry point.
 *
 * The existing/default route is deliberately preserved as DashboardrunV2.
 * Only the explicit ui=students-v2 query parameter selects the candidate.
 */
function doGet(e) {

  const app =
    String(
      e &&
      e.parameter &&
      e.parameter.app
        ? e.parameter.app
        : ''
    )
    .trim()
    .toLowerCase();


  // =====================================================
  // CLINICAL INSTRUCTOR PORTAL
  // =====================================================

  if (app === 'ci') {

    return HtmlService
      .createHtmlOutputFromFile(
        'CI_Portal_Release1_NavigationFixed'
      )
      .setTitle(
        'Clinical Instructor Portal'
      )
      .setXFrameOptionsMode(
        HtmlService
          .XFrameOptionsMode
          .ALLOWALL
      );
  }


  // =====================================================
  // CCMS COORDINATOR DASHBOARD
  // Existing production default.
  // =====================================================

  // Working coordinator dashboard in the approved V2 design
  // (DashboardUI + ccms-v2-shell). ?ui=mock shows the original visual study.
  const ui = String(e && e.parameter && e.parameter.ui || '').trim().toLowerCase();
  return HtmlService
    .createHtmlOutputFromFile(
      ui === 'mock' ? 'DashboardrunV2' : 'DashboardUI'
    )
    .setTitle(
      'Clinical Competency Tracker'
    )
    .setXFrameOptionsMode(
      HtmlService
        .XFrameOptionsMode
        .ALLOWALL
    );
}
