/** Sole web entry point. The Login and Home stay inside the Apps Script iframe. */
function doGet() {
  const template = HtmlService.createTemplateFromFile('CI_Portal_Login_Release1');
  return template.evaluate().setTitle('CCMS | Clinical Instructor Portal');
}
/**
 * Render Home as an EVALUATED template. Never use createHtmlOutputFromFile
 * for Home: it would display <?!= ... ?> includes literally and break routes.
 * The template contains the separately governed Students and Menus inner views.
 */
function ciPortalRenderHome(token) {
  ciR1Require_(token);
  const home = HtmlService.createTemplateFromFile('CI_Portal_Home_Release1').evaluate().getContent();
  // This is a read-only, fail-fast validation of the composition. It does not
  // load or modify any academic source, submission record, or Drive document.
  if (home.indexOf('id="ci-v3-root"') < 0 ||
      home.indexOf('id="ci-workspace-view"') < 0 ||
      home.indexOf('id="ci-resources-view"') < 0) {
    throw new Error('Portal Home inner views did not render. Verify CI_Portal_Students_v3_Inner and CI_Portal_Secondary_Menus HTML files and Home template evaluation.');
  }
  return home;
}
