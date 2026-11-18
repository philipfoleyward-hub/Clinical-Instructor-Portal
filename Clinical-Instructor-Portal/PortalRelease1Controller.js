/** CI Portal standalone controller. Candidate milestone 2026-10-15.
 * One web entry point remains PortalEntry.doGet. Academic records are never written here.
 */
function ciPortalGetLoginOptions() {
  return {success:true,evaluators:ciDataEvaluators_()};
}
function ciR1Login(evaluatorId,password) {
  const props=PropertiesService.getScriptProperties();
  const secret=props.getProperty('CI_PORTAL_SHARED_PASSWORD');
  if (!secret) throw new Error('Portal shared password not configured');
  if (String(password||'')!==secret) throw new Error('Incorrect portal password');
  const evaluator=ciDataEvaluators_().find(e=>e.evaluatorId===String(evaluatorId||'').trim());
  if (!evaluator) throw new Error('Instructor access unavailable');
  // Server-side revocable session; no arbitrary elapsed-time logout.
  const token='ci2_'+Utilities.getUuid()+'_'+Utilities.getUuid();
  const key=ciR1SessionKey_(token);
  props.setProperty(key,JSON.stringify({evaluatorId:evaluator.evaluatorId,issuedAt:new Date().toISOString()}));
  return {success:true,token,evaluator};
}
function ciR1SessionKey_(token) {
  if (!/^ci2_[a-f\d-]{36}_[a-f\d-]{36}$/i.test(String(token||''))) throw new Error('Sign in required');
  const digest=Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,String(token));
  return 'CI_PORTAL_SESSION_'+Utilities.base64EncodeWebSafe(digest).replace(/=+$/,'');
}
function ciR1Require_(token) {
  const props=PropertiesService.getScriptProperties();
  const value=props.getProperty(ciR1SessionKey_(token));
  if (!value) throw new Error('Sign in required');
  let session;
  try {session=JSON.parse(value);} catch(e){throw new Error('Sign in required');}
  const evaluator=ciDataEvaluators_().find(e=>e.evaluatorId===session.evaluatorId);
  if (!evaluator) throw new Error('Instructor access unavailable');
  return evaluator;
}
function ciR1Logout(token) {
  // Server session is revoked before client-local context is discarded.
  try {PropertiesService.getScriptProperties().deleteProperty(ciR1SessionKey_(token));}
  catch(e){throw new Error('Unable to sign out securely. Please try again.');}
  return {success:true};
}
function ciR1Home(token,dateOverride) {
  const evaluator=ciR1Require_(token);
  const date=dateOverride||Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyy-MM-dd');
  return ciDataHome_(evaluator,date);
}
function ciPortalWorkbookId_() {
  const id=String(PropertiesService.getScriptProperties().getProperty('CCMS_WORKBOOK_ID')||'').trim();
  if(!/^[A-Za-z0-9_-]{20,}$/.test(id)) throw new Error('CCMS_WORKBOOK_ID not configured');
  const submitMode=String(PropertiesService.getScriptProperties().getProperty('CI_PORTAL_EVALUATION_SUBMIT_MODE')||'disabled');
  if(submitMode==='test_only'&&id!=='1QnKQEKFipHvDz7xitjOoVu1IhlHcBCI1Z_SBQUaJF6Y'){
    throw new Error('TEST-001 Portal requires the approved sandbox CCMS workbook');
  }
  return id;
}
/** Login form is returned in the existing Apps Script frame after logout. */
function ciR1LoginHtml() {
  return HtmlService.createHtmlOutputFromFile('CI_Portal_Login_Release1').getContent();
}
