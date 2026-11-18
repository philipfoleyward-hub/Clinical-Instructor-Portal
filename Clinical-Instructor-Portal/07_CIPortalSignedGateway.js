/**
 * OPTIONAL, DISABLED-BY-DEFAULT server-to-server ingress to ORIGINAL CCMS.
 * Existing coordinator doGet() is unchanged. No public access is authorized
 * merely by installing this file. This doPost requires an approved deployment,
 * IT/security review, and Script Properties: CI_PORTAL_GATEWAY_ENABLED=true,
 * CI_PORTAL_GATEWAY_SECRET = same long random secret set on BOTH projects.
 * Never place the secret in HTML/client JavaScript.
 *
 * Important: this gateway is a transport adapter, not a business orchestrator.
 * The authoritative submission pipeline remains 03_DigitalEvaluationIntake.
 */
function ccmsPortalGateway_secret_() {
  const props=PropertiesService.getScriptProperties();
  if(props.getProperty('CI_PORTAL_GATEWAY_ENABLED')!=='true') throw new Error('Portal evaluation gateway is disabled');
  const key=props.getProperty('CI_PORTAL_GATEWAY_SECRET')||'';
  if(key.length<48) throw new Error('Portal gateway configuration is incomplete');
  return key;
}
function ccmsPortalGateway_sign_(message,key) {
  return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(message,key)).replace(/=+$/,'');
}
function ccmsPortalGateway_equal_(x,y) {
  const a=String(x||''),b=String(y||'');
  let diff=a.length^b.length;
  for(let i=0;i<Math.max(a.length,b.length);i++)diff|=(a.charCodeAt(i)||0)^(b.charCodeAt(i)||0);
  return diff===0;
}
function ccmsPortalGateway_verify_(envelope) {
  const key=ccmsPortalGateway_secret_();
  const timestamp=Number(envelope.timestamp),nonce=String(envelope.nonce||'');
  if (!Number.isInteger(timestamp)||Math.abs(Date.now()-timestamp)>120000||!/^[a-f0-9-]{36}$/i.test(nonce)) throw new Error('Invalid or expired request');
  if(typeof envelope.payloadJson!=='string'||envelope.payloadJson.length>300000) throw new Error('Invalid request body');
  const action=String(envelope.action||'');
  if(!['schema','preview','submit','draft_save','draft_load','draft_stop','draft_list'].includes(action)) throw new Error('Unsupported gateway action');
  const signed=action+'\n'+timestamp+'\n'+nonce+'\n'+envelope.payloadJson;
  const expected=ccmsPortalGateway_sign_(signed,key);
  if(!ccmsPortalGateway_equal_(expected,envelope.signature)) throw new Error('Unauthorized request');
  // Enforce replay rejection atomically. Cache entries expire with the
  // maximum tolerated request time skew.
  const lock=LockService.getScriptLock();lock.waitLock(15000);
  try {
    const cache=CacheService.getScriptCache();
    const cacheKey='portal_nonce_'+nonce;
    if(cache.get(cacheKey)) throw new Error('Request already used');
    cache.put(cacheKey,'1',180);
  }finally{lock.releaseLock();}
  let payload;
  try{payload=JSON.parse(envelope.payloadJson);}catch(e){throw new Error('Malformed payload JSON');}
  if(!payload||typeof payload!=='object'||Array.isArray(payload))throw new Error('Invalid payload');
  return {action:action,payload:payload};
}
function doPost(e) {
  // Do not register a second doGet. Existing coordinator dashboard is kept.
  try {
    const raw=e&&e.postData&&e.postData.contents||'';
    if(!raw||raw.length>400000)throw new Error('Invalid submission body');
    const verified=ccmsPortalGateway_verify_(JSON.parse(raw));
    let result;
    if(verified.action==='schema') {
      ccmsPortal_resolveIdentity_(verified.payload);
      result={success:true,formItems:ccmsPortal_schema_()};
    } else if(verified.action==='draft_list') {
      const id=ccmsPortal_clean_(verified.payload.authenticatedEvaluatorId);
      const r=ccmsPortal_reference_(SpreadsheetApp.getActiveSpreadsheet(),APP_CONFIG.SHEETS.EVALUATORS,'Evaluator ID',id,['Evaluator Type','Active','Portal Access']);
      if(!ccmsPortal_true_(r.Active)||!ccmsPortal_true_(r['Portal Access'])||ccmsPortal_norm_(r['Evaluator Type'])!=='college ci')throw new Error('Evaluator draft access unavailable');
      result={success:true,drafts:driveService_ciDraftList_(id)};
    } else if(verified.action==='draft_load') {
      const identity=ccmsPortal_resolveIdentity_(verified.payload);
      result=driveService_ciDraftLoad_(identity,verified.payload.submissionId);
    } else if(verified.action==='draft_save'||verified.action==='draft_stop') {
      const identity=ccmsPortal_resolveIdentity_(verified.payload);
      // Draft saves may include incomplete answers; grades are computed only in preview/finalize.
      result=driveService_ciDraftSave_(identity,verified.payload);
    } else if(verified.action==='preview') {
      const outcome=gradeClinicalEvaluationFromPortalPayload(verified.payload);
      result={success:true,grade:outcome.grade};
    } else {
      result=digitalEvaluation_submitPortal_(verified.payload);
    }
    return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
  }catch(error) {
    console.error('Authenticated portal gateway request rejected: '+error.message);
    // Don't leak stack traces or internal configuration to a web caller.
    return ContentService.createTextOutput(JSON.stringify({success:false,error:String(error.message||'Request rejected')})).setMimeType(ContentService.MimeType.JSON);
  }
}
