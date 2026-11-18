/** Portal-only read model for authorized radiographic view choices.
 * Authority: existing CCMS workbook / Views sheet (ID, Criteria).
 * No academic formula, data write, or additional workbook is introduced.
 */
function ciPortalViewCatalog_() {
  const book=ciDataBook_();
  const cache=CacheService.getScriptCache();
  const cacheKey='CI_PORTAL_VIEWS:v1:'+book.getId();
  try {
    const raw=cache.get(cacheKey);
    if(raw){const parsed=JSON.parse(raw);if(Array.isArray(parsed)&&parsed.length)return parsed;}
  }catch(e){/* A cache error must never fabricate reference data. */}
  const sheet=book.getSheetByName('Views');
  if(!sheet)throw new Error('The authoritative Views worksheet is missing from CCMS.');
  const last=sheet.getLastRow();
  if(last<2)throw new Error('The authoritative Views worksheet has no entries.');
  const values=sheet.getRange(1,1,last,2).getDisplayValues();
  if(String(values[0][0]).trim()!=='ID'||String(values[0][1]).trim()!=='Criteria')
    throw new Error('Views worksheet must contain ID and Criteria in columns A and B.');
  const seen=new Set();
  const result=values.slice(1).map(r=>({viewId:String(r[0]||'').trim(),name:String(r[1]||'').trim()}))
    .filter(r=>r.viewId&&r.name);
  for(const view of result){
    if(seen.has(view.viewId))throw new Error('Duplicate stable view ID '+view.viewId+' in Views worksheet.');
    seen.add(view.viewId);
  }
  if(!result.length)throw new Error('No usable radiographic views in the Views worksheet.');
  result.sort((a,b)=>a.name.localeCompare(b.name,undefined,{numeric:true})||a.viewId.localeCompare(b.viewId));
  try{cache.put(cacheKey,JSON.stringify(result),120);}catch(e){/* Best-effort cache. */}
  return result;
}
function ciR1EvaluationViewOptions(token,siteId,date,studentId){
  const access=ciHWAccess_(token,siteId,date);
  if(!ciHWGetStudent_(access,studentId))throw new Error('Student access unavailable for view selection.');
  return {success:true,views:ciPortalViewCatalog_()};
}
/** Grade preview rejects projections that are not in the current authorized
 * reference list. Existing drafts may contain display-only names; those are
 * accepted only if the name uniquely resolves to a current workbook view. */
function ciPortalCheckViewSelections_(uiState){
  if(!uiState||!Array.isArray(uiState.images))return;
  const views=ciPortalViewCatalog_();
  const byId=new Map(views.map(v=>[v.viewId,v]));
  const byName=new Map();
  for(const v of views){const key=v.name.toLowerCase();byName.set(key,[...(byName.get(key)||[]),v]);}
  uiState.images.forEach((item,i)=>{
    if(!item)return;
    const name=String(item.name||'').trim(),id=String(item.viewId||'').trim();
    if(!name&&!id)return;
    const resolved=id?byId.get(id):null;
    if(id&&(!resolved||resolved.name!==name))
      throw new Error('Image view '+(i+1)+' no longer matches the CCMS Views worksheet. Re-select the view.');
    if(!id && (!byName.has(name.toLowerCase())||byName.get(name.toLowerCase()).length!==1))
      throw new Error('Image view '+(i+1)+' must be selected from the official Views dropdown.');
  });
}
