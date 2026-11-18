/**
 * Clinical evaluation Form schema read model (read-only).
 * The official existing Google Form remains the authority. Its structural
 * metadata is cached for up to six hours, not copied into another editable
 * academic source. Set CI_PORTAL_FORM_SCHEMA_REVISION to a new value when
 * changing the authoritative Form to invalidate existing cached definitions.
 * Explicit coordinator refresh: ciPortalRefreshEvaluationSchema().
 * Never writes clinical records, Forms, or grades.
 */
let ciPortalEvaluationSchemaForRequest_ = null;
function ciPortalOfficialFormId_(){
  const id=String(PropertiesService.getScriptProperties().getProperty('CI_PORTAL_OFFICIAL_EVALUATION_FORM_ID')||'1TGm4ciMz3YkKv902gON-FmnofNUORqMmKySVA5pPrsY').trim();
  if(!/^[A-Za-z0-9_-]{20,}$/.test(id))throw new Error('Official clinical evaluation Form ID is invalid');
  return id;
}
function ciPortalFormSchemaKey_(){
  const p=PropertiesService.getScriptProperties();
  const fingerprint=Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,
    ciPortalOfficialFormId_()+'|'+String(p.getProperty('CI_PORTAL_FORM_SCHEMA_REVISION')||'1'))).replace(/=+$/,'').slice(0,32);
  return 'CI_FORM_SCHEMA_V2_'+fingerprint;
}
function ciPortalLoadOfficialFormSchema_(force){
  if(!force && ciPortalEvaluationSchemaForRequest_)return ciPortalEvaluationSchemaForRequest_;
  const cache=CacheService.getScriptCache(),key=ciPortalFormSchemaKey_();
  if(!force){
    try{
      const raw=cache.get(key);
      if(raw){const parsed=JSON.parse(raw);if(parsed&&parsed.version===2&&Array.isArray(parsed.formItems)&&parsed.formItems.length>0){
        ciPortalEvaluationSchemaForRequest_=parsed;return parsed;
      }}
    }catch(e){console.warn('Evaluation schema cache miss: '+String(e));}
  }
  let items;
  try{items=FormApp.openById(ciPortalOfficialFormId_()).getItems();}
  catch(e){throw new Error('Unable to read the authoritative clinical evaluation Form: '+String(e.message||e));}
  const types=FormApp.ItemType;
  const formItems=items.map(item=>{
    const type=item.getType(),entry={itemId:String(item.getId()),title:String(item.getTitle()),type:String(type)};
    if(type===types.GRID){entry.rows=item.asGridItem().getRows().map(String);entry.choices=item.asGridItem().getColumns().map(String);}
    else if(type===types.CHECKBOX)entry.choices=item.asCheckboxItem().getChoices().map(c=>String(c.getValue()));
    else if(type===types.MULTIPLE_CHOICE)entry.choices=item.asMultipleChoiceItem().getChoices().map(c=>String(c.getValue()));
    else if(type===types.LIST)entry.choices=item.asListItem().getChoices().map(c=>String(c.getValue()));
    return entry;
  });
  const names=formItems.filter(x=>x.type==='GRID').map(x=>x.title.trim().toLowerCase());
  for(const required of ['exam preparation','patient care','radiation protection','equipment manipulation','positioning skills','portable / contrast']){
    if(names.filter(x=>x===required).length!==1)throw new Error('Official Form schema is incomplete or duplicated: '+required);
  }
  const schema={version:2,source:'official Google Form',formId:ciPortalOfficialFormId_(),loadedAt:new Date().toISOString(),formItems};
  const raw=JSON.stringify(schema);
  // Script Cache has a 100-KB per-key maximum. If exceeded, use the source
  // directly rather than cache truncated or unauthoritative definitions.
  if(Utilities.newBlob(raw).getBytes().length<90000){try{cache.put(key,raw,21600);}catch(e){console.warn('Schema cache write skipped: '+String(e));}}
  ciPortalEvaluationSchemaForRequest_=schema;
  return schema;
}
/** Apps Script editor/manual diagnostics only; not exposed by any Portal UI. */
function ciPortalRefreshEvaluationSchema(){
  const result=ciPortalLoadOfficialFormSchema_(true);
  return {success:true,formId:result.formId,itemCount:result.formItems.length,loadedAt:result.loadedAt};
}
/** Lightweight adapter for the existing authoritative Form-ID grading mapping.
 * Behaves like the subset of FormApp Item read methods that mapping uses.
 * Academic rules and item identifiers remain unchanged. */
function ciPortalGradingFormItems_(){
  const items=ciPortalLoadOfficialFormSchema_(false).formItems;
  return items.map(entry=>({
    getId:()=>Number(entry.itemId),
    getTitle:()=>entry.title,
    getType:()=>FormApp.ItemType[entry.type],
    asGridItem:()=>({getRows:()=>(entry.rows||[]).slice(),getColumns:()=>(entry.choices||[]).slice()}),
    asCheckboxItem:()=>({getChoices:()=>(entry.choices||[]).map(value=>({getValue:()=>value}))}),
    asMultipleChoiceItem:()=>({getChoices:()=>(entry.choices||[]).map(value=>({getValue:()=>value}))}),
    asListItem:()=>({getChoices:()=>(entry.choices||[]).map(value=>({getValue:()=>value}))})
  }));
}
