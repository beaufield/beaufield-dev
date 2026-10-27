// 日本語コメント: Googleフォームは読むだけ。同期先の通知設定はこのGASから変更しない。
const VERSION='v0.1.1';
function surveyConfig_(inspectOnly){
 const p=PropertiesService.getScriptProperties();
 const forms=JSON.parse(p.getProperty('SURVEY_CONFIG')||'[]');
 if(forms.length!==2||new Set(forms.map(f=>f.key)).size!==2||forms.some(f=>!['maker','customer'].includes(f.key)||!/^[A-Za-z0-9_-]+$/.test(f.formId||'')))throw Error('SURVEY_FORMS_NOT_CONFIGURED');
 const endpoint=p.getProperty('SURVEY_ENDPOINT')||'',secret=p.getProperty('SURVEY_SOURCE_SECRET')||'';
 if(!inspectOnly&&(!/^https:\/\/[A-Za-z0-9-]+\.supabase\.co\/functions\/v1\/[A-Za-z0-9_-]+$/.test(endpoint)||secret.length<32))throw Error('SURVEY_ENDPOINT_NOT_CONFIGURED');
 return {forms,endpoint,secret,enabled:p.getProperty('SURVEY_ENABLED')==='true'};
}
function surveyHex_(bytes){return bytes.map(b=>('0'+((b+256)%256).toString(16)).slice(-2)).join('');}
function surveySort_(x){if(Array.isArray(x))return x.map(surveySort_);if(x&&typeof x==='object'){const y={};Object.keys(x).sort().forEach(k=>y[k]=surveySort_(x[k]));return y;}return x;}
function surveyHash_(x){return surveyHex_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,JSON.stringify(surveySort_(x)),Utilities.Charset.UTF_8));}
function surveyCall_(config,action,body){
 const text=JSON.stringify({action,body}),time=String(Math.floor(Date.now()/1000)),rid=Utilities.getUuid();
 if(Utilities.newBlob(text).getBytes().length>262144)throw Error('SURVEY_PAYLOAD_TOO_LARGE');
 const sig=surveyHex_(Utilities.computeHmacSha256Signature(time+'\n'+rid+'\n'+text,config.secret,Utilities.Charset.UTF_8));
 const response=UrlFetchApp.fetch(config.endpoint,{method:'post',contentType:'application/json',headers:{'x-survey-time':time,'x-survey-request':rid,'x-survey-signature':sig},payload:text,muteHttpExceptions:true,followRedirects:false});
 const code=response.getResponseCode();if(code!==200)throw Error('SURVEY_HTTP_'+code);
 let data;try{data=JSON.parse(response.getContentText());}catch(e){throw Error('SURVEY_BAD_ACK');}
 return data;
}
function surveyItem_(item){
 const type=String(item.getType());const out={id:String(item.getId()),title:item.getTitle(),type,rows:[],choices:[]};
 if(type==='CHECKBOX')out.choices=item.asCheckboxItem().getChoices().map(c=>c.getValue());
 if(type==='MULTIPLE_CHOICE')out.choices=item.asMultipleChoiceItem().getChoices().map(c=>c.getValue());
 if(type==='LIST')out.choices=item.asListItem().getChoices().map(c=>c.getValue());
 if(type==='GRID'){out.rows=item.asGridItem().getRows();out.choices=item.asGridItem().getColumns();}
 if(type==='CHECKBOX_GRID'){out.rows=item.asCheckboxGridItem().getRows();out.choices=item.asCheckboxGridItem().getColumns();}
 return out;
}
function surveyItems_(form){return form.getItems().filter(i=>!['PAGE_BREAK','SECTION_HEADER','IMAGE','VIDEO'].includes(String(i.getType()))).map(surveyItem_);}
function surveySnapshot_(form,response,schema){
 const answered={};response.getItemResponses().forEach(a=>answered[String(a.getItem().getId())]=a.getResponse());
 const items=(schema||surveyItems_(form)).map(i=>{
  let value=Object.prototype.hasOwnProperty.call(answered,i.id)?answered[i.id]:null;
  if(value!==null&&!Array.isArray(value))value=String(value);
  return Object.assign({},i,{answer:value});
 });
 const submitted_at=response.getTimestamp().toISOString();
 const hashItems=items.map(i=>Object.assign({},i,{answer:i.type==='CHECKBOX'&&i.answer?i.answer.slice().sort():i.answer}));
 return {source_id:response.getId(),submitted_at,items,content_hash:surveyHash_({submitted_at,items:hashItems})};
}
function surveySequence_(minimum){
 const p=PropertiesService.getScriptProperties();const n=Math.max(Number(p.getProperty('SURVEY_SEQUENCE')||0),Number(minimum||0))+1;
 if(!Number.isSafeInteger(n))throw Error('SURVEY_SEQUENCE_INVALID');p.setProperty('SURVEY_SEQUENCE',String(n));return n;
}
function surveySave_(config,f,form,response,minimum,snapshot){
 const snap=snapshot||surveySnapshot_(form,response);
 const body=Object.assign({},snap,{survey_key:f.key,form_id:f.formId,observed_at:new Date().toISOString(),source_sequence:surveySequence_(minimum)});
 delete body.content_hash;
 const ack=surveyCall_(config,'ingest',body);if(ack.saved!==true)throw Error('SURVEY_UNSAVED');return body.source_sequence;
}
// 日本語コメント: 初回はこの読取関数で設問IDと型を確認する。回答原文をログへ出さない。
function inspectSurveyConfiguration(){
 const config=surveyConfig_(true);
 return {version:VERSION,enabled:config.enabled,forms:config.forms.map(f=>{
  const form=FormApp.openById(f.formId);
  return {key:f.key,title:form.getTitle(),count:form.getResponses().length,items:surveyItems_(form)};
 })};
}
function surveyOnSubmit(e){
 const config=surveyConfig_();if(!config.enabled)return;
 if(!e||!e.source||!e.response)throw Error('FORM_EVENT_REQUIRED');
 const f=config.forms.find(x=>x.formId===e.source.getId());if(!f)throw Error('FORM_FORBIDDEN');
 const lock=LockService.getScriptLock();if(!lock.tryLock(1000))throw Error('SURVEY_RECONCILE_REQUIRED');
 try{
  const form=FormApp.openById(f.formId),manifest=surveyCall_(config,'source_manifest',{survey_key:f.key,form_id:f.formId});
  // 遅れて実行されたイベントも、現在の回答を読み直して古い内容の上書きを防ぐ。
  surveySave_(config,f,form,form.getResponse(e.response.getId()),manifest.sequence);
 }finally{lock.releaseLock();}
}
function surveyReconcile(){
 const config=surveyConfig_();if(!config.enabled)return;
 const start=Date.now(),lock=LockService.getScriptLock();if(!lock.tryLock(1000))return;
 try{
  for(const f of config.forms){
   const form=FormApp.openById(f.formId),manifest=surveyCall_(config,'source_manifest',{survey_key:f.key,form_id:f.formId});
   const existing={};manifest.responses.forEach(r=>existing[r.source_id]=r.hash);
   const responses=form.getResponses(),schema=surveyItems_(form);const ids=[];let sequence=Number(manifest.sequence);
   for(const response of responses){
    if(Date.now()-start>240000)throw Error('SURVEY_RECONCILE_INCOMPLETE');
    const snap=surveySnapshot_(form,response,schema);ids.push(snap.source_id);
    if(existing[snap.source_id]!==snap.content_hash)sequence=surveySave_(config,f,form,response,sequence,snap);
   }
   const ack=surveyCall_(config,'sync_complete',{survey_key:f.key,form_id:f.formId,ids,source_sequence:surveySequence_(sequence)});
   if(ack.complete!==true)throw Error('SURVEY_RECONCILE_NOT_CONFIRMED');
  }
 }finally{lock.releaseLock();}
}
// 日本語コメント: 本番反映承認後に明示実行。既存の他用途トリガーには触れない。
function installSurveyTriggers(){
 const config=surveyConfig_();if(!config.enabled)throw Error('ENABLE_SYNC_AFTER_APPROVAL');
 const existing=ScriptApp.getProjectTriggers();
 for(const f of config.forms)if(!existing.some(t=>t.getHandlerFunction()==='surveyOnSubmit'&&t.getTriggerSourceId()===f.formId))ScriptApp.newTrigger('surveyOnSubmit').forForm(FormApp.openById(f.formId)).onFormSubmit().create();
 if(!existing.some(t=>t.getHandlerFunction()==='surveyReconcile'))ScriptApp.newTrigger('surveyReconcile').timeBased().everyMinutes(5).create();
 return {version:VERSION,installed:true,worksEnabledByThisFunction:false};
}
function stopSurveySync(){
 PropertiesService.getScriptProperties().setProperty('SURVEY_ENABLED','false');
 ScriptApp.getProjectTriggers().filter(t=>['surveyOnSubmit','surveyReconcile'].includes(t.getHandlerFunction())).forEach(t=>ScriptApp.deleteTrigger(t));
 return {stopped:true};
}
