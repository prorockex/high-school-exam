const vm=require('node:vm'),fs=require('node:fs'),path=require('node:path'),assert=require('node:assert/strict');
const source=fs.readFileSync(path.resolve(__dirname,'../Code.gs'),'utf8');
const questions=JSON.parse(fs.readFileSync(path.resolve(__dirname,'../questions.json'),'utf8')).slice(0,5).map(q=>({...q,difficulty:'段考進階題'}));
function fixture(){
 const rows=[],props={AI_API_KEY:'test-only',EXAM_SCOPES_JSON:'{"英文":"文法"}'};let released=0,apiCalls=0,status=200,payload={questions};const triggers=[{getHandlerFunction:()=> 'other'}],deleted=[];
 const sheet={getLastRow:()=>rows.length,getLastColumn:()=>rows[0]?.length||0,setFrozenRows(){},getRange(r,c,n,m){return {setValues(values){assert.equal(values.length,n);for(let i=0;i<n;i++){rows[r-1+i]??=[];for(let j=0;j<m;j++)rows[r-1+i][c-1+j]=values[i][j];}return this;},getDisplayValues(){return Array.from({length:n},(_,i)=>Array.from({length:m},(_,j)=>String(rows[r-1+i]?.[c-1+j]??'')));},setFontWeight(){}};}};
 const created={};const chain={};for(const key of ['timeBased','atHour','nearMinute','everyDays','inTimezone'])chain[key]=value=>{created[key]=value??true;return chain;};chain.create=()=>created;
 const ctx={console:{log(){}},Set,Math,Date,JSON,Error,PropertiesService:{getScriptProperties:()=>({getProperty:k=>props[k],setProperty:(k,v)=>props[k]=v})},SpreadsheetApp:{getActiveSpreadsheet:()=>({getSheetByName:()=>sheet,insertSheet:()=>sheet}),flush(){}},LockService:{getScriptLock:()=>({tryLock:()=>true,releaseLock:()=>released++})},Utilities:{formatDate:()=> '20261008',getUuid:()=> 'uuid'},UrlFetchApp:{fetch(url,options){apiCalls++;assert.equal(url,'https://api.openai.com/v1/chat/completions');assert.equal(JSON.parse(options.payload).messages.length,2);return {getResponseCode:()=>status,getContentText:()=>JSON.stringify({choices:[{message:{content:JSON.stringify(payload)}}]})};}},ScriptApp:{getProjectTriggers:()=>triggers,deleteTrigger:t=>deleted.push(t),newTrigger:handler=>{created.handler=handler;return chain;}}};
 vm.createContext(ctx);vm.runInContext(source,ctx);
 return {ctx,rows,props,created,triggers,deleted,setPayload:p=>payload=p,setStatus:s=>status=s,getCalls:()=>apiCalls,getReleased:()=>released};
}
const f=fixture();f.ctx.dailyGenerateQuestions();assert.equal(f.rows.length,6);assert.equal(f.rows[1].length,11);assert.equal(f.getCalls(),1);assert.equal(f.props.LAST_SUCCESS_DATE,'20261008');
f.ctx.dailyGenerateQuestions();assert.equal(f.rows.length,6);assert.equal(f.getCalls(),1);assert.equal(f.getReleased(),2);
f.triggers.push({getHandlerFunction:()=> 'dailyGenerateQuestions'});f.ctx.installDailyTrigger();assert.equal(f.deleted.length,1);assert.equal(f.created.atHour,6);assert.equal(f.created.nearMinute,0);assert.equal(f.created.everyDays,1);assert.equal(f.created.inTimezone,'Asia/Taipei');
const bad=fixture();bad.setPayload({questions:[]});assert.throws(()=>bad.ctx.dailyGenerateQuestions(),/5 題/);assert.equal(bad.rows.length,1);assert.equal(bad.getReleased(),1);
const http=fixture();http.setStatus(401);assert.throws(()=>http.ctx.dailyGenerateQuestions(),/HTTP 401/);assert.equal(http.rows.length,1);
const duplicate=fixture();duplicate.setPayload({questions:Array(5).fill(questions[0])});assert.throws(()=>duplicate.ctx.dailyGenerateQuestions(),/重複題目/);assert.equal(duplicate.rows.length,1);
const partial=fixture();partial.ctx.setupQuestionSheet();partial.rows.push(['GAS-20261008-partial']);assert.throws(()=>partial.ctx.dailyGenerateQuestions(),/不完整/);assert.equal(partial.getCalls(),0);
assert.equal(f.ctx.safeSheetText_('=IMPORTXML("x")'),"'=IMPORTXML(\"x\")");assert.equal(f.ctx.safeSheetText_('普通文字'),'普通文字');
console.log('PASS: GAS mocked append, idempotence, malformed API, HTTP error, duplicate questions, partial batch, trigger, formula protection.');
