const {chromium}=require('playwright');
const fs=require('node:fs');
const path=require('node:path');
const assert=require('node:assert/strict');
const {pathToFileURL}=require('node:url');
const http=require('node:http');
(async()=>{
 const browser=await chromium.launch({executablePath:process.env.CHROMIUM_PATH||'/usr/bin/chromium',headless:true,args:['--no-sandbox']});
 const context=await browser.newContext({viewport:{width:1440,height:1000}});
 let page=await context.newPage();const errors=[];page.on('pageerror',e=>errors.push(e.message));
 const server=http.createServer((req,res)=>{res.setHeader('Content-Type','text/html;charset=utf-8');res.end(fs.readFileSync(path.resolve(__dirname,'../index.html')));});
 await new Promise(resolve=>server.listen(0,'127.0.0.1',resolve));
 try { await page.goto(pathToFileURL(path.resolve(__dirname,'../index.html')).href);console.log('Validated file:// loading.'); }
 catch(e) { if(!e.message.includes('ERR_BLOCKED_BY_ADMINISTRATOR'))throw e;console.log('Managed Chromium blocks file://; testing identical single HTML via local HTTP.');await page.close();page=await context.newPage();page.on('pageerror',e=>errors.push(e.message));await page.goto(`http://127.0.0.1:${server.address().port}`); }
 await page.evaluate(()=>localStorage.clear());await page.reload();
 const data=JSON.parse(fs.readFileSync(path.resolve(__dirname,'../questions.json'),'utf8'));
 assert.equal(data.length,30);assert.equal(new Set(data.map(q=>q.id)).size,30);
 await page.evaluate(()=>validateQuestions(seed));
 // Settings and exact allocation, including a small daily budget across all subjects.
 await page.locator('[data-page=settings]').click();
 await page.locator('#daily-minutes').fill('11');
 for(const subject of ['國文','自然','社會'])await page.locator(`[name=subject][value="${subject}"]`).check();
 for(const input of await page.locator('[name=weekday]').all())await input.check();
 await page.locator('#settings-form button[type=submit]').click();
 assert.equal(await page.evaluate(()=>buildPlan()[0].tasks.reduce((n,t)=>n+t.minutes,0)),11);
 await page.locator('[data-page=plan]').click();
 const task=page.locator('[data-task]').first();await task.check();const taskId=await task.getAttribute('data-task');
 await page.reload();assert.equal(await page.evaluate(id=>state.done[id],taskId),true);
 // Reject no selected subjects, retain prior valid settings.
 await page.locator('[data-page=settings]').click();
 for(const input of await page.locator('[name=subject]').all())await input.uncheck();
 await page.locator('#settings-form button[type=submit]').click();
 assert.match(await page.locator('#settings-status').innerText(),/至少選一個考科/);
 assert.equal(await page.evaluate(()=>state.settings.subjects.length),5);
 await page.locator('[name=subject][value="英文"]').check();await page.locator('[name=subject][value="數學"]').check();
 await page.locator('#advanced-settings').evaluate(e=>e.open=true);await page.locator('#api-key').fill('test-key-not-real');await page.locator('#settings-form button[type=submit]').click();
 assert.ok(!(await page.evaluate(()=>localStorage.getItem(STORE))).includes('test-key-not-real'));
 // Offline exam, one wrong and one right answer.
 await page.locator('[data-page=exam]').click();await page.locator('#exam-number').fill('2');await page.locator('#start-exam').click();
 const questions=await page.evaluate(()=>activeExam.questions);
 const wrong=['A','B','C','D'].find(l=>l!==questions[0].answer);
 await page.locator(`[name=answer-0][value=${wrong}]`).check();await page.locator(`[name=answer-1][value=${questions[1].answer}]`).check();
 await page.locator('#answer-form button[type=submit]').click();
 assert.match(await page.locator('.result-banner').innerText(),/50%/);assert.equal(await page.evaluate(()=>Object.keys(state.mistakes).length),1);
 await page.locator('#view-mistakes').click();await page.locator('#retry-mistakes').click();
 await page.locator(`[name=answer-0][value=${questions[0].answer}]`).check();await page.locator('#answer-form button[type=submit]').click();
 assert.equal(await page.evaluate(()=>Object.keys(state.mistakes).length),0);
 // Blank answers become mistakes and survive reload.
 await page.locator('#exam-number').fill('1');await page.locator('#start-exam').click();
 page.once('dialog',d=>d.accept());await page.locator('#answer-form button[type=submit]').click();
 assert.equal(await page.evaluate(()=>Object.keys(state.mistakes).length),1);await page.reload();assert.equal(await page.evaluate(()=>Object.keys(state.mistakes).length),1);
 // CSV round trip with commas, quotes, newlines, BOM and hostile text rendered literally.
 const csv=await page.evaluate(()=>{const q=structuredClone(seed[0]);q.id='TEST-CSV';q.question='換行\n"quoted", <img src=x onerror="window.injected=true">';return csvString([q]);});
 await page.locator('[data-page=bank]').click();await page.locator('#csv-file').setInputFiles({name:'test.csv',mimeType:'text/csv',buffer:Buffer.from(csv)});
 await page.waitForFunction(()=>state.bank.length===31);assert.equal(await page.evaluate(()=>state.bank.length),31);
 await page.locator('#bank-search').fill('quoted');assert.equal(await page.locator('#bank-list img').count(),0);assert.equal(await page.evaluate(()=>window.injected),undefined);
 assert.match(await page.locator('#bank-list').innerText(),/quoted/);
 const before=await page.evaluate(()=>JSON.stringify(state.bank));
 await page.locator('#csv-file').setInputFiles({name:'bad.csv',mimeType:'text/csv',buffer:Buffer.from('id,bad\n1,2')});
 assert.match(await page.locator('#csv-status').innerText(),/欄位/);assert.equal(await page.evaluate(()=>JSON.stringify(state.bank)),before);
 await assert.rejects(page.evaluate(()=>parseCSV('"unfinished')));
 // Reimport all 30 seed questions updates IDs without duplication.
 await page.locator('#csv-file').setInputFiles(path.resolve(__dirname,'../鳳新高中各科題庫.csv'));
 await page.waitForFunction(()=>state.bank.length===31);assert.equal(await page.evaluate(()=>state.bank.length),31);
 // Mock CSV fetch success and failure preserving existing bank.
 await page.locator('[data-page=settings]').click();await page.locator('#advanced-settings').evaluate(e=>e.open=true);await page.locator('#csv-url').fill('https://example.test/bank.csv');await page.locator('#api-key').fill('test-key-not-real');await page.locator('#settings-form button[type=submit]').click();
 await page.route('https://example.test/bank.csv',route=>route.fulfill({status:200,contentType:'text/csv',body:csv}));
 await page.locator('[data-page=bank]').click();await page.locator('#sync-csv').click();await page.waitForFunction(()=>!document.querySelector('#sync-csv').disabled);assert.match(await page.locator('#csv-status').innerText(),/已合併/);
 await page.unroute('https://example.test/bank.csv');await page.route('https://example.test/bank.csv',route=>route.fulfill({status:403,body:'Denied'}));
 await page.locator('#sync-csv').click();await page.waitForFunction(()=>!document.querySelector('#sync-csv').disabled);assert.match(await page.locator('#csv-status').innerText(),/HTTP 403/);
 // AI confirmation and format validation, without paid network calls.
 const aiQuestions=data.slice(0,5).map(q=>({...q,subject:'英文',difficulty:'基礎'}));
 await page.route('https://api.openai.com/v1/chat/completions',route=>route.fulfill({status:200,contentType:'application/json',body:JSON.stringify({choices:[{message:{content:JSON.stringify({questions:aiQuestions})}}]})}));
 page.once('dialog',d=>d.accept());await page.locator('#generate-ai').click();await page.waitForFunction(()=>document.querySelector('#ai-status').textContent.includes('已新增'));
 assert.equal(await page.evaluate(()=>state.bank.length),36);
 await page.unroute('https://api.openai.com/v1/chat/completions');
 await page.route('https://api.openai.com/v1/chat/completions',route=>route.fulfill({status:200,contentType:'application/json',body:JSON.stringify({choices:[{message:{content:'{"questions":[]}'}}]})}));
 page.once('dialog',d=>d.accept());await page.locator('#generate-ai').click();await page.waitForFunction(()=>document.querySelector('#ai-status').textContent.includes('失敗'));
 assert.equal(await page.evaluate(()=>state.bank.length),36);
 await page.reload();assert.equal(await page.locator('#api-key').inputValue(),'');assert.equal(await page.evaluate(()=>apiKey),'');
 // Date boundary logic.
 assert.equal(await page.evaluate(()=>{state.settings.examDate=dateKey();return buildPlan().length;}),0);
 assert.equal(await page.evaluate(()=>countdown()),0);
 // Screenshots and mobile overflow in every page.
 await page.reload();await page.screenshot({path:'/Users/prorock/.hermes/cache/scratch/fengxin-desktop.png',fullPage:false});
 await page.setViewportSize({width:390,height:844});
 for(const section of ['settings','plan','exam','mistakes','bank']){
  await page.evaluate(s=>navigate(s),section);
  assert.ok(await page.evaluate(()=>document.documentElement.scrollWidth<=innerWidth),`mobile overflow: ${section}`);
 }
 await page.evaluate(()=>navigate('plan'));await page.screenshot({path:'/Users/prorock/.hermes/cache/scratch/fengxin-mobile.png',fullPage:false});
 assert.deepEqual(errors,[]);await browser.close();await new Promise(resolve=>server.close(resolve));console.log('PASS: offline UI, settings, planning, exams, mistake retry, persistence, CSV, mocked AI, responsive layout, escaping.');
})().catch(e=>{console.error(e);process.exit(1);});
