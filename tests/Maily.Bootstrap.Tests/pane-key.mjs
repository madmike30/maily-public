import fs from 'node:fs';
import vm from 'node:vm';
import assert from 'node:assert/strict';
const source=fs.readFileSync('services/Maily.OutlookAddin.Web/wwwroot/addin/js/maily-office.js','utf8');
function setup(statuses) {
 const elements=[],requests=[];
 function element(tag){const el={tag,value:'',listeners:{},append(){},setAttribute(){},addEventListener(n,f){this.listeners[n]=f;},focus(){},showModal(){this.open=true;},close(){this.open=false;},remove(){this.removed=true;}};elements.push(el);return el;}
 const forbidden=()=>{throw Error('key persistence/logging prohibited');};
 const ctx=vm.createContext({document:{createElement:element,body:element('body'),activeElement:{focus(){}}},fetch:async(path,options)=>{requests.push({path,headers:{...options.headers},body:options.body,redirect:options.redirect});const status=await (statuses.shift()??200);return{status,ok:status===200,text:async()=>'{"error":"Unauthorized."}',json:async()=>({ok:true})};},describeFetchError:()=> 'Request failed.',localStorage:{setItem:forbidden},sessionStorage:{setItem:forbidden},console:{log:forbidden}});
 // Execute the real object-copy helpers; substituting a variadic mock masks API-key regressions.
 vm.runInContext(source.slice(source.indexOf('  function copyObject('),source.indexOf('  function getElementText(')),ctx);
 const start=source.indexOf('  let backendAccessKey');
 vm.runInContext(source.slice(start<0?source.indexOf('  async function callJson('):start,source.indexOf('  function formatReadinessMessage'))+source.slice(source.indexOf('  async function logClientEvent('),source.indexOf('  window.MailyOffice ='))+'\nglobalThis.call=callJson;globalThis.log=logClientEvent;',ctx);
 return{ctx,elements,requests};
}
async function begin(s){const pending=s.ctx.call('/api/outlook/synthetic',{method:'POST',headers:{'X-Synthetic-Request':'preserved'},body:'{"synthetic":true}'}).then(value=>({value}),error=>({error}));for(let i=0;i<8;i++)await Promise.resolve();assert.ok(s.elements.find(e=>e.tag==='dialog'&&e.open),'401 must offer an accessible key dialog');return pending;}
// Do not await begin's returned operation until interacting with the dialog.
async function start(s){let resolve;const result=new Promise(r=>resolve=r);begin(s).then(resolve);for(let i=0;i<12;i++)await Promise.resolve();assert.ok(s.elements.find(e=>e.tag==='dialog'&&e.open),'401 must show dialog');return{result};}
let s=setup([401,200]), operation=await start(s);
let input=s.elements.find(e=>e.tag==='input');assert.equal(input.type,'password');input.value='synthetic-session-key';s.elements.find(e=>e.tag==='form').listeners.submit({preventDefault(){}});assert.ok((await operation.result).value.ok);assert.equal(s.requests[1].headers['X-Maily-Api-Key'],'synthetic-session-key');assert.equal(s.requests[1].body,s.requests[0].body);assert.equal(input.value,'');await s.ctx.call('/api/health');assert.equal(s.requests[2].headers['X-Maily-Api-Key'],'synthetic-session-key');
for(const request of s.requests) assert.equal(request.redirect,'error','all API attempts must refuse redirects');
assert.equal(s.requests[1].headers['X-Synthetic-Request'],'preserved');
assert.equal(s.requests[1].headers['Content-Type'],'application/json');
assert.equal(s.requests[2].redirect,'error');await s.ctx.call('https://example.invalid/api/test');assert.equal(s.requests[3].headers['X-Maily-Api-Key'],undefined);
await s.ctx.call('//example.invalid/api/test');assert.equal(s.requests.at(-1).headers['X-Maily-Api-Key'],undefined);
await s.ctx.call('/assets/example');assert.equal(s.requests.at(-1).headers['X-Maily-Api-Key'],undefined);
s=setup([401]);operation=await start(s);s.elements.find(e=>e.tag==='button'&&e.textContent==='Cancel').listeners.click();assert.ok((await operation.result).error);assert.equal(s.requests.length,1);
s=setup([401,401]);operation=await start(s);s.elements.find(e=>e.tag==='input').value='incorrect-synthetic';s.elements.find(e=>e.tag==='form').listeners.submit({preventDefault(){}});assert.ok((await operation.result).error);assert.equal(s.requests.length,2);

s=setup([401]);operation=await start(s);s.elements.find(e=>e.tag==='form').listeners.submit({preventDefault(){}});assert.equal(s.requests.length,1);s.elements.find(e=>e.tag==='button'&&e.textContent==='Cancel').listeners.click();await operation.result;
s=setup([200]);await s.ctx.call('/api/health');assert.equal(s.requests[0].headers['X-Maily-Api-Key'],undefined);
console.log('PASS protected pane retry, cancellation, rejected/empty key, and session-only key');
// A startup request rejected after another request authenticated must reuse the current key.
let releaseLate;
const late401=new Promise(resolve=>{releaseLate=resolve;});
s=setup([401,late401,200,200]);operation=await start(s);
const late=s.ctx.call('/api/outlook/settings');
s.elements.find(e=>e.tag==='input').value='synthetic-race-key';s.elements.find(e=>e.tag==='form').listeners.submit({preventDefault(){}});
await operation.result;releaseLate(401);
for(let i=0;i<16;i++) await Promise.resolve();
assert.equal(s.elements.filter(e=>e.tag==='dialog').length,1,'late 401 must not prompt again');
await late;assert.equal(s.requests[3].headers['X-Maily-Api-Key'],'synthetic-race-key');
// Actual startup telemetry must neither request credentials nor repeat a cancelled prompt.
s=setup([401,401,401]);
const openingLog=s.ctx.log('taskpane-open','ok','','taskpane');
for(let i=0;i<12;i++) await Promise.resolve();
assert.equal(s.elements.filter(e=>e.tag==='dialog').length,0,'opening telemetry must not prompt');
await openingLog;operation=await start(s);
s.elements.find(e=>e.tag==='button'&&e.textContent==='Cancel').listeners.click();await operation.result;
const failureLog=s.ctx.log('taskpane-start-failed','error','','taskpane');
for(let i=0;i<12;i++) await Promise.resolve();
assert.equal(s.elements.filter(e=>e.tag==='dialog').length,1,'failure telemetry must not repeat cancelled prompt');
await failureLog;
console.log('PASS delayed authentication race and noninteractive startup telemetry');
let releaseLog;
const delayedLog=new Promise(resolve=>{releaseLog=resolve;});
s=setup([401,delayedLog,200,200]);operation=await start(s);
const background=s.ctx.log('taskpane-open','ok','','taskpane');
s.elements.find(e=>e.tag==='input').value='synthetic-accepted-key';s.elements.find(e=>e.tag==='form').listeners.submit({preventDefault(){}});
await operation.result;releaseLog(401);await background;
await s.ctx.call('/api/health');
assert.equal(s.requests[3].headers['X-Maily-Api-Key'],'synthetic-accepted-key','late telemetry rejection must not clear the accepted key');
assert.equal(s.elements.filter(e=>e.tag==='dialog').length,1);
console.log('PASS late client-log rejection preserves authenticated session');
