import fs from 'node:fs';
import vm from 'node:vm';
import assert from 'node:assert/strict';
const html=fs.readFileSync('services/Maily.OutlookAddin.Web/wwwroot/addin/taskpane.html','utf8');
const nodes=new Map([...html.matchAll(/id="([^"]+)"/g)].map(m=>[m[1],{textContent:'',hidden:false,disabled:false,open:false,listeners:{},className:'',classList:{add(){},remove(){},toggle(){}},setAttribute(){},addEventListener(n,fn){this.listeners[n]=fn;}}]));
let item={id:'synthetic-item'},body='We can supply 37 on Friday.',inserts=0,undos=0,undoAvailable=false,requests=[],deferred=null,lastWritten='';
const originalBody=body;
const office={
 officeReady:async()=>{},ensureBackendAvailable:async()=>{},logClientEvent:async()=>{},getMailboxAddress:()=> 'author@example.test',getMailboxDisplayName:()=> 'Synthetic Author',detectPreferredLanguage:()=> 'Auto',applyDirectionToElement(){},
 captureComposeContext:()=>item,assertComposeContext(ctx){if(ctx!==item)throw Error('Item changed');},
 async getComposeSnapshot(defaults,mode){return {request:{compose:{subject:'Synthetic delivery',bodyText:body},operation:{mode,currentDraft:body},userPreferences:defaults},composeState:{body,bodyType:'text',userDraftText:body,canRewriteUserDraft:true,userDraftSegment:body,preservedSuffix:'',composeType:'reply'}};},
 async callJson(url,options){if(url.includes('/settings'))return {values:{tone:'professional',language:'English'}};const request=JSON.parse(options.body);requests.push(request);if(deferred)return await deferred;return {draft:{text:'We can supply 37 on Friday. Please confirm.',html:'<p>We can supply 37 on Friday. Please confirm.</p>'},meta:{}};},
 async insertGeneratedDraft(draft,options){assert.equal(options.expectedContext,item);if(options.expectedBody!==body)throw Error('Draft changed');body=draft.text;lastWritten=body;inserts++;undoAvailable=true;return {mode:'replace-user-draft',strategy:'office'};},
 canUndoGeneratedDraft:()=>undoAvailable,
 async undoGeneratedDraft(){if(body!==lastWritten){undoAvailable=false;throw Error("Draft changed; Undo refused");}body=originalBody;undos++;undoAvailable=false;}
};
const context={window:{MailyOffice:office},document:{getElementById:id=>nodes.get(id)},Office:{},console,Error};
vm.runInNewContext(fs.readFileSync('services/Maily.OutlookAddin.Web/wwwroot/addin/js/taskpane.js','utf8'),context);
const settle=async()=>{await new Promise(setImmediate);await new Promise(setImmediate);};
const click=async id=>{assert.ok(nodes.has(id),`${id} exists`);if(!nodes.get(id).disabled)await nodes.get(id).listeners.click();await settle();};
await settle();
assert.equal(requests.length,1);
assert.equal(inserts,0,'Generating preview must not change Outlook');
assert.equal(nodes.get('seedText').textContent,originalBody);
assert.match(nodes.get('draftText').textContent,/37 on Friday/);
await click('applyBtn');
assert.equal(inserts,1);assert.equal(nodes.get('applyBtn').disabled,true);assert.equal(nodes.get('undoBtn').disabled,false);
await click('undoBtn');assert.equal(undos,1);assert.equal(nodes.get('undoBtn').disabled,true);
await click('shorterBtn');
assert.equal(inserts,1,'Refinement is preview-only');
assert.equal(requests.at(-1).operation.currentDraft,body,'Refinement uses actual Outlook draft');
assert.match(requests.at(-1).operation.userInstruction,/shorter/i);
assert.match(requests.at(-1).operation.userInstruction,/commitments/i);
body='New manual edit';await click('applyBtn');assert.equal(inserts,1);assert.match(nodes.get('status').textContent,/changed/);
await click('clearerBtn');assert.match(requests.at(-1).operation.userInstruction,/clearer/i);
await click('warmerBtn');assert.match(requests.at(-1).operation.userInstruction,/warmer/i);assert.equal(inserts,1);
let finish;deferred=new Promise(resolve=>finish=resolve);
const generating=click('generateBtn');await settle();item={id:'other-synthetic-item'};finish({draft:{text:'Stale result'},meta:{}});await generating;
assert.equal(inserts,1);assert.equal(nodes.get('applyBtn').disabled,true);assert.match(nodes.get('status').textContent,/changed/);assert.equal(nodes.get('draftText').textContent,'No generated draft yet.');
console.log('PASS preview-before-apply, comparison, refinements, authoritative current draft, stale-body/item rejection and Undo UI');
