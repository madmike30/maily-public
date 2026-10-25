import fs from 'node:fs';
import vm from 'node:vm';
import assert from 'node:assert/strict';

// No network or COM capability is supplied to this context. All mailbox state is synthetic.
const mailbox = 'maily-test@example.invalid';
const folder = 'Maily Synthetic Tests';
const source = fs.readFileSync(new URL('../../services/Maily.OutlookAddin.Web/wwwroot/addin/js/maily-office.js', import.meta.url), 'utf8');
const results = [];
function fixture(type = 'text') {
  const state = { body: type === 'html' ? '<p>Regards,</p><p>Test Sender</p>' : 'Regards,\nTest Sender', writes: 0, sends: 0 };
  const ok = value => ({ status: 'succeeded', value });
  const item = {
    subject: 'Synthetic שלום 🧪', from: { displayName: 'Test Sender', emailAddress: mailbox }, to: [], cc: [],
    body: {
      getTypeAsync: cb => cb(ok(type)),
      getAsync: (_type, cb) => cb(ok(state.body)),
      setAsync: (value, _options, cb) => { state.body = value; state.writes++; cb(ok()); },
      prependAsync: (value, _options, cb) => { state.body = value + state.body; state.writes++; cb(ok()); }
    },
    sendAsync: () => { state.sends++; throw new Error('LIVE SEND DENIED'); }
  };
  const context = vm.createContext({ window: { location: { hostname: 'localhost', search: '' } },
    Office: { context: { mailbox: { item, userProfile: { emailAddress: mailbox, displayName: 'Test Sender' } }, platform: 'OfficeOnline' },
      CoercionType: { Html: 'html', Text: 'text' }, AsyncResultStatus: { Succeeded: 'succeeded' } },
    ReadableStream, TextDecoder, Uint8Array, URLSearchParams,
    fetch: async () => { throw new Error('No fixture response configured; network forbidden'); } });
  vm.runInContext(source, context);
  return { state, context, office: context.window.MailyOffice, item };
}
async function test(name, action) {
  try { await action(); results.push({ name, status: 'PASS' }); console.log(`PASS ${name}`); }
  catch (error) { results.push({ name, status: 'FAIL', error: String(error.stack) }); console.error(`FAIL ${name}: ${error.stack}`); }
}
const json = value => ({ ok: true, json: async () => value });
await test('compose summary separates incoming email from typed reply and signature', async () => {
  for (const [draft, incoming] of [['תודה אבדוק ואחזור מחר.', 'Please confirm the project schedule tomorrow.'], ['I will confirm tomorrow.', 'אנא אשרו את התוכנית המעודכנת מחר.']]) {
    const f = fixture();
    f.item.subject = { getAsync: cb => cb({status:'succeeded',value:'Re: Plan'}) };
    for (const name of ['to','cc']) f.item[name] = { getAsync: cb => cb({status:'succeeded',value:[]}) };
    f.state.body = `${draft}\nRegards,\nTest Sender\nFrom: Other <other@example.invalid>\nSubject: Plan\n${incoming}`;
    const snapshot = await f.office.getComposeSnapshot({userName:'Test Sender'}, 'generate-initial');
    assert.ok(snapshot.composeState.summaryBodyText.includes(incoming));
    assert.ok(!snapshot.composeState.summaryBodyText.includes(draft));
    assert.ok(!snapshot.composeState.summaryBodyText.includes('Regards,'));
    assert.equal(snapshot.request.compose.bodyText,f.state.body,'Reply generation keeps full exchange context');
    const pane = fs.readFileSync(new URL('../../services/Maily.OutlookAddin.Web/wwwroot/addin/js/summary.js', import.meta.url), 'utf8');
    f.context.document = {getElementById: () => ({})};
    // Expose the real request builder without starting the asynchronous pane lifecycle.
    vm.runInContext(pane.slice(0,pane.indexOf('  start().catch')) + 'window.buildSummaryRequest = buildRequest;})();', f.context);
    const payload = await f.context.window.buildSummaryRequest();
    assert.equal(payload.bodyText, snapshot.composeState.summaryBodyText);
    assert.ok(payload.userDraftText.includes(draft));
    assert.equal(f.office.detectPreferredLanguage(payload.bodyText), incoming.includes('אנא') ? 'Hebrew' : 'Auto');
    assert.equal(f.state.writes,0);
  }
});
await test('native rewrite rejection cannot fall back to overwriting a newer draft', async () => {
  const f = fixture(); f.state.body = 'Updated draft B\nRegards,\nTest Sender';
  f.context.Office.context.platform = 'PC'; f.context.Office.context.diagnostics = { hostName: 'Outlook' };
  f.context.fetch = async url => url.includes('client-log') ? json({}) : { ok: false, status: 409, text: async () => '{"error":"Compose changed; regenerate"}' };
  const before = f.state.body;
  await assert.rejects(f.office.insertGeneratedDraft({ text: 'Generated A' }, {
    rewriteComposeBody: true, seedText: 'Original draft A', originalUserDraftSegment: 'Original draft A',
    preservedSuffix: '\nRegards,\nTest Sender', composeDefaults: { userName: 'Test Sender' }
  }), /changed|rewrite|regenerate/i);
  assert.equal(f.state.body, before); assert.equal(f.state.writes, 0);
});
await test('web rewrite rejects a body changed since generation snapshot', async () => {
  const f = fixture(); f.state.body = 'Original draft with new edits\nRegards,\nTest Sender';
  const before = f.state.body;
  await assert.rejects(f.office.insertGeneratedDraft({ text: 'Generated A' }, {
    rewriteComposeBody: true, seedText: 'Original draft', originalUserDraftSegment: 'Original draft',
    expectedBody: 'Original draft\nRegards,\nTest Sender', expectedBodyType: 'text',
    preservedSuffix: '\nRegards,\nTest Sender', composeDefaults: { userName: 'Test Sender' }
  }), /changed|regenerate/i);
  assert.equal(f.state.body, before); assert.equal(f.state.writes, 0);
});
await test('web rewrite preserves edits made between the two body reads', async () => {
  const f = fixture(); f.state.body = 'Original draft\nRegards,\nTest Sender'; const expectedBody = f.state.body;
  let reads = 0;
  f.item.body.getAsync = (_type, cb) => { if (++reads === 2) f.state.body += '\nNew user edit'; cb({ status: 'succeeded', value: f.state.body }); };
  await assert.rejects(f.office.insertGeneratedDraft({ text: 'Generated A' }, {
    rewriteComposeBody: true, seedText: 'Original draft', originalUserDraftSegment: 'Original draft',
    expectedBody, expectedBodyType: 'text', preservedSuffix: '\nRegards,\nTest Sender'
  }), /changed/);
  assert.equal(f.state.writes, 0); assert.ok(f.state.body.endsWith('New user edit'));
});
await test('unchanged web draft still rewrites with signature preserved', async () => {
  const f = fixture(); f.state.body = 'Original draft\nRegards,\nTest Sender'; const expectedBody = f.state.body;
  await f.office.insertGeneratedDraft({ text: 'Generated replacement' }, {
    rewriteComposeBody: true, seedText: 'Original draft', originalUserDraftSegment: 'Original draft',
    expectedBody, expectedBodyType: 'text', preservedSuffix: '\nRegards,\nTest Sender'
  });
  assert.equal(f.state.writes, 1); assert.ok(!f.state.body.includes('Original draft')); assert.ok(f.state.body.endsWith('Regards,\nTest Sender'));
});
await test('generation result cannot enter another item with an identical body', async () => {
  const f = fixture(); f.item.itemId = 'synthetic-item-a';
  const expectedContext = f.office.captureComposeContext ? f.office.captureComposeContext() : { item: f.item };
  f.context.Office.context.mailbox.item = { ...f.item, itemId: 'synthetic-item-b' };
  await assert.rejects(f.office.insertGeneratedDraft({ text: 'Generated for A' }, { expectedContext }), /changed|current draft/i);
  assert.equal(f.state.writes, 0);
});
await test('item switch during body read prevents Office write', async () => {
  const f = fixture(); const expectedContext = f.office.captureComposeContext();
  f.item.body.getAsync = (_type, cb) => { f.context.Office.context.mailbox.item = { ...f.item }; cb({ status: 'succeeded', value: f.state.body }); };
  await assert.rejects(f.office.insertGeneratedDraft({ text: 'Generated A' }, { expectedContext }), /item changed/);
  assert.equal(f.state.writes, 0);
});
await test('ItemChanged epoch invalidates a reused unsaved item object', async () => {
  const f = fixture(); let onChange;
  f.context.Office.EventType = { ItemChanged: 'itemChanged' };
  f.context.Office.context.mailbox.addHandlerAsync = (_event, handler) => { onChange = handler; };
  const expectedContext = f.office.captureComposeContext(); onChange();
  await assert.rejects(f.office.insertGeneratedDraft({ text: 'Generated A' }, { expectedContext }), /item changed/);
  assert.equal(f.state.writes, 0);
});
await test('classic generation writes the captured Office item without selecting a native window', async () => {
  const f = fixture(); f.state.body = 'Original draft\nRegards,\nTest Sender';
  f.context.Office.context.platform = 'PC'; f.context.Office.context.diagnostics = { hostName: 'Outlook' };
  const expectedContext = f.office.captureComposeContext();
  const result = await f.office.insertGeneratedDraft({ text: 'Generated replacement' }, {
    expectedContext, rewriteComposeBody: true, expectedBody: f.state.body, expectedBodyType: 'text',
    originalUserDraftSegment: 'Original draft', preservedSuffix: '\nRegards,\nTest Sender'
  });
  assert.equal(result.strategy, 'office-js'); assert.equal(f.state.writes, 1);
});
await test('repeated plaintext generation preserves exact signature spacing', async () => {
  const f = fixture();
  const draft = { text: 'Synthetic generated reply' };
  await f.office.insertGeneratedDraft(draft, { replaceExisting: true });
  const first = f.state.body;
  for (let iteration = 0; iteration < 4; iteration++) {
    await f.office.insertGeneratedDraft(draft, { replaceExisting: true });
    assert.equal(f.state.body, first);
  }
  assert.ok(f.state.body.endsWith('Regards,\nTest Sender'));
});
await test('stale insertion retry uses the current edited body and discards the rejected preview', async () => {
  const f = fixture();
  const nodes = new Map();
  for (const id of ['status', 'generateBtn', 'insertBtn', 'applyBtn', 'undoBtn', 'shorterBtn', 'clearerBtn', 'warmerBtn', 'replyLanguageEnglishBtn', 'replyLanguageHebrewBtn', 'seedPanel', 'seedMeta', 'seedText', 'draftText', 'draftMeta']) {
    nodes.set(id, { textContent: '', disabled: false, classList: { add() {}, remove() {}, toggle() {} },
      setAttribute() {}, addEventListener(event, handler) { this[event] = handler; } });
  }
  f.context.document = { getElementById: id => nodes.get(id) || null };
  f.context.Office.onReady = callback => callback();
  f.item.subject = { getAsync: callback => callback({ status: 'succeeded', value: 'Synthetic meeting' }) };
  for (const field of ['to', 'cc']) f.item[field] = { getAsync: callback => callback({ status: 'succeeded', value: [] }) };
  // Snapshot and insertion use the production bridge; all external boundaries are mocked.
  f.office.ensureBackendAvailable = async () => {};
  f.office.logClientEvent = async () => {};
  const requests = [];
  f.office.callJson = async (url, options) => {
    if (url.includes('settings')) return { values: { userName: 'Test Sender' } };
    assert.equal(url, '/api/outlook/generate-reply');
    requests.push(JSON.parse(options.body));
    if (requests.length === 1) f.state.body = 'Meet Tuesday\nRegards,\nTest Sender';
    return { draft: { text: requests.length === 1 ? 'Obsolete Monday answer' : 'Confirmed Tuesday answer' } };
  };
  f.state.body = 'Meet Monday\nRegards,\nTest Sender';
  vm.runInContext(fs.readFileSync(new URL('../../services/Maily.OutlookAddin.Web/wwwroot/addin/js/taskpane.js', import.meta.url), 'utf8'), f.context);
  async function waitIdle() {
    for (let i = 0; i < 100; i++) {
      await new Promise(resolve => setImmediate(resolve));
      if (requests.length && !nodes.get('generateBtn').disabled) return;
    }
    throw new Error('Task pane failed to settle: ' + nodes.get('status').textContent);
  }
  await waitIdle();
  assert.equal(requests.length, 1);
  assert.equal(f.state.writes, 0, 'Preview cannot edit Outlook before explicit Apply');
  await nodes.get('applyBtn').click();
  assert.match(nodes.get('status').textContent, /changed|regenerate/i);
  assert.equal(f.state.writes, 0);
  assert.equal(f.state.body, 'Meet Tuesday\nRegards,\nTest Sender');
  assert.equal(nodes.get('draftText').textContent, 'No generated draft yet.');
  await nodes.get('applyBtn').click();
  assert.equal(f.state.writes, 0, 'Consumed stale preview cannot be applied again');
  assert.equal(nodes.get('applyBtn').disabled, true);
  nodes.get('generateBtn').click();
  await waitIdle();
  assert.equal(requests.length, 2);
  assert.equal(requests[1].operation.currentDraft, 'Meet Tuesday');
  assert.equal(requests[1].operation.mode, 'generate-initial');
  assert.equal(f.state.writes, 0, 'Retry also waits for explicit Apply');
  await nodes.get('applyBtn').click();
  assert.equal(f.state.writes, 1);
  assert.ok(f.state.body.includes('Confirmed Tuesday answer'));
  assert.ok(!f.state.body.includes('Monday'));
  assert.ok(!f.state.body.includes('Meet Tuesday'));
  assert.equal(f.state.sends, 0);
});
await test('regeneration uses the live marked reply, including edits around the generated block', async () => {
  const f = fixture();
  f.item.subject = { getAsync: cb => cb({ status: 'succeeded', value: 'Synthetic planning' }) };
  for (const name of ['to','cc']) f.item[name] = { getAsync: cb => cb({ status: 'succeeded', value: [] }) };
  f.state.body = 'Please retain the revised date.\n[Maily Reply Start]\nMeet Tuesday at 11.\n[Maily Reply End]\nBring the signed copy.\nRegards,\nTest Sender\nFrom: Other <other@example.invalid>\nSubject: Plan\nQuoted Monday request';
  const snapshot = await f.office.getComposeSnapshot({ userName: 'Test Sender', currentDraft: 'Obsolete Monday answer' }, 'regenerate');
  assert.equal(snapshot.request.operation.currentDraft, 'Please retain the revised date.\n\nMeet Tuesday at 11.\n\nBring the signed copy.');
  assert.ok(!snapshot.request.operation.currentDraft.includes('Monday'));
  assert.ok(!snapshot.request.operation.currentDraft.includes('Maily Reply'));
  assert.ok(snapshot.composeState.preservedSuffix.includes('Quoted Monday request'));
  assert.equal(f.state.writes, 0);
});
await test('each rewrite replaces the current editable reply, not a cached result or only the marked block', async () => {
  const f = fixture();
  const nodes = new Map();
  for (const id of ['status','generateBtn','insertBtn','applyBtn','undoBtn','shorterBtn','clearerBtn','warmerBtn','replyLanguageEnglishBtn','replyLanguageHebrewBtn','seedPanel','seedMeta','seedText','draftText','draftMeta']) {
    nodes.set(id,{textContent:'',disabled:false,classList:{add(){},remove(){},toggle(){}},setAttribute(){},addEventListener(event,handler){this[event]=handler;}});
  }
  f.context.document={getElementById:id=>nodes.get(id)||null};
  f.context.Office.onReady=callback=>callback();
  f.item.subject={getAsync:cb=>cb({status:'succeeded',value:'Synthetic planning'})};
  for(const name of ['to','cc']) f.item[name]={getAsync:cb=>cb({status:'succeeded',value:[]})};
  f.office.ensureBackendAvailable=async()=>{}; f.office.logClientEvent=async()=>{};
  const requests=[];
  f.office.callJson=async(url,options)=>{
    if(url.includes('settings')) return {values:{userName:'Test Sender'}};
    requests.push(JSON.parse(options.body));
    return {draft:{text:'Polished version '+requests.length}};
  };
  const suffix='Regards,\nTest Sender\nFrom: Other <other@example.invalid>\nSubject: Planning\nOriginal quoted request';
  f.state.body='Meet Tuesday at 11.\n'+suffix;
  vm.runInContext(fs.readFileSync(new URL('../../services/Maily.OutlookAddin.Web/wwwroot/addin/js/taskpane.js',import.meta.url),'utf8'),f.context);
  async function settled(){for(let i=0;i<100;i++){await new Promise(resolve=>setImmediate(resolve));if(requests.length&&!nodes.get('generateBtn').disabled)return;}throw new Error(nodes.get('status').textContent);}
  await settled();
  assert.equal(requests[0].operation.currentDraft,'Meet Tuesday at 11.');
  assert.equal(f.state.writes,0,'Initial preview must not edit Outlook');
  await nodes.get('applyBtn').click();
  assert.ok(!f.state.body.includes('Meet Tuesday at 11.'));
  assert.ok(f.state.body.endsWith(suffix));
  f.state.body='Please keep the revised date.\n'+f.state.body.replace('Polished version 1','Meet Wednesday at 12.');
  nodes.get('generateBtn').click(); await settled();
  assert.match(requests[1].operation.currentDraft,/Please keep the revised date/);
  assert.match(requests[1].operation.currentDraft,/Meet Wednesday at 12/);
  assert.equal(f.state.writes,1,'Regeneration preview must not edit Outlook');
  await nodes.get('applyBtn').click();
  assert.ok(!f.state.body.includes('Please keep the revised date'));
  assert.ok(!f.state.body.includes('Meet Wednesday'));
  assert.ok(f.state.body.endsWith(suffix));
  assert.equal(f.state.body.split('[Maily Reply Start]').length-1,1);
  f.state.body='Final user edit, without any generated markers.\n'+suffix;
  nodes.get('generateBtn').click(); await settled();
  assert.equal(requests[2].operation.currentDraft,'Final user edit, without any generated markers.');
  assert.equal(f.state.writes,2,'Unmarked draft preview must not edit Outlook');
  await nodes.get('applyBtn').click();
  assert.ok(!f.state.body.includes('Final user edit'));
  assert.ok(f.state.body.endsWith(suffix));
  f.state.body=f.state.body.replace('Polished version 3','Latest user wording.');
  await nodes.get('insertBtn').click(); await settled();
  assert.equal(requests[3].operation.currentDraft,'Latest user wording.');
  assert.equal(f.state.writes,3,'Legacy button prepares a preview only');
  await nodes.get('applyBtn').click();
  assert.ok(!f.state.body.includes('Latest user wording'));
  assert.equal(f.state.body.split('[Maily Reply Start]').length-1,1);
  const snapshotReader = f.office.getComposeSnapshot;
  f.office.getComposeSnapshot = async (...args) => {
    const snapshot = await snapshotReader(...args);
    snapshot.composeState.canRewriteUserDraft = false;
    return snapshot;
  };
  const writesBeforeAmbiguous = f.state.writes;
  nodes.get('generateBtn').click(); await settled();
  assert.equal(requests.length,4,'ambiguous protected boundary must stop before generation');
  assert.equal(f.state.writes,writesBeforeAmbiguous);
  assert.equal(f.state.sends,0);
});
console.log(JSON.stringify({ mode: 'DRY-RUN', mailbox, folder, liveConnections: 0, results }));
if (results.some(r => r.status === 'FAIL')) process.exitCode = 1;
