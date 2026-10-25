import fs from 'node:fs';
import vm from 'node:vm';
import assert from 'node:assert/strict';
const source = fs.readFileSync(new URL('../../services/Maily.OutlookAddin.Web/wwwroot/addin/js/maily-office.js', import.meta.url), 'utf8');
function fixture(type = 'text') {
  const state = { type, body: type === 'html' ? '<p>Original reply.</p><div>Regards,</div><div>Synthetic Author</div>' : 'Original reply.\nRegards,\nSynthetic Author', writes: 0 };
  const ok = value => ({ status: 'succeeded', value });
  const item = { itemId: 'synthetic-one', subject: 'Synthetic subject', body: {
    getTypeAsync: cb => cb(ok(state.type)),
    getAsync: (_type, cb) => cb(ok(state.body)),
    setAsync: (value, _options, cb) => { state.body = value; state.writes++; cb(ok()); },
    prependAsync: (value, _options, cb) => { state.body = value + state.body; state.writes++; cb(ok()); }
  } };
  const context = vm.createContext({ window: { location: { hostname: 'localhost', search: '' } }, URLSearchParams,
    Office: { context: { mailbox: { item, userProfile: { emailAddress: 'author@example.invalid' } }, platform: 'OfficeOnline' },
      CoercionType: { Html: 'html', Text: 'text' }, AsyncResultStatus: { Succeeded: 'succeeded' } },
    fetch: () => { throw Error('Undo must not use network'); } });
  vm.runInContext(source, context);
  return { state, item, context, office: context.window.MailyOffice };
}
async function insert(f) {
  const before = f.state.body;
  await f.office.insertGeneratedDraft({ text: 'Improved reply.', html: '<p>Improved reply.</p>' }, {
    expectedContext: f.office.captureComposeContext(), expectedBody: before, expectedBodyType: f.state.type,
    rewriteComposeBody: true, originalUserDraftSegment: f.state.type === 'html' ? '<p>Original reply.</p>' : 'Original reply.'
  });
  return before;
}
for (const type of ['text', 'html']) {
  const f = fixture(type);
  const before = await insert(f);
  assert.equal(f.office.canUndoGeneratedDraft(), true);
  await f.office.undoGeneratedDraft();
  assert.equal(f.state.body, before, 'Undo restores original body including exact signature');
  assert.equal(f.state.writes, 2);
  assert.equal(f.office.canUndoGeneratedDraft(), false);
  await assert.rejects(f.office.undoGeneratedDraft(), /undo|restore/i);
}
console.log('PASS explicit same-draft Undo restores exact HTML/text once, without network');
{
  const f = fixture(); await insert(f);
  f.state.body += '\nNew manual edit.';
  const edited = f.state.body;
  await assert.rejects(f.office.undoGeneratedDraft(), /newer edits|changed/i);
  assert.equal(f.state.body, edited); assert.equal(f.state.writes, 1);
  assert.equal(f.office.canUndoGeneratedDraft(), false);
}
console.log('PASS later manual edits prevent Undo and remain intact');
{
  const f = fixture(); await insert(f);
  const other = { ...f.item, itemId: 'synthetic-two' };
  f.context.Office.context.mailbox.item = other;
  assert.equal(f.office.canUndoGeneratedDraft(), false);
  await assert.rejects(f.office.undoGeneratedDraft(), /undo/i);
  assert.equal(f.state.writes, 1);
}
console.log('PASS identical-body item switch cannot receive Undo');
{
  const f = fixture(); await insert(f);
  let reads = 0;
  f.item.body.getAsync = (_type, cb) => {
    if (++reads === 2) f.state.body += '\nEdit during Undo.';
    cb({ status: 'succeeded', value: f.state.body });
  };
  await assert.rejects(f.office.undoGeneratedDraft(), /newer edits|changed/i);
  assert.ok(f.state.body.includes('Edit during Undo.')); assert.equal(f.state.writes, 1);
}
console.log('PASS edit between Undo reads prevents replacement');
{
  const f = fixture(); await insert(f); f.state.type = 'html';
  await assert.rejects(f.office.undoGeneratedDraft(), /changed/i);
  assert.equal(f.state.writes, 1);
}
console.log('PASS body format change prevents Undo');
{
  const f = fixture();
  f.item.body.setAsync = (value, _options, cb) => { f.state.body = value; f.state.writes++; cb({ status: 'failed' }); };
  await assert.rejects(insert(f));
  assert.equal(f.office.canUndoGeneratedDraft(), false);
  await assert.rejects(f.office.undoGeneratedDraft());
  assert.equal(f.state.writes, 1);
}
console.log('PASS ambiguous failed write never becomes an Undo candidate');
{
  const f = fixture();
  f.item.body.setAsync = (value, _options, cb) => { f.state.body = value + '\nEdit after write.'; f.state.writes++; cb({ status: 'succeeded' }); };
  await insert(f);
  assert.equal(f.office.canUndoGeneratedDraft(), false);
  await assert.rejects(f.office.undoGeneratedDraft());
  assert.ok(f.state.body.includes('Edit after write.')); assert.equal(f.state.writes, 1);
}
console.log('PASS unverified post-write body never enables Undo');
{
  const f = fixture(); await insert(f);
  f.item.body.setAsync = (value, _options, cb) => { f.state.body = value; f.state.writes++; cb({ status: 'failed' }); };
  await assert.rejects(f.office.undoGeneratedDraft());
  await assert.rejects(f.office.undoGeneratedDraft());
  assert.equal(f.state.writes, 2);
}
console.log('PASS failed Undo is consumed before mutation and cannot be retried blindly');
{
  const f = fixture(); let complete;
  f.item.body.setAsync = (value, _options, cb) => { f.state.body = value; f.state.writes++; complete = cb; };
  const first = insert(f);
  for (let i = 0; i < 12 && !complete; i++) await Promise.resolve();
  assert.equal(typeof complete, 'function');
  await assert.rejects(insert(f), /finish/i);
  await assert.rejects(f.office.undoGeneratedDraft(), /finish/i);
  complete({ status: 'succeeded' }); await first;
  assert.equal(f.state.writes, 1); assert.equal(f.office.canUndoGeneratedDraft(), true);
}
console.log('PASS concurrent body mutations are rejected');
{
  const f = fixture(); const original = f.state.body;
  await f.office.insertGeneratedDraft({ text: 'Inserted reply.' }, { expectedContext: f.office.captureComposeContext(), expectedBody: original });
  assert.equal(f.office.canUndoGeneratedDraft(), true);
  await f.office.undoGeneratedDraft(); assert.equal(f.state.body, original);
  await insert(f); const first = f.state.body;
  await f.office.insertGeneratedDraft({ text: 'Second reply.' }, { expectedContext: f.office.captureComposeContext(), expectedBody: first, replaceExisting: true });
  await f.office.undoGeneratedDraft(); assert.equal(f.state.body, first, 'Undo restores only the immediately preceding body');
}
console.log('PASS prepend and repeated generated replacement keep one latest Undo');
{
  const f = fixture(); let changed;
  f.context.Office.EventType = { ItemChanged: 'itemChanged' };
  f.context.Office.context.mailbox.addHandlerAsync = (_event, callback) => { changed = callback; };
  await insert(f); changed();
  assert.equal(f.office.canUndoGeneratedDraft(), false);
  await assert.rejects(f.office.undoGeneratedDraft()); assert.equal(f.state.writes, 1);
}
console.log('PASS Outlook item-change event immediately discards Undo');
