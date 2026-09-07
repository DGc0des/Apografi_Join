const { test } = require('node:test');
const assert = require('node:assert/strict');
const vm = require('node:vm');
const fs = require('node:fs');
const path = require('node:path');
const source = fs.readFileSync(path.join(__dirname, '../app.js'), 'utf8')
  .replace(/initThemeFromStorage\(\);\s*init\(\);\s*$/, '');
function harness() {
  const saved = new Map();
  const events = {};
  const inputs = [2,3].map(row => ({ dataset: { row: String(row) }, value: '', closest: () => null }));
  const elements = { app: {innerHTML:''}, 'admin-save-status': {}, 'admin-save-btn': {}, 'admin-back-modal': {remove(){}} };
  const context = vm.createContext({
    window: {addEventListener: (name, fn) => events[name] = fn},
    document: {body:{classList:{contains:()=>false}}, addEventListener(){}, getElementById: id => elements[id] || null,
      querySelectorAll: selector => selector === '.admin-weight-input' ? inputs : [], querySelector:()=>null},
    localStorage: {getItem:()=>null},
    sessionStorage: { getItem:k=>saved.get(k) || null, setItem:(k,v)=>saved.set(k,v), removeItem:k=>saved.delete(k)},
    console: {error(){}}, setTimeout(){},
  });
  vm.runInContext(source, context);
  const run = code => vm.runInContext(code, context);
  run("storeId=1; countDate='2026-09-07'; showToast=()=>{}; adminOriginalValues={2:'',3:''};");
  return {run, inputs, saved, events, elements};
}
test('accented Greek, decomposed accents, final sigma, and whitespace match', () => {
  const h=harness();
  assert.equal(h.run("normalizeSearch('  Γάλα  ')"), 'γαλα');
  assert.equal(h.run("normalizeSearch('γάλα')"), 'γαλα');
  assert.equal(h.run("normalizeSearch('καφές') === normalizeSearch('ΚΑΦΕΣ')"), true);
  h.run("searchQuery='γάλα'");
  assert.match(h.run('buildItemsList()'), /ΓΑΛΑ SPECIALISTAS/);
});
test('settings provides return navigation for an existing connection', () => {
  const h=harness();h.run('sb={}; renderSetupScreen()');
  assert.match(h.elements.app.innerHTML, /cancelSetup/);
  h.run('cancelSetup()');assert.match(h.elements.app.innerHTML,/store-screen/);
});
test('multiple draft values including cleared weights survive rendering', () => {
  const h=harness();h.run("containerWeights={2:0.5}; adminOriginalValues={2:'0.5',3:''}");
  h.inputs[1].value='0.25';h.run("onAdminWeightInput(document.querySelectorAll('.admin-weight-input')[1])");
  h.inputs[0].value='0.5';h.inputs[1].value='';h.run('renderAdminScreen()');
  assert.equal(h.inputs[0].value,'');assert.equal(h.inputs[1].value,'0.25');
  assert.equal(h.run('adminDirty'),true);
});
test('draft is scoped to store and discard removes it', () => {
  const h=harness();h.inputs[0].value='0.25';h.run("onAdminWeightInput(document.querySelectorAll('.admin-weight-input')[0])");
  h.run('storeId=2');assert.equal(h.run('sessionStorage.getItem(adminDraftKey())'),null);
  h.run('storeId=1; renderCountingScreen=()=>{}; discardAdminChanges()');
  assert.equal(h.saved.size,0);assert.equal(h.run('adminDirty'),false);
});
test('unload warning only applies to unsaved changes', () => {
  const h=harness();let prevented=false;const e={preventDefault(){prevented=true}};
  h.events.beforeunload(e);assert.equal(prevented,false);
  h.run('adminDirty=true');h.events.beforeunload(e);assert.equal(prevented,true);
});
test('failed delete keeps draft and dirty status', async () => {
  const h=harness();h.run("containerWeights={2:0.5}; adminOriginalValues={2:'0.5',3:''}; onAdminWeightInput(document.querySelectorAll('.admin-weight-input')[0]); sb={from:()=>({delete:()=>({eq:()=>({eq:async()=>({error:new Error('failed')})})})})}");
  await h.run('saveContainerWeights()');assert.equal(h.run('adminDirty'),true);
  assert.ok(h.run('sessionStorage.getItem(adminDraftKey())'));assert.equal(h.elements['admin-save-btn'].disabled,false);
});
test('successful save clears draft, but retains edits made during a request', async () => {
  const h=harness();h.inputs[0].value='0.25';
  h.run("onAdminWeightInput(document.querySelectorAll('.admin-weight-input')[0]); sb={from:()=>({upsert:async()=>({error:null})})}");
  await h.run('saveContainerWeights()');assert.equal(h.run('adminDirty'),false);assert.equal(h.saved.size,0);
  h.inputs[0].value='0.5';h.run("onAdminWeightInput(document.querySelectorAll('.admin-weight-input')[0])");
  const saving=h.run('saveContainerWeights()');h.inputs[0].value='0.75';await saving;
  assert.equal(h.run('adminDirty'),true);assert.equal(JSON.parse(h.run('sessionStorage.getItem(adminDraftKey())'))[2],'0.75');
});
