const fs = require('node:fs');
const vm = require('node:vm');
const assert = require('node:assert/strict');

class Sheet {
  constructor(rows = []) { this.rows = rows.map(row => [...row]); }
  getLastColumn() { return Math.max(0, ...this.rows.map(row => row.length)); }
  getDataRange() { return { getValues: () => this.rows.map(row => [...row]) }; }
  getRange(r, c, nr = 1, nc = 1) {
    return {
      getValues: () => Array.from({length:nr}, (_, i) =>
        Array.from({length:nc}, (_, j) => this.rows[r-1+i]?.[c-1+j] ?? '')),
      setValue: value => {
        this.rows[r-1] ||= [];
        this.rows[r-1][c-1] = value;
      },
      setValues: values => values.forEach((row, i) => row.forEach((value, j) => {
        this.rows[r-1+i] ||= [];
        this.rows[r-1+i][c-1+j] = value;
      }))
    };
  }
  appendRow(row) { this.rows.push([...row]); }
  setFrozenRows() {}
  deleteRow(index) { this.rows.splice(index-1, 1); }
}
const sheets = {
  Jugadores: new Sheet([
    ['id','nombre','genero','edad','posicion','rating','creado'],
    [7,'Antiguo','M',28,'JC',8,'original']
  ])
};
const backend = vm.createContext({
  SpreadsheetApp: { getActiveSpreadsheet: () => ({
    getSheetByName: name => sheets[name],
    insertSheet: name => sheets[name] = new Sheet()
  }) },
  ContentService: {
    MimeType: {JSON:'json'},
    createTextOutput: value => ({setMimeType: () => JSON.parse(value)})
  }
});
vm.runInContext(fs.readFileSync(__dirname + '/Codigo.gs','utf8'), backend);
const call = (action, body) => backend.handleRequest({
  parameter:{action}, ...(body ? {postData:{contents:JSON.stringify(body)}} : {})
});
let result = call('getAll');
assert.equal(result.players.length, 1, 'creating missing match sheets preserves players');
assert.equal(result.players[0].rating, 8);
assert.equal(result.assessmentSchemaVersion, 1);
backend.initSheets();
assert.equal(sheets.Jugadores.rows[0].filter(h => h === 'estadoFisico').length, 1);
assert.ok(call('savePlayer', {id:'missing',nombre:'Missing'}).error);
assert.ok(call('savePlayer', {id:'bad',estadoFisico:'alta',habilidadBalon:'invalid'}).error);
assert.equal(call('savePlayer',{id:7,aliases:'Viejo'}).ok, true);
assert.equal(call('getAll').players[0].rating, 8);
assert.equal(call('getAll').players[0].creado, 'original');
assert.equal(call('savePlayer',{id:'new',nombre:'Nueva',genero:'F',estadoFisico:'alta',habilidadBalon:'baja'}).ok, true);
assert.equal(call('savePlayer',{id:'new',aliases:'Alias'}).ok, true);
assert.equal(call('getAll').players.find(p => p.id === 'new').estadoFisico, 'alta');
assert.ok(call('savePlayer',{id:'new',estadoFisico:'',habilidadBalon:''}).error);
assert.equal(call('savePlayer',{id:7,estadoFisico:'intermedia',habilidadBalon:'alta'}).ok, true);
assert.equal(call('getAll').players[0].creado, 'original');
assert.ok(backend.handleRequest({parameter:{},postData:{contents:'bad json'}}).error);

const html = fs.readFileSync(__dirname + '/index.html','utf8');
const ids = [...html.matchAll(/id="([^"]+)"/g)].map(m => m[1]);
assert.equal(new Set(ids).size, ids.length, 'unique HTML ids');
assert.ok(!/⭐|setStar|ratingToStars|star-hint|edit-rating/.test(html));
const stack = [];
for (const match of html.replace(/<script>[\s\S]*?<\/script>/g,'').replace(/<style>[\s\S]*?<\/style>/g,'').replace(/<!--[\s\S]*?-->/g,'').matchAll(/<(\/?)([a-z][a-z0-9]*)\b[^>]*>/gi)) {
  const [, closing, tag] = match;
  if (['meta','input','br','hr'].includes(tag)) continue;
  if (closing) assert.equal(stack.pop(), tag, 'balanced HTML: ' + tag);
  else stack.push(tag);
}
assert.equal(stack.length, 0);
const elements = Object.fromEntries(ids.map(id => [id, {
  value:'',style:{},classList:{add(){},remove(){},toggle(){}},
  setAttribute(){},querySelectorAll(){return []},scrollIntoView(){},
  textContent:'',innerHTML:''
}]));
let writes = 0;
const frontend = vm.createContext({
  document: {getElementById: id => elements[id], querySelectorAll: () => []},
  fetch: async (url, options) => {
    const action = new URL(url).searchParams.get('action');
    if (options.method === 'POST') writes++;
    const result = call(action, options.body && JSON.parse(options.body));
    return {ok:true,json:async () => result};
  },
  setTimeout(){},alert(){},console
});
const run = code => vm.runInContext(code, frontend);
vm.runInContext(html.match(/<script>([\s\S]*?)<\/script>/)[1], frontend);
(async () => {
  await run('reloadData()');
  for (const [physical, points] of [['baja',2],['intermedia',6],['alta',10]]) {
    for (const [skill, skillPoints] of [['baja',2],['intermedia',6],['alta',10]]) {
      assert.equal(run("baseScore({estadoFisico:'" + physical + "',habilidadBalon:'" + skill + "'})"), (points+skillPoints)/2);
    }
  }
  assert.equal(run('baseScore({rating:10})'),6);
  elements['inp-nombre'].value = 'Test';
  elements['inp-genero'].value = 'M';
  elements['inp-posicion'].value = 'JC';
  const before = writes;
  await run('addPlayer()');
  assert.equal(writes,before,'required levels prevent save');
  elements['inp-estado-fisico'].value = 'baja';
  await run('addPlayer()');
  assert.equal(writes,before,'both levels required');
  elements['inp-habilidad-balon'].value = 'alta';
  await run('addPlayer()');
  assert.equal(writes,before+1);
  assert.equal(elements['inp-estado-fisico'].value,'');
  await run('reloadData()');
  assert.equal(run("players.find(p => p.nombre === 'Test').habilidadBalon"),'alta');
  elements['edit-id'].value = 'new';
  elements['edit-nombre'].value = 'Edited';
  elements['edit-genero'].value = 'F';
  elements['edit-posicion'].value = 'JC';
  elements['edit-estado-fisico'].value = 'intermedia';
  elements['edit-habilidad-balon'].value = 'baja';
  await run('saveEdit()');
  await run('reloadData()');
  assert.equal(run("players.find(p => p.id === 'new').estadoFisico"),'intermedia');
  assert.equal(run("players.find(p => p.id === 'new').nombre"),'Edited');
  run('assessmentSchemaVersion=0');
  await assert.rejects(run("api('savePlayer',{id:'new'})"),/Actualizá Google Apps Script/);
  assert.equal(run("players.find(p => p.id === 'new').nombre"),'Edited');
  frontend.fetch = async () => ({ok:true,json:async () => ({error:'Rejected'})});
  await assert.rejects(run("apiFetch(null,'savePlayer',{})"),/Rejected/);
  console.log('PASS: migration, validation, persistence, 9 level combinations, legacy profiles, create/edit, API errors and HTML structure.');
})().catch(error => { console.error(error); process.exitCode=1; });
