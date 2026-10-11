// =====================================================
// FÚTBOL 7 — Google Apps Script Backend
// v3 — estado físico y habilidad con el balón
// =====================================================

const SHEET_PLAYERS  = 'Jugadores';
const SHEET_PARTIDOS = 'Partidos';
const SHEET_STATS    = 'Stats';

function doGet(e)  { return handleRequest(e); }
function doPost(e) { return handleRequest(e); }

function handleRequest(e) {
  let result;
  try {
    const params = e.parameter || {};
    const body = e.postData ? JSON.parse(e.postData.contents || '{}') : {};
    const action = params.action || body.action;
    switch(action) {
      case 'getAll':        result = getAll();               break;
      case 'savePlayer':    result = savePlayer(body);       break;
      case 'deletePlayer':  result = deletePlayer(body.id);  break;
      case 'savePartido':   result = savePartido(body);      break;
      case 'deletePartido': result = deletePartido(body.id); break;
      default: result = { error: 'Acción desconocida: ' + action };
    }
  } catch(err) {
    result = { error: err.message };
  }

  const output = ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);

  return output;
}

function doOptions(e) {
  return ContentService
    .createTextOutput('')
    .setMimeType(ContentService.MimeType.JSON);
}

// ── Inicializar / migrar hojas ───────────────────────
function initSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Jugadores — con aliases
  let sp = ss.getSheetByName(SHEET_PLAYERS);
  if (!sp) {
    sp = ss.insertSheet(SHEET_PLAYERS);
    sp.appendRow(['id','nombre','genero','edad','posicion','rating','creado','aliases','estadoFisico','habilidadBalon']);
    sp.setFrozenRows(1);
  } else {
    ensurePlayerColumns(sp);
  }

  // Partidos
  let spt = ss.getSheetByName(SHEET_PARTIDOS);
  if (!spt) {
    spt = ss.insertSheet(SHEET_PARTIDOS);
    spt.appendRow(['id','fecha','golesA','golesB','creado']);
    spt.setFrozenRows(1);
  }

  // Stats
  let sst = ss.getSheetByName(SHEET_STATS);
  if (!sst) {
    sst = ss.insertSheet(SHEET_STATS);
    sst.appendRow(['partidoId','playerId','equipo','goles','asist','mvp']);
    sst.setFrozenRows(1);
  }

  return { ok: true, msg: 'Hojas listas' };
}

// ── Leer todo ────────────────────────────────────────
function getAll() {
  initSheets();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const spS  = ss.getSheetByName(SHEET_PLAYERS);
  const sptS = ss.getSheetByName(SHEET_PARTIDOS);
  const sstS = ss.getSheetByName(SHEET_STATS);

  const players  = sheetToObjects(spS);
  const partidos = sheetToObjects(sptS);
  const stats    = sheetToObjects(sstS);

  const byPartido = {};
  stats.forEach(s => {
    if (!byPartido[s.partidoId]) byPartido[s.partidoId] = {};
    byPartido[s.partidoId][s.playerId] = {
      equipo: s.equipo,
      goles:  parseInt(s.goles)||0,
      asist:  parseInt(s.asist)||0,
      mvp:    s.mvp==='true'||s.mvp===true
    };
  });
  partidos.forEach(p => { p.stats = byPartido[p.id] || {}; });
  return { players, partidos, assessmentSchemaVersion:1 };
}

// ── Jugadores ─────────────────────────────────────────
function savePlayer(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sp = ss.getSheetByName(SHEET_PLAYERS);
  if (!sp) { initSheets(); sp = ss.getSheetByName(SHEET_PLAYERS); }

  ensurePlayerColumns(sp);
  const headers = sp.getRange(1,1,1,sp.getLastColumn()).getValues()[0];

  const rows  = sp.getDataRange().getValues();
  const idIdx = headers.indexOf('id');
  const existingRow = rows.slice(1).find(row => String(row[idIdx]) === String(data.id));
  const levelFields = ['estadoFisico','habilidadBalon'];
  const validLevel = value => ['baja','intermedia','alta'].includes(value);
  // Un perfil antiguo puede guardar un apodo sin inventar sus calificaciones.
  const hasSubmittedAssessment = levelFields.some(field =>
    data[field] !== undefined && data[field] !== null && data[field] !== '');
  if (!existingRow || hasSubmittedAssessment) {
    const complete = levelFields.every(field => validLevel(
      data[field] !== undefined ? data[field] : existingRow && existingRow[headers.indexOf(field)]
    ));
    if (!complete) throw new Error('Seleccioná el estado físico y la habilidad con el balón: baja, intermedia o alta.');
  }
  // Las calificaciones existentes no se pueden borrar enviando valores vacíos.
  if (existingRow && levelFields.some(field =>
    validLevel(existingRow[headers.indexOf(field)]) &&
    data[field] !== undefined && !validLevel(data[field])
  )) throw new Error('Las calificaciones deben ser baja, intermedia o alta.');

  function buildRow(createdVal) {
    const row = existingRow ? existingRow.slice() : new Array(headers.length).fill('');
    headers.forEach((h,i) => {
      if (h === 'creado') row[i] = createdVal;
      else if (h !== 'rating' && data[h] !== undefined && data[h] !== null) row[i] = data[h];
    });
    return row;
  }

  const newId = data.id || Date.now().toString();
  for (let i=1; i<rows.length; i++) {
    if (String(rows[i][idIdx]) === String(data.id)) {
      const row = buildRow(rows[i][headers.indexOf('creado')]);
      row[idIdx] = data.id;
      sp.getRange(i+1,1,1,headers.length).setValues([row]);
      return { ok:true, action:'updated', id:data.id, assessmentSchemaVersion:1 };
    }
  }

  const row = buildRow(new Date().toISOString());
  row[idIdx] = newId;
  sp.appendRow(row);
  return { ok:true, action:'created', id:newId, assessmentSchemaVersion:1 };
}

function deletePlayer(id) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sp = ss.getSheetByName(SHEET_PLAYERS);
  if (!sp) return { ok:false };
  deleteRowById(sp, id);
  return { ok:true };
}

// ── Partidos ──────────────────────────────────────────
function savePartido(data) {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  let spt = ss.getSheetByName(SHEET_PARTIDOS);
  let sst = ss.getSheetByName(SHEET_STATS);
  if (!spt||!sst) { initSheets(); spt=ss.getSheetByName(SHEET_PARTIDOS); sst=ss.getSheetByName(SHEET_STATS); }

  const partidoId = data.id || Date.now().toString();

  const pRows = spt.getDataRange().getValues();
  const pH = pRows[0]; const pidIdx = pH.indexOf('id');
  let found = false;
  for (let i=1; i<pRows.length; i++) {
    if (String(pRows[i][pidIdx])===String(partidoId)) {
      spt.getRange(i+1,1,1,pH.length).setValues([[
        partidoId,data.fecha,data.golesA,data.golesB,pRows[i][pH.indexOf('creado')]
      ]]);
      found=true; break;
    }
  }
  if (!found) spt.appendRow([partidoId,data.fecha,data.golesA,data.golesB,new Date().toISOString()]);

  // Reescribir stats
  const sRows = sst.getDataRange().getValues();
  const sH = sRows[0]; const spidIdx = sH.indexOf('partidoId');
  for (let i=sRows.length-1; i>=1; i--) {
    if (String(sRows[i][spidIdx])===String(partidoId)) sst.deleteRow(i+1);
  }
  const stats = data.stats || {};
  Object.keys(stats).forEach(pid => {
    const s = stats[pid];
    sst.appendRow([partidoId,pid,s.equipo,s.goles||0,s.asist||0,s.mvp?'true':'false']);
  });
  return { ok:true, id:partidoId };
}

function deletePartido(id) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const spt = ss.getSheetByName(SHEET_PARTIDOS);
  const sst = ss.getSheetByName(SHEET_STATS);
  if (spt) deleteRowById(spt, id);
  if (sst) {
    const rows = sst.getDataRange().getValues();
    const pidIdx = rows[0].indexOf('partidoId');
    for (let i=rows.length-1; i>=1; i--) {
      if (String(rows[i][pidIdx])===String(id)) sst.deleteRow(i+1);
    }
  }
  return { ok:true };
}

// ── Helpers ───────────────────────────────────────────
function ensurePlayerColumns(sheet) {
  const headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  ['aliases','estadoFisico','habilidadBalon'].forEach(field => {
    if (!headers.includes(field)) {
      sheet.getRange(1, sheet.getLastColumn()+1).setValue(field);
      headers.push(field);
    }
  });
}

function sheetToObjects(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length<2) return [];
  const headers = data[0];
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h,i) => { obj[h] = row[i]; });
    return obj;
  });
}

function deleteRowById(sheet, id) {
  const rows = sheet.getDataRange().getValues();
  const idIdx = rows[0].indexOf('id');
  for (let i=rows.length-1; i>=1; i--) {
    if (String(rows[i][idIdx])===String(id)) { sheet.deleteRow(i+1); break; }
  }
}
