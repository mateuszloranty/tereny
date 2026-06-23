/**
 * Google Apps Script — synchronizacja tereny.geojson + arkusz Google Sheets
 *
 * 1. Utwórz plik tereny.geojson na Google Drive.
 * 2. Utwórz arkusz z zakładkami: bracia, tereny, grupa 1, grupa 2, Opracowanie 1-20, …
 * 3. Wklej FILE_ID i SPREADSHEET_ID poniżej.
 * 4. Wdróż jako aplikacja internetowa (wykonuj jako: Ja, dostęp: Każdy).
 * 5. URL wdrożenia → CONFIG.DRIVE_SYNC_URL w admin.html i servant.html
 */

var FILE_ID = '1mI6SMecuweOvA1xED0jQKwYcIwM11wgn';
var SPREADSHEET_ID = '1DI8B5w5tvOSb_IMUdVI_edbtOHOxj6ie4hARlv1unbo';

var SHEET = {
  bracia: 'Bracia',
  tereny: 'Tereny',
  grupaPrefix: 'Grupa ',
  opracowaniePrefix: 'Opracowanie '
};

var OPRACOWANIE_CHUNK = 20;

// ── HTTP ────────────────────────────────────────────────────────────────────

function doGet(e) {
  e = e || {};
  var p = e.parameter || {};
  var resource = (p.resource || 'tereny').toLowerCase();

  try {
  if (resource === 'tereny') {
    return jsonOut_(getTerenyGeoJson_());
  }
  if (!SPREADSHEET_ID) {
    return jsonOut_({ error: 'SPREADSHEET_ID nie ustawiony w drive-sync.gs' });
  }
  if (resource === 'bracia') return jsonOut_(readBracia_());
  if (resource === 'grupy') return jsonOut_(detectGrupy_());
  if (resource === 'grupa') return jsonOut_(readGrupa_(p.id));
  if (resource === 'tereny-meta') return jsonOut_(readTerenyMeta_());
  if (resource === 'opracowanie') return jsonOut_(readOpracowanie_(p.range));
  if (resource === 'bundle') return jsonOut_(getBundle_());
  return jsonOut_({ error: 'Nieznany resource: ' + resource });
  } catch (err) {
    return jsonOut_({ error: String(err.message || err) });
  }
}

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);

    if (body.resource === 'tereny') {
      saveTerenyGeoJson_(body.data);
      return jsonOut_({ ok: true, updated: new Date().toISOString() });
    }

    if (!SPREADSHEET_ID && body.resource !== 'tereny') {
      return jsonOut_({ error: 'SPREADSHEET_ID nie ustawiony w drive-sync.gs' });
    }

    if (body.resource === 'bracia') {
      writeBracia_(body.data);
      return jsonOut_({ ok: true, updated: new Date().toISOString() });
    }
    if (body.resource === 'grupa') {
      writeGrupa_(body.id, body.members, body.assignments);
      return jsonOut_({ ok: true, updated: new Date().toISOString() });
    }
    if (body.resource === 'tereny-meta') {
      writeTerenyMeta_(body.data);
      return jsonOut_({ ok: true, updated: new Date().toISOString() });
    }
    if (body.action === 'assign') {
      return jsonOut_(actionAssign_(body));
    }
    if (body.action === 'complete') {
      return jsonOut_(actionComplete_(body));
    }
    if (body.action === 'zdaj') {
      return jsonOut_(actionZdaj_(body));
    }

    // Legacy: raw GeoJSON POST (admin.html)
    if (!body.resource && !body.action) {
      saveTerenyGeoJson_(e.postData.contents);
      return jsonOut_({ ok: true, updated: new Date().toISOString() });
    }

    return jsonOut_({ error: 'Nieznane żądanie POST' });
  } catch (err) {
    return jsonOut_({ error: String(err.message || err) });
  }
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── GeoJSON ─────────────────────────────────────────────────────────────────

function getTerenyGeoJson_() {
  var file = DriveApp.getFileById(FILE_ID);
  return JSON.parse(file.getBlob().getDataAsString());
}

function saveTerenyGeoJson_(data) {
  var content = typeof data === 'string' ? data : JSON.stringify(data, null, 2);
  DriveApp.getFileById(FILE_ID).setContent(content);
}

// ── Bundle ──────────────────────────────────────────────────────────────────

function getBundle_() {
  var bundle = {
    tereny: getTerenyGeoJson_(),
    bracia: readBracia_(),
    grupy: detectGrupy_(),
    terenyMeta: readTerenyMeta_(),
    opracowanie: {}
  };
  var ranges = detectOpracowanieRanges_();
  for (var i = 0; i < ranges.length; i++) {
    bundle.opracowanie[ranges[i]] = readOpracowanie_(ranges[i]);
  }
  var grupaData = {};
  for (var g = 0; g < bundle.grupy.length; g++) {
    var gid = bundle.grupy[g].id;
    grupaData[gid] = readGrupa_(gid);
  }
  bundle.grupaData = grupaData;
  return bundle;
}

// ── Spreadsheet helpers ─────────────────────────────────────────────────────

function ss_() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function sheet_(name) {
  var sh = ss_().getSheetByName(name);
  if (!sh) throw new Error('Brak zakładki: ' + name);
  return sh;
}

function ensureSheet_(name, headers) {
  var sh = ss_().getSheetByName(name);
  if (!sh) {
    sh = ss_().insertSheet(name);
    if (headers && headers.length) sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  return sh;
}

function normalizeHeader_(h) {
  return String(h || '').toLowerCase().replace(/\s+/g, ' ').trim();
}

function findCol_(headers, candidates) {
  for (var c = 0; c < candidates.length; c++) {
    var want = normalizeHeader_(candidates[c]);
    for (var i = 0; i < headers.length; i++) {
      if (normalizeHeader_(headers[i]) === want) return i;
    }
  }
  return -1;
}

function formatDate_(val) {
  if (!val && val !== 0) return '';
  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  var s = String(val).trim();
  if (!s) return '';
  var m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return m[0];
  var d = new Date(s);
  if (!isNaN(d.getTime())) {
    return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return s;
}

function parseTerenNum_(id) {
  var m = String(id).match(/^(\d+)/);
  return m ? parseInt(m[1], 10) : NaN;
}

// ── Bracia ──────────────────────────────────────────────────────────────────

function readBracia_() {
  var sh = ensureSheet_(SHEET.bracia, ['id', 'imie', 'nazwisko', 'grupa_id']);
  var data = sh.getDataRange().getValues();
  if (data.length < 2) return [];
  var headers = data[0].map(String);
  var cId = findCol_(headers, ['id']);
  var cImie = findCol_(headers, ['imie', 'imię']);
  var cNazw = findCol_(headers, ['nazwisko']);
  var cGrupa = findCol_(headers, ['grupa_id', 'grupa']);
  if (cId < 0) cId = 0;
  if (cImie < 0) cImie = 1;
  if (cNazw < 0) cNazw = 2;
  if (cGrupa < 0) cGrupa = 3;

  var out = [];
  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    var id = String(row[cId] || '').trim();
    if (!id) continue;
    out.push({
      id: id,
      imie: String(row[cImie] || '').trim(),
      nazwisko: String(row[cNazw] || '').trim(),
      grupa_id: String(row[cGrupa] || '').trim()
    });
  }
  return out;
}

function writeBracia_(rows) {
  var sh = ensureSheet_(SHEET.bracia, ['id', 'imie', 'nazwisko', 'grupa_id']);
  var values = [['id', 'imie', 'nazwisko', 'grupa_id']];
  for (var i = 0; i < rows.length; i++) {
    var b = rows[i];
    values.push([b.id, b.imie || '', b.nazwisko || '', b.grupa_id || '']);
  }
  sh.clear();
  sh.getRange(1, 1, values.length, 4).setValues(values);
  syncGroupMembersFromBracia_(rows);
}

function syncGroupMembersFromBracia_(rows) {
  var grupy = detectGrupy_();
  for (var g = 0; g < grupy.length; g++) {
    var gid = String(grupy[g].id);
    var members = [];
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i].grupa_id) === gid) {
        members.push({
          id_brata: rows[i].id,
          imie: rows[i].imie || '',
          nazwisko: rows[i].nazwisko || ''
        });
      }
    }
    var existing = readGrupa_(gid);
    writeGrupa_(gid, members, existing.assignments || []);
  }
}

// ── Grupy ───────────────────────────────────────────────────────────────────

function detectGrupy_() {
  var sheets = ss_().getSheets();
  var out = [];
  var prefix = SHEET.grupaPrefix.toLowerCase();
  for (var i = 0; i < sheets.length; i++) {
    var name = sheets[i].getName();
    if (name.toLowerCase().indexOf(prefix) === 0) {
      var id = name.substring(SHEET.grupaPrefix.length).trim();
      out.push({ id: id, name: name });
    }
  }
  out.sort(function (a, b) {
    var na = parseInt(a.id, 10);
    var nb = parseInt(b.id, 10);
    if (!isNaN(na) && !isNaN(nb)) return na - nb;
    return String(a.id).localeCompare(String(b.id), 'pl', { numeric: true });
  });
  return out;
}

function grupaSheetName_(id) {
  return SHEET.grupaPrefix + String(id);
}

function readGrupa_(id) {
  var name = grupaSheetName_(id);
  var sh = ss_().getSheetByName(name);
  if (!sh) return { id: String(id), members: [], assignments: [] };

  var data = sh.getDataRange().getValues();
  var members = [];
  var assignments = [];
  var section = 'members';

  for (var r = 0; r < data.length; r++) {
    var row = data[r];
    var c0 = normalizeHeader_(row[0]);
    if (c0 === 'przydziały' || c0 === 'przydzialy' || c0 === 'tereny') {
      section = 'assignments';
      continue;
    }
    if (c0 === 'id_brata' || c0 === 'id') {
      if (r > 0 && normalizeHeader_(data[r - 1][0]) !== 'id_brata') continue;
      continue;
    }
    if (c0 === 'teren_nr') continue;

    if (section === 'members') {
      var mid = String(row[0] || '').trim();
      if (!mid) continue;
      members.push({
        id_brata: mid,
        imie: String(row[1] || '').trim(),
        nazwisko: String(row[2] || '').trim()
      });
    } else {
      var tid = String(row[0] || '').trim();
      if (!tid) continue;
      assignments.push({
        teren_nr: tid,
        data_przydzialu: formatDate_(row[1]),
        data_opracowania: formatDate_(row[2])
      });
    }
  }
  return { id: String(id), members: members, assignments: assignments };
}

function writeGrupa_(id, members, assignments) {
  members = members || [];
  assignments = assignments || [];
  var name = grupaSheetName_(id);
  var sh = ensureSheet_(name, []);
  var values = [
    ['id_brata', 'imie', 'nazwisko'],
    ['', '', '']
  ];
  for (var m = 0; m < members.length; m++) {
    var mem = members[m];
    values.push([mem.id_brata, mem.imie || '', mem.nazwisko || '']);
  }
  values.push(['', '', '']);
  values.push(['Przydziały', '', '']);
  values.push(['teren_nr', 'data_przydzialu', 'data_opracowania']);
  for (var a = 0; a < assignments.length; a++) {
    var asg = assignments[a];
    values.push([asg.teren_nr, asg.data_przydzialu || '', asg.data_opracowania || '']);
  }
  sh.clear();
  sh.getRange(1, 1, values.length, 3).setValues(values);
}

// ── Tereny meta ─────────────────────────────────────────────────────────────

function readTerenyMeta_() {
  var sh = ensureSheet_(SHEET.tereny, [
    'teren_nr', 'nazwa', 'parent_teren_nr', 'ostatnie_opracowanie',
    'przypisany_typ', 'przypisany_id', 'data_przydzialu', 'opracowania_json'
  ]);
  var data = sh.getDataRange().getValues();
  if (data.length < 2) return [];

  var headers = data[0].map(String);
  var cols = {
    teren_nr: findCol_(headers, ['teren_nr', 'teren nr']),
    nazwa: findCol_(headers, ['nazwa', 'name']),
    parent: findCol_(headers, ['parent_teren_nr', 'parent']),
    ostatnie: findCol_(headers, ['ostatnie_opracowanie', 'data ostatniego opracowania']),
    typ: findCol_(headers, ['przypisany_typ', 'typ']),
    pid: findCol_(headers, ['przypisany_id', 'id']),
    dataP: findCol_(headers, ['data_przydzialu', 'data przydziału']),
    oprJson: findCol_(headers, ['opracowania_json', 'opracowania'])
  };

  var out = [];
  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    var terenNr = String(row[cols.teren_nr >= 0 ? cols.teren_nr : 0] || '').trim();
    if (!terenNr) continue;

    var completions = [];
    if (cols.oprJson >= 0 && row[cols.oprJson]) {
      try {
        completions = JSON.parse(String(row[cols.oprJson]));
        if (!Array.isArray(completions)) completions = [];
      } catch (e) { completions = []; }
    }

    out.push({
      teren_nr: terenNr,
      nazwa: cols.nazwa >= 0 ? String(row[cols.nazwa] || '').trim() : '',
      parent_teren_nr: cols.parent >= 0 ? String(row[cols.parent] || '').trim() : '',
      ostatnie_opracowanie: cols.ostatnie >= 0 ? formatDate_(row[cols.ostatnie]) : '',
      przypisany_typ: cols.typ >= 0 ? String(row[cols.typ] || '').trim() : '',
      przypisany_id: cols.pid >= 0 ? String(row[cols.pid] || '').trim() : '',
      data_przydzialu: cols.dataP >= 0 ? formatDate_(row[cols.dataP]) : '',
      completions: completions
    });
  }
  return out;
}

function writeTerenyMeta_(rows) {
  var sh = ensureSheet_(SHEET.tereny, [
    'teren_nr', 'nazwa', 'parent_teren_nr', 'ostatnie_opracowanie',
    'przypisany_typ', 'przypisany_id', 'data_przydzialu', 'opracowania_json'
  ]);
  var values = [[
    'teren_nr', 'nazwa', 'parent_teren_nr', 'ostatnie_opracowanie',
    'przypisany_typ', 'przypisany_id', 'data_przydzialu', 'opracowania_json'
  ]];
  for (var i = 0; i < rows.length; i++) {
    var t = rows[i];
    values.push([
      t.teren_nr,
      t.nazwa || '',
      t.parent_teren_nr || '',
      t.ostatnie_opracowanie || '',
      t.przypisany_typ || '',
      t.przypisany_id || '',
      t.data_przydzialu || '',
      JSON.stringify(t.completions || [])
    ]);
  }
  sh.clear();
  if (values.length > 0) {
    sh.getRange(1, 1, values.length, 8).setValues(values);
  }
}

function getTerenMetaById_(terenId) {
  var all = readTerenyMeta_();
  for (var i = 0; i < all.length; i++) {
    if (String(all[i].teren_nr) === String(terenId)) return all[i];
  }
  return null;
}

function upsertTerenMeta_(meta) {
  var all = readTerenyMeta_();
  var found = false;
  for (var i = 0; i < all.length; i++) {
    if (String(all[i].teren_nr) === String(meta.teren_nr)) {
      all[i] = meta;
      found = true;
      break;
    }
  }
  if (!found) all.push(meta);
  writeTerenyMeta_(all);
  return meta;
}

function defaultMeta_(terenId, nazwa) {
  return {
    teren_nr: String(terenId),
    nazwa: nazwa || '',
    parent_teren_nr: '',
    ostatnie_opracowanie: '',
    przypisany_typ: '',
    przypisany_id: '',
    data_przydzialu: '',
    completions: []
  };
}

// ── Opracowanie sheets ──────────────────────────────────────────────────────

function detectOpracowanieRanges_() {
  var sheets = ss_().getSheets();
  var out = [];
  var prefix = SHEET.opracowaniePrefix.toLowerCase();
  for (var i = 0; i < sheets.length; i++) {
    var name = sheets[i].getName();
    if (name.toLowerCase().indexOf(prefix) !== 0) continue;
    var rest = name.substring(SHEET.opracowaniePrefix.length).trim();
    var m = rest.match(/(\d+)\s*[-–]\s*(\d+)/);
    if (m) out.push(m[1] + '-' + m[2]);
  }
  out.sort(function (a, b) {
    return parseInt(a.split('-')[0], 10) - parseInt(b.split('-')[0], 10);
  });
  return out;
}

function opracowanieSheetName_(range) {
  return SHEET.opracowaniePrefix + range;
}

function findOpracowanieSheetForTeren_(terenNr) {
  var num = parseTerenNum_(terenNr);
  if (isNaN(num)) return null;
  var start = Math.floor((num - 1) / OPRACOWANIE_CHUNK) * OPRACOWANIE_CHUNK + 1;
  var end = start + OPRACOWANIE_CHUNK - 1;
  var range = start + '-' + end;
  var name = opracowanieSheetName_(range);
  var sh = ss_().getSheetByName(name);
  if (!sh) {
    sh = ensureOpracowanieSheet_(range, start, end);
  }
  return { sheet: sh, range: range, start: start, end: end };
}

function ensureOpracowanieSheet_(range, start, end) {
  var name = opracowanieSheetName_(range);
  var sh = ss_().getSheetByName(name);
  if (!sh) sh = ss_().insertSheet(name);

  var header1 = ['Teren nr', 'Data ostatniego opracowania*'];
  var header2 = ['', ''];
  for (var i = 0; i < 8; i++) {
    header1.push('Przydzielono:');
    header1.push('');
    header2.push('Data przydziału');
    header2.push('Data opracowania');
  }
  sh.clear();
  sh.getRange(1, 1, 1, header1.length).setValues([header1]);
  sh.getRange(2, 1, 1, header2.length).setValues([header2]);

  var row = 3;
  for (var t = start; t <= end; t++) {
    sh.getRange(row, 1).setValue(t);
    row += 2;
  }
  return sh;
}

function terenRowInOpracowanie_(terenNr, start) {
  var num = parseTerenNum_(terenNr);
  if (isNaN(num) || num < start) return -1;
  var idx = num - start;
  return 3 + idx * 2;
}

function readOpracowanie_(range) {
  if (!range) return [];
  var sh = ss_().getSheetByName(opracowanieSheetName_(range));
  if (!sh) return [];

  var parts = String(range).split('-');
  var start = parseInt(parts[0], 10);
  var end = parseInt(parts[1], 10);
  if (isNaN(start) || isNaN(end)) return [];

  var lastCol = Math.max(sh.getLastColumn(), 18);
  var lastRow = Math.max(sh.getLastRow(), 3 + (end - start) * 2);
  var data = sh.getRange(1, 1, lastRow, lastCol).getValues();
  var out = [];

  for (var t = start; t <= end; t++) {
    var rowIdx = terenRowInOpracowanie_(t, start) - 1;
    if (rowIdx < 0 || rowIdx >= data.length) continue;
    var row = data[rowIdx];
    var historia = [];
    for (var col = 2; col < lastCol; col += 2) {
      var dp = formatDate_(row[col]);
      var do_ = formatDate_(row[col + 1]);
      if (!dp && !do_) continue;
      historia.push({ data_przydzialu: dp, data_opracowania: do_ });
    }
    out.push({
      teren_nr: String(t),
      ostatnie_opracowanie: formatDate_(row[1]),
      historia: historia
    });
  }
  return out;
}

function appendOpracowanieHistory_(terenId, dataPrzydzialu, dataOpracowania, ostatnieOpracowanie) {
  var info = findOpracowanieSheetForTeren_(terenId);
  if (!info) return;
  var sh = info.sheet;
  var row = terenRowInOpracowanie_(terenId, info.start);
  if (row < 1) return;

  var lastCol = Math.max(sh.getLastColumn(), 18);
  var rowData = sh.getRange(row, 1, 1, lastCol).getValues()[0];

  if (ostatnieOpracowanie) {
    sh.getRange(row, 2).setValue(ostatnieOpracowanie);
  }

  var insertCol = 3;
  for (var col = 2; col < lastCol; col += 2) {
    var dp = formatDate_(rowData[col]);
    var do_ = formatDate_(rowData[col + 1]);
    if (!dp && !do_) {
      insertCol = col + 1;
      break;
    }
    insertCol = col + 3;
  }

  if (insertCol > lastCol) {
    sh.insertColumnsAfter(lastCol, 2);
  }
  sh.getRange(row, insertCol).setValue(dataPrzydzialu || '');
  sh.getRange(row, insertCol + 1).setValue(dataOpracowania || '');
}

// ── Actions ─────────────────────────────────────────────────────────────────

function actionAssign_(body) {
  var terenId = String(body.terenId);
  var meta = getTerenMetaById_(terenId) || defaultMeta_(terenId, body.nazwa || '');
  meta.przypisany_typ = body.type || '';
  meta.przypisany_id = body.id ? String(body.id) : '';
  meta.data_przydzialu = body.assignedDate ? formatDate_(body.assignedDate) : '';
  meta.completions = [];
  upsertTerenMeta_(meta);
  return { ok: true, meta: meta };
}

function actionComplete_(body) {
  var terenId = String(body.terenId);
  var meta = getTerenMetaById_(terenId) || defaultMeta_(terenId, '');
  if (!meta.completions) meta.completions = [];
  meta.completions.push({
    date: body.date ? formatDate_(body.date) : formatDate_(new Date()),
    subterenId: body.subterenId ? String(body.subterenId) : null
  });
  upsertTerenMeta_(meta);
  return { ok: true, meta: meta };
}

function actionZdaj_(body) {
  var terenId = String(body.terenId);
  var meta = getTerenMetaById_(terenId) || defaultMeta_(terenId, '');

  var lastCompletion = '';
  if (meta.completions && meta.completions.length) {
    lastCompletion = meta.completions[meta.completions.length - 1].date || '';
  }

  var zdajDate = body.date ? formatDate_(body.date) : (lastCompletion || formatDate_(new Date()));

  appendOpracowanieHistory_(
    terenId,
    meta.data_przydzialu,
    zdajDate,
    zdajDate
  );

  meta.ostatnie_opracowanie = zdajDate;
  meta.przypisany_typ = '';
  meta.przypisany_id = '';
  meta.data_przydzialu = '';
  meta.completions = [];
  upsertTerenMeta_(meta);

  return { ok: true, meta: meta, zdajDate: zdajDate };
}
