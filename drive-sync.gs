/**
 * Google Apps Script — synchronizacja tereny.geojson + arkusz Google Sheets
 *
 * 1. Utwórz plik tereny.geojson na Google Drive.
 * 2. Utwórz arkusz z zakładkami: Bracia, Tereny, Grupa 1, Grupa 2, Opracowanie 1-20, …
 *    (zakładki „Grupa X” uzupełniaj ręcznie — skrypt ich nie nadpisuje)
 * 3. Wklej FILE_ID i SPREADSHEET_ID poniżej.
 * 4. Wdróż jako aplikacja internetowa (wykonuj jako: Ja, dostęp: Każdy).
 * 5. Skopiuj URL wdrożenia do CONFIG.DRIVE_SYNC_URL w admin.html i servant.html
 *
 * WAŻNE — pierwsza konfiguracja uprawnień:
 * W edytorze Apps Script uruchom raz funkcję authorizeOnce() i zaakceptuj dostęp
 * do Dysku Google oraz Arkuszy. Bez tego zapis do zakładki Tereny nie zadziała.
 */

var FILE_ID = '1mI6SMecuweOvA1xED0jQKwYcIwM11wgn';
var SPREADSHEET_ID = '1FBZwO6xrmwEvIy8VUhe9dTQfBb6K45XeVBUHo4WYmjo';

var SHEET = {
  bracia: 'Bracia',
  tereny: 'Tereny',
  grupaPrefix: 'Grupa ',
  templateGrupa: 'Grupa 1',
  opracowaniePrefix: 'Opracowanie '
};

var OPRACOWANIE_CHUNK = 20;
var TERENY_MAX_BIEZACE_OPRACOWANIA = 3;

function terenySheetHeaders_() {
  var headers = [
    'Numer terenu',
    'Nazwa',
    'Teren nadrzędny',
    'Data ostatniego opracowania',
    'Typ przydziału',
    'Aktualny przydział',
    'Data przydziału'
  ];
  for (var n = 1; n <= TERENY_MAX_BIEZACE_OPRACOWANIA; n++) {
    headers.push('Opracowanie ' + n + ' – przydział');
    headers.push('Opracowanie ' + n + ' – opracowanie');
  }
  return headers;
}

function braciaSheetHeaders_() {
  return ['ID', 'Imię', 'Nazwisko', 'Grupa'];
}

function bratDisplayName_(b) {
  return (String(b.imie || '') + ' ' + String(b.nazwisko || '')).trim() || ('Brat ' + b.id);
}

function normalizePrzypisanyTyp_(val) {
  var s = String(val || '').trim().toLowerCase();
  if (s === 'brat' || s === 'głosiciel' || s === 'glosiciel') return 'brat';
  if (s === 'grupa' || s === 'group') return 'grupa';
  return s;
}

function formatPrzypisanyTypForSheet_(typ) {
  if (typ === 'brat') return 'Brat';
  if (typ === 'grupa') return 'Grupa';
  return typ || '';
}

function formatAssigneeForSheet_(typ, id, bracia, grupy) {
  if (!typ || !id) return '';
  var idStr = String(id);
  if (typ === 'grupa') {
    for (var g = 0; g < grupy.length; g++) {
      if (String(grupy[g].id) === idStr || String(grupy[g].name) === idStr) {
        return grupy[g].name || String(grupy[g].id);
      }
    }
    if (/^\d+$/.test(idStr)) return 'Grupa ' + idStr;
    return idStr;
  }
  if (typ === 'brat') {
    for (var i = 0; i < bracia.length; i++) {
      if (String(bracia[i].id) === idStr) return bratDisplayName_(bracia[i]);
    }
    return idStr;
  }
  return idStr;
}

function parseAssigneeIdFromSheet_(typ, displayValue, bracia, grupy) {
  var s = String(displayValue || '').trim();
  if (!s) return '';
  if (typ === 'brat') {
    for (var i = 0; i < bracia.length; i++) {
      if (String(bracia[i].id) === s) return String(bracia[i].id);
      if (bratDisplayName_(bracia[i]) === s) return String(bracia[i].id);
    }
    return s;
  }
  if (typ === 'grupa') {
    for (var g = 0; g < grupy.length; g++) {
      var name = grupy[g].name || String(grupy[g].id);
      if (name === s || String(grupy[g].id) === s) return name;
    }
    return s;
  }
  return s;
}

/**
 * Uruchom ręcznie w edytorze Apps Script (▶), zaakceptuj uprawnienia,
 * potem wdróż ponownie aplikację internetową.
 */
function authorizeOnce() {
  var file = DriveApp.getFileById(FILE_ID);
  Logger.log('Drive OK: ' + file.getName());
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  Logger.log('Sheets OK: ' + ss.getName());
  Logger.log('Zakładki: ' + ss.getSheets().map(function (s) { return s.getName(); }).join(', '));
  return { ok: true, spreadsheet: ss.getName() };
}

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
  if (resource === 'grupa') return jsonOut_(readGrupa_(p.name || p.id));
  if (resource === 'tereny-meta') return jsonOut_(readTerenyMeta_());
  if (resource === 'opracowanie') return jsonOut_(readOpracowanie_(p.range));
  if (resource === 'bundle') {
    var light = p.light === '1' || p.light === 'true';
    return jsonOut_(getBundle_(light));
  }
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
      writeGrupa_(body.name || body.id, body.members, body.assignments);
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
    if (body.action === 'unassign') {
      return jsonOut_(actionUnassign_(body));
    }
    if (body.action === 'import-tereny') {
      return jsonOut_(importTerenyMetaFromGeoJson_(body));
    }
    if (body.action === 'ensure-opracowanie') {
      var geoFeatures = getTerenyGeoJson_().features || [];
      return jsonOut_({
        ok: true,
        opracowanieEnsured: ensureOpracowanieSheetsFromGeo_(geoFeatures)
      });
    }
    if (body.action === 'create-grupa') {
      return jsonOut_(createGrupaFromTemplate_(body.name));
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

function safeRead_(fn, fallback) {
  try {
    return fn();
  } catch (e) {
    return fallback;
  }
}

function getBundle_(light) {
  var bundle = {
    tereny: light ? null : safeRead_(getTerenyGeoJson_, null),
    bracia: safeRead_(readBracia_, []),
    grupy: safeRead_(detectGrupy_, []),
    terenyMeta: safeRead_(readTerenyMeta_, []),
    opracowanie: {},
    grupaData: {}
  };

  if (!light) {
    var ranges = safeRead_(detectOpracowanieRanges_, []);
    for (var i = 0; i < ranges.length; i++) {
      var range = ranges[i];
      bundle.opracowanie[range] = safeRead_(function (r) {
        return function () { return readOpracowanie_(r); };
      }(range), []);
    }
  }

  for (var g = 0; g < bundle.grupy.length; g++) {
    var gid = bundle.grupy[g].id;
    bundle.grupaData[gid] = safeRead_(function (id) {
      return function () { return readGrupa_(id); };
    }(gid), { id: String(gid), members: [], assignments: [] });
  }
  return bundle;
}

// ── Spreadsheet helpers ─────────────────────────────────────────────────────

function ss_() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function sheet_(name) {
  var sh = findSheet_(name);
  if (!sh) throw new Error('Brak zakładki: ' + name);
  return sh;
}

/** Dopasowanie zakładki po nazwie (wielkość liter ignorowana). */
function findSheet_(canonicalName) {
  var sh = ss_().getSheetByName(canonicalName);
  if (sh) return sh;
  var want = String(canonicalName).toLowerCase();
  var sheets = ss_().getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getName().toLowerCase() === want) return sheets[i];
  }
  return null;
}

/** Wszystkie zakładki zaczynające się od prefiksu (np. „Grupa ”, „Opracowanie ”). */
function findSheetsByPrefix_(prefix) {
  var want = String(prefix).toLowerCase();
  var sheets = ss_().getSheets();
  var out = [];
  for (var i = 0; i < sheets.length; i++) {
    var name = sheets[i].getName();
    if (name.toLowerCase().indexOf(want) !== 0) continue;
    out.push({
      sheet: sheets[i],
      name: name,
      suffix: name.substring(want.length).trim()
    });
  }
  return out;
}

function ensureSheet_(name, headers) {
  var sh = findSheet_(name);
  if (!sh) {
    sh = ss_().insertSheet(name);
    if (headers && headers.length) sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  return sh;
}

function getSheetOptional_(name) {
  return findSheet_(name);
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
  var sh = getSheetOptional_(SHEET.bracia);
  if (!sh) return [];
  var data = sh.getDataRange().getValues();
  if (data.length < 2) return [];
  var headers = data[0].map(String);
  var cId = findCol_(headers, ['id']);
  var cImie = findCol_(headers, ['imię', 'imie']);
  var cNazw = findCol_(headers, ['nazwisko']);
  var cGrupa = findCol_(headers, ['grupa', 'grupa_id']);
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
  var headers = braciaSheetHeaders_();
  var sh = ensureSheet_(SHEET.bracia, headers);
  var values = [headers.slice()];
  for (var i = 0; i < rows.length; i++) {
    var b = rows[i];
    values.push([b.id, b.imie || '', b.nazwisko || '', b.grupa_id || '']);
  }
  sh.clear();
  sh.getRange(1, 1, values.length, headers.length).setValues(values);
}

// ── Grupy (zakładki w jednym arkuszu Google Sheets) ─────────────────────────

/** Pełna nazwa zakładki grupy — obsługa starego id „1” oraz „Grupa 1”. */
function resolveGrupaSheetName_(idOrName) {
  var s = String(idOrName || '').trim();
  if (!s) return '';

  var direct = findSheet_(s);
  if (direct) return direct.getName();

  var prefixed = SHEET.grupaPrefix + s;
  var byPrefix = findSheet_(prefixed);
  if (byPrefix) return byPrefix.getName();

  return prefixed;
}

function detectGrupy_() {
  var matches = findSheetsByPrefix_(SHEET.grupaPrefix);
  var out = [];
  for (var i = 0; i < matches.length; i++) {
    var sheetName = matches[i].name;
    out.push({ id: sheetName, name: sheetName });
  }
  out.sort(function (a, b) {
    return a.name.localeCompare(b.name, 'pl', { numeric: true });
  });
  return out;
}

/**
 * Nowa grupa = kopia zakładki szablonu (domyślnie „Grupa 1”) w tym samym pliku arkusza.
 */
function createGrupaFromTemplate_(newName) {
  newName = String(newName || '').trim();
  if (!newName) throw new Error('Podaj nazwę zakładki grupy');
  if (findSheet_(newName)) throw new Error('Zakładka „' + newName + '” już istnieje');

  var templateName = SHEET.templateGrupa || (SHEET.grupaPrefix + '1');
  var template = findSheet_(templateName);
  if (!template) {
    throw new Error('Brak szablonu „' + templateName + '” — utwórz zakładkę wzorcową w arkuszu');
  }

  var copy = template.copyTo(ss_());
  copy.setName(newName);

  return { ok: true, id: newName, name: newName };
}

function readGrupa_(idOrName) {
  var sheetName = resolveGrupaSheetName_(idOrName);
  var sh = findSheet_(sheetName);
  if (!sh) return { id: sheetName, name: sheetName, members: [], assignments: [] };

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
  return { id: sheetName, name: sheetName, members: members, assignments: assignments };
}

/**
 * Zakładki „Grupa X” mają niestandardowy układ — uzupełniaj je ręcznie w arkuszu.
 * Zapis z API jest celowo pomijany (tylko odczyt przez readGrupa_).
 */
function writeGrupa_(idOrName, members, assignments) {
  // no-op
}

// ── Tereny meta ─────────────────────────────────────────────────────────────

function terenyCompletionColCandidates_(n, kind) {
  var label = kind === 'przydzial' ? 'przydział' : 'opracowanie';
  var ascii = kind === 'przydzial' ? 'przydzial' : 'opracowanie';
  return [
    'opracowanie ' + n + ' – ' + label,
    'opracowanie ' + n + ' - ' + label,
    'opracowanie ' + n + ' – ' + ascii,
    'opracowanie ' + n + ' - ' + ascii
  ];
}

function parseCompletionOpracowanieCell_(val) {
  var s = String(val || '').trim();
  if (!s) return { date: '', subterenId: null };
  var m = s.match(/^(\d{4}-\d{2}-\d{2})(?:\s*\(([^)]+)\))?/);
  if (m) return { date: m[1], subterenId: m[2] ? String(m[2]).trim() : null };
  return { date: formatDate_(val), subterenId: null };
}

function formatCompletionOpracowanieCell_(completion) {
  if (!completion || !completion.date) return '';
  return completion.subterenId
    ? (completion.date + ' (' + completion.subterenId + ')')
    : completion.date;
}

function readCompletionsFromTerenyRow_(row, cols, dataPrzydzialu) {
  var fromPairs = [];
  for (var i = 0; i < TERENY_MAX_BIEZACE_OPRACOWANIA; i++) {
    var pIdx = cols.oprPrzydzial[i];
    var oIdx = cols.oprOpracowanie[i];
    var dp = pIdx >= 0 ? formatDate_(row[pIdx]) : '';
    var parsed = oIdx >= 0 ? parseCompletionOpracowanieCell_(row[oIdx]) : { date: '', subterenId: null };
    if (!dp && !parsed.date) continue;
    fromPairs.push({
      date: parsed.date,
      subterenId: parsed.subterenId
    });
  }
  if (fromPairs.length) return fromPairs;

  if (cols.oprJson >= 0 && row[cols.oprJson]) {
    try {
      var legacy = JSON.parse(String(row[cols.oprJson]));
      if (Array.isArray(legacy)) return legacy;
    } catch (e) { /* legacy JSON */ }
  }
  return [];
}

function completionsToTerenyRowTail_(meta) {
  var comps = meta.completions || [];
  var defaultPrzydzial = meta.data_przydzialu || '';
  var out = [];
  for (var n = 0; n < TERENY_MAX_BIEZACE_OPRACOWANIA; n++) {
    var c = comps[n];
    if (c && c.date) {
      out.push(defaultPrzydzial);
      out.push(formatCompletionOpracowanieCell_(c));
    } else {
      out.push('');
      out.push('');
    }
  }
  return out;
}

function readTerenyMeta_() {
  var sh = getSheetOptional_(SHEET.tereny);
  if (!sh) return [];
  var data = sh.getDataRange().getValues();
  if (data.length < 2) return [];

  var bracia = readBracia_();
  var grupy = detectGrupy_();
  var headers = data[0].map(String);
  var cols = {
    teren_nr: findCol_(headers, ['numer terenu', 'teren_nr', 'teren nr']),
    nazwa: findCol_(headers, ['nazwa', 'name']),
    parent: findCol_(headers, ['teren nadrzędny', 'teren nadrzedny', 'parent_teren_nr', 'parent']),
    ostatnie: findCol_(headers, ['data ostatniego opracowania', 'ostatnie_opracowanie']),
    typ: findCol_(headers, ['typ przydziału', 'typ przydzialu', 'przypisany_typ', 'typ']),
    pid: findCol_(headers, ['aktualny przydział', 'aktualny przydzial', 'przypisany_id']),
    dataP: findCol_(headers, ['data przydziału', 'data przydzialu', 'data_przydzialu']),
    oprPrzydzial: [],
    oprOpracowanie: [],
    oprJson: findCol_(headers, [
      'opracowania (bieżący przydział)',
      'opracowania (biezacy przydzial)',
      'opracowania_json',
      'opracowania'
    ])
  };
  for (var n = 1; n <= TERENY_MAX_BIEZACE_OPRACOWANIA; n++) {
    cols.oprPrzydzial.push(findCol_(headers, terenyCompletionColCandidates_(n, 'przydzial')));
    cols.oprOpracowanie.push(findCol_(headers, terenyCompletionColCandidates_(n, 'opracowanie')));
  }

  var out = [];
  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    var terenNr = String(row[cols.teren_nr >= 0 ? cols.teren_nr : 0] || '').trim();
    if (!terenNr) continue;

    var dataPrzydzialu = cols.dataP >= 0 ? formatDate_(row[cols.dataP]) : '';
    var completions = readCompletionsFromTerenyRow_(row, cols, dataPrzydzialu);
    var przypisanyTyp = cols.typ >= 0 ? normalizePrzypisanyTyp_(row[cols.typ]) : '';
    var przypisanyDisplay = cols.pid >= 0 ? String(row[cols.pid] || '').trim() : '';
    var przypisanyId = parseAssigneeIdFromSheet_(przypisanyTyp, przypisanyDisplay, bracia, grupy);

    out.push({
      teren_nr: terenNr,
      nazwa: cols.nazwa >= 0 ? String(row[cols.nazwa] || '').trim() : '',
      parent_teren_nr: cols.parent >= 0 ? String(row[cols.parent] || '').trim() : '',
      ostatnie_opracowanie: cols.ostatnie >= 0 ? formatDate_(row[cols.ostatnie]) : '',
      przypisany_typ: przypisanyTyp,
      przypisany_id: przypisanyId,
      data_przydzialu: dataPrzydzialu,
      completions: completions
    });
  }
  return out;
}

function writeTerenyMeta_(rows) {
  var headers = terenySheetHeaders_();
  var sh = ensureSheet_(SHEET.tereny, headers);
  var bracia = readBracia_();
  var grupy = detectGrupy_();
  var values = [headers.slice()];
  for (var i = 0; i < rows.length; i++) {
    var t = rows[i];
    values.push([
      t.teren_nr,
      t.nazwa || '',
      t.parent_teren_nr || '',
      t.ostatnie_opracowanie || '',
      formatPrzypisanyTypForSheet_(t.przypisany_typ),
      formatAssigneeForSheet_(t.przypisany_typ, t.przypisany_id, bracia, grupy),
      t.data_przydzialu || ''
    ].concat(completionsToTerenyRowTail_(t)));
  }
  sh.clear();
  if (values.length > 0) {
    sh.getRange(1, 1, values.length, headers.length).setValues(values);
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

// ── Import GeoJSON → arkusz Tereny ──────────────────────────────────────────

function getFeatureIdFromGeo_(f) {
  if (f.properties && f.properties.id) return String(f.properties.id);
  if (f.id !== undefined && f.id !== null) return String(f.id);
  return '';
}

function getFeatureTitleFromGeo_(f) {
  var p = f.properties || {};
  return String(p.title || p.name || '').trim();
}

function isSubterenFromGeo_(f) {
  var p = f.properties || {};
  return !!(p.isSubteren || p.parentId);
}

function sortTerenyMeta_(rows) {
  rows.sort(function (a, b) {
    var aId = String(a.teren_nr);
    var bId = String(b.teren_nr);
    var aM = aId.match(/^(\d+)(.*)/);
    var bM = bId.match(/^(\d+)(.*)/);
    if (aM && bM) {
      var d = parseInt(aM[1], 10) - parseInt(bM[1], 10);
      if (d !== 0) return d;
      return aM[2].localeCompare(bM[2]);
    }
    return aId.localeCompare(bId);
  });
  return rows;
}

/**
 * Scal tereny.geojson z zakładką Tereny.
 * Zachowuje istniejące przydziały i daty; aktualizuje nazwę i parent_teren_nr z GeoJSON.
 * Uruchom z edytora Apps Script: importTerenyFromGeoJson()
 */
function importTerenyFromGeoJson() {
  var result = importTerenyMetaFromGeoJson_({ ensureOpracowanie: true });
  Logger.log(JSON.stringify(result));
  return result;
}

function importTerenyMetaFromGeoJson_(options) {
  options = options || {};
  var geo = getTerenyGeoJson_();
  var features = geo.features || [];
  var existing = readTerenyMeta_();
  var byId = {};

  for (var i = 0; i < existing.length; i++) {
    byId[String(existing[i].teren_nr)] = existing[i];
  }

  var added = 0;
  var updated = 0;

  for (var f = 0; f < features.length; f++) {
    var feat = features[f];
    var id = getFeatureIdFromGeo_(feat);
    if (!id) continue;

    var p = feat.properties || {};
    var nazwa = getFeatureTitleFromGeo_(feat);
    var parent = p.parentId ? String(p.parentId) : '';

    if (byId[id]) {
      byId[id].nazwa = nazwa;
      byId[id].parent_teren_nr = parent;
      updated++;
    } else {
      byId[id] = defaultMeta_(id, nazwa);
      byId[id].parent_teren_nr = parent;
      added++;
    }
  }

  var merged = [];
  Object.keys(byId).forEach(function (k) { merged.push(byId[k]); });
  sortTerenyMeta_(merged);
  writeTerenyMeta_(merged);

  var opracowanieEnsured = 0;
  if (options.ensureOpracowanie === true) {
    opracowanieEnsured = ensureOpracowanieSheetsFromGeo_(features);
  }

  return {
    ok: true,
    added: added,
    updated: updated,
    total: merged.length,
    opracowanieEnsured: opracowanieEnsured
  };
}

function ensureOpracowanieSheetsFromGeo_(features) {
  var nums = {};
  for (var i = 0; i < features.length; i++) {
    if (isSubterenFromGeo_(features[i])) continue;
    var id = getFeatureIdFromGeo_(features[i]);
    var num = parseTerenNum_(id);
    if (!isNaN(num) && num > 0 && num < 900) nums[num] = true;
  }

  var numList = Object.keys(nums).map(function (n) { return parseInt(n, 10); }).sort(function (a, b) { return a - b; });
  if (!numList.length) return 0;

  var chunkStarts = {};
  for (var j = 0; j < numList.length; j++) {
    var n = numList[j];
    var start = Math.floor((n - 1) / OPRACOWANIE_CHUNK) * OPRACOWANIE_CHUNK + 1;
    chunkStarts[start] = true;
  }

  var ensured = 0;
  Object.keys(chunkStarts).forEach(function (startKey) {
    var start = parseInt(startKey, 10);
    var end = start + OPRACOWANIE_CHUNK - 1;
    var range = start + '-' + end;
    var sh = findSheet_(opracowanieSheetName_(range));
    if (!sh) return;

    for (var t = start; t <= end; t++) {
      if (!nums[t]) continue;
      var row = terenRowInOpracowanie_(t, start);
      var val = sh.getRange(row, 1).getValue();
      if (!val && val !== 0) {
        sh.getRange(row, 1).setValue(t);
        ensured++;
      }
    }
  });
  return ensured;
}

// ── Opracowanie sheets ──────────────────────────────────────────────────────

function detectOpracowanieRanges_() {
  var matches = findSheetsByPrefix_(SHEET.opracowaniePrefix);
  var out = [];
  for (var i = 0; i < matches.length; i++) {
    var m = matches[i].suffix.match(/(\d+)\s*[-–]\s*(\d+)/);
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
  if (isNaN(num) || num < 1 || num >= 900) return null;
  var start = Math.floor((num - 1) / OPRACOWANIE_CHUNK) * OPRACOWANIE_CHUNK + 1;
  var end = start + OPRACOWANIE_CHUNK - 1;
  var range = start + '-' + end;
  var name = opracowanieSheetName_(range);
  var sh = findSheet_(name);
  if (!sh) return null;
  return { sheet: sh, range: range, start: start, end: end };
}

function ensureOpracowanieSheet_(range, start, end) {
  var name = opracowanieSheetName_(range);
  var sh = findSheet_(name);
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
  var sh = findSheet_(opracowanieSheetName_(range));
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

function actionUnassign_(body) {
  var terenId = String(body.terenId);
  var meta = getTerenMetaById_(terenId) || defaultMeta_(terenId, '');

  meta.przypisany_typ = '';
  meta.przypisany_id = '';
  meta.data_przydzialu = '';
  if (body.clearHistory) {
    meta.completions = [];
  }

  upsertTerenMeta_(meta);
  return { ok: true, meta: meta };
}
