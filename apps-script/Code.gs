/**
 * Arbitrage – backend Google Sheet
 * -----------------------------------------------------------
 * À coller dans Extensions > Apps Script, depuis le Google Sheet
 * qui doit contenir les données. Voir README.md pour le déploiement.
 *
 * Le site envoie / reçoit tout en JSON. Ce script :
 *   - refuse toute requête qui n'a pas le bon mot de passe (TOKEN)
 *   - lit les 3 onglets et renvoie les données + un numéro de version
 *   - écrit les 3 onglets (avec les mêmes formules que le fichier Excel)
 *   - refuse d'écraser si le Sheet a changé depuis la dernière lecture
 */

// ▼▼▼ CHANGE CE MOT DE PASSE (long, unique) puis redéploie ▼▼▼
const TOKEN = 'CHANGE-MOI';
// ▲▲▲ Tu le taperas une fois dans le site, sur chaque appareil ▲▲▲

const TAB_ROWS = 'Feuille 1';
const TAB_TARIFFS = 'Liste de prix';
const TAB_ARENAS = 'Liste des arénas';

const ROW_HEADERS = ['Date', 'Aréna', 'Catégorie', 'Prix', 'Déplacement à payer', 'Déplacement',
  'Supplément', 'Total', 'Comptabilisé', 'Commentaire', 'Tournoi', 'Payé'];
const TARIFF_HEADERS = ['Catégorie', 'Variante', 'Description', 'Type', 'Coût', 'Saison'];
const ARENA_HEADERS = ['Aréna', 'Coût de déplacement'];

/* ------------------------------ Entrées HTTP ------------------------------ */

function doGet() {
  return out({ ok: true, message: 'Le service Arbitrage répond. Utilise le site pour accéder aux données.' });
}

function doPost(e) {
  try {
    const req = JSON.parse(e.postData.contents);
    if (req.token !== TOKEN) return out({ ok: false, error: 'auth' });
    if (req.action === 'load') return out(Object.assign({ ok: true }, snapshot()));
    if (req.action === 'save') return out(save(req));
    return out({ ok: false, error: 'Action inconnue' });
  } catch (err) {
    return out({ ok: false, error: String(err) });
  }
}

function out(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* --------------------------------- Lecture -------------------------------- */

function sheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const get = (name) => ss.getSheetByName(name);
  let rows = get(TAB_ROWS);
  if (!rows) {
    // Un Sheet neuf contient un onglet vide : on le renomme plutôt que d'en créer un second.
    const all = ss.getSheets();
    if (all.length === 1 && all[0].getLastRow() === 0) { rows = all[0]; rows.setName(TAB_ROWS); }
    else rows = ss.insertSheet(TAB_ROWS);
  }
  const tariffs = get(TAB_TARIFFS) || ss.insertSheet(TAB_TARIFFS);
  const arenas = get(TAB_ARENAS) || ss.insertSheet(TAB_ARENAS);
  return { ss: ss, rows: rows, tariffs: tariffs, arenas: arenas };
}

function bool(v) { return v === true || String(v).toUpperCase() === 'TRUE'; }
function num(v) { const n = Number(v); return isFinite(n) ? n : 0; }

// Ancien format : case à cocher « Double lettre ». Nouveau format : texte libre (« Simple lettre », vide, etc.).
function variantOf(v) {
  if (v === true || String(v).toUpperCase() === 'TRUE') return 'Double lettre';
  if (v === false || String(v).toUpperCase() === 'FALSE') return 'Simple lettre';
  return String(v);
}

function body(sh, ncols) {
  const last = sh.getLastRow();
  return last < 2 ? [] : sh.getRange(2, 1, last - 1, ncols).getValues();
}

function snapshot() {
  const s = sheets();
  const tz = s.ss.getSpreadsheetTimeZone();

  const rows = body(s.rows, 12)
    .filter(function (r) { return r[0] !== '' || r[1] !== '' || r[2] !== ''; })
    .map(function (r) {
      return {
        date: r[0] instanceof Date ? Utilities.formatDate(r[0], tz, 'yyyy-MM-dd') : String(r[0]),
        arena: String(r[1]),
        category: String(r[2]),
        travel: bool(r[4]),
        extra: num(r[6]),
        counted: bool(r[8]),
        comment: String(r[9]),
        tournament: bool(r[10]),
        paid: bool(r[11])
      };
    });

  const tariffs = body(s.tariffs, 6)
    .filter(function (r) { return r[0] !== '' || r[3] !== ''; })
    .map(function (r) {
      return { season: String(r[5]), cat: String(r[0]), variant: variantOf(r[1]), type: String(r[3]), cost: num(r[4]) };
    });

  const arenas = body(s.arenas, 2)
    .filter(function (r) { return r[0] !== ''; })
    .map(function (r) { return { name: String(r[0]), travelCost: num(r[1]) }; });

  const data = { rows: rows, tariffs: tariffs, arenas: arenas };
  return { data: data, version: hash(data) };
}

function hash(data) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, JSON.stringify(data), Utilities.Charset.UTF_8);
  return bytes.map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
}

/* -------------------------------- Écriture -------------------------------- */

function save(req) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const current = snapshot();
    if (!req.force && req.baseVersion && current.version !== req.baseVersion) {
      return { ok: false, conflict: true, data: current.data, version: current.version };
    }
    writeAll(req.data);
    SpreadsheetApp.flush();
    return { ok: true, version: snapshot().version };
  } finally {
    lock.releaseLock();
  }
}

function resetBody(sh, ncols) {
  const last = Math.max(sh.getLastRow(), 2);
  const rg = sh.getRange(2, 1, last - 1, ncols);
  rg.clearContent();
  rg.clearDataValidations();
}

// Midi plutôt que minuit : la date reste la bonne même si les fuseaux du script et du Sheet diffèrent.
function sheetDate(str, tz) {
  if (!str) return '';
  if (/^\d{4}-\d{2}-\d{2}$/.test(str)) return Utilities.parseDate(str + ' 12:00', tz, 'yyyy-MM-dd HH:mm');
  return str;
}

function header(sh, labels) {
  sh.getRange(1, 1, 1, labels.length).setValues([labels])
    .setFontWeight('bold').setBackground('#10263A').setFontColor('#FFFFFF');
  sh.setFrozenRows(1);
}

function writeAll(d) {
  const s = sheets();
  const tz = s.ss.getSpreadsheetTimeZone();

  /* Liste de prix */
  header(s.tariffs, TARIFF_HEADERS);
  resetBody(s.tariffs, 6);
  if (d.tariffs.length) {
    const vals = d.tariffs.map(function (t, i) {
      const n = i + 2;
      return [t.cat, t.variant || '', '=TEXTJOIN(" - ",TRUE,A' + n + ',B' + n + ',D' + n + ')',
        t.type, num(t.cost), t.season || ''];
    });
    s.tariffs.getRange(2, 1, vals.length, 6).setValues(vals);
    s.tariffs.getRange(2, 5, vals.length, 1).setNumberFormat('#,##0.00 $');
  }
  s.tariffs.setColumnWidths(1, 6, 150);
  s.tariffs.setColumnWidth(3, 330);

  /* Liste des arénas */
  header(s.arenas, ARENA_HEADERS);
  resetBody(s.arenas, 2);
  if (d.arenas.length) {
    const vals = d.arenas.map(function (a) { return [a.name, num(a.travelCost)]; });
    s.arenas.getRange(2, 1, vals.length, 2).setValues(vals);
    s.arenas.getRange(2, 2, vals.length, 1).setNumberFormat('#,##0.00 $');
  }
  s.arenas.setColumnWidth(1, 300);
  s.arenas.setColumnWidth(2, 170);

  /* Feuille 1 : mêmes formules que le fichier Excel d'origine */
  header(s.rows, ROW_HEADERS);
  resetBody(s.rows, 12);
  const n = d.rows.length;
  if (n) {
    const vals = d.rows.map(function (r, i) {
      const k = i + 2;
      return [
        sheetDate(r.date, tz),
        r.arena || '',
        r.category || '',
        '=IF(C' + k + '<>"",IFERROR(VLOOKUP(C' + k + ',\'Liste de prix\'!$C$2:$E$500,3,FALSE),0),0)',
        !!r.travel,
        '=IF(E' + k + '=TRUE,IFERROR(VLOOKUP(B' + k + ',\'Liste des arénas\'!$A$2:$B$200,2,FALSE),0),0)',
        num(r.extra),
        '=SUM(D' + k + ',F' + k + ',G' + k + ')',
        !!r.counted,
        r.comment || '',
        !!r.tournament,
        !!r.paid
      ];
    });
    s.rows.getRange(2, 1, n, 12).setValues(vals);
    [5, 9, 11, 12].forEach(function (c) { s.rows.getRange(2, c, n, 1).insertCheckboxes(); });
    s.rows.getRange(2, 1, n, 1).setNumberFormat('dd/MM/yyyy');
    [4, 6, 7, 8].forEach(function (c) { s.rows.getRange(2, c, n, 1).setNumberFormat('#,##0.00 $'); });
  }

  // Totaux (à droite du tableau, ne bougent pas quand on ajoute des lignes)
  s.rows.getRange('N1:O5').setValues([
    ['Totaux', ''],
    ['Prix', '=SUM(D2:D)'],
    ['Déplacement', '=SUM(F2:F)'],
    ['Supplément', '=SUM(G2:G)'],
    ['Total', '=SUM(H2:H)']
  ]);
  s.rows.getRange('N1:O1').setFontWeight('bold');
  s.rows.getRange('N5:O5').setFontWeight('bold');
  s.rows.getRange('O2:O5').setNumberFormat('#,##0.00 $');

  // Ligne verte quand « Payé » est coché (comme dans l'Excel)
  if (s.rows.getConditionalFormatRules().length === 0) {
    const rule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$L2=TRUE')
      .setBackground('#B7E1CD')
      .setRanges([s.rows.getRange('A2:L1000')])
      .build();
    s.rows.setConditionalFormatRules([rule]);
  }

  s.rows.setFrozenRows(1);
  const widths = [100, 260, 320, 80, 130, 100, 100, 90, 110, 200, 80, 60];
  widths.forEach(function (w, i) { s.rows.setColumnWidth(i + 1, w); });
}
