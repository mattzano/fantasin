/**
 * FantaSIN – backend su Google Fogli (Apps Script)
 *
 * PRIMA VOLTA
 *   1. Incolla questo file (Code.gs) e il file Arbitro.html nel progetto Apps Script del Foglio.
 *   2. Esegui la funzione "setup": crea i fogli, carica relatori, programma, poster e bingo
 *      da dati.json, imposta formule e protezioni.
 *   3. Nel foglio "Arbitri" scrivi nome, PIN, sala e ruolo di ogni arbitro.
 *   4. Deploy → Manage deployments → matita → Version: New version → Deploy.
 *
 * PASSAGGIO ALLA VERSIONE SEMPLIFICATA (12 sessioni, 3 relatori + 2 moderatori, bingo)
 *   Esegui una volta "ricaricaGioco": svuota Relatori, Sessioni, Programma, Squadre, Puntate, Log
 *   e Bingo, poi ricarica tutto da dati.json. Poi pubblica una nuova versione.
 *
 * LINK
 *   Dati per FantaSIN (JSON):  .../exec
 *   App arbitri (telefono):    .../exec?page=arbitro
 */

const TZ = 'Europe/Rome';
const DATES = ['2026-10-24', '2026-10-25', '2026-10-26', '2026-10-27'];
const DAY_LABEL = ['Sab 24', 'Dom 25', 'Lun 26', 'Mar 27'];
const CARD = 24;   // caselle con frase (la centrale è libera)

const HEAD = {
  Config:     ['Chiave', 'Valore', 'Nota'],
  Relatori:   ['ID', 'Nome', 'Città', 'Prezzo', 'Presenze', 'Relatore', 'Moderatore'],
  Sessioni:   ['ID', 'Giorno', 'Inizio', 'Fine', 'Sala', 'Tipo', 'Titolo', 'Comune'],
  Programma:  ['Slot', 'Sessione', 'Giorno', 'Ora', 'Sala', 'ID relatore', 'Relatore', 'Ruolo', 'Titolo', 'Stato'],
  Poster:     ['ID', 'Giorno', 'Numero', 'Titolo', 'Primo autore'],
  Eventi:     ['Codice', 'Descrizione', 'Punti', 'Categoria', 'Tetto giornaliero', 'Per ruolo'],
  Arbitri:    ['Nome', 'PIN', 'Sala', 'Ruolo'],
  Log:        ['Ora', 'Giorno', 'Sala', 'Arbitro', 'Relatore', 'ID relatore', 'Evento', 'Punti base', 'Punti', 'Stato', 'Slot', 'ID'],
  Squadre:    ['Nickname', 'Area', 'Sede', 'Relatore 1', 'Relatore 2', 'Relatore 3', 'Moderatore 1', 'Moderatore 2', 'Capitano',
               'Costo', 'Budget', 'Punti relatori', 'Punti scommesse', 'Punti bingo', 'Totale', 'Aggiornata', 'PID', 'Cartella'],
  Scommesse:  ['ID', 'Domanda', 'Tipo', 'Chiusura', 'Risultato', 'Punti'],
  Puntate:    ['PID', 'Nickname', 'Scommessa', 'Scelta', 'Ora', 'Punti'],
  Bingo:      ['ID', 'Frase', 'Categoria', 'Sentita', 'Ora', 'Arbitro'],
  Classifica: ['Pos', 'Nickname', 'Area', 'Sede', 'Punti']
};
// colonne Squadre (0-based)
const SQ = { nick: 0, area: 1, sede: 2, r1: 3, m1: 6, cap: 8, cost: 9, ok: 10, pr: 11, pb: 12, pg: 13, tot: 14, ts: 15, pid: 16, card: 17 };

const AREE = ['Cefalee', 'Demenze e cognitività', 'Disturbi del movimento', 'Epilessia', 'Malattie cerebrovascolari',
  'Malattie neuromuscolari', 'Neuroimmunologia e SM', 'Neuro-oncologia', 'Neurofisiologia', 'Neuroriabilitazione',
  'Sonno', 'Neurologia generale', 'Altro'];

/* ====================== helpers ====================== */
function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }
function sh_(name) {
  let s = ss_().getSheetByName(name);
  if (!s) { s = ss_().insertSheet(name); header_(s, name); }
  return s;
}
function header_(s, name) {
  s.getRange(1, 1, 1, HEAD[name].length).setValues([HEAD[name]]).setFontWeight('bold'); s.setFrozenRows(1);
}
function rows_(name) {
  const s = sh_(name), n = s.getLastRow();
  if (n < 2) return [];
  const v = s.getRange(2, 1, n - 1, HEAD[name].length).getValues();
  const dc = { Sessioni: 1, Programma: 2, Poster: 1, Log: 1 }[name];
  if (dc != null) v.forEach(r => { r[dc] = dayFix_(r[dc]); });
  return v;
}
function out_(o, cb) {
  const txt = JSON.stringify(o);
  if (cb && /^[A-Za-z_$][\w$]*$/.test(cb))
    return ContentService.createTextOutput(cb + '(' + txt + ')').setMimeType(ContentService.MimeType.JAVASCRIPT);
  return ContentService.createTextOutput(txt).setMimeType(ContentService.MimeType.JSON);
}
function clip_(x, n) {
  const v = String(x == null ? '' : x).replace(/[\u0000-\u001f]/g, ' ').trim().slice(0, n);
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}
function pid_(token) {
  const d = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(token), Utilities.Charset.UTF_8);
  return 'p' + d.map(b => ('0' + (b & 255).toString(16)).slice(-2)).join('').slice(0, 16);
}
function config_() {
  const c = {};
  rows_('Config').forEach(r => { if (r[0]) c[String(r[0])] = r[1]; });
  return c;
}
function num_(v, d) { const n = Number(v); return isFinite(n) && v !== '' ? n : d; }
function todayIdx_() {
  const i = DATES.indexOf(Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'));
  return i < 0 ? (new Date() < new Date(DATES[0] + 'T00:00:00+02:00') ? 0 : 3) : i;
}
function asDate_(v) { return v instanceof Date ? v : (v ? new Date(v) : null); }
function dayFix_(v) {
  if (v instanceof Date) {
    const i = [24, 25, 26, 27].indexOf(v.getDate());
    return i >= 0 ? DAY_LABEL[i] : Utilities.formatDate(v, TZ, 'dd/MM');
  }
  return v;
}
function fixDays_() {
  [['Sessioni', 2], ['Programma', 3], ['Poster', 2], ['Log', 2]].forEach(([name, col]) => {
    const s = sh_(name), n = s.getLastRow();
    if (n < 2) return;
    const rg = s.getRange(2, col, n - 1, 1), v = rg.getValues();
    if (v.some(r => r[0] instanceof Date)) rg.setNumberFormat('@').setValues(v.map(r => [dayFix_(r[0])]));
  });
}
function safe_(o) { return JSON.parse(JSON.stringify(o)); }
function bust_() { CacheService.getScriptCache().remove('state'); }
function eventi_() {
  const m = {};
  rows_('Eventi').forEach(r => { if (r[0]) m[r[0]] = { code: r[0], label: r[1], p: Number(r[2]) || 0, cat: r[3], cap: Math.abs(Number(r[4]) || 0), role: r[5] || '' }; });
  return m;
}
function people_() {
  const m = {};
  rows_('Relatori').forEach(r => { if (r[0]) m[r[0]] = { n: r[1], p: Number(r[3]) || 0, rel: r[5] === 'sì', mod: r[6] === 'sì' }; });
  return m;
}
function setCfg_(key, val, note) {
  const s = sh_('Config'), v = rows_('Config'), i = v.findIndex(r => r[0] === key);
  if (i >= 0) s.getRange(i + 2, 2).setValue(val); else s.appendRow([key, val, note || '']);
}

/* ====================== setup ====================== */
function setup() {
  Object.keys(HEAD).forEach(sh_);
  const def = ss_().getSheetByName('Foglio1') || ss_().getSheetByName('Sheet1');
  if (def && ss_().getSheets().length > 1 && def.getLastRow() === 0) ss_().deleteSheet(def);

  const cfg = sh_('Config');
  if (cfg.getLastRow() < 2) {
    cfg.getRange(2, 1, 5, 3).setValues([
      ['budget', 100, 'Neuroni a disposizione per la squadra'],
      ['scadenza_squadre', new Date('2026-10-25T08:30:00+01:00'), 'Dopo questa ora le squadre non si modificano più'],
      ['dati_url', 'https://fantasin.it/dati.json', 'Da dove setup() carica relatori, programma, poster e bingo'],
      ['email_responsabile', '', 'Facoltativo: email Google del responsabile, che può modificare i fogli'],
      ['bonus_soggettivi', 'solo sessioni comuni', 'Promemoria: i bonus soggettivi valgono solo nelle sessioni comuni']
    ]);
  }
  ensureCfg_();

  const ev = sh_('Eventi');
  if (ev.getLastRow() < 2) {
    ev.getRange(2, 1, 10, 6).setValues([
      ['REL', 'Relazione tenuta', 10, 'oggettivo', 20, 'REL'],
      ['MOD', 'Moderazione', 6, 'oggettivo', 12, 'MOD'],
      ['COM', 'Comunicazione orale', 8, 'oggettivo', 16, 'COM'],
      ['ASS', 'Assente o sostituito', -5, 'oggettivo', 10, ''],
      ['FIR', 'Cita Firenze', 5, 'soggettivo', 10, ''],
      ['MEME', 'Battuta o meme', 5, 'soggettivo', 10, ''],
      ['SFO', 'Sfora di oltre 10 minuti', -5, 'soggettivo', 10, ''],
      ['MIC', 'Microfono in tilt', -3, 'soggettivo', 6, ''],
      ['PREMIO', 'Premio o riconoscimento', 20, 'organizzatori', 20, ''],
      ['POSTER', 'Poster premiato', 30, 'organizzatori', 30, '']
    ]);
  }

  const ar = sh_('Arbitri');
  if (ar.getLastRow() < 2) {
    const pin = () => String(Math.floor(1000 + Math.random() * 9000));
    ar.getRange(2, 1, 5, 4).setValues([
      ['Responsabile', pin(), 'Tutte', 'organizzatore'],
      ['Arbitro Auditorium', pin(), 'Auditorium', 'arbitro'],
      ['Arbitro Sala 4', pin(), 'Sala 4', 'arbitro'],
      ['Arbitro Sala 5', pin(), 'Sala 5', 'arbitro'],
      ['Arbitro Sala 6', pin(), 'Sala 6', 'arbitro']
    ]);
    ar.getRange('C2:C').setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInList(['Tutte', 'Auditorium', 'Sala 1', 'Sala 2', 'Sala 3', 'Sala 4', 'Sala 5', 'Sala 6', 'Sala 7', 'Sala 8', 'Sala 9'], true).build());
    ar.getRange('D2:D').setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInList(['arbitro', 'responsabile', 'organizzatore'], true).build());
  }

  const sc = sh_('Scommesse');
  if (sc.getLastRow() < 2) {
    const close = new Date('2026-10-27T12:00:00+01:00');
    sc.getRange(2, 1, 5, 6).setValues([
      ['POSTER25', 'Miglior poster di domenica 25', 'poster:1', close, '', 30],
      ['POSTER26', 'Miglior poster di lunedì 26', 'poster:2', close, '', 30],
      ['POSTER27', 'Miglior poster di martedì 27', 'poster:3', close, '', 30],
      ['SFORA', 'Sessione che sfora di più', 'sessione', new Date('2026-10-25T08:30:00+01:00'), '', 15],
      ['APPLAUSI', 'Relatore più applaudito', 'relatore', new Date('2026-10-25T08:30:00+01:00'), '', 15]
    ]);
  }

  if (sh_('Relatori').getLastRow() < 2) loadData_();
  formulas_();
  fixDays_();
  protect_();
  SpreadsheetApp.flush();
  Logger.log('Setup completato. PIN arbitri nel foglio "Arbitri".');
}

function ensureCfg_() {
  const c = config_();
  if (c.relatori_squadra === undefined) setCfg_('relatori_squadra', 3, 'Relatori per squadra');
  if (c.moderatori_squadra === undefined) setCfg_('moderatori_squadra', 2, 'Moderatori per squadra');
  if (c.bingo_riga === undefined) setCfg_('bingo_riga', 10, 'Punti per ogni riga, colonna o diagonale del bingo');
  if (c.bingo_piena === undefined) setCfg_('bingo_piena', 30, 'Punti extra per la cartella piena');
}

function loadData_() {
  const data = JSON.parse(UrlFetchApp.fetch(config_().dati_url + '?v=' + Date.now()).getContentText());
  const put = (name, rows) => { if (rows.length) sh_(name).getRange(2, 1, rows.length, rows[0].length).setValues(rows); };
  put('Relatori', data.roster.map(r => [r.id, r.n, r.c, r.p, r.k, r.rel ? 'sì' : 'no', r.mod ? 'sì' : 'no']));
  put('Sessioni', data.sessions.map(s => [s.id, "'" + DAY_LABEL[s.d], "'" + s.t, "'" + s.e, s.sala, s.k, s.ti, s.c ? 'sì' : 'no']));
  const names = {}; data.roster.forEach(r => names[r.id] = r.n);
  put('Programma', data.slots.map(x => [x.id, x.s, "'" + DAY_LABEL[x.d], "'" + x.t, x.sala, x.r, names[x.r] || x.r, x.ro, x.ti, '']));
  if (sh_('Poster').getLastRow() < 2) put('Poster', data.posters.map(p => [p.id, "'" + DAY_LABEL[p.d], p.n, p.ti, p.a]));
  if (sh_('Bingo').getLastRow() < 2 && data.bingo) put('Bingo', data.bingo.map(b => [b.id, b.t, b.cat, '', '', '']));
}

function formulas_() {
  const cl = sh_('Classifica');
  cl.getRange('B2').setFormula('=IFERROR(SORT(FILTER({Squadre!A2:A,Squadre!B2:B,Squadre!C2:C,Squadre!O2:O},Squadre!A2:A<>"",Squadre!K2:K="OK"),4,FALSE),)');
  cl.getRange('A2').setFormula('=ARRAYFORMULA(IF(E2:E="",,COUNTIF(E2:E,">"&E2:E)+1))');
}

/** Passaggio al gioco semplificato: ricarica programma e azzera squadre, puntate, log e bingo. */
function ricaricaGioco() {
  ['Relatori', 'Sessioni', 'Programma', 'Squadre', 'Puntate', 'Log', 'Bingo', 'Classifica'].forEach(n => {
    const s = sh_(n);
    s.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(p => p.remove());
    s.clear(); header_(s, n);
  });
  ensureCfg_();
  loadData_();
  formulas_();
  protect_();
  bust_();
  SpreadsheetApp.flush();
  Logger.log('Gioco ricaricato: ' + (sh_('Relatori').getLastRow() - 1) + ' persone, ' + (sh_('Sessioni').getLastRow() - 1) + ' sessioni, ' + (sh_('Bingo').getLastRow() - 1) + ' frasi bingo.');
}

function protect_() {
  const me = Session.getEffectiveUser();
  const resp = String(config_().email_responsabile || '').trim();
  ss_().getSheets().forEach(s => {
    s.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(p => p.remove());
    const p = s.protect().setDescription('FantaSIN: modifiche solo tramite app arbitri');
    p.removeEditors(p.getEditors());
    p.addEditor(me);
    if (resp) { try { p.addEditor(resp); } catch (e) {} }
    if (p.canDomainEdit()) p.setDomainEdit(false);
  });
}

/* ====================== cap / scoring ====================== */
function recompute_() {
  const s = sh_('Log'), n = s.getLastRow();
  if (n < 2) return;
  const v = rows_('Log');
  const caps = eventi_(), used = {};
  const outCol = v.map(r => {
    if (r[9] !== 'ok') return [0];
    const base = Number(r[7]) || 0, ev = caps[r[6]], cap = ev ? ev.cap : 0;
    if (!cap) return [base];
    const k = r[1] + '|' + r[5] + '|' + r[6], u = used[k] || 0;
    const eff = Math.sign(base) * Math.min(Math.abs(base), Math.max(0, cap - u));
    used[k] = u + Math.abs(eff);
    return [eff];
  });
  s.getRange(2, 9, outCol.length, 1).setValues(outCol);
}

/* ====================== bingo ====================== */
const LINES = (() => {
  const L = [];
  for (let r = 0; r < 5; r++) L.push([0, 1, 2, 3, 4].map(c => r * 5 + c));
  for (let c = 0; c < 5; c++) L.push([0, 1, 2, 3, 4].map(r => r * 5 + c));
  L.push([0, 6, 12, 18, 24]); L.push([4, 8, 12, 16, 20]);
  return L;
})();
function called_() { const c = {}; rows_('Bingo').forEach(r => { if (r[0] && r[3] === 'sì') c[r[0]] = true; }); return c; }
function newCard_() {
  const ids = rows_('Bingo').map(r => r[0]).filter(Boolean);
  for (let i = ids.length - 1; i > 0; i--) { const j = Math.floor(Math.random() * (i + 1)); const t = ids[i]; ids[i] = ids[j]; ids[j] = t; }
  return ids.slice(0, CARD);
}
function bingoScore_(card, called, cfg) {
  if (!card || card.length !== CARD) return { lines: 0, full: false, pts: 0 };
  const cells = card.slice(0, 12).concat(['FREE'], card.slice(12));
  const on = cells.map(id => id === 'FREE' || !!called[id]);
  const lines = LINES.filter(l => l.every(i => on[i])).length;
  const full = on.every(Boolean);
  return { lines: lines, full: full, pts: lines * num_(cfg.bingo_riga, 10) + (full ? num_(cfg.bingo_piena, 30) : 0) };
}
function recomputeBingo_() {
  const s = sh_('Squadre'), v = rows_('Squadre');
  if (!v.length) return;
  const called = called_(), cfg = config_();
  s.getRange(2, SQ.pg + 1, v.length, 1).setValues(v.map(r => [r[SQ.nick] ? bingoScore_(String(r[SQ.card]).split(',').filter(Boolean), called, cfg).pts : '']));
}

/* ====================== public JSON (FantaSIN) ====================== */
function state_() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get('state');
  if (hit) return JSON.parse(hit);
  SpreadsheetApp.flush();
  const cfg = config_();
  const teams = rows_('Squadre').filter(r => r[SQ.nick]).map(r => ({
    pid: r[SQ.pid], nick: r[SQ.nick], area: r[SQ.area], sede: r[SQ.sede],
    rel: [r[SQ.r1], r[SQ.r1 + 1], r[SQ.r1 + 2]], mods: [r[SQ.m1], r[SQ.m1 + 1]], cap: r[SQ.cap],
    cost: Number(r[SQ.cost]) || 0, ok: r[SQ.ok] === 'OK', pr: Number(r[SQ.pr]) || 0, pb: Number(r[SQ.pb]) || 0,
    pg: Number(r[SQ.pg]) || 0, tot: Number(r[SQ.tot]) || 0, card: String(r[SQ.card] || '').split(',').filter(Boolean)
  }));
  const ranked = teams.filter(t => t.ok).sort((a, b) => b.tot - a.tot);
  ranked.forEach(t => { t.pos = 1 + ranked.filter(x => x.tot > t.tot).length; });
  const log = rows_('Log').filter(r => r[9] === 'ok');
  const pts = {};
  log.forEach(r => { pts[r[5]] = (pts[r[5]] || 0) + (Number(r[8]) || 0); });
  const feed = log.slice(-25).reverse().map(r => ({
    t: Utilities.formatDate(asDate_(r[0]) || new Date(), TZ, 'HH:mm'), d: r[1], sala: r[2], n: r[4], r: r[5], e: r[6], p: Number(r[8]) || 0
  }));
  const ev = eventi_();
  const bets = rows_('Scommesse').filter(r => r[0]).map(r => ({
    id: r[0], q: r[1], type: r[2], close: asDate_(r[3]) ? asDate_(r[3]).getTime() : null, result: String(r[4] || ''), p: Number(r[5]) || 0
  }));
  const bingo = rows_('Bingo').filter(r => r[0]).map(r => ({ id: r[0], t: r[1], cat: r[2], on: r[3] === 'sì',
    at: r[3] === 'sì' && asDate_(r[4]) ? Utilities.formatDate(asDate_(r[4]), TZ, 'EEE HH:mm') : '' }));
  const st = {
    ok: true, now: Date.now(),
    config: {
      budget: num_(cfg.budget, 100), nrel: num_(cfg.relatori_squadra, 3), nmod: num_(cfg.moderatori_squadra, 2),
      deadline: asDate_(cfg.scadenza_squadre) ? asDate_(cfg.scadenza_squadre).getTime() : null, aree: AREE,
      bingoRiga: num_(cfg.bingo_riga, 10), bingoPiena: num_(cfg.bingo_piena, 30)
    },
    teams: teams, pts: pts, feed: feed, bets: bets, bingo: bingo,
    eventi: Object.values(ev).map(e => ({ code: e.code, label: e.label, p: e.p, cat: e.cat, cap: e.cap }))
  };
  cache.put('state', JSON.stringify(st), 15);
  return st;
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.page === 'arbitro') {
    return HtmlService.createHtmlOutputFromFile('Arbitro')
      .setTitle('FantaSIN · Arbitri')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
  }
  if (p.action === 'prezzi') {
    return out_({ ok: true, prices: Object.fromEntries(Object.entries(people_()).map(([k, v]) => [k, v.p])) }, p.callback);
  }
  const st = state_();
  if (p.pid) {
    st.mybets = rows_('Puntate').filter(r => r[0] === p.pid).map(r => ({ id: r[2], choice: String(r[3]), p: r[5] }));
  }
  return out_(st, p.callback);
}

/* ====================== player actions (POST, text/plain) ====================== */
function doPost(e) {
  let q;
  try { q = JSON.parse(e.postData.contents); } catch (err) { return out_({ ok: false, error: 'Richiesta non valida' }); }
  if (!q.token || String(q.token).length < 16) return out_({ ok: false, error: 'Richiesta non valida' });
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return out_({ ok: false, error: 'Server occupato, riprova tra qualche secondo' });
  try {
    const pid = pid_(q.token);
    if (q.action === 'saveTeam') return out_(saveTeam_(pid, q));
    if (q.action === 'placeBet') return out_(placeBet_(pid, q));
    return out_({ ok: false, error: 'Azione sconosciuta' });
  } finally { lock.releaseLock(); }
}

function saveTeam_(pid, q) {
  const cfg = config_(), nrel = num_(cfg.relatori_squadra, 3), nmod = num_(cfg.moderatori_squadra, 2), budget = num_(cfg.budget, 100);
  const dl = asDate_(cfg.scadenza_squadre);
  if (dl && new Date() > dl) return { ok: false, error: 'Tempo scaduto: le squadre non si possono più modificare' };
  const nick = clip_(q.nick, 24).replace(/^'/, '');
  if (nick.length < 2) return { ok: false, error: 'Scegli un nickname di almeno 2 caratteri' };
  const area = AREE.indexOf(q.area) >= 0 ? q.area : 'Altro';
  const sede = clip_(q.sede, 40);
  const rel = (q.rel || []).map(String), mods = (q.mods || []).map(String), all = rel.concat(mods);
  const pe = people_();
  if (rel.length !== nrel || rel.some(p => !pe[p] || !pe[p].rel)) return { ok: false, error: 'Scegli ' + nrel + ' relatori dalla lista' };
  if (mods.length !== nmod || mods.some(p => !pe[p] || !pe[p].mod)) return { ok: false, error: 'Scegli ' + nmod + ' moderatori dalla lista' };
  if (new Set(all).size !== all.length) return { ok: false, error: 'Ogni persona può comparire una sola volta' };
  if (all.indexOf(String(q.cap)) < 0) return { ok: false, error: 'Il capitano deve essere uno dei tuoi cinque' };
  const cost = all.reduce((s, p) => s + pe[p].p, 0);
  if (cost > budget) return { ok: false, error: 'Budget superato: ' + cost + ' su ' + budget + ' Neuroni' };

  const s = sh_('Squadre'), v = rows_('Squadre');
  const clash = v.findIndex(r => String(r[SQ.nick]).toLowerCase() === nick.toLowerCase() && r[SQ.pid] !== pid);
  if (clash >= 0) return { ok: false, error: 'Nickname già usato da un altro giocatore' };
  const idx = v.findIndex(r => r[SQ.pid] === pid);
  const r = idx >= 0 ? idx + 2 : Math.max(2, s.getLastRow() + 1);
  const card = idx >= 0 && String(v[idx][SQ.card]).split(',').filter(Boolean).length === CARD ? String(v[idx][SQ.card]) : newCard_().join(',');
  const priceSum = ['D', 'E', 'F', 'G', 'H'].map(c => `SUMIFS(Relatori!$D:$D,Relatori!$A:$A,${c}${r})`).join('+');
  const logSum = ['D', 'E', 'F', 'G', 'H', 'I'].map(c => `SUMIFS(Log!$I:$I,Log!$F:$F,${c}${r},Log!$J:$J,"ok")`).join('+');
  s.getRange(r, 1, 1, 18).setValues([[
    clip_(nick, 24), area, sede, rel[0], rel[1], rel[2], mods[0], mods[1], q.cap,
    '=' + priceSum,
    `=IF(J${r}<=VLOOKUP("budget",Config!$A:$B,2,FALSE),"OK","FUORI BUDGET")`,
    '=' + logSum,
    `=SUMIFS(Puntate!$F:$F,Puntate!$A:$A,Q${r})`,
    bingoScore_(card.split(','), called_(), cfg).pts,
    `=IF(K${r}="OK",L${r}+M${r}+N${r},0)`,
    new Date(), pid, card
  ]]);
  bust_();
  return { ok: true, pid: pid, state: state_() };
}

function placeBet_(pid, q) {
  const bet = rows_('Scommesse').find(r => r[0] === q.bet);
  if (!bet) return { ok: false, error: 'Scommessa non trovata' };
  const close = asDate_(bet[3]);
  if (close && new Date() > close) return { ok: false, error: 'Scommessa chiusa' };
  if (bet[4] !== '' && bet[4] != null) return { ok: false, error: 'Risultato già inserito' };
  const team = rows_('Squadre').find(r => r[SQ.pid] === pid);
  if (!team) return { ok: false, error: 'Prima salva la tua squadra' };
  const choice = clip_(q.choice, 40);
  if (!choice) return { ok: false, error: 'Scegli un’opzione' };
  const s = sh_('Puntate'), v = rows_('Puntate');
  let row = v.findIndex(r => r[0] === pid && r[2] === q.bet);
  row = row >= 0 ? row + 2 : Math.max(2, s.getLastRow() + 1);
  const r = row;
  s.getRange(r, 1, 1, 6).setValues([[pid, team[SQ.nick], q.bet, choice, new Date(),
    `=IFERROR(IF(VLOOKUP(C${r},Scommesse!$A:$F,5,FALSE)="",0,IF(TO_TEXT(VLOOKUP(C${r},Scommesse!$A:$F,5,FALSE))=TO_TEXT(D${r}),VLOOKUP(C${r},Scommesse!$A:$F,6,FALSE),0)),0)`]]);
  bust_();
  return { ok: true, mybets: rows_('Puntate').filter(x => x[0] === pid).map(x => ({ id: x[2], choice: String(x[3]), p: x[5] })) };
}

/* ====================== referee app (google.script.run) ====================== */
function ref_(pin) {
  const r = rows_('Arbitri').find(x => String(x[1]).trim() !== '' && String(x[1]).trim() === String(pin).trim());
  if (!r) throw new Error('PIN non valido');
  const role = String(r[3] || 'arbitro');
  return { name: String(r[0]), sala: String(r[2] || ''), role: role, all: role !== 'arbitro' || r[2] === 'Tutte', org: role === 'organizzatore' };
}

function arbLogin(pin) {
  const me = ref_(pin);
  fixDays_();
  const roster = rows_('Relatori').map(r => [r[0], r[1], r[2]]);
  const sale = ['Auditorium', 'Sala 1', 'Sala 2', 'Sala 3', 'Sala 4', 'Sala 5', 'Sala 6', 'Sala 7', 'Sala 8', 'Sala 9'];
  const sessions = rows_('Sessioni').map(r => [r[0], r[1], String(r[2]), r[4], r[5], r[6]]);
  const days = DAY_LABEL.map((d, i) => i);
  const used = {}; sessions.forEach(s => used[s[3]] = true);
  return safe_({ me: me, roster: roster, sale: sale.filter(s => used[s]), today: todayIdx_(), days: DAY_LABEL, eventi: Object.values(eventi_()), sessions: sessions });
}

function arbData(pin, day, sala) {
  const me = ref_(pin);
  if (!me.all) sala = me.sala;
  const dl = DAY_LABEL[day];
  const ses = rows_('Sessioni').filter(r => r[1] === dl && r[4] === sala)
    .map(r => ({ id: r[0], t: String(r[2]), e: String(r[3]), k: r[5], ti: r[6], c: r[7] === 'sì' }));
  const names = {}; rows_('Relatori').forEach(r => names[r[0]] = r[1]);
  const log = rows_('Log');
  const bySlot = {}, subj = {};
  log.forEach(r => {
    if (r[9] !== 'ok') return;
    if (r[10]) {
      if (String(r[10]).startsWith('SL')) bySlot[r[10]] = { e: r[6], by: r[3] };
      else { const k = r[10] + '|' + r[5]; (subj[k] = subj[k] || []).push(r[6]); }
    }
  });
  const slots = rows_('Programma').filter(r => r[2] === dl && r[4] === sala).map(r => ({
    id: r[0], s: r[1], t: String(r[3]), r: r[5], n: names[r[5]] || r[6], ro: r[7], ti: r[8], done: bySlot[r[0]] || null
  }));
  slots.forEach(x => { x.subj = subj[x.s + '|' + x.r] || []; });
  const recent = log.map(r => ({
    id: r[11], t: Utilities.formatDate(asDate_(r[0]) || new Date(), TZ, 'HH:mm'), d: r[1], sala: r[2], by: r[3], n: r[4], e: r[6], p: Number(r[8]) || 0, st: r[9]
  })).filter(x => me.all || x.by === me.name).slice(-40).reverse();
  const bets = me.org ? rows_('Scommesse').map(r => ({ id: r[0], q: r[1], type: r[2], result: String(r[4] || '') })) : [];
  const bingo = rows_('Bingo').filter(r => r[0]).map(r => ({ id: r[0], t: r[1], cat: r[2], on: r[3] === 'sì', by: r[5] || '' }));
  return safe_({ sala: sala, day: day, sessions: ses, slots: slots, recent: recent, bets: bets, bingo: bingo });
}

function arbRegistra(pin, req) {
  const me = ref_(pin);
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const ev = eventi_()[req.code];
    if (!ev) throw new Error('Evento sconosciuto');
    let day, sala, rid = req.r, slotKey = '';
    if (ev.cat === 'organizzatori') {
      if (!me.org) throw new Error('Premi e poster li inseriscono solo gli organizzatori');
      day = DAY_LABEL[todayIdx_()]; sala = '—';
      if (!people_()[rid]) throw new Error('Relatore non trovato');
    } else if (ev.cat === 'oggettivo') {
      const sl = rows_('Programma').find(r => r[0] === req.slot);
      if (!sl) throw new Error('Intervento non trovato');
      if (!me.all && sl[4] !== me.sala) throw new Error('Puoi registrare solo nella tua sala (' + me.sala + ')');
      if (ev.role && ev.role !== sl[7]) throw new Error('Evento non adatto a questo intervento');
      const dup = rows_('Log').find(r => r[9] === 'ok' && r[10] === req.slot);
      if (dup) throw new Error('Già registrato da ' + dup[3] + ': annulla prima di correggere');
      day = sl[2]; sala = sl[4]; rid = sl[5]; slotKey = sl[0];
    } else {
      const se = rows_('Sessioni').find(r => r[0] === req.sess);
      if (!se) throw new Error('Sessione non trovata');
      if (se[7] !== 'sì') throw new Error('I bonus soggettivi valgono solo nelle sessioni comuni');
      if (!me.all && se[4] !== me.sala) throw new Error('Puoi registrare solo nella tua sala (' + me.sala + ')');
      const inSess = rows_('Programma').some(r => r[1] === se[0] && r[5] === rid);
      if (!inSess) throw new Error('Questo relatore non è in questa sessione');
      day = se[1]; sala = se[4]; slotKey = se[0];
    }
    const name = (people_()[rid] || {}).n || rid;
    const id = Utilities.getUuid().slice(0, 8);
    sh_('Log').appendRow([new Date(), "'" + dayFix_(day), sala, me.name, name, rid, ev.code, ev.p, ev.p, 'ok', slotKey, id]);
    if (slotKey.startsWith('SL')) setSlotState_(slotKey, ev.code === 'ASS' ? 'assente' : 'confermato');
    recompute_(); bust_();
    const last = sh_('Log').getRange(sh_('Log').getLastRow(), 9).getValue();
    return { ok: true, msg: (last > 0 ? '+' : '') + last + ' · ' + ev.label + ' · ' + name + (last !== ev.p ? ' (tetto giornaliero)' : '') };
  } finally { lock.releaseLock(); }
}

function setSlotState_(slot, state) {
  const s = sh_('Programma'), v = s.getRange(2, 1, s.getLastRow() - 1, 1).getValues();
  const i = v.findIndex(r => r[0] === slot);
  if (i >= 0) s.getRange(i + 2, 10).setValue(state);
}

function annullaRow_(i, v) {
  const s = sh_('Log');
  s.getRange(i + 2, 10).setValue('annullato');
  if (String(v[i][10]).startsWith('SL')) setSlotState_(v[i][10], '');
  recompute_(); bust_();
  return { ok: true, msg: 'Annullato: ' + v[i][6] + ' · ' + v[i][4] };
}

function arbAnnullaUltimo(pin) {
  const me = ref_(pin);
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const v = rows_('Log');
    for (let i = v.length - 1; i >= 0; i--) if (v[i][3] === me.name && v[i][9] === 'ok') return annullaRow_(i, v);
    throw new Error('Niente da annullare');
  } finally { lock.releaseLock(); }
}

function arbAnnulla(pin, id) {
  const me = ref_(pin);
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const v = rows_('Log'), i = v.findIndex(r => r[11] === id);
    if (i < 0) throw new Error('Registrazione non trovata');
    if (!me.all && v[i][3] !== me.name) throw new Error('Puoi annullare solo le tue registrazioni');
    if (v[i][9] !== 'ok') throw new Error('Già annullata');
    return annullaRow_(i, v);
  } finally { lock.releaseLock(); }
}

function arbBingo(pin, id, on) {
  const me = ref_(pin);
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const s = sh_('Bingo'), v = rows_('Bingo'), i = v.findIndex(r => r[0] === id);
    if (i < 0) throw new Error('Frase non trovata');
    if (!on && !me.all && v[i][5] && v[i][5] !== me.name) throw new Error('Può toglierla solo ' + v[i][5] + ' o il responsabile');
    s.getRange(i + 2, 4, 1, 3).setValues([on ? ['sì', new Date(), me.name] : ['', '', '']]);
    recomputeBingo_(); bust_();
    return { ok: true, msg: (on ? 'Bingo, segnata: ' : 'Bingo, tolta: ') + v[i][1] };
  } finally { lock.releaseLock(); }
}

function arbRisultato(pin, betId, value) {
  const me = ref_(pin);
  if (!me.org) throw new Error('Solo gli organizzatori inseriscono i risultati');
  const s = sh_('Scommesse'), v = rows_('Scommesse'), i = v.findIndex(r => r[0] === betId);
  if (i < 0) throw new Error('Scommessa non trovata');
  s.getRange(i + 2, 5).setValue(clip_(value, 40));
  bust_();
  return { ok: true, msg: 'Risultato salvato: ' + v[i][1] + ' → ' + value };
}
