/**
 * FantaSIN – backend su Google Fogli (Apps Script)
 * 1. Cambia REF_CODE qui sotto con un codice segreto per gli arbitri.
 * 2. Distribuisci → Nuova distribuzione → App web, Esegui come: Me, Accesso: Chiunque.
 */
const REF_CODE = 'CAMBIAMI-arbitri';   // <-- codice arbitri (non lo vede nessun giocatore)
const TEAM_SIZE = 5;

const SHEETS = {
  teams:  ['pid', 'token', 'player', 'name', 'picks', 'cap', 'ts'],
  events: ['id', 'a', 'b', 'p', 'w', 'note', 'byName', 'ts'],
  extra:  ['id', 'n', 'c', 'ts'],
  config: ['key', 'value']
};

function sheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); sh.appendRow(SHEETS[name]); sh.setFrozenRows(1); }
  return sh;
}
function rows_(name) {
  const v = sheet_(name).getDataRange().getValues();
  const head = v.shift();
  return v.filter(r => r[0] !== '').map(r => Object.fromEntries(head.map((h, i) => [h, r[i]])));
}
function findRow_(name, key) {
  const v = sheet_(name).getRange(1, 1, sheet_(name).getLastRow(), 1).getValues();
  for (let i = 1; i < v.length; i++) if (String(v[i][0]) === String(key)) return i + 1;
  return -1;
}
function upsert_(name, obj) {
  const sh = sheet_(name), row = SHEETS[name].map(h => obj[h] === undefined ? '' : obj[h]);
  const r = findRow_(name, row[0]);
  if (r > 0) sh.getRange(r, 1, 1, row.length).setValues([row]); else sh.appendRow(row);
}
function remove_(name, key) { const r = findRow_(name, key); if (r > 0) sheet_(name).deleteRow(r); }
function pid_(token) {
  const d = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(token), Utilities.Charset.UTF_8);
  return 'p' + d.map(b => ('0' + (b & 255).toString(16)).slice(-2)).join('').slice(0, 16);
}
function locked_() { return rows_('config').some(r => r.key === 'locked' && String(r.value) === 'true'); }
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function str_(x, n) { const v = String(x == null ? '' : x).slice(0, n); return /^[=+\-@]/.test(v) ? "'" + v : v; }

function state_() {
  return {
    ok: true,
    teams: rows_('teams').map(t => ({ id: t.pid, player: t.player, name: t.name, picks: String(t.picks).split(',').filter(Boolean), cap: t.cap, ts: t.ts })),
    events: rows_('events').map(e => ({ id: String(e.id), a: e.a, b: e.b, p: Number(e.p), w: e.w, note: e.note, byName: e.byName, ts: Number(e.ts) })),
    extra: rows_('extra').map(x => ({ id: x.id, n: x.n, c: x.c })),
    config: { locked: locked_() }
  };
}

function doGet(e) {
  return out_(state_());
}

function doPost(e) {
  let q;
  try { q = JSON.parse(e.postData.contents); } catch (err) { return out_({ ok: false, error: 'bad_request' }); }
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const isRef = q.code === REF_CODE;
    switch (q.action) {
      case 'checkRef':
        return out_({ ok: isRef, error: isRef ? null : 'denied' });

      case 'saveTeam': {
        if (!q.token || String(q.token).length < 16) return out_({ ok: false, error: 'denied' });
        if (locked_()) return out_({ ok: false, error: 'locked' });
        const picks = (q.picks || []).map(String);
        if (picks.length !== TEAM_SIZE || new Set(picks).size !== TEAM_SIZE || picks.indexOf(String(q.cap)) < 0)
          return out_({ ok: false, error: 'invalid' });
        upsert_('teams', { pid: pid_(q.token), token: q.token, player: str_(q.player, 40), name: str_(q.name, 40),
          picks: picks.join(','), cap: String(q.cap), ts: Date.now() });
        break;
      }
      case 'addEvent': {
        if (!isRef) return out_({ ok: false, error: 'denied' });
        const p = Math.round(Number(q.p));
        if (!isFinite(p) || Math.abs(p) > 100 || !q.a) return out_({ ok: false, error: 'invalid' });
        upsert_('events', { id: Utilities.getUuid(), a: str_(q.a, 80), b: str_(q.b, 20), p: p, w: str_(q.w, 60),
          note: str_(q.note, 80), byName: str_(q.byName, 40), ts: Date.now() });
        break;
      }
      case 'delEvent':
        if (!isRef) return out_({ ok: false, error: 'denied' });
        remove_('events', q.id);
        break;
      case 'setLock':
        if (!isRef) return out_({ ok: false, error: 'denied' });
        upsert_('config', { key: 'locked', value: q.locked ? 'true' : 'false' });
        break;
      case 'addExtra':
        if (!isRef) return out_({ ok: false, error: 'denied' });
        upsert_('extra', { id: str_(q.id, 80), n: str_(q.n, 40), c: str_(q.c, 30), ts: Date.now() });
        break;
      default:
        return out_({ ok: false, error: 'bad_request' });
    }
    return out_(state_());
  } finally {
    lock.releaseLock();
  }
}
