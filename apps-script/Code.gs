/**
 * Магазин Чорсу — сервер базы в Google Таблице.
 * Хранит все записи приложения на листах этой таблицы.
 *
 * Установка (один раз):
 *   1. Расширения → Apps Script → стереть всё → вставить этот код → Сохранить.
 *   2. Начать развертывание → Новое развертывание → Веб-приложение →
 *      Запуск от имени: «Я», Доступ: «Все» → Начать развертывание → разрешить доступ.
 *   3. Нажать на «URL веб-приложения». Откроется страница с QR-кодом и ссылкой
 *      для телефонов (её видит только владелец таблицы).
 */

// Адрес приложения на GitHub Pages
const APP_URL = 'https://pulatovs-s.github.io/chorsu/';
// Не обязательно: адрес …/exec, если страница подключения не открывается сама
const WEB_APP_URL = '';

const TABLES = {
  products: 'Товары', prices: 'Цены', docs: 'Приходы', prihod: 'Приход_строки',
  inv: 'Инвентаризации', invlines: 'Инвент_строки', extra: 'Вне_списка', cash: 'Касса',
  adj: 'Корректировки', debtors: 'Должники', moves: 'Долги', settings: 'Настройки'
};
// Колонки, которые всегда хранятся как текст (id, даты, названия) — чтобы Таблица не превращала их в даты/числа
const TEXT_COLS = {
  products: [0, 2, 3, 4], prices: [0, 1, 2, 3], docs: [0, 1, 2], prihod: [0, 1, 2, 3, 4, 5],
  inv: [0, 1, 2, 3], invlines: [0, 1, 2, 3, 4], extra: [0, 1, 2], cash: [0, 1], adj: [0, 1, 3],
  debtors: [0, 1, 2], moves: [0, 1, 2, 3, 5], settings: [0, 1]
};

function props_() { return PropertiesService.getScriptProperties(); }

/** Запустить один раз после развертывания: создаёт ключ и выдаёт ссылку подключения. */
function setup() {
  const p = props_();
  let key = p.getProperty('KEY');
  if (!key) { key = Utilities.getUuid().replace(/-/g, ''); p.setProperty('KEY', key); }
  if (!p.getProperty('REV')) p.setProperty('REV', '0');
  const ss = SpreadsheetApp.getActive();
  Object.keys(TABLES).forEach(t => {
    const sh = ss.getSheetByName(TABLES[t]);
    if (!sh) throw new Error('Нет листа «' + TABLES[t] + '». Таблица должна быть создана из файла «Chorsu база.xlsx».');
    const width = sh.getLastColumn();
    if (sh.getMaxRows() < 2) sh.insertRowsAfter(1, 200);
    TEXT_COLS[t].concat([width - 1]).forEach(c => sh.getRange(2, c + 1, sh.getMaxRows() - 1, 1).setNumberFormat('@'));
  });
  const url = WEB_APP_URL || ScriptApp.getService().getUrl() || '';
  if (!/\/exec$/.test(url)) throw new Error('Вставьте в WEB_APP_URL «URL веб-приложения» из окна развертывания (он заканчивается на /exec) и снова запустите setup.');
  const link = APP_URL + '#connect=' + Utilities.base64EncodeWebSafe(JSON.stringify({u: url, k: key}));
  Logger.log('Ссылка для подключения телефонов (откройте её на каждом телефоне):\n' + link);
  return link;
}

/** Если ссылку надо отключить (например, потеряли телефон): новый ключ, старые ссылки перестанут работать. */
function resetKey() { props_().deleteProperty('KEY'); return setup(); }

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (!p.key && !p.action) return ownerPage_();
  return handle_(p);
}

/** Страница подключения: показывает ссылку и QR-код только владельцу таблицы. */
function ownerPage_() {
  let me = '', owner = '';
  try { me = Session.getActiveUser().getEmail(); owner = Session.getEffectiveUser().getEmail(); } catch (err) {}
  if (!me || me !== owner) {
    return HtmlService.createHtmlOutput('<div style="font:18px/1.45 sans-serif;padding:20px;max-width:560px">' +
      '<b>Магазин Чорсу.</b> Приложение открывается по личной ссылке подключения.<br><br>' +
      '<span style="color:#56645D">Если вы владелец таблицы и не видите здесь QR-код: в редакторе Apps Script выберите вверху функцию <b>setup</b> и нажмите «Выполнить» — ссылка появится внизу, в «Журнале выполнения».</span></div>');
  }
  const link = setup();
  const qr = 'https://api.qrserver.com/v1/create-qr-code/?size=320x320&margin=10&data=' + encodeURIComponent(link);
  const html = '<!doctype html><meta name="viewport" content="width=device-width,initial-scale=1">' +
    '<div style="font:18px/1.45 -apple-system,Segoe UI,Roboto,sans-serif;max-width:560px;margin:0 auto;padding:20px;color:#16201B">' +
    '<h1 style="font-size:26px;margin:0 0 6px">Готово! База подключена</h1>' +
    '<p style="margin:0 0 18px;color:#56645D">Таблица «' + SpreadsheetApp.getActive().getName() + '» работает как база приложения.</p>' +
    '<h2 style="font-size:20px;margin:0 0 8px">Как подключить телефон</h2>' +
    '<p style="margin:0 0 10px"><b>Способ 1.</b> Наведите камеру телефона на этот QR-код и откройте ссылку.</p>' +
    '<img src="' + qr + '" width="320" height="320" alt="QR-код" style="display:block;border:1px solid #D3DCD5;border-radius:12px;margin:0 0 16px">' +
    '<p style="margin:0 0 8px"><b>Способ 2.</b> Скопируйте ссылку и отправьте себе и папе в Telegram (лично, не в группу):</p>' +
    '<textarea readonly onclick="this.select()" style="width:100%;height:110px;font:14px monospace;padding:10px;border:2px solid #1D5C45;border-radius:10px;box-sizing:border-box">' + link + '</textarea>' +
    '<button onclick="var t=document.querySelector(\'textarea\');t.select();document.execCommand(\'copy\');this.textContent=\'Скопировано ✓\'" style="margin:10px 0 18px;font:bold 18px sans-serif;background:#1D5C45;color:#fff;border:0;border-radius:12px;padding:14px 20px;width:100%">Скопировать ссылку</button>' +
    '<p style="margin:0 0 6px"><b>Потом поставьте иконку на экран:</b></p>' +
    '<p style="margin:0">iPhone — в Safari кнопка «Поделиться» → «На экран Домой».<br>Android — в Chrome ⋮ → «Добавить на главный экран».</p>' +
    '<p style="margin:18px 0 0;color:#8A5A00;background:#FFF1D2;padding:12px;border-radius:10px">Эта ссылка — ключ к вашей базе. Не пересылайте её посторонним.</p></div>';
  return HtmlService.createHtmlOutput(html).setTitle('Магазин Чорсу — подключение').addMetaTag('viewport', 'width=device-width, initial-scale=1');
}
function doPost(e) {
  let body = {};
  try { body = JSON.parse(e.postData.contents); } catch (err) { return out_({ok: false, error: 'bad json'}); }
  return handle_(body);
}

function handle_(p) {
  const key = props_().getProperty('KEY');
  if (!key || p.key !== key) return out_({ok: false, error: 'key'});
  try {
    const U = users_(), has = Object.keys(U).length > 0;
    let me = null;
    if (p.token) { const l = props_().getProperty('T.' + p.token); if (l && U[l] && U[l].active) me = U[l]; }
    if (p.action === 'login') return out_(login_(p, U));
    if (p.action === 'ping') return out_({ok: true, rev: Number(props_().getProperty('REV') || 0), users: has, user: me && pub_(me)});
    if (has && !me) return out_({ok: false, error: 'login'});
    const role = me ? me.role : 'админ';            // пока пользователей нет — всё как раньше
    if (p.action === 'pull') return out_(pull_(Number(p.since == null ? -1 : p.since), role));
    if (p.action === 'push') return out_(push_(p.ops || [], role));
    if (p.action === 'report') return out_(report_(p.sheets || []));
    if (p.action === 'logout') { props_().deleteProperty('T.' + p.token); return out_({ok: true}); }
    if (p.action === 'passwd') return out_(passwd_(me, p));
    if (p.action === 'users') { if (role !== 'админ') return out_({ok: false, error: 'role'}); return out_({ok: true, users: Object.keys(U).map(l => pub_(U[l]))}); }
    if (p.action === 'usersave') { if (role !== 'админ') return out_({ok: false, error: 'role'}); return out_(userSave_(p.user || {}, U, !has)); }
    return out_({ok: false, error: 'action'});
  } catch (err) {
    return out_({ok: false, error: String(err && err.message || err)});
  }
}

/* ===================== Пользователи =====================
 * Лист «Пользователи»: логин, имя, роль (админ / управляющий / продавец), соль, хэш пароля, активен.
 * Пароли не хранятся — только хэш. Вход выдаёт личный токен телефона; «Активен: нет» — вход и токены перестают работать.
 */
const USERS_SHEET = 'Пользователи', ROLES = ['админ', 'управляющий', 'продавец'];
const SELLER_TABLES = {docs: 1, prihod: 1, inv: 1, invlines: 1, extra: 1, cash: 1, debtors: 1, moves: 1};
function usersSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(USERS_SHEET);
  if (!sh) {
    sh = ss.insertSheet(USERS_SHEET, ss.getNumSheets());
    sh.getRange(1, 1, 1, 6).setValues([['Логин', 'Имя', 'Роль', 'соль', 'хэш пароля', 'Активен']]).setFontWeight('bold');
    sh.setFrozenRows(1); sh.getRange(2, 1, sh.getMaxRows() - 1, 6).setNumberFormat('@');
  }
  return sh;
}
function users_() {
  const c = CacheService.getScriptCache(), v = c.get('USERS');
  if (v) return JSON.parse(v);
  const sh = SpreadsheetApp.getActive().getSheetByName(USERS_SHEET), map = {};
  if (sh && sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 6).getValues().forEach(r => {
    const l = String(r[0]).trim().toLowerCase();
    if (l) map[l] = {login: l, name: String(r[1]), role: ROLES.indexOf(String(r[2])) >= 0 ? String(r[2]) : 'продавец', salt: String(r[3]), hash: String(r[4]), active: String(r[5]).trim().toLowerCase() !== 'нет'};
  });
  c.put('USERS', JSON.stringify(map), 300);
  return map;
}
const pub_ = u => ({login: u.login, name: u.name, role: u.role, active: u.active});
function hash_(salt, pw) { return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + '|' + pw, Utilities.Charset.UTF_8)); }
function login_(p, U) {
  const u = U[String(p.login || '').trim().toLowerCase()];
  if (!u || !u.active || hash_(u.salt, String(p.password || '')) !== u.hash) { Utilities.sleep(1200); return {ok: false, error: 'badlogin'}; }
  const token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '').slice(0, 8);
  props_().setProperty('T.' + token, u.login);
  return {ok: true, token, user: pub_(u)};
}
function writeUser_(u) {
  const sh = usersSheet_(), n = Math.max(sh.getLastRow() - 1, 0);
  const vals = n ? sh.getRange(2, 1, n, 1).getValues() : [];
  let row = vals.findIndex(r => String(r[0]).trim().toLowerCase() === u.login) + 2;
  if (row < 2) row = n + 2;
  sh.getRange(row, 1, 1, 6).setNumberFormat('@').setValues([[u.login, u.name, u.role, u.salt, u.hash, u.active ? 'да' : 'нет']]);
  CacheService.getScriptCache().remove('USERS');
}
function userSave_(x, U, first) {
  const login = String(x.login || '').trim().toLowerCase();
  if (!/^[a-zа-яё0-9._-]{2,30}$/.test(login)) return {ok: false, error: 'Логин: 2–30 букв или цифр, без пробелов'};
  const old = U[login], pw = String(x.password || '');
  if (!old && pw.length < 4) return {ok: false, error: 'Пароль — не короче 4 знаков'};
  if (pw && pw.length < 4) return {ok: false, error: 'Пароль — не короче 4 знаков'};
  const role = first ? 'админ' : (ROLES.indexOf(x.role) >= 0 ? x.role : 'продавец');
  const u = {login, name: String(x.name || login).trim(), role, salt: old ? old.salt : Utilities.getUuid(), hash: old ? old.hash : '', active: x.active !== false};
  if (pw) { u.salt = Utilities.getUuid(); u.hash = hash_(u.salt, pw); }
  writeUser_(u);
  const res = {ok: true, user: pub_(u)};
  if (first) { const t = login_({login, password: pw}, users_()); res.token = t.token; }
  return res;
}
function passwd_(me, p) {
  if (!me) return {ok: false, error: 'login'};
  if (hash_(me.salt, String(p.old || '')) !== me.hash) { Utilities.sleep(800); return {ok: false, error: 'Старый пароль неверный'}; }
  const pw = String(p.password || ''); if (pw.length < 4) return {ok: false, error: 'Пароль — не короче 4 знаков'};
  me.salt = Utilities.getUuid(); me.hash = hash_(me.salt, pw); writeUser_(me);
  return {ok: true};
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function norm_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return v;
}

function pull_(since, role) {
  const ss = SpreadsheetApp.getActive(), res = {};
  Object.keys(TABLES).forEach(t => {
    const sh = ss.getSheetByName(TABLES[t]);
    const n = sh.getLastRow() - 1;
    if (n < 1) { res[t] = []; return; }
    const w = sh.getLastColumn(), vals = sh.getRange(2, 1, n, w).getValues();
    const rows = [];
    for (let i = 0; i < vals.length; i++) {
      const r = vals[i];
      if (r[0] === '' || r[0] === null) continue;
      const rev = Number(r[w - 2]) || 0;
      if (rev > since && !(role === 'продавец' && t === 'settings' && String(r[0]) === 'pin')) rows.push(r.map(norm_));
    }
    res[t] = rows;
  });
  return {ok: true, rev: Number(props_().getProperty('REV') || 0), tables: res};
}

function push_(ops, role) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = SpreadsheetApp.getActive(), p = props_();
    let rev = Number(p.getProperty('REV') || 0);
    // keep only the last change of each record
    const last = {};
    ops.forEach((o, i) => { last[o.t + '\u0001' + o.id] = i; });
    const byTable = {};
    ops.forEach((o, i) => { if (last[o.t + '\u0001' + o.id] === i && (role !== 'продавец' || SELLER_TABLES[o.t])) (byTable[o.t] = byTable[o.t] || []).push(o); });
    Object.keys(byTable).forEach(t => {
      if (!TABLES[t]) return;
      const sh = ss.getSheetByName(TABLES[t]);
      const w = sh.getLastColumn(), n = Math.max(sh.getLastRow() - 1, 0);
      const vals = n ? sh.getRange(2, 1, n, w).getValues() : [];
      const idx = {};
      vals.forEach((r, i) => { if (r[0] !== '') idx[String(r[0])] = i; });
      let lo = Infinity, hi = -1; const appends = [];
      byTable[t].forEach(o => {
        rev++;
        const i = idx[String(o.id)];
        if (o.del) {
          if (i == null) return;
          vals[i][w - 2] = rev; vals[i][w - 1] = 'x';
          lo = Math.min(lo, i); hi = Math.max(hi, i);
          return;
        }
        const row = (o.row || []).slice(0, w - 2);
        while (row.length < w - 2) row.push('');
        row.push(rev, '');
        if (i == null) { idx[String(o.id)] = vals.length + appends.length; appends.push(row); }
        else { vals[i] = row; lo = Math.min(lo, i); hi = Math.max(hi, i); }
      });
      if (hi >= 0) sh.getRange(2 + lo, 1, hi - lo + 1, w).setValues(vals.slice(lo, hi + 1));
      if (appends.length) {
        const start = n + 2;
        if (sh.getMaxRows() < start + appends.length) sh.insertRowsAfter(sh.getMaxRows(), start + appends.length - sh.getMaxRows() + 100);
        const rg = sh.getRange(start, 1, appends.length, w);
        TEXT_COLS[t].concat([w - 1]).forEach(c => sh.getRange(start, c + 1, appends.length, 1).setNumberFormat('@'));
        rg.setValues(appends);
      }
    });
    p.setProperty('REV', String(rev));
    return {ok: true, rev};
  } finally {
    lock.releaseLock();
  }
}

/* ===================== Отчёты =====================
 * Приложение считает отчёты само и присылает готовые таблицы.
 * Здесь они записываются на листы и красиво оформляются.
 * Листы «📊 …» каждый раз перезаписываются; листы для офиса («Qarz18.07», «18.07»)
 * создаются по одному на неделю, как в отчёте офиса.
 */
const TPL_SHEET = 'Шаблон офиса';
const GREEN = '#1D5C45', LINE = '#D3DCD5', TOT_BG = '#EEF3EF';

function report_(sheets) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = SpreadsheetApp.getActive();
    const done = [];
    // сначала листы «📊» (впереди, по порядку), потом недельные для офиса — сразу за ними, новые первыми
    sheets.filter(s => s.pos !== 'week').forEach((s, i) => { writeSheet_(ss, s, i); done.push(s.name); });
    const dash = ss.getSheets().filter(x => x.getName().indexOf('📊') === 0).length;
    sheets.filter(s => s.pos === 'week').reverse().forEach(s => {
      if (s.kind === 'office') s = officeSheet_(ss, s);
      writeSheet_(ss, s, dash); done.push(s.name);
    });
    SpreadsheetApp.flush();
    return {ok: true, sheets: done};
  } finally {
    lock.releaseLock();
  }
}

/** 'yyyy-mm-dd' → номер дня Google Таблицы (не зависит от языка и часового пояса таблицы) */
function serial_(iso) { return (Date.UTC(+iso.slice(0, 4), +iso.slice(5, 7) - 1, +iso.slice(8, 10)) - Date.UTC(1899, 11, 30)) / 86400000; }

function writeSheet_(ss, s, pos) {
  let sh = ss.getSheetByName(s.name);
  if (!sh) sh = ss.insertSheet(s.name, Math.min(Number(pos) || 0, ss.getNumSheets()));
  else {
    const f = sh.getFilter(); if (f) f.remove();
    sh.setFrozenRows(0); sh.setFrozenColumns(0);
    sh.clear(); sh.clearNotes();
  }
  const rows = s.rows || [[]];
  const w = Math.max(1, ...rows.map(r => r.length));
  const data = rows.map(r => { const x = r.map(v => v == null ? '' : v); while (x.length < w) x.push(''); return x; });
  const n = data.length;
  if (sh.getMaxRows() < n + 1) sh.insertRowsAfter(sh.getMaxRows(), n + 1 - sh.getMaxRows());
  if (sh.getMaxColumns() < w) sh.insertColumnsAfter(sh.getMaxColumns(), w - sh.getMaxColumns());
  (s.text || []).forEach(c => sh.getRange(1, c + 1, n, 1).setNumberFormat('@'));
  sh.getRange(1, 1, n, w).setValues(data).setFontFamily('Arial').setFontSize(10).setVerticalAlignment('middle');
  const head = s.head || 0, body = n - head;
  if (body > 0) {
    (s.num || []).forEach(c => sh.getRange(head + 1, c + 1, body, 1).setNumberFormat('#,##0;[Red]-#,##0;""'));
    (s.dates || []).forEach(c => sh.getRange(head + 1, c + 1, body, 1).setNumberFormat('dd.MM.yyyy').setHorizontalAlignment('left'));
    sh.getRange(head + 1, 1, body, w).setBorder(true, true, true, true, true, true, LINE, SpreadsheetApp.BorderStyle.SOLID);
  }
  if (head) {
    const hr = sh.getRange(head, 1, 1, w);
    hr.setFontWeight('bold').setBackground(s.headBg || GREEN).setFontColor(s.headBg ? '#000000' : '#FFFFFF')
      .setVerticalAlignment('middle').setWrap(!!s.wrapHead).setBorder(true, true, true, true, true, true, s.headBg ? '#999999' : GREEN, SpreadsheetApp.BorderStyle.SOLID);
    if (s.wrapHead) sh.setRowHeight(head, 46);
    if (head > 1) sh.getRange(1, 1, head - 1, w).setFontWeight('bold');
  }
  (s.tot || []).forEach(r => sh.getRange(r, 1, 1, w).setFontWeight('bold').setBackground(TOT_BG));
  (s.marks || []).forEach(m => { const rg = sh.getRange(m.r, m.c || 1, 1, m.w || w); if (m.bold) rg.setFontWeight('bold'); if (m.bg) rg.setBackground(m.bg); if (m.color) rg.setFontColor(m.color); });
  if (s.freezeRows) sh.setFrozenRows(s.freezeRows);
  if (s.freezeCols) sh.setFrozenColumns(s.freezeCols);
  for (let c = 0; c < w; c++) sh.setColumnWidth(c + 1, (s.widths && s.widths[c]) || 104);
  if (s.filter && head && body > 0) sh.getRange(head, 1, body + 1, w).createFilter();
  (s.cells || []).forEach(c => { const rg = sh.getRange(c.a1); if (c.f) rg.setNumberFormat(c.f); if (c.note) rg.setNote(c.note); if (c.bold) rg.setFontWeight('bold'); });
  if (s.name.indexOf('📊') === 0) sh.setTabColor(GREEN);
  return sh;
}

/* «Отчёт для офиса»: строки в порядке и с названиями из листа «Шаблон офиса». */
const TRL_ = {'а':'a','б':'b','в':'v','г':'g','д':'d','е':'e','ё':'e','ж':'j','з':'z','и':'i','й':'y','к':'k','л':'l','м':'m','н':'n','о':'o','п':'p','р':'r','с':'s','т':'t','у':'u','ф':'f','х':'x','ц':'c','ч':'ch','ш':'sh','щ':'sh','ъ':'','ы':'i','ь':'','э':'e','ю':'yu','я':'ya'};
function key_(s) {
  s = String(s || '').toLowerCase().replace(/[а-яё]/g, ch => TRL_[ch]).replace(/[^0-9a-z]/g, '');
  s = s.replace(/(\d)(o+)/g, (m, d, o) => d + '0'.repeat(o.length));
  return s.replace(/superslims?/g, 'ss').replace(/dona/g, 'sht').replace(/^sigar/, 'cigar').replace(/compact/g, 'com').replace(/navi/g, 'navy');
}
function lev_(a, b) {
  const m = a.length, n = b.length; if (!m) return n; if (!n) return m;
  let p = []; for (let j = 0; j <= n; j++) p[j] = j;
  for (let i = 1; i <= m; i++) { const c = [i]; for (let j = 1; j <= n; j++) c[j] = Math.min(p[j] + 1, c[j - 1] + 1, p[j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1)); p = c; }
  return p[n];
}
/** tpl: [[офис, название в приложении]], names: названия в приложении → [{j, how}] */
function matchOffice_(tpl, names) {
  const keys = names.map(key_), used = {}, res = tpl.map(() => ({j: -1, how: ''}));
  const take = (k) => { for (let j = 0; j < keys.length; j++) if (!used[j] && keys[j] === k) { used[j] = 1; return j; } return -1; };
  tpl.forEach((t, i) => { const b = String(t[1] || '').trim(); if (b) { const j = take(key_(b)); res[i] = {j, how: j >= 0 ? 'по столбцу B' : 'B: нет в приложении'}; } });
  tpl.forEach((t, i) => { if (res[i].how) return; const j = take(key_(t[0])); if (j >= 0) res[i] = {j, how: 'точно'}; });
  tpl.forEach((t, i) => {
    if (res[i].how) return;
    const k = key_(t[0]); let best = -1, bd = 1e9, second = 1e9;
    keys.forEach((x, j) => { if (used[j]) return; const d = lev_(k, x); if (d < bd) { second = bd; bd = d; best = j; } else if (d < second) second = d; });
    if (best >= 0 && bd <= Math.max(1, Math.floor(Math.min(k.length, keys[best].length) * 0.15)) && second > bd) { used[best] = 1; res[i] = {j: best, how: 'похоже — проверьте'}; }
    else res[i] = {j: -1, how: 'нет в приложении'};
  });
  return res;
}
function officeSheet_(ss, s) {
  const items = s.items || [], names = items.map(x => x[0]);
  const tsh = ss.getSheetByName(TPL_SHEET);
  let shop = '', order = null, map = null;
  if (tsh && tsh.getLastRow() > 1) {
    shop = String(tsh.getRange('F1').getValue() || '');
    const v = tsh.getRange(2, 1, tsh.getLastRow() - 1, 2).getValues();
    const last = v.map(r => String(r[0]).trim() !== '').lastIndexOf(true);
    order = v.slice(0, last + 1);
    map = matchOffice_(order, names);
    const col = map.map((m, i) => [String(order[i][0]).trim() === '' ? '' : (m.j >= 0 && m.how !== 'точно' && m.how !== 'по столбцу B' ? m.how + ': ' + names[m.j] : m.how)]);
    if (col.length) {
      const rg = tsh.getRange(2, 3, col.length, 1);
      rg.setValues(col).setBackgrounds(col.map(c => [/^похоже/.test(c[0]) ? '#FFF1D2' : /^нет|^B:/.test(c[0]) ? '#F1F1F1' : null]));
    }
  }
  const rows = [['', "Do'kon nomi:", 'Sana:'], ['', shop, s.date ? serial_(s.date) : ''], ['Номи', 'ККМ', 'Омбор']];
  const usedItem = {};
  if (order) order.forEach((t, i) => { const j = map[i].j; if (j >= 0) usedItem[j] = 1; const it = j >= 0 ? items[j] : null; rows.push([String(t[0]).trim(), it && it[1] ? it[1] : '', it && it[2] ? it[2] : '']); });
  else items.forEach((it, j) => { usedItem[j] = 1; rows.push([it[0], it[1] || '', it[2] || '']); });
  const lastTpl = rows.length;
  const extra = items.filter((it, j) => !usedItem[j] && (it[1] || it[2]));
  const marks = [];
  if (extra.length) {
    rows.push(['', '', '']);
    rows.push(['Есть остаток, но нет в шаблоне офиса — допишите вручную:', '', '']);
    marks.push({r: rows.length, bold: true, bg: '#FFF1D2'});
    extra.forEach(it => rows.push([it[0], it[1] || '', it[2] || '']));
  }
  let tk = 0, to = 0; items.forEach(it => { tk += it[1] || 0; to += it[2] || 0; });
  rows.push(['', '', '']);
  rows.push(['Всего пачек (для проверки, не копировать)', tk, to]);
  marks.push({r: rows.length, bold: true, bg: TOT_BG});
  return {name: s.name, pos: 'week', rows, head: 3, freezeRows: 3, num: [1, 2], dates: [], text: [0], widths: [290, 80, 110], headBg: '#D8D8D8', marks,
    cells: [{a1: 'C2', f: 'dd.MM.yyyy'}, {a1: 'B2', bold: true}, {a1: 'A3', note: 'Для отчёта офиса скопируйте B4:C' + lastTpl + ' и вставьте в B4 листа «' + s.name + '» офиса (Правка → Специальная вставка → Только значения).'}]};
}
