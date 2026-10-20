/**
 * Магазин Чорсу — сервер базы в Google Таблице.
 * Хранит все записи приложения на листах этой таблицы.
 *
 * Установка (один раз):
 *   1. Расширения → Apps Script → вставить этот код → Сохранить.
 *   2. Развернуть → Новое развертывание → Тип: Веб-приложение →
 *      Запуск от имени: «Я», Доступ: «Все» → Развернуть → разрешить доступ.
 *   3. Скопировать «URL веб-приложения» (…/exec) и вставить в WEB_APP_URL ниже → Сохранить.
 *   4. Выбрать функцию setup → Выполнить. В «Журнале выполнения» появится
 *      ссылка для подключения телефонов. Откройте её на каждом телефоне.
 */

// Адрес приложения на GitHub Pages
const APP_URL = 'https://pulatovs-s.github.io/chorsu/';
// Вставьте сюда «URL веб-приложения» из окна развертывания (заканчивается на /exec)
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

function doGet(e) { return handle_(e.parameter || {}); }
function doPost(e) {
  let body = {};
  try { body = JSON.parse(e.postData.contents); } catch (err) { return out_({ok: false, error: 'bad json'}); }
  return handle_(body);
}

function handle_(p) {
  const key = props_().getProperty('KEY');
  if (!key || p.key !== key) return out_({ok: false, error: 'key'});
  try {
    if (p.action === 'pull') return out_(pull_(Number(p.since == null ? -1 : p.since)));
    if (p.action === 'push') return out_(push_(p.ops || []));
    if (p.action === 'ping') return out_({ok: true, rev: Number(props_().getProperty('REV') || 0)});
    return out_({ok: false, error: 'action'});
  } catch (err) {
    return out_({ok: false, error: String(err && err.message || err)});
  }
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function norm_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return v;
}

function pull_(since) {
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
      if (rev > since) rows.push(r.map(norm_));
    }
    res[t] = rows;
  });
  return {ok: true, rev: Number(props_().getProperty('REV') || 0), tables: res};
}

function push_(ops) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = SpreadsheetApp.getActive(), p = props_();
    let rev = Number(p.getProperty('REV') || 0);
    // keep only the last change of each record
    const last = {};
    ops.forEach((o, i) => { last[o.t + '\u0001' + o.id] = i; });
    const byTable = {};
    ops.forEach((o, i) => { if (last[o.t + '\u0001' + o.id] === i) (byTable[o.t] = byTable[o.t] || []).push(o); });
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
