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
