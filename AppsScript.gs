/**
 * まなびステップ — Google Sheets 安全な回答記録
 * スプレッドシートの「拡張機能 > Apps Script」に貼り付ける。
 * setup() をエディタで1回実行して承認した後、ウェブアプリとしてデプロイ。
 * 連携キーは非公開シート「設定・使い方」!B12から読む。GitHubや公開HTMLへ転記しない。
 */
const LOG_TAB = '回答ログ';
const CONFIG_TAB = '設定・使い方';
const MAX_BATCH = 12;

function setup() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) throw new Error('必ず学習記録スプレッドシートから「拡張機能 > Apps Script」を開いて実行してください。');
  if (!spreadsheet.getSheetByName(LOG_TAB) || !spreadsheet.getSheetByName(CONFIG_TAB)) {
    throw new Error('必要なタブが見つかりません。');
  }
  const secret = String(spreadsheet.getSheetByName(CONFIG_TAB).getRange('B12').getValue()).trim();
  if (secret.length < 20) throw new Error('連携キーを設定してください。');
  PropertiesService.getScriptProperties().setProperty('MANABI_SPREADSHEET_ID', spreadsheet.getId());
  return '準備完了: このスプレッドシートのIDを保存しました。';
}

function sheetBook_() {
  const id = PropertiesService.getScriptProperties().getProperty('MANABI_SPREADSHEET_ID');
  if (!id) throw new Error('setup() を先に1回実行してください。');
  return SpreadsheetApp.openById(id);
}

function json_(payload) {
  return ContentService.createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  const callback = String(p.callback || '');
  // JSONP コールバックはサイトが生成する固定形式のみ許可する。
  if (!/^__manabi_ack_[a-z0-9]{1,30}$/.test(callback)) {
    return json_({ok: false});
  }
  let ok = false;
  if (p.action === 'status' && /^[a-zA-Z0-9_-]{8,180}$/.test(String(p.batch || ''))) {
    ok = CacheService.getScriptCache().get('ack:' + p.batch) === '1';
  }
  return ContentService.createTextOutput(callback + '(' + JSON.stringify({ok: ok}) + ');')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function safeText_(v, limit) {
  let t = String(v === undefined || v === null ? '' : v).slice(0, limit);
  // セルに入力された文字列から数式が実行されないようにする。
  if (/^[=+\-@\t\r]/.test(t)) t = "'" + t;
  return t;
}

function validEvent_(e) {
  return e && typeof e === 'object' &&
    /^[a-z0-9_-]{8,80}$/i.test(String(e.id || '')) &&
    Number.isInteger(Number(e.unit)) && Number(e.unit) >= 0 && Number(e.unit) <= 12 &&
    Number.isInteger(Number(e.number)) && Number(e.number) >= 0 && Number(e.number) <= 60 &&
    typeof e.question === 'string' && typeof e.answer === 'string' &&
    typeof e.selected === 'string' && typeof e.correct === 'boolean';
}

function doPost(e) {
  try {
    const raw = String((e && e.postData && e.postData.contents) || '');
    if (!raw || raw.length > 50000) return json_({ok: false});
    const data = JSON.parse(raw);
    const batch = String(data.batch || '');
    const events = data.events;
    if (!/^[a-zA-Z0-9_-]{8,180}$/.test(batch) ||
        !Array.isArray(events) || events.length < 1 || events.length > MAX_BATCH ||
        !events.every(validEvent_)) return json_({ok: false});

    const book = sheetBook_();
    const expectedKey = String(book.getSheetByName(CONFIG_TAB).getRange('B12').getValue()).trim();
    if (expectedKey.length < 20 || String(data.key || '') !== expectedKey) {
      return json_({ok: false});
    }

    const lock = LockService.getScriptLock();
    if (!lock.tryLock(20000)) return json_({ok: false});
    try {
      const target = book.getSheetByName(LOG_TAB);
      const lastRow = target.getLastRow();
      const existingIds = lastRow >= 2
        ? new Set(target.getRange(2, 2, lastRow - 1, 1).getValues().map(r => String(r[0])))
        : new Set();
      const rows = [];
      for (const ev of events) {
        if (existingIds.has(String(ev.id))) continue;
        const date = new Date(String(ev.time || ''));
        const jst = isNaN(date.getTime())
          ? Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss')
          : Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
        rows.push([
          jst,
          safeText_(ev.id, 80),
          safeText_(ev.session, 60),
          '学習者1',
          Number(ev.unit),
          safeText_(ev.unitTitle, 100),
          Number(ev.number),
          safeText_(ev.question, 250),
          safeText_(ev.selected, 80),
          safeText_(ev.answer, 80),
          ev.correct ? '正解' : '不正解',
          Math.max(0, Math.min(3600, Number(ev.seconds) || 0)),
          safeText_(ev.version, 30)
        ]);
        existingIds.add(String(ev.id));
      }
      if (rows.length) target.getRange(lastRow + 1, 1, rows.length, 13).setValues(rows);
      SpreadsheetApp.flush();
      // 受信確認の問い合わせのみ許可。記録本体は公開しない。
      CacheService.getScriptCache().put('ack:' + batch, '1', 3600);
      return json_({ok: true, appended: rows.length});
    } finally {
      lock.releaseLock();
    }
  } catch (error) {
    console.error('まなびステップの受信処理エラー', error);
    return json_({ok: false});
  }
}
