// Beads owner trial: receives the sign-up form and appends one row to this sheet.
// Deploy: Apps Script > Deploy > New deployment > Web app, execute as me, access: anyone.
const SHEET_ID = '1-CnHJFHW8wQhL7UFSJVa1jAWfWyvMg2sJwSO9XEBlBA';
const HEAD = ['받은 시각', '가게', '메뉴 수', '휴대폰', '동의'];

// Leading = + - @ would run as a formula in Sheets; force text.
const cell = (v, n) => { const s = String(v || '').slice(0, n); return /^[=+\-@]/.test(s) ? "'" + s : s; };

function doPost(e) {
  const p = e.parameter || {};
  const tel = String(p.tel || '').replace(/\D/g, '');
  const ok = p.agree === 'yes' && /^0\d{8,10}$/.test(tel) && String(p.store || '').trim() && /^([1-5]|6\+)$/.test(p.menus || '');
  if (ok) {
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      const sh = SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
      if (sh.getLastRow() === 0) sh.appendRow(HEAD);
      sh.appendRow([Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss'), cell(p.store.trim(), 60), "'" + p.menus, "'" + cell(p.tel, 13), '동의']);
    } finally { lock.releaseLock(); }
  }
  return ContentService.createTextOutput(JSON.stringify({ ok: !!ok })).setMimeType(ContentService.MimeType.JSON);
}
