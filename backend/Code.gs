/**
 * 見面邀請卡後台（Google Apps Script）
 * 功能：收朋友的回覆 → 寫進試算表 → 寄通知給你 → 後台排日曆 → 一鍵寄確認信
 * 設定步驟請看同資料夾的 README.md
 */

// ★ 1. 把這串改成你自己的密碼（後台網址要帶這串才打得開，請不要太簡單）
const ADMIN_KEY = '請改成你自己的密碼';

const SHEET_REPLY = '回覆';
const SHEET_ROSTER = '名單';
const HEAD = ['ID', '時間', '暱稱', '關係', '城市', '可約時段', '見面方式', '想邀請', 'Email', 'IG',
  '角色', '猜中', '訊息', 'slots', '狀態', '確認日期', '確認時間', 'EventId', '確認城市'];
const COL = {}; HEAD.forEach((h, i) => COL[h] = i);
const TZ = 'Asia/Kuala_Lumpur';
const MAX_ROWS = 1000;

/* ---------- 工具 ---------- */
function sheet_(name, head) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(name);
  if (!s) {
    s = ss.insertSheet(name);
    if (head) { s.appendRow(head); s.setFrozenRows(1); }
  }
  return s;
}
function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
// 避免試算表把「=開頭」的文字當成公式執行
function safe_(v, max) {
  let t = String(v == null ? '' : v).replace(/\r/g, '').trim().slice(0, max || 200);
  if (/^[=+\-@]/.test(t)) t = "'" + t;
  return t;
}
function auth_(key) {
  if (key !== ADMIN_KEY || ADMIN_KEY === '請改成你自己的密碼') throw new Error('沒有權限');
}
function adminUrl_() {
  return ScriptApp.getService().getUrl() + '?key=' + encodeURIComponent(ADMIN_KEY);
}
function norm_(n) { return String(n || '').replace(/\s+/g, '').toLowerCase(); }
function isDate_(v) { return Object.prototype.toString.call(v) === '[object Date]'; }
function fmtDate_(v) {
  return isDate_(v) ? Utilities.formatDate(v, TZ, 'yyyy-MM-dd') : String(v || '');
}
function rowObj_(r, rowNo) {
  let slots = [];
  try { slots = JSON.parse(r[COL.slots] || '[]'); } catch (e) {}
  return {
    row: rowNo, id: String(r[COL.ID]),
    time: isDate_(r[COL['時間']]) ? Utilities.formatDate(r[COL['時間']], TZ, 'M/d HH:mm') : String(r[COL['時間']]),
    name: String(r[COL['暱稱']]), relation: String(r[COL['關係']]), city: String(r[COL['城市']]),
    dates: String(r[COL['可約時段']]), format: String(r[COL['見面方式']]), invitees: String(r[COL['想邀請']]),
    email: String(r[COL.Email]), ig: String(r[COL.IG]), role: String(r[COL['角色']]),
    guess: String(r[COL['猜中']]), message: String(r[COL['訊息']]), slots: slots,
    status: String(r[COL['狀態']] || '待安排'),
    confDate: fmtDate_(r[COL['確認日期']]),
    confTime: isDate_(r[COL['確認時間']]) ? Utilities.formatDate(r[COL['確認時間']], TZ, 'HH:mm') : String(r[COL['確認時間']] || ''),
    confCity: String(r[COL['確認城市']] || ''), eventId: String(r[COL.EventId] || '')
  };
}
function allRows_() {
  const s = sheet_(SHEET_REPLY, HEAD);
  const n = s.getLastRow() - 1;
  if (n < 1) return [];
  return s.getRange(2, 1, n, HEAD.length).getValues().map((r, i) => rowObj_(r, i + 2));
}
function rosterNames_() {
  const s = sheet_(SHEET_ROSTER, ['暱稱（把你要邀請的朋友一個一個填在這一欄）']);
  const n = s.getLastRow() - 1;
  if (n < 1) return [];
  return s.getRange(2, 1, n, 1).getValues().map(r => String(r[0]).trim()).filter(Boolean);
}
function missing_() {
  const done = {};
  allRows_().forEach(r => done[norm_(r.name)] = true);
  return rosterNames_().filter(n => !done[norm_(n)]);
}

/* ---------- 第一次執行：建立分頁並授權 ---------- */
function setup() {
  sheet_(SHEET_REPLY, HEAD);
  sheet_(SHEET_ROSTER, ['暱稱（把你要邀請的朋友一個一個填在這一欄）']);
  CalendarApp.getDefaultCalendar();
  Logger.log('完成。請繼續照 README 部署成網頁應用程式。');
}

/* ---------- 朋友送出回覆（邀請卡網頁呼叫） ---------- */
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const d = JSON.parse(e.postData.contents);
    const name = safe_(d.name, 40);
    if (!name) return json_({ ok: false, error: 'name' });
    const s = sheet_(SHEET_REPLY, HEAD);
    if (s.getLastRow() > MAX_ROWS) return json_({ ok: false, error: 'full' });
    const email = safe_(d.email, 100);
    if (email && !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) return json_({ ok: false, error: 'email' });
    const slots = (Array.isArray(d.slots) ? d.slots : []).slice(0, 40).map(x => ({
      city: String(x.city || '').slice(0, 20), d: String(x.d || '').slice(0, 10),
      to: String(x.to || '').slice(0, 10), note: String(x.note || '').slice(0, 30)
    }));
    const id = Utilities.getUuid().slice(0, 8);
    const row = [
      id, new Date(), name, safe_(d.relation, 60), safe_(d.city, 40),
      safe_((d.dates || []).join ? d.dates.join('\n') : d.dates, 600), safe_(d.format, 60),
      safe_((d.invitees || []).join ? d.invitees.join('、') : d.invitees, 300), email, safe_(d.ig, 60),
      safe_(d.role, 30), safe_(d.guess, 20), safe_(d.message, 2000), JSON.stringify(slots), '待安排', '', '', '', ''
    ];
    s.appendRow(row);
    SpreadsheetApp.flush();
    try { notify_(allRows_().pop()); } catch (err) { console.error(err); }
    return json_({ ok: true });
  } catch (err) {
    console.error(err);
    return json_({ ok: false, error: 'bad' });
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}

function notify_(r) {
  const me = Session.getEffectiveUser().getEmail();
  const all = allRows_();
  const miss = missing_();
  const lines = [
    `${r.name} 剛剛填了見面邀請卡！`, '',
    `關係：${r.relation}`, `城市：${r.city}`, `可約時段：\n${r.dates}`, `見面方式：${r.format}`,
    r.invitees ? `想邀請：${r.invitees}` : '', `聯絡：${[r.email, r.ig && 'IG ' + r.ig].filter(Boolean).join(' / ') || '（沒留）'}`,
    '', `目前已填 ${all.length} 人`,
    rosterNames_().length ? (miss.length ? `還沒填（${miss.length}）：${miss.join('、')}` : '名單上的人都填完了 🎉') : '（「名單」分頁還沒填，填了就能看到誰還沒填）',
    '', `到後台排時間：${adminUrl_()}`
  ].filter(x => x !== '');
  MailApp.sendEmail({ to: me, subject: `【見面邀請】${r.name} 填好了（${all.length} 人）`, body: lines.join('\n') });
}

/* ---------- 後台頁面 ---------- */
function doGet(e) {
  if (e && e.parameter && e.parameter.key === ADMIN_KEY && ADMIN_KEY !== '請改成你自己的密碼') {
    const t = HtmlService.createTemplateFromFile('Admin');
    t.key = ADMIN_KEY;
    return t.evaluate().setTitle('見面排程後台')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  return ContentService.createTextOutput('OK');
}

function adminData(key) {
  auth_(key);
  return { rows: allRows_(), roster: rosterNames_(), missing: missing_() };
}

function saveRoster(key, names) {
  auth_(key);
  const s = sheet_(SHEET_ROSTER, ['暱稱（把你要邀請的朋友一個一個填在這一欄）']);
  if (s.getLastRow() > 1) s.getRange(2, 1, s.getLastRow() - 1, 1).clearContent();
  const list = names.map(n => safe_(n, 40)).filter(Boolean).slice(0, 300);
  if (list.length) s.getRange(2, 1, list.length, 1).setValues(list.map(n => [n]));
  return { roster: rosterNames_(), missing: missing_() };
}

function findRow_(id) {
  const s = sheet_(SHEET_REPLY, HEAD);
  const n = s.getLastRow() - 1;
  const ids = n > 0 ? s.getRange(2, 1, n, 1).getValues() : [];
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return { s: s, no: i + 2 };
  throw new Error('找不到這筆回覆');
}

/** o = {date:'2026-12-11', start:'12:00', end:'14:00', city, force, sendMail, subject, body} */
function confirmBooking(key, id, o) {
  auth_(key);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(o.date) || !/^\d{2}:\d{2}$/.test(o.start) || !/^\d{2}:\d{2}$/.test(o.end))
    throw new Error('日期或時間格式不對');
  const f = findRow_(id);
  const r = rowObj_(f.s.getRange(f.no, 1, 1, HEAD.length).getValues()[0], f.no);
  const start = new Date(o.date + 'T' + o.start + ':00+08:00');
  const end = new Date(o.date + 'T' + o.end + ':00+08:00');
  if (!(end > start)) throw new Error('結束時間要比開始時間晚');
  const cal = CalendarApp.getDefaultCalendar();
  if (!o.force) {
    const clash = cal.getEvents(start, end).filter(ev => ev.getId() !== r.eventId).map(ev => ev.getTitle());
    if (clash.length) return { conflict: clash };
  }
  if (r.eventId) { try { const old = cal.getEventById(r.eventId); if (old) old.deleteEvent(); } catch (e) {} }
  const ev = cal.createEvent('見面：' + r.name, start, end, {
    location: o.city || r.city,
    description: ['關係：' + r.relation, '見面方式：' + r.format, r.invitees && ('想邀請：' + r.invitees),
      r.email && ('Email：' + r.email), r.ig && ('IG：' + r.ig)].filter(Boolean).join('\n')
  });
  f.s.getRange(f.no, COL['狀態'] + 1, 1, 5).setValues([['已確認', o.date, o.start + '–' + o.end, ev.getId(), safe_(o.city, 40)]]);
  let mailed = false;
  if (o.sendMail && r.email) {
    MailApp.sendEmail({ to: r.email, subject: safe_(o.subject, 100) || '我們確定見面囉！', body: String(o.body || '').slice(0, 3000) });
    mailed = true;
  }
  return { ok: true, mailed: mailed, rows: allRows_() };
}

function cancelBooking(key, id) {
  auth_(key);
  const f = findRow_(id);
  const r = rowObj_(f.s.getRange(f.no, 1, 1, HEAD.length).getValues()[0], f.no);
  if (r.eventId) { try { const ev = CalendarApp.getDefaultCalendar().getEventById(r.eventId); if (ev) ev.deleteEvent(); } catch (e) {} }
  f.s.getRange(f.no, COL['狀態'] + 1, 1, 5).setValues([['待安排', '', '', '', '']]);
  return { rows: allRows_() };
}
