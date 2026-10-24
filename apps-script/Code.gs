/**
 * 圈出詐騙｜嶺口國小防詐騙宣導網頁 成績收集程式（Google Apps Script）
 *
 * 用法：在 Google 試算表的「擴充功能 → Apps Script」貼上這整個檔案，
 * 先執行一次 setup，再「部署 → 新增部署作業 → 網頁應用程式」
 * （執行身分：我；誰可以存取：所有人）。詳細步驟見 README。
 */

// 要和 index.html 裡的 REPORT_KEY 一樣
const KEY = 'lingkou-165';

const SHEET_SUMMARY = '成績總表';
const SHEET_ITEMS = '逐題紀錄';
const SHEET_WRONG = '錯題統計';
const SHEET_CLASS = '班級統計';
const GAMES = ['防詐全壘打', '真假快判', '找出破綻', '來電應變'];

const SUMMARY_HEAD = ['時間', '班級', '座號', '匿名代號', '遊戲', '關卡／模式', '得分', '滿分', '答對', '答錯', '獲得章', '備註'];
const ITEMS_HEAD = ['時間', '班級', '座號', '匿名代號', '遊戲', '關卡／模式', '題目', '學生作答', '結果'];

/** 建立四個分頁和統計公式。第一次部署前執行一次即可，重複執行不會刪資料。 */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSheet_(ss, SHEET_SUMMARY, SUMMARY_HEAD);
  ensureSheet_(ss, SHEET_ITEMS, ITEMS_HEAD);

  const wrong = ss.getSheetByName(SHEET_WRONG) || ss.insertSheet(SHEET_WRONG);
  wrong.getRange('A1').setFormula(
    "=QUERY('" + SHEET_ITEMS + "'!A:I," +
    "\"select E, G, count(I) where I = '答錯' or I = '沒找到' group by E, G order by count(I) desc " +
    "label E '遊戲', G '題目', count(I) '答錯人次'\", 1)"
  );

  const cls = ss.getSheetByName(SHEET_CLASS) || ss.insertSheet(SHEET_CLASS);
  cls.getRange('A1').setFormula(
    "=QUERY('" + SHEET_SUMMARY + "'!A:L," +
    "\"select B, E, count(E), avg(G) where B <> '' group by B, E order by B " +
    "label B '班級', E '遊戲', count(E) '遊玩次數', avg(G) '平均得分'\", 1)"
  );

  const first = ss.getSheetByName('工作表1') || ss.getSheetByName('Sheet1');
  if (first && first.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(first);
}

/** 用瀏覽器打開部署網址時，確認程式有在運作。 */
function doGet() {
  return ContentService.createTextOutput('圈出詐騙成績收集：運作中');
}

/** 網頁每玩完一關送來一筆資料：summary 一列寫入成績總表，items 每題一列寫入逐題紀錄。 */
function doPost(e) {
  try {
    const raw = (e && e.postData && e.postData.contents) || '';
    if (!raw || raw.length > 30000) return out_('rejected');
    const data = JSON.parse(raw);
    if (data.key !== KEY) return out_('rejected');
    const s = data.summary || {};
    if (GAMES.indexOf(s.game) < 0) return out_('rejected');

    const now = new Date();
    const who = [clean_(data.cls, 20), clean_(data.seat, 10), clean_(data.anon, 20)];
    const level = clean_(s.level, 40);
    const summaryRow = [now].concat(who, [
      s.game, level, num_(s.score), num_(s.max), num_(s.correct), num_(s.wrong),
      s.stamp ? '是' : '', clean_(s.note, 300)
    ]);
    const items = (Array.isArray(data.items) ? data.items : []).slice(0, 60).map(function (it) {
      return [now].concat(who, [s.game, level, clean_(it.q, 200), clean_(it.a, 120), clean_(it.r, 10)]);
    });

    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      if (!ss.getSheetByName(SHEET_SUMMARY) || !ss.getSheetByName(SHEET_ITEMS)) setup();
      append_(ss.getSheetByName(SHEET_SUMMARY), [summaryRow]);
      if (items.length) append_(ss.getSheetByName(SHEET_ITEMS), items);
    } finally {
      lock.releaseLock();
    }
    return out_('ok');
  } catch (err) {
    return out_('error');
  }
}

function ensureSheet_(ss, name, head) {
  const sh = ss.getSheetByName(name) || ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function append_(sh, rows) {
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
}

// 文字欄位：去掉換行、限制長度；開頭是 = + - @ 的加上單引號，避免被當成公式
function clean_(v, max) {
  let s = String(v === null || v === undefined ? '' : v).replace(/[\r\n\t]+/g, ' ').trim().slice(0, max);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}

function num_(v) {
  if (v === '' || v === null || v === undefined) return '';
  const n = Number(v);
  return isFinite(n) ? n : '';
}

function out_(msg) {
  return ContentService.createTextOutput(msg);
}
