/**
 * 歌詞省字王 - 成績紀錄 (Google Apps Script)
 *
 * 部署成「網頁應用程式」：
 *   執行身分：我
 *   誰可以存取：<學校網域> 中的所有人
 * 學生用學校帳號開啟網址即可闖關，每完成一首歌，
 * 成績會寫入這份試算表的「成績」工作表 (每位學生一列，只保留各首最佳成績)。
 */

const SHEET_NAME = '成績';

// [代碼, 歌名, 最多可省字數(完美), 原始歌詞字數]
const SONGS = [
  ['star', '小星星', 9, 42],
  ['congrats', '恭喜恭喜', 119, 247],
  ['bird', '飛鳥和蟬', 221, 396],
  ['content', '知足', 124, 311],
  ['balloon', '告白氣球', 136, 306]
];

const HEADERS = ['學生 Email']
  .concat(SONGS.map(s => s[1]))
  .concat(['完成首數', '完美首數', '總省字', '總分', '最後更新']);

// 每首歌的分數：未完成 0 分；還原成功但沒省到字 (≤0) 0 分；省字 15 分；完美 20 分
function songPoints_(best, maxSave) {
  if (best === null || best === '' || typeof best !== 'number' || best <= 0) return 0;
  return best >= maxSave ? 20 : 15;
}

function totalPoints_(bestList) {
  return bestList.reduce((sum, v, i) => sum + songPoints_(v, SONGS[i][2]), 0);
}

// 開啟網頁
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('歌詞省字王 - 模組化互動教具')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

// 取得目前學生的學校帳號 (不同網域或未登入時會是空字串)
function getUserEmail() {
  return Session.getActiveUser().getEmail() || '';
}

// 儲存成績：scores = { star: 9, congrats: 112, ... } (本次闖關的各首最佳)
// 與試算表中原有成績比較，每首只保留較高的分數。
function saveScores(scores) {
  const email = Session.getActiveUser().getEmail();
  if (!email) {
    throw new Error('無法取得學生帳號，請確認已用學校帳號登入。');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(20000); // 避免全班同時送出時互相覆蓋
  try {
    const sheet = getSheet_();
    const lastRow = sheet.getLastRow();
    const emails = lastRow > 1
      ? sheet.getRange(2, 1, lastRow - 1, 1).getValues().map(r => String(r[0]).trim().toLowerCase())
      : [];
    const idx = emails.indexOf(email.toLowerCase());
    const rowNum = idx === -1 ? lastRow + 1 : idx + 2;
    const old = idx === -1 ? [] : sheet.getRange(rowNum, 2, 1, SONGS.length).getValues()[0];

    const best = SONGS.map(([key, , , originalCount], i) => {
      const prev = typeof old[i] === 'number' ? old[i] : null;
      const v = scores ? scores[key] : null;
      // 只接受合理的整數 (不可能省超過原始字數)
      const cur = (typeof v === 'number' && isFinite(v) && v <= originalCount) ? Math.round(v) : null;
      if (prev === null) return cur;
      if (cur === null) return prev;
      return Math.max(prev, cur);
    });

    const done = best.filter(v => v !== null).length;
    const perfect = best.filter((v, i) => v !== null && v >= SONGS[i][2]).length;
    const total = best.reduce((sum, v) => sum + (v || 0), 0);

    const row = [email]
      .concat(best.map(v => (v === null ? '' : v)))
      .concat([done, perfect, total, totalPoints_(best), new Date()]);
    sheet.getRange(rowNum, 1, 1, row.length).setValues([row]);
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

// 取得 (或建立) 成績工作表
function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
  } else {
    upgradeSheet_(sheet);
  }
  return sheet;
}

// 舊版工作表沒有「總分」欄：在「最後更新」前插入一欄，並幫既有學生算好總分
function upgradeSheet_(sheet) {
  const totalCol = HEADERS.indexOf('總分') + 1; // 第 10 欄 (J)
  const header = sheet.getRange(1, totalCol).getValue();
  if (header === '總分') return;
  if (header !== '最後更新') return; // 不是預期的舊格式，就不動它
  sheet.insertColumnBefore(totalCol);
  sheet.getRange(1, totalCol).setValue('總分').setFontWeight('bold');
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  const scores = sheet.getRange(2, 2, lastRow - 1, SONGS.length).getValues();
  sheet.getRange(2, totalCol, lastRow - 1, 1).setValues(scores.map(r => [totalPoints_(r)]));
}
