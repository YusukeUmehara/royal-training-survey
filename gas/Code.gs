/** @OnlyCurrentDoc */
/**
 * ROYAL 新人研修 振り返りアンケート — 回答受け取り用（Google Apps Script）
 * 回答用スプレッドシートの「拡張機能 → Apps Script」に貼り付けて、ウェブアプリとしてデプロイします。
 */
const SHEET_NAME = '回答';
const HEADERS = [
  '送信日時', '名前', '所属会社', '参加した研修',
  '一番印象に残った研修', 'その理由',
  'その他印象に残った研修', 'その理由（その他）',
  '4日間を通しての感想', 'こんな研修あったらいいな',
  '回答ID', '端末',
];
const SESSIONS = ['商材説明', '金融', '結果の出る新人のスタンス', '具体と抽象', 'チームビルディング', 'ロジカルシンキング',
  '鬼速PDCA', 'データ分析', 'タスクジュール', '世界基準の部下の育て方', '1on1研修'];
const DAYS = ['Day1', 'Day2', 'Day3', 'Day4'];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const d = JSON.parse(e.postData.contents);
    const sh = getSheet_();

    // 同じ回答IDの二重送信は保存しない
    const id = str_(d.id, 80);
    const last = sh.getLastRow();
    if (id && last > 1) {
      const ids = sh.getRange(2, 11, last - 1, 1).getValues().flat();
      if (ids.indexOf(id) !== -1) return json_({ ok: true, duplicate: true });
    }

    const name = str_(d.name, 40);
    const best = SESSIONS.indexOf(d.best) !== -1 ? d.best : '';
    if (!name || !best) return json_({ ok: false, error: 'missing required fields' });

    sh.appendRow([
      new Date(),
      safe_(name),
      safe_(str_(d.company, 60)),
      list_(d.days, DAYS).join('、'),
      best,
      safe_(str_(d.bestReason, 1000)),
      list_(d.others, SESSIONS).join('、'),
      safe_(str_(d.othersReason, 1000)),
      safe_(str_(d.impression, 2000)),
      safe_(str_(d.idea, 1000)),
      id,
      str_(d.ua, 200),
    ]);
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return json_({ ok: true, service: 'royal-training-survey' });
}

/** 初回だけ手動で実行：シート名・見出し・書式を整えます（承認画面が出たら許可） */
function setup() {
  const sh = getSheet_();
  sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS])
    .setFontWeight('bold').setBackground('#1A1A1A').setFontColor('#F2D21D');
  sh.setFrozenRows(1);
  sh.setColumnWidth(1, 150);
  [6, 8, 9, 10].forEach(c => sh.setColumnWidth(c, 360));
  sh.getRange('A:A').setNumberFormat('yyyy/mm/dd HH:mm');
  sh.getRange('B:L').setWrap(true).setVerticalAlignment('top');
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.getSheets()[0];
    sh.setName(SHEET_NAME);
  }
  if (sh.getLastRow() === 0) sh.appendRow(HEADERS);
  return sh;
}

function str_(v, max) { return (v == null ? '' : String(v)).trim().slice(0, max); }
function list_(v, allowed) { return Array.isArray(v) ? v.filter(x => allowed.indexOf(x) !== -1) : []; }
// 数式として解釈されないようにする
function safe_(s) { return /^[=+\-@]/.test(s) ? "'" + s : s; }
function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
