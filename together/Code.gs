/**
 * 같이 가는 길 - 기록 모으기 (Google Apps Script)
 * 구글 스프레드시트의 [확장 프로그램 > Apps Script]에 붙여 넣어 사용합니다.
 * '기록' 시트에 모든 기록이 쌓이고, '반별 1등' 시트가 자동으로 갱신됩니다.
 */
const SHEET_LOG = '기록';
const SHEET_TOP = '반별 1등';
const MAX_GRADE = 6;
const MAX_CLASS = 15;

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('같이 가는 길')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

/** 게임을 GitHub Pages 등 다른 곳에 올렸을 때 기록을 받고, 반 순위를 돌려줍니다. */
function doPost(e) {
  let out;
  try {
    out = { ok: true, res: saveRecord(JSON.parse(e.postData.contents)) };
  } catch (err) {
    out = { ok: false, error: String((err && err.message) || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('같이 가는 길')
    .addItem('반별 1등 다시 계산', 'rebuildTop')
    .addToUi();
}

function fmt_(sec) {
  const m = Math.floor(sec / 60), r = sec - m * 60;
  return m + ':' + (r < 10 ? '0' : '') + r.toFixed(1);
}

function logSheet_(ss) {
  let sh = ss.getSheetByName(SHEET_LOG);
  if (!sh) {
    sh = ss.insertSheet(SHEET_LOG);
    sh.appendRow(['기록 시각', '학년', '반', '별명', '걸린 시간(초)', '걸린 시간', '문제 방식']);
    sh.setFrozenRows(1);
  }
  return sh;
}

/** 게임 화면에서 호출: 기록 1건을 저장하고 그 반의 순위를 돌려줍니다. */
function saveRecord(r) {
  const grade = Number(r && r.grade), cls = Number(r && r.cls);
  const sec = Math.round(Number(r && r.sec) * 10) / 10;
  // 수식으로 입력되지 않도록 별명 앞머리의 = + - @ 기호는 지웁니다.
  const nick = String((r && r.nick) || '').replace(/^[=+\-@\s]+/, '').slice(0, 12);
  const mode = r && r.mode === 'easy' ? '쉬움' : '보통';
  if (!(grade >= 1 && grade <= MAX_GRADE) || !(cls >= 1 && cls <= MAX_CLASS) || !nick || !(sec > 0 && sec < 36000)) {
    throw new Error('학년, 반, 별명을 다시 확인해 주세요.');
  }
  const lock = LockService.getScriptLock();
  // 여러 학생이 한꺼번에 보내면 차례를 기다립니다. (최대 25초)
  lock.waitLock(25000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = logSheet_(ss);
    // 별명 열을 '일반 텍스트'로 두어 007, 1/2 같은 별명이 숫자·날짜로 바뀌지 않게 합니다.
    sh.getRange('D:D').setNumberFormat('@');
    sh.appendRow([new Date(), grade, cls, nick, sec, fmt_(sec), mode]);
    const rows = sh.getDataRange().getValues().slice(1);
    // '반별 1등' 시트 갱신이 실패해도 기록 저장과 순위 안내는 그대로 진행합니다.
    try { writeTop_(ss, rows); } catch (e) { console.error(e); }
    const meIdx = rows.length - 1;
    const mine = rows.map(function (v, i) { return { v: v, i: i }; })
      .filter(function (o) { return Number(o.v[1]) === grade && Number(o.v[2]) === cls; })
      .sort(function (a, b) { return Number(a.v[4]) - Number(b.v[4]) || a.i - b.i; });
    const rank = mine.findIndex(function (o) { return o.i === meIdx; }) + 1;
    return {
      rank: rank, total: mine.length,
      top: mine.slice(0, 5).map(function (o) { return { nick: String(o.v[3]), sec: Number(o.v[4]), me: o.i === meIdx }; })
    };
  } finally {
    lock.releaseLock();
  }
}

/** 반별로 가장 빠른 1명을 '반별 1등' 시트에 정리합니다. (같은 시간이면 먼저 기록한 학생) */
function writeTop_(ss, rows) {
  const best = {};
  rows.forEach(function (v) {
    if (!v[1] || !v[2] || !v[4]) return;
    const key = v[1] + '-' + v[2];
    if (!best[key] || Number(v[4]) < Number(best[key][4])) best[key] = v;
  });
  const out = Object.keys(best).map(function (k) { return best[k]; })
    .sort(function (a, b) { return a[1] - b[1] || a[2] - b[2]; })
    .map(function (v) { return [v[1], v[2], v[3], fmt_(Number(v[4])), v[4], v[0], v[6]]; });
  let top = ss.getSheetByName(SHEET_TOP);
  if (!top) top = ss.insertSheet(SHEET_TOP);
  top.clearContents();
  top.getRange(1, 1, 1, 7).setValues([['학년', '반', '1등 별명', '걸린 시간', '걸린 시간(초)', '기록 시각', '문제 방식']]);
  top.setFrozenRows(1);
  if (out.length) top.getRange(2, 1, out.length, 7).setValues(out);
}

/** 시트 메뉴에서 실행: '기록' 시트를 직접 고친 뒤 1등을 다시 계산할 때 씁니다. */
function rebuildTop() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  writeTop_(ss, logSheet_(ss).getDataRange().getValues().slice(1));
}
