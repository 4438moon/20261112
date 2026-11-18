/**
 * 안전 탐험대 - 학습 진전도 모으기 (Google Apps Script)
 * 구글 스프레드시트의 [확장 프로그램 > Apps Script]에 붙여 넣어 사용합니다.
 * '기록' 시트에 학생이 보낸 기록이 모두 쌓이고,
 * '진전도' 시트에 학생 번호별 진행 상황이, '많이 틀린 수칙' 시트에 다시 지도할 수칙이 자동으로 정리됩니다.
 */
const SHEET_LOG = '기록';
const SHEET_SUM = '진전도';
const SHEET_MISS = '많이 틀린 수칙';
const MAX_GRADE = 6;
const MAX_CLASS = 15;
const MAX_NUM = 40;
const ZONE_COUNT = 4;   // 구간 수
const TOTAL = 16;       // 찾아야 하는 위험 수 (보통 방식)
const TOTAL_EASY = 8;   // 찾아야 하는 위험 수 (쉬움 방식)

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('안전 탐험대')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

/** 게임을 다른 곳(예: GitHub Pages)에 올렸을 때 기록을 받는 통로입니다. */
function doPost(e) {
  try {
    saveProgress(JSON.parse(e.postData.contents));
    return ContentService.createTextOutput('ok');
  } catch (err) {
    return ContentService.createTextOutput('error');
  }
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('안전 탐험대')
    .addItem('진전도 다시 계산', 'rebuildSummary')
    .addToUi();
}

function fmt_(sec) {
  const m = Math.floor(sec / 60), r = Math.round(sec - m * 60);
  return m + '분 ' + r + '초';
}

function int_(v) {
  const n = Number(v);
  return Math.floor(n) === n ? n : NaN;
}

function logSheet_(ss) {
  let sh = ss.getSheetByName(SHEET_LOG);
  if (!sh) {
    sh = ss.insertSheet(SHEET_LOG);
    sh.appendRow(['기록 시각', '학년', '반', '번호', '통과한 구간 수', '찾은 위험 수', '틀린 횟수', '걸린 시간(초)', '걸린 시간', '틀렸던 수칙', '상태', '방식']);
    sh.setFrozenRows(1);
  }
  return sh;
}

/** 게임 화면에서 호출: 학생의 현재 진행 상황 1건을 저장합니다. */
function saveProgress(r) {
  r = r || {};
  const grade = int_(r.grade), cls = int_(r.cls), num = int_(r.num);
  const zone = int_(r.zone), found = int_(r.found), wrong = int_(r.wrong);
  const sec = Math.round(Number(r.sec));
  if (!(grade >= 1 && grade <= MAX_GRADE) || !(cls >= 1 && cls <= MAX_CLASS) || !(num >= 1 && num <= MAX_NUM)) {
    throw new Error('학년, 반, 번호를 다시 확인해 주세요.');
  }
  const easy = r.mode === 'easy';
  const total = easy ? TOTAL_EASY : TOTAL;
  if (!(zone >= 0 && zone <= ZONE_COUNT) || !(found >= 0 && found <= total) || !(wrong >= 0 && wrong < 1000) || !(sec >= 0 && sec < 36000)) {
    throw new Error('기록 값이 올바르지 않아요.');
  }
  // 수식으로 입력되지 않도록 앞머리의 = + - @ 기호는 지웁니다.
  const missed = String(r.missed || '').replace(/^[=+\-@\s]+/, '').slice(0, 600);
  const done = r.done === true && zone === ZONE_COUNT && found === total;
  const lock = LockService.getScriptLock();
  // 여러 학생이 한꺼번에 보내면 차례를 기다립니다. (최대 25초)
  lock.waitLock(25000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = logSheet_(ss);
    sh.getRange('J:J').setNumberFormat('@');
    sh.appendRow([new Date(), grade, cls, num, zone, found, wrong, sec, fmt_(sec), missed, done ? '완료' : '진행 중', easy ? '쉬움' : '보통']);
    // 정리 시트 갱신이 실패해도 기록 저장은 그대로 둡니다.
    try { writeSummary_(ss, sh.getDataRange().getValues().slice(1)); } catch (e) { console.error(e); }
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

/** 학생 번호별 진행 상황과, 많이 틀린 수칙을 정리합니다. */
function writeSummary_(ss, rows) {
  const by = {};
  rows.forEach(function (v) {
    if (!v[1] || !v[2] || !v[3]) return;
    // 같은 학생이 보통과 쉬움을 모두 했으면 방식별로 한 줄씩 정리합니다.
    const mode = v[11] === '쉬움' ? '쉬움' : '보통';
    const key = v[1] + '-' + v[2] + '-' + v[3] + '-' + mode;
    const s = by[key] || (by[key] = { g: v[1], c: v[2], n: v[3], mode: mode, zone: 0, found: 0, wrong: 0, time: '', done: false, miss: {}, tries: 0, last: v[0] });
    const zone = Number(v[4]), found = Number(v[5]);
    if (zone === 0) s.tries++;
    if (zone > s.zone || (zone === s.zone && found >= s.found)) { s.zone = zone; s.found = found; s.wrong = v[6]; }
    if (v[10] === '완료' && !s.done) { s.done = true; s.time = v[8]; }
    String(v[9] || '').split(' / ').forEach(function (m) { if (m) s.miss[m] = true; });
    s.last = v[0];
  });
  const list = Object.keys(by).map(function (k) { return by[k]; })
    .sort(function (a, b) { return a.g - b.g || a.c - b.c || a.n - b.n || (a.mode < b.mode ? -1 : 1); });
  const mark = function (s, i) { return s.zone >= i ? '●' : '○'; };
  const out = list.map(function (s) {
    return [s.g, s.c, s.n, s.mode, s.done ? '완료' : (s.zone ? s.zone + '구간 통과' : '시작함'),
      mark(s, 1), mark(s, 2), mark(s, 3), mark(s, 4),
      s.found + '개 (' + (s.mode === '쉬움' ? TOTAL_EASY : TOTAL) + '개 중)', s.wrong, s.time,
      Object.keys(s.miss).join(' / '), Math.max(1, s.tries), s.last];
  });
  const head = ['학년', '반', '번호', '방식', '진행 상태', '1구간 버스', '2구간 걸어서 이동', '3구간 체험 장소', '4구간 돌아가는 길',
    '찾은 위험', '틀린 횟수', '완료까지 걸린 시간', '틀렸던 수칙', '시작한 횟수', '마지막 기록 시각'];
  let sum = ss.getSheetByName(SHEET_SUM);
  if (!sum) sum = ss.insertSheet(SHEET_SUM);
  sum.clearContents();
  sum.getRange('J:J').setNumberFormat('@');
  sum.getRange('M:M').setNumberFormat('@');
  sum.getRange(1, 1, 1, head.length).setValues([head]);
  sum.setFrozenRows(1);
  if (out.length) sum.getRange(2, 1, out.length, head.length).setValues(out);

  const tally = {};
  list.forEach(function (s) { Object.keys(s.miss).forEach(function (m) { tally[m] = (tally[m] || 0) + 1; }); });
  const rank = Object.keys(tally).map(function (m) { return [m, tally[m]]; })
    .sort(function (a, b) { return b[1] - a[1]; });
  let miss = ss.getSheetByName(SHEET_MISS);
  if (!miss) miss = ss.insertSheet(SHEET_MISS);
  miss.clearContents();
  miss.getRange('A:A').setNumberFormat('@');
  miss.getRange(1, 1, 1, 2).setValues([['수칙', '틀린 학생 수']]);
  miss.setFrozenRows(1);
  if (rank.length) miss.getRange(2, 1, rank.length, 2).setValues(rank);
}

/** 시트 메뉴에서 실행: '기록' 시트를 직접 고친 뒤 다시 정리할 때 씁니다. */
function rebuildSummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  writeSummary_(ss, logSheet_(ss).getDataRange().getValues().slice(1));
}
