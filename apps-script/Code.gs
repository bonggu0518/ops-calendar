/**
 * 생산운영본부 일정 캘린더 — 구글 시트 저장용 Apps Script
 *
 * 처음 한 번:
 *   1) 아래 INITIAL_PASSWORD 를 본부 공용 비밀번호로 바꾼다
 *   2) 상단 함수 선택에서 setup 을 고르고 [실행] → 권한 승인
 *   3) [배포] → [새 배포] → 유형: 웹 앱 / 실행: 나 / 액세스: 모든 사용자 → URL 복사
 *
 * 비밀번호를 바꿀 때: INITIAL_PASSWORD 수정 → setPassword 실행 (재배포 필요 없음)
 * 팀·구분 목록을 바꿀 때: '설정' 시트에서 직접 수정 (재배포 필요 없음)
 */

const INITIAL_PASSWORD = 'change-me-1234';

const SHEET_EVENTS = '일정';
const SHEET_CONFIG = '설정';
const HEADERS = ['id', '팀', '구분', '제목', '담당자', '시작일', '종료일', '시작시간', '종료시간', '장소', '메모', '등록자', '등록일시', '수정일시'];
const FIELDS  = ['id', 'team', 'category', 'title', 'person', 'start', 'end', 'startTime', 'endTime', 'location', 'memo', 'author', 'createdAt', 'updatedAt'];

const DEFAULT_TEAMS = [
  ['생산기획팀', '#2f6fde'],
  ['생산팀', '#e0782b'],
  ['생산물류팀', '#1f9d6a'],
  ['품질관리팀', '#9b51e0'],
  ['공무팀', '#d2416e'],
];
const DEFAULT_CATEGORIES = ['미팅', '출장', '외근', '교육', '행사', '연차', '반차', '병가', '기타'];

/* ---------- 초기 설정 ---------- */

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let ev = ss.getSheetByName(SHEET_EVENTS);
  if (!ev) ev = ss.insertSheet(SHEET_EVENTS);
  if (ev.getLastRow() === 0) {
    ev.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold').setBackground('#eef1f5');
    ev.setFrozenRows(1);
    ev.getRange(1, 1, ev.getMaxRows(), HEADERS.length).setNumberFormat('@'); // 날짜를 글자로 유지
    ev.setColumnWidth(1, 120);
    ev.setColumnWidth(4, 240);
    ev.setColumnWidth(11, 240);
  }

  let cf = ss.getSheetByName(SHEET_CONFIG);
  if (!cf) cf = ss.insertSheet(SHEET_CONFIG);
  if (cf.getLastRow() === 0) {
    cf.getRange(1, 1, 1, 4).setValues([['팀명', '색상(선택, #RRGGBB)', '', '구분']]).setFontWeight('bold').setBackground('#eef1f5');
    cf.getRange(2, 1, DEFAULT_TEAMS.length, 2).setValues(DEFAULT_TEAMS);
    cf.getRange(2, 4, DEFAULT_CATEGORIES.length, 1).setValues(DEFAULT_CATEGORIES.map(c => [c]));
    cf.setColumnWidth(1, 160);
    cf.setColumnWidth(2, 160);
    cf.setColumnWidth(4, 120);
  }

  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('PASSWORD')) props.setProperty('PASSWORD', INITIAL_PASSWORD);

  const sheet1 = ss.getSheetByName('시트1') || ss.getSheetByName('Sheet1');
  if (sheet1 && ss.getSheets().length > 2 && sheet1.getLastRow() === 0) ss.deleteSheet(sheet1);
}

function setPassword() {
  PropertiesService.getScriptProperties().setProperty('PASSWORD', INITIAL_PASSWORD);
}

/* ---------- 웹 요청 ---------- */

function doGet() {
  return json_({ ok: true, message: '일정 캘린더 서버가 동작 중입니다.' });
}

function doPost(e) {
  try {
    const req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const saved = PropertiesService.getScriptProperties().getProperty('PASSWORD');
    if (!saved || String(req.password || '') !== saved) {
      return json_({ ok: false, code: 'AUTH', error: '비밀번호가 맞지 않습니다.' });
    }
    switch (req.action) {
      case 'auth':   return json_({ ok: true, data: true });
      case 'config': return json_({ ok: true, data: readConfig_() });
      case 'list':   return json_({ ok: true, data: listEvents_(req.from, req.to) });
      case 'create': return json_({ ok: true, data: withLock_(() => createEvent_(req.event)) });
      case 'update': return json_({ ok: true, data: withLock_(() => updateEvent_(req.event)) });
      case 'delete': return json_({ ok: true, data: withLock_(() => deleteEvent_(req.id)) });
      default:       return json_({ ok: false, error: '알 수 없는 요청입니다: ' + req.action });
    }
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

/* ---------- 데이터 처리 ---------- */

function readConfig_() {
  const cf = sheet_(SHEET_CONFIG);
  const rows = cf.getLastRow() > 1 ? cf.getRange(2, 1, cf.getLastRow() - 1, 4).getDisplayValues() : [];
  const teams = rows.filter(r => r[0].trim()).map(r => ({ name: r[0].trim(), color: r[1].trim() }));
  const categories = rows.filter(r => r[3].trim()).map(r => r[3].trim());
  return {
    teams: teams.length ? teams : DEFAULT_TEAMS.map(t => ({ name: t[0], color: t[1] })),
    categories: categories.length ? categories : DEFAULT_CATEGORIES,
  };
}

function allRows_() {
  const ev = sheet_(SHEET_EVENTS);
  const n = ev.getLastRow() - 1;
  if (n < 1) return [];
  return ev.getRange(2, 1, n, FIELDS.length).getDisplayValues();
}

function rowToEvent_(r) {
  const o = {};
  FIELDS.forEach((f, i) => { o[f] = r[i]; });
  return o;
}

function listEvents_(from, to) {
  return allRows_()
    .map(rowToEvent_)
    .filter(e => e.id && (!from || (e.end || e.start) >= from) && (!to || e.start <= to));
}

function clean_(ev) {
  const s = v => String(v == null ? '' : v).trim().slice(0, 500);
  const o = {
    team: s(ev.team), category: s(ev.category), title: s(ev.title), person: s(ev.person),
    start: s(ev.start), end: s(ev.end) || s(ev.start),
    startTime: s(ev.startTime), endTime: s(ev.endTime),
    location: s(ev.location), memo: s(ev.memo), author: s(ev.author),
  };
  const d = /^\d{4}-\d{2}-\d{2}$/, t = /^(\d{2}:\d{2})?$/;
  if (!o.team) throw new Error('팀을 선택해 주세요.');
  if (!o.category) throw new Error('구분을 선택해 주세요.');
  if (!o.title && !o.person) throw new Error('제목이나 담당자 중 하나는 입력해 주세요.');
  if (!d.test(o.start) || !d.test(o.end)) throw new Error('날짜 형식이 올바르지 않습니다.');
  if (o.end < o.start) throw new Error('종료일이 시작일보다 빠릅니다.');
  if (!t.test(o.startTime) || !t.test(o.endTime)) throw new Error('시간 형식이 올바르지 않습니다.');
  return o;
}

function now_() {
  return Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm');
}

function writeRow_(rowIndex, o) {
  const ev = sheet_(SHEET_EVENTS);
  const range = ev.getRange(rowIndex, 1, 1, FIELDS.length);
  range.setNumberFormat('@');
  range.setValues([FIELDS.map(f => o[f] || '')]);
}

function createEvent_(input) {
  const o = clean_(input || {});
  o.id = Utilities.getUuid().slice(0, 8) + Date.now().toString(36);
  o.createdAt = now_();
  o.updatedAt = o.createdAt;
  writeRow_(sheet_(SHEET_EVENTS).getLastRow() + 1, o);
  return o;
}

function findRow_(id) {
  if (!id) throw new Error('일정 id가 없습니다.');
  const rows = allRows_();
  for (let i = 0; i < rows.length; i++) if (rows[i][0] === id) return { index: i + 2, event: rowToEvent_(rows[i]) };
  throw new Error('일정을 찾지 못했습니다. 다른 사람이 먼저 삭제했을 수 있습니다.');
}

function updateEvent_(input) {
  const found = findRow_(input && input.id);
  const o = clean_(input);
  o.id = found.event.id;
  o.createdAt = found.event.createdAt;
  o.updatedAt = now_();
  writeRow_(found.index, o);
  return o;
}

function deleteEvent_(id) {
  const found = findRow_(id);
  sheet_(SHEET_EVENTS).deleteRow(found.index);
  return { id: id };
}

/* ---------- 공통 ---------- */

function sheet_(name) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sh) throw new Error("'" + name + "' 시트가 없습니다. Apps Script에서 setup 을 먼저 실행해 주세요.");
  return sh;
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
