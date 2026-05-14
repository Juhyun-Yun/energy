/**
 * 🌍 지구지킴이 에너지 탐험대 — Google Apps Script 백엔드
 *
 * 📌 설정 순서:
 *
 *   1. Google Drive에서 새 스프레드시트 만들기
 *
 *   2. 첫 시트 이름을 "학생명단" 으로 바꾸기
 *        A1 칸 : 번호    B1 칸 : 이름   (헤더)
 *        A2~  : 1, 2, 3, ...           B2~ : 학생 이름들
 *        예)
 *          | 번호 | 이름   |
 *          |  1  | 김민준 |
 *          |  2  | 이서연 |
 *          |  3  | 박지호 |
 *
 *   3. 스프레드시트 안에서 메뉴: 확장 프로그램 > Apps Script
 *      (반드시 "그 스프레드시트 안에서" 열어야 자동 연결됩니다)
 *
 *   4. 기본으로 열린 Code.gs 파일 내용을 모두 지우고,
 *      이 파일(Code.gs) 내용을 통째로 복사해서 붙여넣기
 *
 *   5. SHEET_ID는 비워두셔도 됩니다 (자동으로 현재 스프레드시트 사용)
 *
 *   6. 디스크 모양 아이콘으로 저장 (Ctrl+S)
 *
 *   7. 우측 상단 "배포" > "새 배포"
 *        - 톱니바퀴 모양 ⚙️ 클릭 → 유형: "웹 앱"
 *        - 설명: 자원과 에너지 단원 (자유)
 *        - 실행 계정: 본인
 *        - 액세스: "모든 사용자"  (학교 환경이면 "도메인 내 사용자")
 *        - "배포" 클릭 → 권한 승인
 *        - 배포 후 표시되는 "웹 앱 URL"을 복사
 *
 *   8. config.js 파일을 열어 API_URL 항목에 7번 URL 붙여넣기
 *
 *   ✅ "진행기록" 시트는 자동으로 만들어집니다 (코드가 알아서 생성).
 */

// 💡 Apps Script를 스프레드시트의 "확장 프로그램 > Apps Script"로 열었다면
//    아래 SHEET_ID를 비워둬도 자동으로 그 스프레드시트를 사용합니다.
//    (다른 시트의 데이터를 쓰고 싶을 때만 ID를 넣어주세요)
const SHEET_ID = '';

const STUDENTS_SHEET = '학생명단';
const RECORDS_SHEET  = '진행기록';

/* ============ 메뉴 (스프레드시트 열 때 자동 생성) ============ */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📊 학습 대시보드')
    .addItem('📡 실시간 현황 (사이드바)', 'showSidebar')
    .addItem('📈 전체 현황 (큰 팝업)', 'showDashboard')
    .addSeparator()
    .addItem('🔄 진행기록 시트 초기화', 'resetRecords')
    .addToUi();
}

/* 큰 팝업 — 전체 현황 (탭별 상세) */
function showDashboard() {
  const html = HtmlService.createHtmlOutputFromFile('Dashboard')
    .setWidth(1100)
    .setHeight(680);
  SpreadsheetApp.getUi().showModalDialog(html, '우리 반 학습 통계');
}

/* 사이드바 — 실시간 현황 (자동 새로고침) */
function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('SidebarLive')
    .setTitle('📡 실시간 학습 현황');
  SpreadsheetApp.getUi().showSidebar(html);
}

/* 대시보드/사이드바에서 google.script.run 으로 호출하는 공개 함수 */
function getDashboardStats() {
  return getStats_();
}

/* 사이드바 — 최근 N개 활동 (시간 역순) */
function getRecentRecords(n) {
  const limit = Number(n) || 8;
  const sh = ensureRecordsSheet_();
  const last = sh.getLastRow();
  if (last < 2) return [];
  const start = Math.max(2, last - 200);                       // 최대 200개 안에서 추리기
  const rows = sh.getRange(start, 1, last - start + 1, 6).getValues();
  return rows
    .filter(r => r[1] && r[2])                                 // 학생/미션 있는 것만
    .map(r => ({
      time:    r[0] instanceof Date ? r[0].toISOString() : String(r[0]),
      student: String(r[1]),
      mission: String(r[2]),
      status:  String(r[3] || ''),
      score:   r[4]
    }))
    .sort((a, b) => new Date(b.time) - new Date(a.time))
    .slice(0, limit);
}

function resetRecords() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.alert('확인', '진행기록 시트의 모든 데이터를 지울까요?\n(되돌릴 수 없습니다)', ui.ButtonSet.YES_NO);
  if (r !== ui.Button.YES) return;
  const sh = ss_().getSheetByName(RECORDS_SHEET);
  if (sh) {
    const last = sh.getLastRow();
    if (last > 1) sh.getRange(2, 1, last - 1, sh.getLastColumn()).clearContent();
    ui.alert('완료', '진행기록을 초기화했어요.', ui.ButtonSet.OK);
  }
}

/* ============ 진입점 (웹 앱) ============ */
function doGet(e)  { return reply_(handle_(e)); }
function doPost(e) { return reply_(handle_(e)); }

function handle_(e) {
  try {
    // POST 요청: 진행 기록
    if (e && e.postData && e.postData.contents) {
      const data = JSON.parse(e.postData.contents);
      recordEvent_(data);
      return { ok: true };
    }
    // GET 요청: 조회
    const action = e && e.parameter && e.parameter.action;
    if (action === 'students') return { ok: true, students: getStudents_() };
    if (action === 'progress') return { ok: true, progress: getProgress_(e.parameter.student) };
    if (action === 'stats')    return { ok: true, stats:    getStats_() };
    if (action === 'ping')     return { ok: true, message: 'pong' };
    return { ok: false, error: 'unknown_action' };
  } catch (err) {
    return { ok: false, error: String(err && err.message || err) };
  }
}

function reply_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ============ 시트 접근 ============ */
function ss_() {
  // 1) SHEET_ID가 채워져 있으면 그걸로 열기
  if (SHEET_ID) return SpreadsheetApp.openById(SHEET_ID);
  // 2) 비어있으면 바인딩된 스프레드시트(스크립트를 만든 그 시트)를 자동 사용
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) return active;
  throw new Error('스프레드시트를 찾을 수 없어요. 스프레드시트의 "확장 프로그램 > Apps Script"로 열었는지 확인하거나, SHEET_ID를 수동 입력해 주세요.');
}

function ensureRecordsSheet_() {
  const ss = ss_();
  let sh = ss.getSheetByName(RECORDS_SHEET);
  if (!sh) {
    sh = ss.insertSheet(RECORDS_SHEET);
    sh.appendRow(['시간', '학생', '미션', '상태', '점수', '세부']);
    sh.setFrozenRows(1);
    sh.setColumnWidth(1, 160);
    sh.setColumnWidth(2, 100);
    sh.setColumnWidth(3, 60);
    sh.setColumnWidth(4, 60);
    sh.setColumnWidth(5, 60);
    sh.setColumnWidth(6, 360);
  }
  return sh;
}

/* ============ 데이터 함수 ============ */
function getStudents_() {
  const sh = ss_().getSheetByName(STUDENTS_SHEET);
  if (!sh) throw new Error('"' + STUDENTS_SHEET + '" 시트를 찾을 수 없습니다.');
  const last = sh.getLastRow();
  if (last < 2) return [];
  // A열은 번호, B열에서 이름 읽기
  return sh.getRange(2, 2, last - 1, 1).getValues()
    .map(r => String(r[0] || '').trim())
    .filter(v => v !== '');
}

function recordEvent_(data) {
  const sh = ensureRecordsSheet_();
  const time = data.timestamp ? new Date(data.timestamp) : new Date();
  const score = (data.score != null && data.score !== '') ? data.score : '';
  const detail = Object.assign({}, data);
  delete detail.student;
  delete detail.mission;
  delete detail.status;
  delete detail.score;
  delete detail.timestamp;
  sh.appendRow([
    time,
    String(data.student || ''),
    String(data.mission || ''),
    String(data.status || '완료'),
    score,
    JSON.stringify(detail)
  ]);
}

function getProgress_(student) {
  if (!student) return [];
  const sh = ensureRecordsSheet_();
  const data = sh.getDataRange().getValues();
  if (data.length < 2) return [];
  const target = String(student).trim();
  return data.slice(1)
    .filter(r => String(r[1]).trim() === target)
    .map(r => ({
      time:    r[0],
      student: r[1],
      mission: String(r[2]),
      status:  r[3],
      score:   r[4],
      detail:  r[5]
    }));
}

function getStats_() {
  const students = getStudents_();
  const sh = ensureRecordsSheet_();
  const data = sh.getDataRange().getValues();
  const records = data.slice(1);

  const seen = {};                  // 중복 제거
  const missionCompletions = {};    // mission -> Set(student)
  const studentProgress    = {};    // student -> Set(mission)
  const studentScores      = {};    // student -> {mission -> score}
  const quizScores         = [];    // [{student, score}]

  records.forEach(r => {
    const student = String(r[1] || '').trim();
    const mission = String(r[2] || '').trim();
    const score   = r[4];
    if (!student || !mission) return;
    const key = student + '|' + mission;

    if (!missionCompletions[mission]) missionCompletions[mission] = {};
    missionCompletions[mission][student] = true;

    if (!studentProgress[student]) studentProgress[student] = {};
    studentProgress[student][mission] = true;

    if (!studentScores[student]) studentScores[student] = {};
    // 가장 높은 점수 유지
    if (score !== '' && score != null) {
      const num = Number(score);
      if (!isNaN(num)) {
        const prev = studentScores[student][mission];
        studentScores[student][mission] = (prev == null) ? num : Math.max(prev, num);
      }
    }
    seen[key] = true;
  });

  // 미션별 집계
  const perMission = {};
  for (let m = 1; m <= 7; m++) {
    const set = missionCompletions[String(m)] || {};
    perMission[m] = {
      count: Object.keys(set).length,
      students: Object.keys(set)
    };
  }

  // 학생별 진행
  const perStudent = students.map(name => ({
    name,
    completed: Object.keys(studentProgress[name] || {}).map(Number).sort((a,b)=>a-b),
    scores:    studentScores[name] || {}
  }));

  // 7번 미션(골든벨) 점수만 따로
  Object.keys(studentScores).forEach(s => {
    if (studentScores[s]['7'] != null) {
      quizScores.push({ student: s, score: studentScores[s]['7'] });
    }
  });

  return {
    totalStudents: students.length,
    totalRecords: records.length,
    perMission,
    perStudent,
    quizScores
  };
}

/* ============ 테스트용 (Apps Script 에디터에서 직접 실행 가능) ============ */
function testGetStudents() { Logger.log(getStudents_()); }
function testGetStats()    { Logger.log(JSON.stringify(getStats_(), null, 2)); }
