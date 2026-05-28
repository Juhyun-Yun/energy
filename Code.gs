/**
 * 🌍 지구지킴이 에너지 탐험대 — Google Apps Script 백엔드
 *
 * 📌 다른 선생님께서 쓰실 때 (3단계만!)
 *
 *   1. 이 스프레드시트의 [파일] → [사본 만들기] (본인 구글 드라이브에 복사)
 *   2. 사본의 [확장 프로그램] → [Apps Script] → [배포] → [새 배포] → 웹 앱
 *        - 액세스: "모든 사용자"
 *        - 배포 후 나오는 "웹 앱 URL" 복사
 *   3. 학습 사이트 (만든이의 GitHub 주소) 우측 상단 [🧑‍🏫] 버튼 → URL 붙여넣기
 *
 *   ✅ config.js 같은 파일을 직접 수정하실 필요 없어요!
 *   ✅ "학생명단" 시트에 우리 반 학생을 입력하시고, 나머지 시트는 자동 생성됩니다.
 *
 *   📖 자세한 안내는 사본 첫 번째 탭의 "🧑‍🏫 선생님 설정 안내" 시트를 보세요.
 */

// 💡 Apps Script를 스프레드시트의 "확장 프로그램 > Apps Script"로 열었다면
//    아래 SHEET_ID를 비워둬도 자동으로 그 스프레드시트를 사용합니다.
//    (다른 시트의 데이터를 쓰고 싶을 때만 ID를 넣어주세요)
const SHEET_ID = '';

const STUDENTS_SHEET = '학생명단';
const RECORDS_SHEET  = '진행기록';
const WORKS_SHEET    = '에너지마을작품';
const COMMENTS_SHEET = '에너지마을댓글';
const GUIDE_SHEET    = '🧑‍🏫 선생님 설정 안내';
const WORKS_FOLDER   = '에너지 마을 작품';

/* ============ 메뉴 (스프레드시트 열 때 자동 생성) ============ */
function onOpen() {
  // 안내 시트가 없으면 자동으로 만들고, 있으면 현재 상태만 갱신
  try { ensureGuideSheet_(); } catch (e) { /* 권한 문제 등은 무시 */ }

  SpreadsheetApp.getUi()
    .createMenu('📊 학습 대시보드')
    .addItem('📡 실시간 현황 (사이드바)', 'showSidebar')
    .addItem('📈 전체 현황 (큰 팝업)', 'showDashboard')
    .addSeparator()
    .addItem('🔑 작품 업로드 권한 승인하기', 'authorizeDrive')
    .addItem('📖 선생님 설정 안내 다시 만들기', 'rebuildGuideSheet')
    .addItem('🔄 진행기록 시트 초기화', 'resetRecords')
    .addToUi();
}

/* 학생이 작품을 업로드하려면 Drive 권한이 필요해요.
 * 이 함수는 메뉴에서도, 편집기에서도 실행할 수 있지만
 * 권한 다이얼로그는 "편집기에서 직접 실행"해야 안정적으로 떠요. */
function authorizeDrive() {
  const ui = SpreadsheetApp.getUi();

  // 1단계: Drive 접근 권한 확인
  try {
    DriveApp.getRootFolder();
  } catch (e) {
    ui.alert(
      '🔑 Drive 권한 승인이 필요해요',
      '【중요】 메뉴 클릭만으로는 권한 창이 안 뜰 때가 많아요.\n' +
      '아래 방법으로 "편집기에서 직접 실행"해 주세요.\n\n' +
      '━━━━━━━━━━━━━━━━━━━\n' +
      '1️⃣  메뉴: [확장 프로그램] → [Apps Script]\n' +
      '2️⃣  좌측 파일 목록의 "Code.gs" 클릭\n' +
      '3️⃣  상단 함수 드롭다운에서 "authorizeDrive" 선택\n' +
      '4️⃣  ▶ "실행" 버튼 클릭\n' +
      '5️⃣  뜨는 권한 요청 창에서:\n' +
      '       → 본인 Google 계정 선택\n' +
      '       → "고급" 클릭\n' +
      '       → "(안전하지 않음)..." 링크 클릭\n' +
      '       → 모든 권한 "허용"\n' +
      '━━━━━━━━━━━━━━━━━━━\n\n' +
      '【필수】 권한 승인 후 반드시 새 배포!\n' +
      '편집기 우상단 [배포] → [배포 관리]\n' +
      '→ 기존 배포 옆 ✏️(편집) → 버전 "새 버전" 선택 → [배포]\n' +
      '(URL은 그대로 유지됨)\n\n' +
      '오류 원본: ' + (e.message || e),
      ui.ButtonSet.OK
    );
    return;
  }

  // 2단계: 작품 폴더 생성/조회
  let folder;
  try {
    folder = getWorksFolder_();
  } catch (e) {
    ui.alert(
      '⚠️ 폴더 접근 실패',
      'Drive 권한은 있는데 폴더 생성에서 실패했어요.\n오류: ' + (e.message || e),
      ui.ButtonSet.OK
    );
    return;
  }

  // 3단계: 매니페스트(appsscript.json) 확인
  let manifestNote = '';
  try {
    const mf = JSON.parse(DriveApp.getFileById(ScriptApp.getScriptId()).getBlob().getDataAsString());
    const scopes = (mf && mf.oauthScopes) || [];
    const hasDrive = scopes.some(s => /\/auth\/drive$|\/auth\/drive\.readonly$/.test(s));
    if (!hasDrive && scopes.length > 0) {
      manifestNote = '\n\n⚠️ 매니페스트에 Drive 권한이 빠져 있어요. ' +
        '편집기 ⚙️ 설정에서 매니페스트를 보이게 한 뒤, oauthScopes에 ' +
        '"https://www.googleapis.com/auth/drive"를 추가하고 새 배포하세요.';
    }
  } catch (e) { /* 매니페스트 자체를 못 읽으면 무시 */ }

  ui.alert(
    '✅ 권한 승인 완료',
    '작품 업로드 권한이 정상이에요.\n작품 저장 폴더: ' + folder.getName() +
    '\n\n⚠️ 그래도 학생이 업로드 실패한다면 새 배포가 필요할 수 있어요:\n' +
    '편집기 [배포] → [배포 관리] → 기존 배포 옆 ✏️\n' +
    '→ 버전 "새 버전" 선택 → [배포]' +
    manifestNote,
    ui.ButtonSet.OK
  );
}

/* 메뉴에서 안내 시트를 강제로 다시 만들고 싶을 때 */
function rebuildGuideSheet() {
  ensureGuideSheet_(true);
  SpreadsheetApp.getUi().alert('완료', '"' + GUIDE_SHEET + '" 시트를 새로 만들었어요.', SpreadsheetApp.getUi().ButtonSet.OK);
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
    // POST 요청: 진행 기록 또는 에너지 마을 액션
    if (e && e.postData && e.postData.contents) {
      const data = JSON.parse(e.postData.contents);
      const type = data && data.type;
      if (type === 'upload-work') return uploadWork_(data);
      if (type === 'comment')     return saveComment_(data);
      // 기본: 진행 기록
      recordEvent_(data);
      return { ok: true };
    }
    // GET 요청: 조회
    const action = e && e.parameter && e.parameter.action;
    if (action === 'students') return { ok: true, students: getStudents_() };
    if (action === 'progress') return { ok: true, progress: getProgress_(e.parameter.student) };
    if (action === 'stats')    return { ok: true, stats:    getStats_() };
    if (action === 'gallery')  return { ok: true, works:    getWorks_() };
    if (action === 'comments') return { ok: true, comments: getComments_(e.parameter.workId) };
    if (action === 'mywork')   return { ok: true, work:     getMyWork_(e.parameter.student) };
    if (action === 'ping')     return { ok: true, message: 'pong' };
    return { ok: false, error: 'unknown_action' };
  } catch (err) {
    const msg = String(err && err.message || err);
    // Drive 권한 미승인 오류를 친근한 메시지로 변환
    if (/DriveApp|drive\.readonly|googleapis\.com\/auth\/drive/i.test(msg)) {
      return {
        ok: false,
        errorType: 'drive_not_authorized',
        error: '선생님의 Drive 권한 승인이 필요해요.\n\n선생님께 알려주세요:\n1. 우리 반 스프레드시트 열기\n2. 메뉴 → 📊 학습 대시보드 → 🔑 작품 업로드 권한 승인하기\n3. (이미 했다면) Apps Script에서 새 배포(새 버전)로 다시 배포'
      };
    }
    return { ok: false, error: msg };
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

  // 에너지 마을 작품·댓글 통계
  const village = getVillageStats_();

  return {
    totalStudents: students.length,
    totalRecords: records.length,
    perMission,
    perStudent,
    village
  };
}

/* ============ 에너지 마을: 작품 & 댓글 ============ */

function ensureWorksSheet_() {
  const ss = ss_();
  let sh = ss.getSheetByName(WORKS_SHEET);
  if (!sh) {
    sh = ss.insertSheet(WORKS_SHEET);
    sh.appendRow(['시간', '학생', '작품ID', '버전', '이미지URL', '드라이브파일ID', '설명', '자기평가(JSON)', '한줄소감']);
    sh.setFrozenRows(1);
    sh.setColumnWidth(1, 160); sh.setColumnWidth(2, 90);  sh.setColumnWidth(3, 160);
    sh.setColumnWidth(4, 50);  sh.setColumnWidth(5, 320); sh.setColumnWidth(6, 200);
    sh.setColumnWidth(7, 320); sh.setColumnWidth(8, 240); sh.setColumnWidth(9, 240);
  }
  return sh;
}

function ensureCommentsSheet_() {
  const ss = ss_();
  let sh = ss.getSheetByName(COMMENTS_SHEET);
  if (!sh) {
    sh = ss.insertSheet(COMMENTS_SHEET);
    sh.appendRow(['시간', '작품ID', '작품학생', '댓글학생', '동료평가(JSON)', '한줄평']);
    sh.setFrozenRows(1);
    sh.setColumnWidth(1, 160); sh.setColumnWidth(2, 160); sh.setColumnWidth(3, 90);
    sh.setColumnWidth(4, 90);  sh.setColumnWidth(5, 240); sh.setColumnWidth(6, 320);
  }
  return sh;
}

function getWorksFolder_() {
  const it = DriveApp.getFoldersByName(WORKS_FOLDER);
  if (it.hasNext()) return it.next();
  const folder = DriveApp.createFolder(WORKS_FOLDER);
  folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return folder;
}

/* 작품 업로드 / 수정 재업로드 (같은 작품ID로 들어오면 버전 증가) */
function uploadWork_(data) {
  const student = String(data.student || '').trim();
  if (!student) return { ok: false, error: 'no_student' };
  if (!data.imageBase64) return { ok: false, error: 'no_image' };

  // base64 → Drive 업로드
  const m = String(data.imageBase64).match(/^data:(.+?);base64,(.*)$/);
  const mime = m ? m[1] : 'image/jpeg';
  const b64  = m ? m[2] : String(data.imageBase64);
  const bytes = Utilities.base64Decode(b64);
  const ext = mime.indexOf('png') >= 0 ? 'png' : 'jpg';
  const folder = getWorksFolder_();
  const fileName = student + '_' + new Date().getTime() + '.' + ext;
  const blob = Utilities.newBlob(bytes, mime, fileName);
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  const fileId = file.getId();
  const url = 'https://drive.google.com/thumbnail?id=' + fileId + '&sz=w1200';

  // 작품ID 결정 — 처음이면 새로 생성, 수정이면 기존 ID 유지
  const sh = ensureWorksSheet_();
  const last = sh.getLastRow();
  let workId = String(data.workId || '').trim();
  let version = 1;
  if (workId) {
    // 기존 ID가 있으면 최신 버전 찾기
    if (last >= 2) {
      const rows = sh.getRange(2, 1, last - 1, 9).getValues();
      rows.forEach(r => {
        if (String(r[2]) === workId) {
          const v = Number(r[3]);
          if (!isNaN(v) && v >= version) version = v + 1;
        }
      });
    }
  } else {
    workId = 'W' + new Date().getTime() + '_' + Math.floor(Math.random() * 1000);
  }

  sh.appendRow([
    new Date(),
    student,
    workId,
    version,
    url,
    fileId,
    String(data.description || ''),
    JSON.stringify(data.selfEval || {}),
    String(data.reflection || '')
  ]);

  // 진행기록에도 표시 (미션 7 완료)
  ensureRecordsSheet_().appendRow([
    new Date(), student, '7', '완료', '',
    JSON.stringify({ workId, version, kind: 'energy-village' })
  ]);

  return { ok: true, workId, version, url };
}

/* 갤러리 조회 — 작품ID별 최신 버전만 반환 */
function getWorks_() {
  const sh = ensureWorksSheet_();
  const last = sh.getLastRow();
  if (last < 2) return [];
  const rows = sh.getRange(2, 1, last - 1, 9).getValues();
  const latest = {};
  rows.forEach(r => {
    const workId = String(r[2]); if (!workId) return;
    const ver = Number(r[3]) || 1;
    const cur = latest[workId];
    if (!cur || ver > cur.version) {
      latest[workId] = {
        time:        r[0] instanceof Date ? r[0].toISOString() : String(r[0]),
        student:     String(r[1]),
        workId:      workId,
        version:     ver,
        url:         String(r[4]),
        description: String(r[6]),
        selfEval:    safeParse_(r[7]),
        reflection:  String(r[8])
      };
    }
  });
  return Object.values(latest)
    .sort((a, b) => new Date(b.time) - new Date(a.time));
}

/* 내 작품 찾기 (학생명 기준 최신) */
function getMyWork_(student) {
  if (!student) return null;
  const works = getWorks_();
  const target = String(student).trim();
  return works.find(w => w.student === target) || null;
}

/* 댓글 저장 */
function saveComment_(data) {
  const sh = ensureCommentsSheet_();
  sh.appendRow([
    new Date(),
    String(data.workId || ''),
    String(data.workStudent || ''),
    String(data.commenter || ''),
    JSON.stringify(data.peerEval || {}),
    String(data.text || '')
  ]);
  return { ok: true };
}

/* 작품별 댓글 조회 */
function getComments_(workId) {
  if (!workId) return [];
  const sh = ensureCommentsSheet_();
  const last = sh.getLastRow();
  if (last < 2) return [];
  const rows = sh.getRange(2, 1, last - 1, 6).getValues();
  const target = String(workId);
  return rows
    .filter(r => String(r[1]) === target)
    .map(r => ({
      time:      r[0] instanceof Date ? r[0].toISOString() : String(r[0]),
      commenter: String(r[3]),
      peerEval:  safeParse_(r[4]),
      text:      String(r[5])
    }))
    .sort((a, b) => new Date(a.time) - new Date(b.time));
}

function getVillageStats_() {
  const works = getWorks_();
  const sh = ensureCommentsSheet_();
  const last = sh.getLastRow();
  const totalComments = Math.max(0, last - 1);
  return {
    totalWorks: works.length,
    totalComments,
    works
  };
}

function safeParse_(v) {
  try { return v ? JSON.parse(v) : {}; } catch (e) { return {}; }
}

/* ============================================================
 * 🧑‍🏫 선생님 설정 안내 시트 — 사본 만든 선생님이 바로 보고 따라할 수 있도록
 *    스프레드시트 안에 자동으로 만들어주는 시각적 가이드 시트입니다.
 *
 *  • onOpen() 에서 자동 호출 — 시트가 없으면 만들고, 있으면 그대로 둠
 *  • 메뉴의 "📖 선생님 설정 안내 다시 만들기" 로 강제 재생성 가능
 *  • "현재 상태" 칸들은 수식으로 들어가 있어서 학생이 활동하면 자동 갱신됨
 * ============================================================ */
function ensureGuideSheet_(forceRebuild) {
  const ss = ss_();
  let sh = ss.getSheetByName(GUIDE_SHEET);
  if (sh && !forceRebuild) return sh;
  if (sh) ss.deleteSheet(sh);

  sh = ss.insertSheet(GUIDE_SHEET, 0);  // 가장 첫 번째 탭으로
  sh.setHiddenGridlines(true);
  sh.setColumnWidth(1, 30);    // 왼쪽 여백
  sh.setColumnWidth(2, 90);    // 단계/아이콘
  sh.setColumnWidth(3, 240);   // 항목 제목
  sh.setColumnWidth(4, 560);   // 자세한 설명
  sh.setColumnWidth(5, 30);    // 오른쪽 여백

  let r = 1;

  /* ---------- 큰 제목 배너 ---------- */
  sh.getRange(r, 1, 1, 5).merge()
    .setValue('🌍 지구지킴이 에너지 탐험대 — 선생님 설정 안내')
    .setFontSize(20).setFontWeight('bold')
    .setBackground('#16a34a').setFontColor('#ffffff')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.setRowHeight(r++, 56);

  /* ---------- 서브 배너 ---------- */
  sh.getRange(r, 1, 1, 5).merge()
    .setValue('🙋‍♀️ 만든이의 학습 사이트 (GitHub) 는 그대로 쓰시고, 데이터(우리 반 시트)만 본인 걸로 연결하시면 돼요.\n아래 3단계만 따라하면 약 5분이면 끝납니다!')
    .setFontSize(11).setBackground('#dcfce7').setFontColor('#14532d')
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setWrap(true);
  sh.setRowHeight(r++, 56);

  // 빈 줄
  sh.setRowHeight(r++, 12);

  /* ---------- 현재 상태 섹션 ---------- */
  sh.getRange(r, 2, 1, 3).merge()
    .setValue('📊 우리 반 현재 상태  —  학생이 활동하면 실시간으로 자동 갱신돼요')
    .setFontSize(13).setFontWeight('bold')
    .setBackground('#e0f2fe').setFontColor('#0369a1')
    .setHorizontalAlignment('left').setVerticalAlignment('middle');
  sh.setRowHeight(r++, 32);

  const statusRows = [
    ['👥', '학생 명단 (' + STUDENTS_SHEET + ' 시트)',
      '=IFERROR(IF(COUNTA(\'' + STUDENTS_SHEET + '\'!B2:B)=0,"⚠️ 아직 학생을 입력하지 않았어요",COUNTA(\'' + STUDENTS_SHEET + '\'!B2:B)&"명 입력됨 ✓"),"⚠️ \'' + STUDENTS_SHEET + '\' 시트가 없어요")'],
    ['📝', '쌓인 학습 기록 (' + RECORDS_SHEET + ' 시트)',
      '=IFERROR(IF(COUNTA(\'' + RECORDS_SHEET + '\'!A2:A)=0,"아직 0건 (학생이 활동하면 자동 생성)",COUNTA(\'' + RECORDS_SHEET + '\'!A2:A)&"건 ✓"),"아직 0건 (학생이 활동하면 자동 생성)")'],
    ['🎨', '에너지 마을 작품 (' + WORKS_SHEET + ' 시트)',
      '=IFERROR(IF(COUNTA(\'' + WORKS_SHEET + '\'!A2:A)=0,"아직 0점 (학생이 업로드하면 자동 생성)",COUNTA(\'' + WORKS_SHEET + '\'!A2:A)&"점 ✓"),"아직 0점 (학생이 업로드하면 자동 생성)")'],
    ['💬', '서로 남긴 댓글 (' + COMMENTS_SHEET + ' 시트)',
      '=IFERROR(IF(COUNTA(\'' + COMMENTS_SHEET + '\'!A2:A)=0,"아직 0개",COUNTA(\'' + COMMENTS_SHEET + '\'!A2:A)&"개 ✓"),"아직 0개")']
  ];
  statusRows.forEach(arr => {
    const [icon, label, formula] = arr;
    sh.getRange(r, 2).setValue(icon).setFontSize(16)
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
    sh.getRange(r, 3).setValue(label).setFontSize(11)
      .setVerticalAlignment('middle');
    sh.getRange(r, 4).setFormula(formula).setFontSize(11).setFontWeight('bold')
      .setFontColor('#0f766e').setVerticalAlignment('middle');
    sh.setRowHeight(r++, 30);
  });

  // 빈 줄
  sh.setRowHeight(r++, 16);

  /* ---------- 3단계 설정 섹션 ---------- */
  sh.getRange(r, 2, 1, 3).merge()
    .setValue('🛠️ 딱 3단계만! 우리 반 시트와 학습 사이트 연결하기')
    .setFontSize(13).setFontWeight('bold')
    .setBackground('#fef3c7').setFontColor('#92400e')
    .setHorizontalAlignment('left').setVerticalAlignment('middle');
  sh.setRowHeight(r++, 32);

  const steps = [
    {
      num: '1단계',
      title: '학생 명단 입력하기  (지금 이 스프레드시트에서)',
      body:
        '아래쪽 시트 탭에서 "학생명단" 시트를 클릭하세요.\n\n' +
        '• A1 칸 = "번호",  B1 칸 = "이름"  → 이 두 글자는 절대 바꾸지 마세요!\n' +
        '   (프로그램이 헤더로 인식해요)\n' +
        '• A2 칸부터 1, 2, 3, ... 번호를 적고\n' +
        '• B2 칸부터 우리 반 학생 이름을 한 명씩 적어주세요.\n\n' +
        '예)   A2 = 1     B2 = 김민준\n        A3 = 2     B3 = 이서연\n        A4 = 3     B4 = 박지호  ...\n\n' +
        '✅ 다 적은 뒤 위쪽 "현재 상태"에서 학생 수를 확인!'
    },
    {
      num: '2단계',
      title: 'Apps Script에서 웹 앱 URL 받기  ★가장 중요해요!',
      body:
        '이 스프레드시트 위쪽 메뉴에서:\n' +
        '       [확장 프로그램]  →  [Apps Script]\n' +
        '를 클릭하세요. 새 탭에 코드 편집기가 열립니다.\n\n' +
        '편집기 우측 상단 파란색 [배포] 버튼을 누르세요:\n' +
        '       [배포]  →  [새 배포]  →  톱니바퀴(⚙️) 클릭  →  "웹 앱" 선택\n\n' +
        '아래 항목을 정확히 채워주세요:\n' +
        '   • 설명: 자유 (예: "5학년 자원과 에너지")\n' +
        '   • 다음 사용자 인증 정보로 실행: 나 (선생님 본인 계정)\n' +
        '   • 액세스 권한: 모든 사용자  ← 꼭 이걸로!\n\n' +
        '[배포] 클릭 → 처음에는 "권한 검토" 창이 떠요. 놀라지 마세요!\n' +
        '   ① 본인 구글 계정 선택\n' +
        '   ② "Google에서 확인하지 않은 앱" 경고 → 왼쪽 아래 [고급] 클릭\n' +
        '   ③ "안전하지 않음으로 이동(Unsafe)" 클릭\n' +
        '   ④ [허용] 클릭\n\n' +
        '※ 권한 목록에 다음 항목들이 나와요. 본인 사본이 쓰는 권한이라 안전합니다.\n' +
        '   • Google 시트  → 학습 기록·작품 정보 저장용\n' +
        '   • Google 드라이브  → 학생 작품 사진 저장용 (선생님 본인 드라이브의 "에너지 마을 작품" 폴더)\n' +
        '   • 외부 서비스 연결  → 학생 페이지 ↔ 시트 통신용\n\n' +
        '🎉 배포 끝나면 "웹 앱 URL"이 나와요:\n' +
        '       https://script.google.com/macros/s/AKfyc........./exec\n\n' +
        '   → [복사] 버튼을 누르세요. 3단계에서 씁니다!'
    },
    {
      num: '3단계',
      title: '학생들에게 줄 "우리 반 전용 링크" 만들기',
      body:
        '학생들에게 학습 사이트 주소만 알려드리면 데이터가 만든이의 시트로 가버려요.\n' +
        '우리 반 시트로 보내려면 사이트 주소 끝에 2단계의 URL을 붙여서 알려주세요.\n\n' +
        '📋 만드는 법 (그냥 두 주소를 이어 붙이면 끝):\n\n' +
        '   원래 사이트 주소 :  https://○○○.github.io/\n' +
        '   2단계의 URL :     https://script.google.com/macros/.../exec\n\n' +
        '   이어 붙인 주소 :  https://○○○.github.io/?api=https://script.google.com/macros/.../exec\n' +
        '                                         ↑ ?api= 만 사이에 끼워 넣으면 돼요\n\n' +
        '✅ 이 긴 주소를 학생들에게 한 번만 알려주세요!\n' +
        '   학생이 그 링크로 한 번 들어가면 그 컴퓨터·브라우저가 우리 반 시트를\n' +
        '   자동으로 기억해요. 다음부턴 짧은 사이트 주소만 들어가도 우리 반에 기록됩니다.\n\n' +
        '📌 학생들에게 전달하는 팁\n' +
        '   • 학교 컴퓨터마다 한 번씩 그 긴 링크로 열어두면 학생들은 그냥 짧은 주소로 들어와도 OK\n' +
        '   • 링크가 너무 길면 네이버 me, 비틀리(bit.ly) 같은 단축 URL로 만들어서 주셔도 좋아요\n' +
        '   • QR 코드로 만들어 칠판에 띄워주시는 것도 추천!\n\n' +
        '🧪 직접 테스트해 보고 싶으면 선생님이 그 긴 링크로 한 번 들어가서\n' +
        '    학생 이름이 우리 반 명단으로 뜨는지 확인해 보세요.'
    }
  ];

  steps.forEach(s => {
    sh.getRange(r, 2)
      .setValue(s.num).setFontSize(14).setFontWeight('bold')
      .setBackground('#fbbf24').setFontColor('#7c2d12')
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
    sh.getRange(r, 3)
      .setValue(s.title).setFontSize(13).setFontWeight('bold')
      .setBackground('#fffbeb').setFontColor('#78350f')
      .setVerticalAlignment('middle').setWrap(true);
    sh.getRange(r, 4)
      .setValue(s.body).setFontSize(11)
      .setBackground('#fffbeb').setFontColor('#1f2937')
      .setVerticalAlignment('top').setWrap(true);
    // 줄 수에 따라 행 높이 자동 계산 (한 줄 약 18px + 여유)
    const lines = s.body.split('\n').length;
    sh.setRowHeight(r++, Math.max(80, lines * 18 + 16));
  });

  // 빈 줄
  sh.setRowHeight(r++, 16);

  /* ---------- 자주 묻는 질문 섹션 ---------- */
  sh.getRange(r, 2, 1, 3).merge()
    .setValue('❓ 자주 묻는 질문 (Q&A)')
    .setFontSize(13).setFontWeight('bold')
    .setBackground('#fae8ff').setFontColor('#86198f')
    .setHorizontalAlignment('left').setVerticalAlignment('middle');
  sh.setRowHeight(r++, 32);

  const qas = [
    ['학생이 활동했는데 우리 반 시트에 안 쌓여요',
     '학생이 들어온 주소가 우리 반 전용 링크 (?api=...가 붙은 긴 주소)였는지 확인하세요.\n그 링크 없이 짧은 주소만으로 처음 들어오면 만든이의 기본 시트로 데이터가 가요.\n그 컴퓨터에서 한 번만 긴 링크로 들어오면 다음부턴 짧은 주소도 우리 반으로 자동 연결됩니다.'],
    ['학생 명단을 바꿨는데 학습 화면에 반영이 안 돼요',
     '학생 컴퓨터에서 브라우저를 새로고침(F5) 하면 다시 불러와요.\n캐시가 남아 있으면 Ctrl + F5 (강제 새로고침)을 누르세요.'],
    ['"권한이 거부되었습니다" 오류가 나요',
     '2단계 배포에서 "액세스 권한"이 "모든 사용자"가 아닌 경우예요.\nApps Script에서 [배포] → [배포 관리]를 열어 액세스 권한을 다시 확인해 주세요.'],
    ['Code.gs(이 코드)를 수정했는데 학생 화면에 반영이 안 돼요',
     '저장만으로는 안 되고 "재배포"가 필요해요. 단, 새 배포를 또 만들면 URL이 바뀌어서 학생들에게 새 링크를 다시 알려야 해요.\n→ [배포] → [배포 관리] → 연필(✏️) 아이콘 → 버전 "새 버전" 선택 → [배포]\n   이렇게 하면 URL이 그대로 유지돼서 학생 링크를 안 바꿔도 됩니다!'],
    ['학생 작품 사진은 어디에 저장되나요?',
     '선생님 본인의 구글 드라이브 안 "에너지 마을 작품" 폴더에 자동 저장돼요.\n시트에는 사진의 링크만 들어가요. (시트가 무거워지지 않게 하려고요)'],
    ['학년이 끝나서 진행기록을 한꺼번에 지우고 싶어요',
     '위쪽 [📊 학습 대시보드] 메뉴 → [🔄 진행기록 시트 초기화] 클릭.\n(작품/댓글은 따로 보관되니, 필요 없으면 시트 탭에서 직접 우클릭 → 삭제)'],
    ['학생이 한 번에 한 명씩만 활동할 수 있나요?',
     '아니에요! 여러 학생이 동시에 활동해도 모두 따로따로 기록돼요.\n반 전체가 동시에 사용해도 괜찮아요.'],
    ['긴 링크가 너무 길어서 학생들에게 알려주기 힘들어요',
     '세 가지 방법 추천:\n   • 네이버 me, bit.ly 같은 단축 URL 서비스로 짧게 만들어 주세요.\n   • 그 링크로 QR 코드를 만들어 칠판이나 모니터에 띄워 학생이 휴대폰·태블릿으로 찍게 하세요.\n   • 학교 컴퓨터마다 선생님이 한 번씩 그 링크로 열어두면, 학생들은 그냥 짧은 주소만 알려주셔도 돼요.']
  ];

  qas.forEach(arr => {
    const [q, a] = arr;
    sh.getRange(r, 2).setValue('Q&A').setFontSize(11).setFontWeight('bold')
      .setBackground('#f5d0fe').setFontColor('#86198f')
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
    const text = 'Q. ' + q + '\nA. ' + a;
    sh.getRange(r, 3, 1, 2).merge()
      .setValue(text).setFontSize(11)
      .setBackground('#fdf4ff').setFontColor('#1f2937')
      .setVerticalAlignment('top').setWrap(true);
    const lines = text.split('\n').length;
    sh.setRowHeight(r++, Math.max(50, lines * 18 + 14));
  });

  // 빈 줄
  sh.setRowHeight(r++, 16);

  /* ---------- 문제 해결 섹션 ---------- */
  sh.getRange(r, 2, 1, 3).merge()
    .setValue('🆘 문제가 생겼을 때  —  가장 흔한 4가지')
    .setFontSize(13).setFontWeight('bold')
    .setBackground('#fee2e2').setFontColor('#991b1b')
    .setHorizontalAlignment('left').setVerticalAlignment('middle');
  sh.setRowHeight(r++, 32);

  const troubles = [
    ['학습 화면에서 학생 이름이 안 떠요',
     '"학생명단" 시트를 열어 B열 이름이 비어있지 않은지 확인하세요.\n시트 이름이 정확히 "학생명단"인지도 확인 (공백·오타 X).'],
    ['우리 반 시트에 진행기록이 안 쌓여요 (= 다른 분 시트로 가요)',
     '학생들이 우리 반 전용 링크 (?api=...가 붙은 긴 주소)로 한 번이라도 들어왔는지 확인하세요.\n그 링크 없이 짧은 사이트 주소만으로 처음 접속한 컴퓨터는 만든이 시트로 데이터가 가요.\n해결: 그 컴퓨터에서 한 번만 긴 링크로 들어와 주세요. 그러면 그 컴퓨터·브라우저가 우리 반 시트를 자동으로 기억합니다.'],
    ['작품 업로드(7차시)가 실패해요',
     '2단계 배포에서 "액세스 권한"이 "모든 사용자"인지 다시 확인하세요.\nApps Script [배포] → [배포 관리]에서 변경 가능합니다.'],
    ['시트가 갑자기 사라졌어요',
     '시트 이름을 임의로 바꾸지 마세요.\n   (학생명단 / 진행기록 / 에너지마을작품 / 에너지마을댓글 — 이 4개 이름은 고정)\n실수로 삭제했다면 좌측 하단 "시트 탭"에서 [+] 눌러 같은 이름으로 다시 만들면\n프로그램이 자동으로 헤더를 채워줍니다.']
  ];

  troubles.forEach(arr => {
    const [problem, fix] = arr;
    sh.getRange(r, 2).setValue('🔧').setFontSize(16)
      .setBackground('#fecaca').setFontColor('#991b1b')
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
    const text = '❗ ' + problem + '\n→ ' + fix;
    sh.getRange(r, 3, 1, 2).merge()
      .setValue(text).setFontSize(11)
      .setBackground('#fef2f2').setFontColor('#1f2937')
      .setVerticalAlignment('top').setWrap(true);
    const lines = text.split('\n').length;
    sh.setRowHeight(r++, Math.max(54, lines * 18 + 14));
  });

  // 빈 줄
  sh.setRowHeight(r++, 16);

  /* ---------- 마무리 안내 ---------- */
  sh.getRange(r, 1, 1, 5).merge()
    .setValue('🎉 설정이 끝나면 이 시트는 그대로 두셔도 되고, 시트 탭을 우클릭해서 숨기셔도 돼요.\n다시 보고 싶을 땐 [📊 학습 대시보드] 메뉴 → [📖 선생님 설정 안내 다시 만들기] 를 누르세요.')
    .setFontSize(11).setBackground('#f1f5f9').setFontColor('#334155')
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setWrap(true);
  sh.setRowHeight(r++, 50);

  // 안내 시트를 활성화 (사본 만들고 처음 열었을 때 바로 보이도록)
  ss.setActiveSheet(sh);
  sh.getRange('A1').activate();

  return sh;
}

/* ============ 테스트용 (Apps Script 에디터에서 직접 실행 가능) ============ */
function testGetStudents() { Logger.log(getStudents_()); }
function testGetStats()    { Logger.log(JSON.stringify(getStats_(), null, 2)); }
function testBuildGuide()  { ensureGuideSheet_(true); }
