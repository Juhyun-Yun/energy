/* 🌍 지구지킴이 에너지 탐험대 — 공통 설정 + 헬퍼
 *
 * 모든 페이지(index, lesson1~7)에서 공유합니다.
 * 학생 정보·진행 기록·통계는 Google 스프레드시트와 연동됩니다.
 *
 * 설정 방법은 Code.gs 파일을 참고하세요.
 */

window.APP_CONFIG = {
  // ⬇️⬇️⬇️  여기에 본인의 Google Apps Script 웹앱 URL을 붙여넣으세요  ⬇️⬇️⬇️
  //
  //   1) Google Drive에서 새 스프레드시트 만들기
  //   2) 첫 시트 이름을 "학생명단"으로 바꾸고:
  //        A1 = 번호,  B1 = 이름    (헤더)
  //        A2~ = 1, 2, 3, ...     B2~ = 학생 이름들
  //   3) 확장 프로그램 > Apps Script 열기
  //   4) Code.gs 내용을 통째로 붙여넣고 SHEET_ID 채우기
  //   5) 배포 > 새 배포 > 웹 앱  (액세스: 모든 사용자) → URL 복사
  //   6) 아래 따옴표 사이에 그 URL을 붙여넣기
  //
  //   비워두면 자동으로 오프라인 모드(브라우저 LocalStorage만)로 동작합니다.
  API_URL: 'https://script.google.com/macros/s/AKfycbwBjQR6qVTInT1Sh9k8OeUuWB0qRAEA6p-cNIcEyaqmORqY02xPe8GyjfTZtTt67o0D/exec'
};

window.appHelpers = {
  /* ===== 현재 학생 (세션 저장소 사용 — 브라우저 탭 닫으면 자동 로그아웃) ===== */
  getStudent() {
    return sessionStorage.getItem('app-student') || '';
  },
  setStudent(name) {
    if (name && name.trim()) sessionStorage.setItem('app-student', name.trim());
    else sessionStorage.removeItem('app-student');
  },
  clearStudent() {
    sessionStorage.removeItem('app-student');
    // 진행 상황도 함께 정리 (다음 학생이 잘못된 데이터 안 보도록)
    localStorage.removeItem('energy-explorer-v1');
  },

  /* ===== API 연결 여부 ===== */
  hasApi() {
    return !!(window.APP_CONFIG && window.APP_CONFIG.API_URL);
  },

  /* ===== GET 호출 (학생 목록·진행·통계) ===== */
  async apiGet(action, params = {}) {
    if (!this.hasApi()) throw new Error('API URL이 설정되지 않았습니다.');
    const url = new URL(window.APP_CONFIG.API_URL);
    url.searchParams.set('action', action);
    Object.entries(params).forEach(([k, v]) => url.searchParams.set(k, v));
    const ctrl = new AbortController();
    const timer = setTimeout(() => ctrl.abort(), 8000);
    try {
      const res = await fetch(url.toString(), { signal: ctrl.signal });
      const data = await res.json();
      clearTimeout(timer);
      return data;
    } catch (e) {
      clearTimeout(timer);
      throw e;
    }
  },

  /* ===== 내부 페이지 이동 표시 =====
   * 같은 탭 안에서 페이지가 바뀔 때만 학생 세션을 이어갑니다.
   * 외부 진입(새 탭, 닫고 다시 열기 등)은 강제 재선택.
   */
  markInternalNav() {
    try { sessionStorage.setItem('app-internal-nav', '1'); } catch {}
  },

  /* ===== 진행 기록 전송 (POST, 백그라운드) =====
   * navigator.sendBeacon으로 fire-and-forget 전송 → 페이지 이동 직전에도 안전.
   */
  reportProgress(missionId, detail = {}) {
    const student = this.getStudent();
    const url = window.APP_CONFIG && window.APP_CONFIG.API_URL;
    if (!student || !url) return false;
    const payload = JSON.stringify({
      student,
      mission: String(missionId),
      status: '완료',
      ...detail,
      timestamp: new Date().toISOString()
    });
    try {
      const blob = new Blob([payload], { type: 'text/plain;charset=utf-8' });
      if (navigator.sendBeacon && navigator.sendBeacon(url, blob)) return true;
      // fallback
      fetch(url, {
        method: 'POST',
        body: payload,
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        keepalive: true
      }).catch(() => { });
      return true;
    } catch (e) {
      console.warn('reportProgress failed', e);
      return false;
    }
  }
};

/* 페이지를 떠날 때(다른 페이지로 이동·새로고침) 자동으로 "내부 이동" 표시.
 * 탭이 닫히면 sessionStorage 자체가 사라지므로 이 플래그도 함께 사라집니다. */
window.addEventListener('pagehide', () => {
  try { sessionStorage.setItem('app-internal-nav', '1'); } catch {}
});
