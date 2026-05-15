/* 🌍 지구지킴이 에너지 탐험대 — 공통 설정 + 헬퍼
 *
 * 모든 페이지(index, lesson1~7)에서 공유합니다.
 * 학생 정보·진행 기록·통계는 Google 스프레드시트와 연동됩니다.
 *
 * ─────────────────────────────────────────────
 * 📌 다른 선생님께서 사용하실 때 — config.js 손댈 필요 없어요!
 *
 *    선생님은 본인 시트의 [선생님 설정 안내] 시트를 보고 "우리 반 전용 링크"를 만듭니다.
 *    예) https://○○○.github.io/?api=https://script.google.com/macros/.../exec
 *
 *    학생이 그 긴 링크로 한 번 들어오면, 그 컴퓨터·브라우저는 우리 반 시트를
 *    자동으로 기억합니다. 다음부터는 짧은 사이트 주소로만 들어와도 우리 반 시트로 기록돼요.
 *    (자세한 안내는 사본 스프레드시트의 첫 번째 탭 참고)
 * ─────────────────────────────────────────────
 */

window.APP_CONFIG = {
  // 제(만든이)의 기본 URL — 다른 선생님이 ?api=... 링크로 따로 연결 안 하면 이 시트로 데이터가 가요.
  API_URL_DEFAULT: 'https://script.google.com/macros/s/AKfycbwBjQR6qVTInT1Sh9k8OeUuWB0qRAEA6p-cNIcEyaqmORqY02xPe8GyjfTZtTt67o0D/exec',

  // ?api=... 로 한 번 들어온 URL을 저장해 두는 브라우저 저장소 키
  STORAGE_KEY: 'teacher-api-url'
};

/* API_URL 은 동적으로 결정됩니다.
 *  우선순위:  ① 주소창 ?api=...  →  ② 브라우저에 저장된 선생님 URL  →  ③ 기본 URL
 *  (이렇게 해 두면 기존 코드의 window.APP_CONFIG.API_URL 호출이 그대로 작동해요)
 */
Object.defineProperty(window.APP_CONFIG, 'API_URL', {
  get() {
    // ① 주소 뒤에 ?api=... 가 붙어 있으면 그걸 자동 저장
    try {
      const fromParam = new URLSearchParams(location.search).get('api');
      if (fromParam) {
        const u = decodeURIComponent(fromParam).trim();
        if (u) {
          localStorage.setItem(window.APP_CONFIG.STORAGE_KEY, u);
          return u;
        }
      }
    } catch (e) { /* 무시 */ }
    // ② 브라우저에 저장된 선생님 URL
    const fromStorage = (function () {
      try { return localStorage.getItem(window.APP_CONFIG.STORAGE_KEY); }
      catch (e) { return null; }
    })();
    if (fromStorage) return fromStorage;
    // ③ 기본 URL (제 스프레드시트)
    return window.APP_CONFIG.API_URL_DEFAULT || '';
  }
});

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
