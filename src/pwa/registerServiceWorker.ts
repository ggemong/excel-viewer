/**
 * public/sw.js 등록 — 홈 화면에 설치됐을 때 네이티브 앱처럼 동작하게(오프라인에서도
 * 이미 연 적 있으면 shell이 뜸) 한다. telemetry/reportError.ts의 installErrorReporting()과
 * 같은 이유로 PROD 빌드에서만 등록한다 — 개발 서버(`vite dev`)에서 등록하면 HMR로 받는
 * 최신 코드 대신 캐시된 옛 버전이 뜨는 혼란을 겪기 쉽다.
 */
export function registerServiceWorker() {
  if (!import.meta.env.PROD) return
  if (!('serviceWorker' in navigator)) return

  window.addEventListener('load', () => {
    navigator.serviceWorker.register('/sw.js').catch((err) => {
      console.error('서비스워커 등록 실패:', err)
    })
  })
}
