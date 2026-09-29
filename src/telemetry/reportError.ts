/**
 * 못 잡은 에러를 우리 도메인의 /report-error(Cloudflare Pages Function, M6)로
 * 보낸다. 그 함수가 withvibe API 키를 서버 쪽에서만 들고 중계한다. 파일
 * 내용·셀 값·사용자 입력은 여기 전혀 안 들어간다 — 에러 메시지와 현재 경로뿐.
 * 로컬 개발(`vite dev`)에서는 그 라우트 자체가 없어서 설치하지 않는다 —
 * 개발 중 에러가 실제 사용자 리포트로 잘못 잡히는 걸 막는다.
 */
export function installErrorReporting() {
  if (!import.meta.env.PROD) return

  window.addEventListener('error', (event) => {
    reportError(event.error instanceof Error ? event.error.message : String(event.message))
  })
  window.addEventListener('unhandledrejection', (event) => {
    const reason = event.reason as unknown
    reportError(reason instanceof Error ? reason.message : String(reason))
  })
}

/** ErrorBoundary처럼, 렌더링 중 React가 직접 잡은 에러를 수동으로 보고할 때 쓴다. */
export function reportError(message: string) {
  if (!import.meta.env.PROD) return
  const payload = JSON.stringify({ message, page: window.location.pathname })

  try {
    if (navigator.sendBeacon('/report-error', new Blob([payload], { type: 'application/json' }))) {
      return
    }
  } catch {
    // sendBeacon이 없거나 실패하면 fetch로 한 번 더 시도
  }

  fetch('/report-error', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: payload,
    keepalive: true,
  }).catch(() => {
    // 텔레메트리 실패는 사용자에게 보여줄 이유가 없다 — best-effort.
  })
}
