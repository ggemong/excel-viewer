/**
 * Cloudflare Pages Function — POST /report-error.
 *
 * withvibe 허브의 에러 리포트 API(docs: home-withvibe/docs/telemetry-api.md)는
 * API 키를 서버 쪽에서만 호출하도록 못 박아뒀다. 이 앱은 완전 클라이언트
 * 전용이 원칙이라 브라우저가 키를 들고 있을 수 없어서, 이 작은 중계
 * 함수 하나만 서버 역할을 한다 — 파일 내용은 절대 이 함수를 거치지
 * 않는다(에러 메시지 + 현재 경로만 받는다).
 *
 * WITHVIBE_API_KEY는 Cloudflare Pages 프로젝트의 secret으로 넣는다:
 *   wrangler pages secret put WITHVIBE_API_KEY
 */

interface Env {
  WITHVIBE_API_KEY: string
}

interface ReportBody {
  message?: unknown
  page?: unknown
}

export const onRequestPost: PagesFunction<Env> = async (context) => {
  const { request, env } = context

  if (!env.WITHVIBE_API_KEY) {
    return new Response('Not configured', { status: 501 })
  }

  let body: ReportBody
  try {
    body = await request.json()
  } catch {
    return new Response('Bad Request', { status: 400 })
  }

  if (typeof body.message !== 'string' || typeof body.page !== 'string') {
    return new Response('Bad Request', { status: 400 })
  }

  try {
    const upstream = await fetch('https://www.withvibe.kr/api/v1/events/', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        Authorization: `Bearer ${env.WITHVIBE_API_KEY}`,
      },
      body: JSON.stringify({
        project: 'excel-viewer',
        type: 'error',
        page: body.page.slice(0, 300),
        message: body.message.slice(0, 500),
      }),
    })
    return new Response(null, { status: upstream.status })
  } catch {
    return new Response('Upstream unreachable', { status: 502 })
  }
}
