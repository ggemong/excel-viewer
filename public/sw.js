/*
 * excel-viewer 서비스워커 — withvibe 허브(home-withvibe/public/sw.js)와 같은 원칙으로
 * 라이브러리 없이 직접 작성했지만, 이 앱은 Vite로 빌드되는 순수 SPA라 두 가지를
 * 다르게 가져간다:
 *
 * 1) 이 앱은 index.html 자체가 곧 앱 전체다(페이지 이동이 없음) — 그래서 허브처럼
 *    내비게이션을 네트워크 전용으로 두면 오프라인에서 아예 못 연다. 대신 네트워크
 *    우선 + 성공하면 캐시에 저장 + 실패(오프라인)하면 캐시에서 꺼내는 방식을 쓴다.
 *    이렇게 하면 캐시된 index.html과 그 시점의 해시 자산들이 항상 같은
 *    CACHE_VERSION 아래 함께 저장돼서, 배포 후 "캐시된 shell이 가리키는 자산은
 *    이미 사라짐" 같은 불일치가 안 생긴다.
 * 2) Vite는 public/*를 빌드 때 그대로 복사할 뿐이라 여기에 해시 파일명을 미리
 *    적어 둘 수 없다 — 그래서 설치 시점에 정적 자산 목록을 캐시해 두는 대신,
 *    /assets/ 밑 요청이 처음 성공할 때마다 그 자리에서 캐시에 채워 넣는다
 *    (파일명에 내용 해시가 있어 내용이 바뀌면 주소도 바뀌므로 이 방식이 안전하다).
 *
 * CACHE_VERSION은 빌드마다 vite.config.ts의 injectSwCacheVersion 플러그인이
 * 자동으로 채운다(수동으로 올리다 깜빡하는 걸 피하려고) — 개발 서버(`vite dev`)에서는
 * 치환 전 플레이스홀더 문자열 그대로 남는데, 서비스워커는 프로덕션 빌드에서만
 * 등록하므로(src/pwa/registerServiceWorker.ts) 개발 중엔 이 값이 실제로 안 쓰인다.
 */
const CACHE_VERSION = '__SW_CACHE_VERSION__'
const CACHE_NAME = `excel-viewer-${CACHE_VERSION}`
const STATIC_PREFIXES = ['/assets/', '/icons/']

self.addEventListener('install', (event) => {
  event.waitUntil(self.skipWaiting())
})

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches
      .keys()
      .then((keys) => Promise.all(keys.filter((k) => k.startsWith('excel-viewer-') && k !== CACHE_NAME).map((k) => caches.delete(k))))
      .then(() => self.clients.claim()),
  )
})

self.addEventListener('fetch', (event) => {
  const { request } = event
  if (request.method !== 'GET') return
  const url = new URL(request.url)
  if (url.origin !== self.location.origin) return

  if (request.mode === 'navigate') {
    // 네트워크 우선 — 새 배포가 있으면 항상 최신 shell을 받고, 오프라인일 때만
    // 마지막으로 성공했던 shell로 대신한다(비행기 모드에서도 이미 연 적 있으면 열림).
    event.respondWith(
      fetch(request)
        .then((res) => {
          if (res.ok) {
            const copy = res.clone()
            caches.open(CACHE_NAME).then((cache) => cache.put(request, copy))
          }
          return res
        })
        .catch(() => caches.match(request).then((cached) => cached || caches.match('/'))),
    )
    return
  }

  if (STATIC_PREFIXES.some((p) => url.pathname.startsWith(p))) {
    event.respondWith(
      caches.match(request).then(
        (cached) =>
          cached ||
          fetch(request).then((res) => {
            if (res.ok) {
              const copy = res.clone()
              caches.open(CACHE_NAME).then((cache) => cache.put(request, copy))
            }
            return res
          }),
      ),
    )
  }
})
