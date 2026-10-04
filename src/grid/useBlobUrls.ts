import { useCallback, useEffect, useState } from 'react'

/**
 * Blob 목록에 대한 object URL을 만들어 주고, 목록이 바뀌거나 컴포넌트가 사라지면 반드시
 * 해제한다(해제하지 않으면 시트를 오갈 때마다 이미지 메모리가 누수된다).
 *
 * URL 생성을 useMemo가 아니라 effect 안에서 하는 이유: React StrictMode(개발)는 effect를
 * 마운트->정리->재마운트로 두 번 돌리는데, useMemo로 만든 URL을 정리 단계에서 해제하면
 * 재마운트 후에는 이미 해제된 URL을 계속 쓰게 되어 이미지가 깨진다. effect 안에서 만들면
 * 재마운트 때 새로 만들어진다.
 *
 * @param blobs 참조가 안정적인(useMemo 등) 배열이어야 한다 — 매 렌더 새 배열이면 URL을 계속 다시 만든다.
 * @returns blob -> URL 조회 함수. 아직 만들어지기 전이면 undefined.
 */
export function useBlobUrls(blobs: Blob[]): (blob: Blob) => string | undefined {
  const [urls, setUrls] = useState<Map<Blob, string>>(() => new Map())

  useEffect(() => {
    const created = new Map<Blob, string>()
    for (const blob of new Set(blobs)) created.set(blob, URL.createObjectURL(blob))
    // 외부 자원(object URL) 동기화라 effect 안 setState가 맞다 — 위 설명처럼 렌더 중에 만들면 StrictMode에서 깨진다.
    // oxlint-disable-next-line react/set-state-in-effect
    setUrls(created)
    return () => {
      for (const url of created.values()) URL.revokeObjectURL(url)
    }
  }, [blobs])

  return useCallback((blob: Blob) => urls.get(blob), [urls])
}
