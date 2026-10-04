import { useEffect, useState } from 'react'
import { resetTextMetrics } from '../xlsx/textMetrics'

/** 글꼴 파일이 연달아 도착할 때마다 다시 재지 않고, 잠잠해진 뒤 한 번만 다시 재기 위한 대기(ms). */
const FONTS_SETTLE_MS = 120

/**
 * 웹폰트 로딩이 끝나면 바뀌는 번호. 글자 폭·줄 수를 잰 값(행 높이, 글자 넘침)은 글꼴이 도착하기 전에는
 * 대체 글꼴 기준이라 틀리므로, 이 번호를 의존성으로 걸어 두면 글꼴이 도착한 뒤 다시 계산된다.
 *
 * @remarks 캐시된 글자 폭도 여기서 비운다(resetTextMetrics) — 안 비우면 다시 계산해도 옛 글꼴의 폭이 나온다.
 * Pretendard는 글자 범위별 조각으로 쪼개져 있어서, 파일에 처음 나오는 글자가 있으면 그 조각이 그때 도착한다.
 */
export function useFontsVersion(): number {
  const [version, setVersion] = useState(0)

  useEffect(() => {
    const fonts = typeof document === 'undefined' ? undefined : document.fonts
    if (!fonts) return
    let timer: ReturnType<typeof setTimeout> | undefined
    const settle = () => {
      clearTimeout(timer)
      timer = setTimeout(() => {
        resetTextMetrics()
        setVersion((v) => v + 1)
      }, FONTS_SETTLE_MS)
    }
    fonts.addEventListener('loadingdone', settle)
    // 이미 로딩 중이면 끝나는 때에도 한 번 다시 잰다. 이미 다 도착해 있다면 할 일이 없다(불필요한 재계산 방지).
    if (fonts.status === 'loading') void fonts.ready.then(settle)
    return () => {
      clearTimeout(timer)
      fonts.removeEventListener('loadingdone', settle)
    }
  }, [])

  return version
}
