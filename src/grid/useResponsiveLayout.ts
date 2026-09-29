import { useEffect, useState } from 'react'

const MOBILE_BREAKPOINT = '(max-width: 768px)'

/** 768px 미만이면 모바일 카드뷰, 아니면 데스크톱 그리드. */
export function useResponsiveLayout(): 'mobile' | 'desktop' {
  const [isMobile, setIsMobile] = useState(
    () => typeof window !== 'undefined' && window.matchMedia(MOBILE_BREAKPOINT).matches,
  )

  useEffect(() => {
    // matchMedia의 'change' 리스너는 일부 환경(예: CDP 기반 뷰포트 에뮬레이션)에서
    // 실제로 폭이 바뀌어도 발화하지 않는 경우가 있어(fresh query는 정확한데
    // 이벤트만 안 옴), window의 'resize'에 얹어 매번 새로 질의하는 쪽이 더 안전하다.
    const recompute = () => setIsMobile(window.matchMedia(MOBILE_BREAKPOINT).matches)
    window.addEventListener('resize', recompute)
    return () => window.removeEventListener('resize', recompute)
  }, [])

  return isMobile ? 'mobile' : 'desktop'
}
