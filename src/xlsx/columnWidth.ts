/**
 * 엑셀 열너비(문자폭 단위)/행높이(포인트)를 이 앱이 실제로 쓰는 px로 바꾼다.
 * Excel과 픽셀 단위로 똑같아지는 건 애초에 불가능하다(다른 OS/폰트 렌더링) — "비율이
 * 맞아 보이는 정도"가 현실적 목표다.
 */

/** 1pt = 1/72인치, CSS px는 96dpi 기준 — 고정 공식이라 폰트와 무관하다. */
export function excelPointsToPx(points: number): number {
  return (points * 96) / 72
}

/**
 * ECMA-376 공식: px = trunc(((256*charWidth + trunc(128/MDW)) / 256) * MDW).
 * MDW(Max Digit Width, 기본 폰트 숫자 '0'의 폭)는 보통 Calibri 11 기준 7로
 * 고정해서 쓰는데, 이 앱은 Pretendard를 쓰고 실사용 파일은 맑은 고딕이 기본인
 * 경우가 많아 그 값대로 고정하면 또 틀어진다 — 하드코딩 대신 이 앱이 실제
 * 그리드 렌더링에 쓰는 폰트/크기(--font, --fs-sm 토큰)로 캔버스에서 직접 잰다.
 */
export function excelColumnWidthToPx(charWidth: number, mdw: number): number {
  return Math.trunc(((256 * charWidth + Math.trunc(128 / mdw)) / 256) * mdw)
}

let cachedMdw: number | null = null

/**
 * 그리드 셀과 같은 폰트/크기로 숫자 '0' 하나의 폭을 잰다(MDW) — CSS 커스텀
 * 프로퍼티(--font/--fs-sm)를 그대로 읽어서, 토큰이 바뀌어도 여기 값을 따로
 * 맞춰줄 필요가 없게 한다. 캔버스를 못 쓰는 환경(SSR 등)이면 Calibri 11의
 * 통상적인 근사값(7)으로 대신한다 — 이 앱은 브라우저 전용이라 실제로는 항상
 * 캔버스 경로를 탄다.
 */
export function measureDefaultFontWidth(): number {
  if (cachedMdw !== null) return cachedMdw
  if (typeof document === 'undefined') return 7

  const canvas = document.createElement('canvas')
  const ctx = canvas.getContext('2d')
  if (!ctx) return 7

  const rootStyle = getComputedStyle(document.documentElement)
  const fontFamily = rootStyle.getPropertyValue('--font').trim() || 'sans-serif'
  const fontSize = rootStyle.getPropertyValue('--fs-sm').trim() || '14px'

  ctx.font = `${fontSize} ${fontFamily}`
  const width = ctx.measureText('0').width
  cachedMdw = width > 0 ? width : 7
  return cachedMdw
}
