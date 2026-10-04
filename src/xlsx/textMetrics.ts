/**
 * 칸이 실제로 그리는 글꼴 기준의 글자 크기·폭 측정 — 열너비(MDW), 줄바꿈 행높이 맞춤, 글자 넘침이 전부
 * "화면에 그려질 크기 그대로" 재야 해서 한 곳에 모았다. 글꼴·기본 크기는 CSS 토큰(--font, --fs-sm)을 읽어
 * 오기 때문에 토큰이 바뀌어도 여기를 따로 맞출 필요가 없다.
 *
 * 경계: 이 파일은 "잰다"까지만 한다. 어떤 칸을 어떻게 키우고 넘치게 할지는 src/grid/rowFit.ts / textSpill.ts 몫이다.
 * columnWidth.ts(열너비의 기준 폭 MDW)도 같은 글꼴을 재므로 이 도구를 쓴다.
 */
import type { CellStyle } from './cellStyle'

/** 토큰을 못 읽는 환경(CSS가 없는 테스트·SSR)에서만 쓰는 값 — 브라우저에서는 항상 토큰을 읽는다. */
const FALLBACK_FONT_FAMILY = 'sans-serif'
const FALLBACK_BASE_FONT_PX = 14
/** 캐시가 이 개수를 넘으면 비운다(아주 큰 시트에서 글 하나하나가 쌓여 메모리가 불어나는 걸 막는다). */
const MAX_CACHE_ENTRIES = 20000
const FONT_FAMILY_TOKEN = '--font'
const BASE_FONT_SIZE_TOKEN = '--fs-sm'

/** 글자 크기·굵기에 영향을 주는 서식만 — 이 셋이 같으면 같은 폭으로 잰다. */
export type TextFace = Pick<CellStyle, 'fontScale' | 'bold' | 'italic'>

interface MeasureEnv {
  family: string
  basePx: number
  ctx: CanvasRenderingContext2D | null
}

let cachedEnv: MeasureEnv | null = null
let widthCache = new Map<string, number>()

function readEnv(): MeasureEnv {
  if (cachedEnv) return cachedEnv
  let family = FALLBACK_FONT_FAMILY
  let basePx = FALLBACK_BASE_FONT_PX
  let ctx: CanvasRenderingContext2D | null = null
  if (typeof document !== 'undefined') {
    const rootStyle = getComputedStyle(document.documentElement)
    family = rootStyle.getPropertyValue(FONT_FAMILY_TOKEN).trim() || family
    const parsed = Number.parseFloat(rootStyle.getPropertyValue(BASE_FONT_SIZE_TOKEN))
    if (Number.isFinite(parsed) && parsed > 0) basePx = parsed
    ctx = document.createElement('canvas').getContext('2d')
  }
  cachedEnv = { family, basePx, ctx }
  return cachedEnv
}

/** 칸의 기본 글자 크기(px) — CSS 토큰 --fs-sm. */
export function baseFontPx(): number {
  return readEnv().basePx
}

/** 서식이 반영된 글자 크기(px). */
export function fontPxOf(face: TextFace | null): number {
  return baseFontPx() * (face?.fontScale ?? 1)
}

/** canvas `font` 속성에 넣는 문자열(굵기·기울임·크기·글꼴 계열). */
export function canvasFontOf(face: TextFace | null): string {
  const { family } = readEnv()
  return `${face?.italic ? 'italic ' : ''}${face?.bold ? '700 ' : ''}${fontPxOf(face)}px ${family}`
}

/**
 * 글 한 줄의 폭(px). 캔버스를 못 쓰는 환경이면 null — 호출부가 "모른다"로 다루게 한다(눈대중 폭을 만들어
 * 넘치기·줄 수 판단에 쓰면 조용히 틀린다).
 */
export function textWidth(text: string, face: TextFace | null): number | null {
  return measureWithFont(canvasFontOf(face), text)
}

/**
 * 칸 서식이 아닌 임의 크기(px)로 쓴 글 한 줄의 폭 — 차트의 제목·축·범례 글자처럼 칸과 따로 크기가 정해진 글을 잴 때.
 * 글꼴 계열은 칸과 같다. 캔버스를 못 쓰면 null.
 */
export function textWidthAt(text: string, fontPx: number, bold = false): number | null {
  return measureWithFont(`${bold ? '700 ' : ''}${fontPx}px ${readEnv().family}`, text)
}

function measureWithFont(font: string, text: string): number | null {
  const { ctx } = readEnv()
  if (!ctx) return null
  const key = `${font}\u0000${text}`
  const hit = widthCache.get(key)
  if (hit !== undefined) return hit
  ctx.font = font
  const width = ctx.measureText(text).width
  if (widthCache.size >= MAX_CACHE_ENTRIES) widthCache.clear()
  widthCache.set(key, width)
  return width
}

/**
 * 측정 결과를 버리고 CSS 토큰도 다시 읽는다. 웹폰트가 늦게 도착하면 그 전에 잰 폭은 대체 글꼴 기준이라
 * 틀리므로, 글꼴 로딩이 끝날 때마다 부른다.
 */
export function resetTextMetrics(): void {
  cachedEnv = null
  widthCache = new Map()
}
