/**
 * 셀 서식(배경색·글자색·굵게/기울임) 추출 + 화면에 입히는 변환 — "엑셀처럼 보여주기"
 * 위한 표시 전용 정보다. 이 프로젝트는 읽기 전용 뷰어라 저장 경로 자체가 없다(D-006) —
 * 여기서 읽은 값은 오직 렌더링에만 쓰인다.
 *
 * v1 범위: 단색 배경(fill pattern "solid")과 글자색/굵게/기울임만 다룬다. 테두리·정렬·
 * 조건부서식은 범위 밖이다 — formatValue.ts가 numFmt를 "자주 쓰이는 것만" 다루는
 * 것과 같은 절충이다. 색상은 직접 RGB(argb)와 테마 색상(theme+tint, src/xlsx/
 * themeColor.ts) 둘 다 다룬다 — 실사용 파일 대부분이 테마 색상이라 이것만 빠지면
 * 색이 거의 안 보인다.
 *
 * cellStyleProps는 Grid.tsx가 쓴다 — "배경은 선택/비교 강조색에 밀리고 글자색/굵기는
 * 유지한다"는 규칙을 셀 렌더링 코드에 직접 섞지 않고 여기 모아뒀다.
 */
import type { CSSProperties } from 'react'
import { resolveThemeColor, type ThemeColors } from './themeColor'

export interface CellStyle {
  bg: string | null
  color: string | null
  bold: boolean
  italic: boolean
}

interface ExcelColor {
  argb?: string
  theme?: number
  tint?: number
}

interface CellFill {
  type?: string
  pattern?: string
  fgColor?: ExcelColor
}

interface CellFont {
  bold?: boolean
  italic?: boolean
  color?: ExcelColor
}

/** ExcelJS의 ARGB(8자리 hex, 앞 2자리는 투명도) -> CSS #RRGGBB. */
function argbToHex(argb: string | undefined): string | null {
  if (!argb || argb.length < 6) return null
  return `#${argb.slice(-6)}`
}

/** 직접 RGB(argb)를 우선 쓰고, 없으면 테마 인덱스(theme+tint)로 해석한다. 둘 다 없으면 null. */
function resolveColor(color: ExcelColor | undefined, theme: ThemeColors | null): string | null {
  if (!color) return null
  const direct = argbToHex(color.argb)
  if (direct) return direct
  if (theme && color.theme !== undefined) return resolveThemeColor(theme, color.theme, color.tint)
  return null
}

function luminance(hex: string): number {
  const channel = (i: number) => parseInt(hex.slice(i, i + 2), 16) / 255
  const linear = (v: number) => (v <= 0.03928 ? v / 12.92 : ((v + 0.055) / 1.055) ** 2.4)
  return 0.2126 * linear(channel(1)) + 0.7152 * linear(channel(3)) + 0.0722 * linear(channel(5))
}

/**
 * 배경색에 맞춰 읽히는 글자색을 고른다(WCAG 상대 휘도 기준) — 파일이 글자색을 따로
 * 지정 안 했는데 배경만 밝으면, 앱이 다크 테마라고 밝은 기본 글자색을 그대로 쓰면
 * 밝은 배경 위에 밝은 글씨가 돼서 안 보인다. 앱 테마와 무관하게 항상 읽히게 한다.
 */
function contrastColor(bgHex: string): string {
  return luminance(bgHex) > 0.5 ? '#1a1a1a' : '#f5f5f5'
}

/**
 * @param cell ExcelJS Cell — fill/font만 쓰므로 그 둘만 받는 최소 타입으로 받는다(테스트에서 실제 Cell 없이도 검증 가능).
 * @param theme 워크북 하나당 한 번만 파싱해서 넘겨받는다(src/xlsx/read.ts) — 셀마다 테마 XML을 다시 파싱하지 않기 위해서.
 */
export function extractCellStyle(cell: { fill?: CellFill; font?: CellFont }, theme: ThemeColors | null = null): CellStyle | null {
  const fill = cell.fill
  const bg = fill?.type === 'pattern' && fill.pattern === 'solid' ? resolveColor(fill.fgColor, theme) : null

  const font = cell.font
  const explicitColor = resolveColor(font?.color, theme)
  const color = explicitColor ?? (bg ? contrastColor(bg) : null)
  const bold = Boolean(font?.bold)
  const italic = Boolean(font?.italic)

  if (!bg && !color && !bold && !italic) return null
  return { bg, color, bold, italic }
}

/**
 * 파일 서식을 인라인 스타일로 바꾼다. `suppressBg`가 true면(선택 중이거나 비교 모드로
 * 다른 점이 있을 때) 배경은 비워서 CSS의 선택/비교 강조색이 보이게 하고, 글자색/굵기는
 * 그대로 유지한다 — 강조 중에도 원래 어떤 글자가 꾸며져 있었는지는 알아볼 수 있게.
 */
export function cellStyleProps(style: CellStyle | null, suppressBg: boolean): CSSProperties {
  if (!style) return {}
  return {
    background: suppressBg ? undefined : (style.bg ?? undefined),
    color: style.color ?? undefined,
    fontWeight: style.bold ? 700 : undefined,
    fontStyle: style.italic ? 'italic' : undefined,
  }
}
