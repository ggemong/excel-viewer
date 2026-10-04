/**
 * 셀 서식(배경색·글자색·굵게/기울임·테두리·정렬) 추출 + 화면에 입히는 변환 —
 * "엑셀처럼 보여주기" 위한 표시 전용 정보다. 이 프로젝트는 읽기 전용 뷰어라
 * 저장 경로 자체가 없다(D-006) — 여기서 읽은 값은 오직 렌더링에만 쓰인다.
 *
 * v1 범위: 단색 배경(fill pattern "solid")과 글자색/굵게/기울임/테두리/정렬·
 * 줄바꿈만 다룬다. 조건부서식은 범위 밖이다 — formatValue.ts가 numFmt를
 * "자주 쓰이는 것만" 다루는 것과 같은 절충이다. 색상(배경·글자·테두리 전부)은
 * 직접 RGB(argb)와 테마 색상(theme+tint, src/xlsx/themeColor.ts) 둘 다 다룬다
 * — 실사용 파일 대부분이 테마 색상이라 이것만 빠지면 색이 거의 안 보인다.
 *
 * cellStyleProps는 Grid.tsx가 쓴다 — "배경은 선택/비교 강조색에 밀리고 글자색/굵기는
 * 유지한다"는 규칙을 셀 렌더링 코드에 직접 섞지 않고 여기 모아뒀다.
 */
import type { CSSProperties } from 'react'
import type { CfStyle } from './conditionalFormat'
import { resolveThemeColor, type ThemeColors } from './themeColor'

export interface CellBorderSide {
  width: string
  style: string
  color: string
}

export interface CellBorder {
  top: CellBorderSide | null
  right: CellBorderSide | null
  bottom: CellBorderSide | null
  left: CellBorderSide | null
}

export interface CellAlign {
  h: 'left' | 'center' | 'right' | null
  v: 'top' | 'middle' | 'bottom' | null
  wrap: boolean
}

export interface CellStyle {
  bg: string | null
  color: string | null
  bold: boolean
  italic: boolean
  border: CellBorder | null
  align: CellAlign | null
  /** 취소선/밑줄 — 있을 때만 true로 둔다(없으면 키 자체가 없음). */
  strike?: true
  underline?: true
  /**
   * 글자색이 파일에 명시된 값인가(true) 아니면 배경에서 자동 계산한 대비색인가(없음). 조건부서식이
   * 배경만 바꿀 때, 명시된 글자색은 지키고 자동 계산된 색만 새 배경에 맞춰 다시 고르기 위해 구분한다.
   */
  colorExplicit?: true
}

export interface ExcelColor {
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
  strike?: boolean
  underline?: boolean | string
  color?: ExcelColor
}

interface ExcelBorderSide {
  style?: string
  color?: ExcelColor
}

interface CellBorderModel {
  top?: ExcelBorderSide
  right?: ExcelBorderSide
  bottom?: ExcelBorderSide
  left?: ExcelBorderSide
}

interface CellAlignment {
  horizontal?: string
  vertical?: string
  wrapText?: boolean
}

/** ExcelJS의 ARGB(8자리 hex, 앞 2자리는 투명도) -> CSS #RRGGBB. */
function argbToHex(argb: string | undefined): string | null {
  if (!argb || argb.length < 6) return null
  return `#${argb.slice(-6)}`
}

/** 직접 RGB(argb)를 우선 쓰고, 없으면 테마 인덱스(theme+tint)로 해석한다. 둘 다 없으면 null. */
export function resolveColor(color: ExcelColor | undefined, theme: ThemeColors | null): string | null {
  if (!color) return null
  const direct = argbToHex(color.argb)
  if (direct) return direct
  if (theme && color.theme !== undefined) return resolveThemeColor(theme, color.theme, color.tint)
  return null
}

/**
 * ExcelJS/Excel는 셀에 테두리·채우기 등 아무 서식 하나라도 쓰이면, 사용자가 글자색을
 * 건드린 적 없어도 `font.color = {theme: 1}`(dk1, 보정 없음)인 기본 폰트 레코드를
 * XML에 같이 써버린다(직접 round-trip으로 확인됨) — Excel 자신의 캔버스가 항상 밝은
 * 배경이라 "자동" 검정 글자가 늘 멀쩡히 보여서 드러나지 않던 값이다. 이 앱은 실제
 * 다크 테마가 있어서 그대로 믿으면 어두운 배경 위에 검정 글자가 그대로 깔려 안 보인다.
 * tint 보정 없는 순수 dk1은 "사용자가 고른 색"과 "그냥 기본값"을 API로는 구분할 수
 * 없으므로, 이 경우만 명시적 지정이 아니라고 보고 null로 돌려 앱의 테마 적응형 기본
 * 글자색(CSS)이 대신 적용되게 한다.
 */
function isAutomaticDefaultFontColor(color: ExcelColor | undefined): boolean {
  return color !== undefined && color.argb === undefined && color.theme === 1 && !color.tint
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
export function contrastColor(bgHex: string): string {
  return luminance(bgHex) > 0.5 ? '#1a1a1a' : '#f5f5f5'
}

/** Excel 테두리 스타일 -> CSS 두께. 안 쓰는 값(hair 등 가는 변형)은 가장 가까운 걸로 뭉뚱그린다. */
const BORDER_WIDTH: Record<string, string> = {
  hair: '1px',
  thin: '1px',
  dashed: '1px',
  dotted: '1px',
  dashDot: '1px',
  dashDotDot: '1px',
  slantDashDot: '1px',
  medium: '2px',
  mediumDashed: '2px',
  mediumDashDot: '2px',
  mediumDashDotDot: '2px',
  thick: '3px',
  double: '3px',
}
/** Excel 테두리 스타일 -> CSS border-style. CSS엔 없는 "점선+파선 섞임" 계열은 dashed로 근사. */
const BORDER_CSS_STYLE: Record<string, string> = {
  dotted: 'dotted',
  dashed: 'dashed',
  mediumDashed: 'dashed',
  dashDot: 'dashed',
  dashDotDot: 'dashed',
  slantDashDot: 'dashed',
  mediumDashDot: 'dashed',
  mediumDashDotDot: 'dashed',
  double: 'double',
}

function extractBorderSide(side: ExcelBorderSide | undefined, theme: ThemeColors | null): CellBorderSide | null {
  if (!side?.style) return null
  return {
    width: BORDER_WIDTH[side.style] ?? '1px',
    style: BORDER_CSS_STYLE[side.style] ?? 'solid',
    // 파일이 테두리 색을 안 정하면 Excel은 검정("자동")으로 그리는데, 이 앱은 어두운
    // 테마라 그대로 쓰면 배경에 묻혀 안 보인다 — 그럴 때만 테마에 맞는 중립 회색으로 대신한다.
    color: resolveColor(side.color, theme) ?? 'var(--line-strong)',
  }
}

const ALIGN_H: Record<string, CellAlign['h']> = { left: 'left', center: 'center', right: 'right' }
const ALIGN_V: Record<string, CellAlign['v']> = { top: 'top', middle: 'middle', bottom: 'bottom' }

/**
 * @param cell ExcelJS Cell — fill/font/border/alignment만 쓰므로 그것만 받는 최소 타입으로 받는다(테스트에서 실제 Cell 없이도 검증 가능).
 * @param theme 워크북 하나당 한 번만 파싱해서 넘겨받는다(src/xlsx/read.ts) — 셀마다 테마 XML을 다시 파싱하지 않기 위해서.
 */
export function extractCellStyle(
  cell: { fill?: CellFill; font?: CellFont; border?: CellBorderModel; alignment?: CellAlignment },
  theme: ThemeColors | null = null,
): CellStyle | null {
  const fill = cell.fill
  const bg = fill?.type === 'pattern' && fill.pattern === 'solid' ? resolveColor(fill.fgColor, theme) : null

  const font = cell.font
  const explicitColor = isAutomaticDefaultFontColor(font?.color) ? null : resolveColor(font?.color, theme)
  const color = explicitColor ?? (bg ? contrastColor(bg) : null)
  const bold = Boolean(font?.bold)
  const italic = Boolean(font?.italic)
  const strike = Boolean(font?.strike)
  const underline = Boolean(font?.underline) && font?.underline !== 'none'

  const borderModel = cell.border
  const border = borderModel
    ? {
        top: extractBorderSide(borderModel.top, theme),
        right: extractBorderSide(borderModel.right, theme),
        bottom: extractBorderSide(borderModel.bottom, theme),
        left: extractBorderSide(borderModel.left, theme),
      }
    : null
  const hasBorder = border && (border.top || border.right || border.bottom || border.left)

  const alignment = cell.alignment
  const align = alignment
    ? {
        h: alignment.horizontal ? (ALIGN_H[alignment.horizontal] ?? null) : null,
        v: alignment.vertical ? (ALIGN_V[alignment.vertical] ?? null) : null,
        wrap: Boolean(alignment.wrapText),
      }
    : null
  const hasAlign = align && (align.h || align.v || align.wrap)

  if (!bg && !color && !bold && !italic && !strike && !underline && !hasBorder && !hasAlign) return null
  return {
    bg,
    color,
    bold,
    italic,
    border: hasBorder ? border : null,
    align: hasAlign ? align : null,
    ...(strike ? { strike: true as const } : {}),
    ...(underline ? { underline: true as const } : {}),
    ...(explicitColor ? { colorExplicit: true as const } : {}),
  }
}

const JUSTIFY_CONTENT: Record<NonNullable<CellAlign['h']>, CSSProperties['justifyContent']> = {
  left: 'flex-start',
  center: 'center',
  right: 'flex-end',
}
const ALIGN_ITEMS: Record<NonNullable<CellAlign['v']>, CSSProperties['alignItems']> = {
  top: 'flex-start',
  middle: 'center',
  bottom: 'flex-end',
}

/**
 * 파일 서식을 인라인 스타일로 바꾼다. `suppressBg`가 true면(선택 중이거나 비교 모드로
 * 다른 점이 있을 때) 배경은 비워서 CSS의 선택/비교 강조색이 보이게 하고, 글자색/굵기는
 * 그대로 유지한다 — 강조 중에도 원래 어떤 글자가 꾸며져 있었는지는 알아볼 수 있게.
 *
 * @remarks 정렬(justifyContent/alignItems)·줄바꿈(whiteSpace) 키는 파일에 실제
 * 지정이 있을 때만 결과 객체에 넣는다 — `undefined`로라도 넣으면, Grid.tsx가 이
 * 객체를 스프레드해서 만드는 style에서 그 앞에 미리 넣어둔 숫자형 추측 정렬값을
 * 덮어써 버린다(스프레드는 값이 undefined여도 키가 있으면 덮어씀).
 */
export function cellStyleProps(style: CellStyle | null, suppressBg: boolean): CSSProperties {
  if (!style) return {}
  const props: CSSProperties = {
    background: suppressBg ? undefined : (style.bg ?? undefined),
    color: style.color ?? undefined,
    fontWeight: style.bold ? 700 : undefined,
    fontStyle: style.italic ? 'italic' : undefined,
    textDecoration: [style.strike && 'line-through', style.underline && 'underline'].filter(Boolean).join(' ') || undefined,
  }

  if (style.align?.h) props.justifyContent = JUSTIFY_CONTENT[style.align.h]
  if (style.align?.v) props.alignItems = ALIGN_ITEMS[style.align.v]
  if (style.align?.wrap) props.whiteSpace = 'normal'

  const b = style.border
  if (b?.top) props.borderTop = `${b.top.width} ${b.top.style} ${b.top.color}`
  if (b?.right) props.borderRight = `${b.right.width} ${b.right.style} ${b.right.color}`
  if (b?.bottom) props.borderBottom = `${b.bottom.width} ${b.bottom.style} ${b.bottom.color}`
  if (b?.left) props.borderLeft = `${b.left.width} ${b.left.style} ${b.left.color}`

  return props
}

/**
 * 조건부서식 결과(CfStyle)를 셀 서식 위에 덮어쓴다. 조건부서식이 정한 속성만 바뀌고 나머지는 셀 서식이 그대로 남는다.
 *
 * 배경이 바뀌는데 글자색은 안 정해졌으면: 파일이 명시한 글자색은 지키고, 배경에서 자동으로 고른 색(또는
 * 색 없음)이었다면 새 배경에 맞춰 다시 고른다 — 안 그러면 연한 새 배경 위에 앱의 밝은 기본 글자색이 얹혀
 * 읽히지 않는다(다크 테마).
 */
export function applyConditionalStyle(base: CellStyle | null, cf: CfStyle): CellStyle {
  const bg = cf.bg ?? base?.bg ?? null
  let color = cf.color ?? null
  let colorExplicit = cf.color !== undefined
  if (color === null) {
    if (base?.colorExplicit) {
      color = base.color
      colorExplicit = true
    } else if (cf.bg !== undefined) {
      color = contrastColor(cf.bg)
    } else {
      color = base?.color ?? null
    }
  }
  const strike = cf.strike ?? base?.strike
  const underline = cf.underline ?? base?.underline
  return {
    bg,
    color,
    bold: cf.bold ?? base?.bold ?? false,
    italic: cf.italic ?? base?.italic ?? false,
    border: base?.border ?? null,
    align: base?.align ?? null,
    ...(strike ? { strike: true as const } : {}),
    ...(underline ? { underline: true as const } : {}),
    ...(colorExplicit ? { colorExplicit: true as const } : {}),
  }
}
