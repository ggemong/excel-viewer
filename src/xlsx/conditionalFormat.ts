/**
 * 조건부서식을 ExcelJS 모델에서 화면용 모델로 바꾼다.
 *
 * 왜 중요한가: 업무 현황 시트는 상태 열 기준으로 행 전체를 칠한다(완료 회색, 불가/보류 취소선, 기한 지난
 * 행 빨간 글자). 이 색은 장식이 아니라 **정보**라서, 빠지면 시트의 의미가 사라진다.
 *
 * 범위(의도적 제한): 수식 기반 규칙(`expression` 및 수식을 같이 주는 containsText 계열)과 `cellIs`만 지원한다.
 * 색조·데이터 막대·아이콘 집합·상위/하위 N·평균 기준·중복 값은 지원하지 않고, 지원 함수 밖의 수식도
 * 마찬가지로 **조용히 틀리게 칠하지 않고 건너뛴 뒤 종류별 개수를 skipped로 돌려준다**(화면이 안내 문구로 알림).
 * 스타일은 채우기·글자색·굵게·기울임·취소선·밑줄만 반영한다(테두리·표시 형식은 아직).
 */
import { resolveColor, type ExcelColor } from './cellStyle'
import { collectFunctionNames, parseFormula, UnsupportedFormulaError } from '../formula/parse'
import { SUPPORTED_FUNCTIONS } from '../formula/evaluate'
import type { ThemeColors } from './themeColor'
import { parseCellAddress } from './cellRef'

export interface CfRange {
  r0: number
  c0: number
  r1: number
  c1: number
}

export interface CfStyle {
  bg?: string
  color?: string
  bold?: true
  italic?: true
  strike?: true
  underline?: true
}

export interface CfRule {
  /** 작을수록 먼저(우선) 적용된다 — 시트 전체에서 비교한다. */
  priority: number
  stopIfTrue: boolean
  kind: 'expression' | 'cellIs'
  /** cellIs의 비교 연산(greaterThan 등). expression이면 null. */
  operator: string | null
  formulae: string[]
  style: CfStyle
}

export interface ConditionalFormat {
  /** 적용 범위들. 수식의 상대 참조는 첫 범위의 왼쪽 위 칸을 기준으로 한다. */
  ranges: CfRange[]
  rules: CfRule[]
}

export interface ExtractedConditionalFormats {
  formats: ConditionalFormat[]
  /** 표시하지 못해 건너뛴 규칙의 종류별 개수(예: { colorScale: 2, unsupportedFormula: 1 }). */
  skipped: Record<string, number>
}

interface RawDxf {
  fill?: { pattern?: string; fgColor?: ExcelColor; bgColor?: ExcelColor } | null
  font?: { bold?: boolean; italic?: boolean; strike?: boolean; underline?: boolean | string; color?: ExcelColor } | null
  border?: unknown
  numFmt?: unknown
}
interface RawRule {
  type: string
  priority?: number
  stopIfTrue?: boolean
  operator?: string
  formulae?: string[]
  style?: RawDxf | null
}
export interface RawConditionalFormatting {
  ref: string
  rules: RawRule[]
}

/** 수식을 같이 주는 규칙 종류 — 수식 하나로 판정 논리가 이미 들어 있어 expression처럼 계산하면 된다. */
const FORMULA_BACKED_TYPES = new Set([
  'expression',
  'containsText',
  'notContainsText',
  'beginsWith',
  'endsWith',
  'containsBlanks',
  'notContainsBlanks',
  'containsErrors',
  'notContainsErrors',
  'timePeriod',
])

/** sqref `A1:N3 H4:N5 F4`를 범위 목록으로. 해석 못 하는 조각은 건너뛴다. */
export function parseSqref(sqref: string): CfRange[] {
  const out: CfRange[] = []
  for (const part of sqref.split(/\s+/).filter(Boolean)) {
    const [a, b] = part.split(':')
    try {
      const from = parseCellAddress(a.replace(/\$/g, ''))
      const to = parseCellAddress((b ?? a).replace(/\$/g, ''))
      out.push({ r0: Math.min(from.row, to.row), c0: Math.min(from.col, to.col), r1: Math.max(from.row, to.row), c1: Math.max(from.col, to.col) })
    } catch {
      // 해석 못 하는 범위 조각은 무시(같은 sqref의 나머지 범위는 살린다)
    }
  }
  return out
}

function convertStyle(dxf: RawDxf | null | undefined, theme: ThemeColors | null): { style: CfStyle; unsupportedParts: boolean } {
  const style: CfStyle = {}
  // dxf의 채우기 색은 bgColor에 들어 있다(patternFill 안에서 fg/bg 의미가 일반 셀 서식과 반대).
  if (dxf?.fill && dxf.fill.pattern !== 'none') {
    const bg = resolveColor(dxf.fill.bgColor, theme) ?? resolveColor(dxf.fill.fgColor, theme)
    if (bg) style.bg = bg
  }
  const font = dxf?.font
  if (font) {
    const color = resolveColor(font.color, theme)
    if (color) style.color = color
    if (font.bold) style.bold = true
    if (font.italic) style.italic = true
    if (font.strike) style.strike = true
    if (font.underline && font.underline !== 'none') style.underline = true
  }
  return { style, unsupportedParts: Boolean(dxf?.border || dxf?.numFmt) }
}

/** 이 수식을 계산할 수 있는가. 못 하면 사유를 돌려준다. */
export function formulaSupport(formula: string): { ok: true } | { ok: false; reason: string } {
  try {
    const unknown = [...collectFunctionNames(parseFormula(formula))].filter((name) => !SUPPORTED_FUNCTIONS.has(name))
    return unknown.length === 0 ? { ok: true } : { ok: false, reason: `함수: ${unknown.join(', ')}` }
  } catch (err) {
    return { ok: false, reason: err instanceof UnsupportedFormulaError ? err.message : '수식 해석 실패' }
  }
}

const count = (skipped: Record<string, number>, kind: string) => {
  skipped[kind] = (skipped[kind] ?? 0) + 1
}

export function extractConditionalFormats(raw: RawConditionalFormatting[], theme: ThemeColors | null): ExtractedConditionalFormats {
  const formats: ConditionalFormat[] = []
  const skipped: Record<string, number> = {}

  for (const block of raw) {
    const ranges = parseSqref(block.ref)
    const rules: CfRule[] = []

    for (const r of block.rules) {
      const kind = FORMULA_BACKED_TYPES.has(r.type) ? 'expression' : r.type === 'cellIs' ? 'cellIs' : null
      if (!kind) {
        count(skipped, r.type)
        continue
      }
      const formulae = r.formulae ?? []
      if (formulae.length === 0 || !formulae.every((f) => formulaSupport(f).ok)) {
        count(skipped, 'unsupportedFormula')
        continue
      }
      const { style, unsupportedParts } = convertStyle(r.style, theme)
      if (Object.keys(style).length === 0) {
        // 칠할 게 없는 규칙: 테두리·표시 형식만 바꾸는 규칙이면 "못 보여준 것"으로 알리고, 아니면 효과가 없는 규칙이다.
        if (unsupportedParts) count(skipped, 'borderOrNumberFormat')
        continue
      }
      rules.push({ priority: r.priority ?? Number.MAX_SAFE_INTEGER, stopIfTrue: Boolean(r.stopIfTrue), kind, operator: r.operator ?? null, formulae, style })
    }

    if (ranges.length > 0 && rules.length > 0) formats.push({ ranges, rules })
  }

  return { formats, skipped }
}
