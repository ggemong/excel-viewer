/**
 * 칸의 링크 규칙 한 곳: 어떤 외부 주소를 클릭 가능하게 둘지, 그리고 `=HYPERLINK(...)` 수식을 링크로 읽는 법.
 *
 * 왜 수식을 따로 읽나: 업무 파일은 시트 사이를 오가는 링크를 `=HYPERLINK("#'참조 2607-1'!A1", "참조 2607-1")` 수식으로 많이 만든다.
 * 실제 파일 하나에 76개였고, 이것이 그 파일의 핵심 작업 흐름(요청 시트 ↔ 참조 시트 이동)이다. ExcelJS는 수식 칸의 결과 글자만 주고
 * 링크는 주지 않는다. 같은 통합문서 안으로 가는 링크(`<hyperlink location=...>`)도 ExcelJS가 읽지 않지만, 그쪽은 시트 XML을 전부 다시
 * 읽어야 해서 아직 다루지 않는다.
 */
import { parseCellAddress } from './cellRef'

/** 같은 통합문서 안의 칸으로 가는 링크. sheet가 null이면 지금 시트. row/col은 1부터. */
export interface InternalLink {
  sheet: string | null
  row: number
  col: number
}

/** 사용자가 올린 임의 파일을 그대로 열어주는 뷰어라, 외부 링크는 이 스킴만 신뢰한다 —
 * javascript:/file: 같은 스킴으로 된 악성 링크가 그대로 클릭 가능한 <a>가 되는 걸 막는다. */
const ALLOWED_HYPERLINK_SCHEMES = ['http:', 'https:', 'mailto:']

/** 허용된 스킴이면 그대로, 아니면 null. */
export function sanitizeHyperlink(url: string): string | null {
  try {
    return ALLOWED_HYPERLINK_SCHEMES.includes(new URL(url).protocol) ? url : null
  } catch {
    return null
  }
}

const HYPERLINK_FORMULA_RE = /^\s*HYPERLINK\s*\(\s*"((?:[^"]|"")*)"\s*[,)]/i
const CELL_REF_RE = /^\$?([A-Za-z]{1,3})\$?(\d+)$/

/** `'시트 이름'` 또는 `시트이름`에서 따옴표를 벗긴다(따옴표 안의 `''`는 `'` 하나). */
function unquoteSheet(raw: string): string {
  return raw.startsWith("'") && raw.endsWith("'") && raw.length >= 2 ? raw.slice(1, -1).replace(/''/g, "'") : raw
}

/**
 * `#시트!A1`, `#'시트 이름'!B2:C3`, `#A1` 형태의 링크 대상을 읽는다. 범위면 왼쪽 위 칸.
 * 이름 정의(named range)처럼 칸 주소가 아닌 대상은 어디로 가야 할지 알 수 없어 null.
 */
export function parseInternalTarget(target: string): InternalLink | null {
  if (!target.startsWith('#')) return null
  const body = target.slice(1)
  const bang = body.lastIndexOf('!')
  const sheet = bang >= 0 ? unquoteSheet(body.slice(0, bang)) : null
  const first = (bang >= 0 ? body.slice(bang + 1) : body).split(':')[0].trim()
  const match = CELL_REF_RE.exec(first)
  if (!match || (bang >= 0 && sheet === '')) return null
  const { row, col } = parseCellAddress(`${match[1].toUpperCase()}${match[2]}`)
  return row >= 1 && col >= 1 ? { sheet, row, col } : null
}

/** `HYPERLINK("글자 그대로의 대상", ...)`처럼 첫 인자가 리터럴인 수식인가(대상을 계산할 필요가 없는 경우). */
export function isLiteralHyperlinkFormula(formula: string): boolean {
  return HYPERLINK_FORMULA_RE.test(formula)
}

export type FormulaLink = { kind: 'external'; url: string } | { kind: 'internal'; link: InternalLink }

/**
 * `HYPERLINK("대상", ...)` 수식의 링크. 첫 인자가 글자 그대로(리터럴)일 때만 여기서 읽는다 — `HYPERLINK("#'"&A1&"'!A1")`처럼 셀 값으로
 * 만들어지는 대상은 수식을 계산해야 알 수 있어서 null이고, 그런 칸은 눌렀을 때 src/formula/linkTarget.ts가 계산한다.
 *
 * @param formula '=' 없는 수식 글자
 */
export function linkFromFormula(formula: string): FormulaLink | null {
  const match = HYPERLINK_FORMULA_RE.exec(formula)
  if (!match) return null
  const target = match[1].replace(/""/g, '"')
  if (target.startsWith('#')) {
    const link = parseInternalTarget(target)
    return link ? { kind: 'internal', link } : null
  }
  const url = sanitizeHyperlink(target)
  return url ? { kind: 'external', url } : null
}
