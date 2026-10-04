/**
 * `=HYPERLINK(대상, 이름)` 수식의 이동 대상을 **클릭하는 순간** 계산한다.
 *
 * 왜 클릭 시점인가: 실제 업무 파일의 "돌아가기" 링크는 대상이 계산으로 만들어진다 — 지금 시트 이름(CELL("filename"))을 읽어서
 * 요청 시트의 열(INDIRECT)에서 그 이름이 몇 번째 행인지(MATCH) 찾는 식이다. 모든 칸을 열 때 계산하면 열기가 느려지고 대부분은
 * 안 눌린다. 눌릴 때 한 칸만 계산하면 비용이 없고 결과도 그때의 값으로 정확하다.
 *
 * 계산에 실패하면(지원하지 않는 함수, 오류 결과, 이해 못 하는 대상) 틀린 곳으로 보내지 않고 이유를 돌려준다 — 호출부가 사용자에게 알린다.
 */
import { parseInternalTarget, sanitizeHyperlink, type InternalLink } from '../xlsx/hyperlink'
import type { CellModel, SheetModel, WorkbookModel } from '../xlsx/types'
import { evaluateFormula, type EvalContext, type Scalar } from './evaluate'
import { parseFormula, type Node } from './parse'

export type LinkTarget = { kind: 'internal'; link: InternalLink } | { kind: 'external'; url: string }
export type LinkResolution = { ok: true; target: LinkTarget } | { ok: false; reason: string }

const LINK_FORMULA_RE = /^\s*(?:IFERROR\s*\(\s*)?HYPERLINK\s*\(/i

/** 수식이 `HYPERLINK(...)`이거나 `IFERROR(HYPERLINK(...), ...)`로 시작하는가 — 칸을 링크로 보일지 정하는 가벼운 검사. */
export function isLinkFormula(formula: string): boolean {
  return LINK_FORMULA_RE.test(formula)
}

type CallNode = Extract<Node, { t: 'call' }>

/** 수식에서 HYPERLINK 호출을 찾는다. 오류를 빈 글자로 바꾸는 IFERROR 겉껍질은 벗겨서 본다. */
function findHyperlinkCall(ast: Node): CallNode | null {
  if (ast.t !== 'call') return null
  if (ast.name === 'HYPERLINK') return ast
  return ast.name === 'IFERROR' && ast.args[0] ? findHyperlinkCall(ast.args[0]) : null
}

/**
 * @param formula '=' 없는 수식 글자
 * @param ctx 수식이 쓰인 칸 기준의 계산 문맥(workbookEvalContext)
 */
export function resolveLinkFormula(formula: string, ctx: EvalContext): LinkResolution {
  let call: CallNode | null
  try {
    call = findHyperlinkCall(parseFormula(formula))
  } catch {
    return { ok: false, reason: '이 링크의 수식을 해석하지 못했어요' }
  }
  if (!call?.args[0]) return { ok: false, reason: '이 링크의 수식을 해석하지 못했어요' }

  const value = evaluateFormula(call.args[0], ctx)
  if (typeof value === 'object' && value !== null) return { ok: false, reason: `이동할 곳을 계산하지 못했어요(${value.error})` }
  if (typeof value !== 'string' || value === '') return { ok: false, reason: '이동할 곳이 비어 있어요' }

  if (value.startsWith('#')) {
    const link = parseInternalTarget(value)
    return link ? { ok: true, target: { kind: 'internal', link } } : { ok: false, reason: `이동할 곳을 이해하지 못했어요(${value})` }
  }
  const url = sanitizeHyperlink(value)
  return url ? { ok: true, target: { kind: 'external', url } } : { ok: false, reason: '열 수 없는 형식의 주소예요' }
}

const scalarOf = (cell: CellModel | undefined): Scalar => cell?.value ?? null

/**
 * 한 칸을 기준으로 한 계산 문맥: 그 칸이 있는 시트의 값, 다른 시트의 값(INDIRECT), 파일/시트 이름(CELL).
 * 시트 이름은 Excel처럼 대소문자를 구분하지 않고 찾는다.
 */
export function workbookEvalContext(workbook: WorkbookModel, sheet: SheetModel, row: number, col: number): EvalContext {
  const byName = new Map(workbook.sheets.map((s) => [s.name.toLowerCase(), s]))
  return {
    row,
    col,
    anchorRow: row,
    anchorCol: col,
    getValue: (r, c) => scalarOf(sheet.rows[r - 1]?.[c - 1]),
    sheetName: sheet.name,
    fileName: workbook.fileName,
    sheetValue: (name, r, c) => scalarOf(byName.get(name.toLowerCase())?.rows[r - 1]?.[c - 1]),
    sheetExtent: (name) => {
      const found = byName.get(name.toLowerCase())
      return found ? { rows: found.rowCount, cols: found.colCount } : null
    },
  }
}
