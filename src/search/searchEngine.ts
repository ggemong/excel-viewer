/**
 * 시트 검색의 순수 계산부 — 화면(React)을 모르는 함수들. 어떤 셀/도형이 검색어와 일치하는지,
 * 결과에 무엇을 보여줄지만 책임진다. 검색 상태·키 처리·이동은 useSearch.ts, 그리기는
 * SearchBar/SearchResults/Grid가 맡는다.
 *
 * 일치 기준(합의한 설계):
 *  - 화면에 **보이는 글자**와 **원래 값**을 둘 다 비교한다. 그래야 `1000`으로 `1,000원`을 찾을 수
 *    있다(Excel은 보이는 글자만 비교해서 이게 불편하다). 단, 날짜 셀의 원래 값은 직렬번호(46299)라
 *    사용자에게 의미가 없어 비교하지 않는다.
 *  - 병합 셀은 왼쪽 위 칸만 센다 — ExcelJS가 병합 범위 전체에 값을 복제해서 주기 때문에, 안 거르면
 *    같은 글자가 칸 수만큼 중복으로 잡힌다(그리드도 마스터만 그린다).
 *  - 숨기거나 필터로 가려진 행/열의 일치는 결과에서 빼되 개수는 센다("숨겨진 N건 제외"로 알리기 위해).
 *  - 도형/텍스트상자 안 글자도 찾는다(요청서류는 말풍선에 수정 지시가 들어 있다). 이미지 속 글자는 못 찾는다.
 *
 * 성능: 200만 셀을 전부 서식 변환하며 찾으면 약 3.4초(문자열 원본만 훑으면 26ms)라서, 숫자는 원래
 * 값을 먼저 비교하고 서식 변환은 그다음에만 하며, 일정 개수마다 이벤트 루프에 양보하고 취소를 확인한다.
 */
import { cellAddress } from '../xlsx/cellRef'
import type { DrawingNode } from '../xlsx/drawingTypes'
import { formatCellValue } from '../xlsx/formatValue'
import { parseMergeRanges } from '../xlsx/mergeRange'
import { isDateTimeFormat } from '../xlsx/numberFormat'
import type { CellModel, SheetModel } from '../xlsx/types'

/** 이 개수의 셀을 훑을 때마다 이벤트 루프에 양보한다 — 입력/스크롤이 멈추지 않게. */
const YIELD_EVERY_CELLS = 20_000
/** 한 번의 검색이 모을 결과 상한 — 한 글자 검색이 수십만 건을 메모리/화면에 올리지 않게. */
export const MAX_HITS = 10_000
/** 결과 목록 발췌에서 일치 앞뒤로 보여줄 글자 수. */
const SNIPPET_CONTEXT = 24
/** 열 번호 상한(16384)보다 크게 잡아 (행,열)을 한 숫자 키로 합친다. */
const COL_KEY_STRIDE = 20_000

export interface SearchOptions {
  query: string
  matchCase: boolean
  /** 셀 전체가 검색어와 같을 때만 일치 */
  wholeCell: boolean
}

interface TextMatch {
  start: number
  length: number
}

interface HitBase {
  sheetIndex: number
  /** 결과 목록/이동에 쓰는 화면 글자(서식 적용 후). */
  text: string
  /** text 안의 일치 위치. 원래 값으로만 일치했으면 -1. */
  matchStart: number
  matchLength: number
}

export type SearchHit =
  | (HitBase & { kind: 'cell'; row: number; col: number })
  | (HitBase & { kind: 'shape'; drawingIndex: number; /** 도형이 놓인 칸 주소(절대 위치 도형은 빈 문자열) */ location: string })

export interface SheetSearchResult {
  hits: SearchHit[]
  /** 일치했지만 숨김/필터로 가려져 결과에서 뺀 건수. */
  hiddenExcluded: number
  truncated: boolean
  cancelled: boolean
}

export interface SearchControl {
  cancelled: () => boolean
  /** 이 시트에서 더 모을 수 있는 결과 수 */
  maxHits: number
}

export function createMatcher(opts: SearchOptions): (text: string) => TextMatch | null {
  const needle = opts.matchCase ? opts.query : opts.query.toLowerCase()
  return (text) => {
    const hay = opts.matchCase ? text : text.toLowerCase()
    if (opts.wholeCell) return hay === needle ? { start: 0, length: text.length } : null
    const i = hay.indexOf(needle)
    return i < 0 ? null : { start: i, length: needle.length }
  }
}

type Found = { text: string; start: number; length: number }

/** 셀 하나가 일치하는지: 보이는 글자와 원래 값을 모두 본다. 일치하지 않으면 null. */
function matchCell(cell: CellModel, match: (text: string) => TextMatch | null): Found | null {
  const v = cell.value
  if (v === null || v === undefined) return null

  if (typeof v === 'string') {
    const shown = cell.numFmt && cell.numFmt !== 'General' ? formatCellValue(cell) : v
    const m = match(shown)
    if (m) return { text: shown, start: m.start, length: m.length }
    if (shown !== v && match(v)) return { text: shown, start: -1, length: 0 }
    return null
  }

  if (typeof v === 'boolean') {
    const shown = v ? 'TRUE' : 'FALSE'
    const m = match(shown)
    return m ? { text: shown, start: m.start, length: m.length } : null
  }

  // 숫자: 원래 값이 더 싸게 비교되므로 먼저 보고, 서식 변환은 그다음에만 한다.
  if (!isDateTimeFormat(cell.numFmt) && match(String(v))) {
    const shown = formatCellValue(cell)
    const m = match(shown)
    return { text: shown, start: m ? m.start : -1, length: m ? m.length : 0 }
  }
  const shown = formatCellValue(cell)
  const m = match(shown)
  return m ? { text: shown, start: m.start, length: m.length } : null
}

/** 병합 범위에서 왼쪽 위(마스터)를 뺀 나머지 칸의 키 집합. */
function nonMasterMergedCells(sheet: SheetModel): Set<number> {
  const covered = new Set<number>()
  for (const range of parseMergeRanges(sheet.merges)) {
    const lastRow = Math.min(range.r1, sheet.rows.length)
    for (let r = range.r0; r <= lastRow; r++) {
      for (let c = range.c0; c <= range.c1; c++) {
        if (r === range.r0 && c === range.c0) continue
        covered.add(r * COL_KEY_STRIDE + c)
      }
    }
  }
  return covered
}

/** 도형/묶음 안의 모든 글자를 한 줄로(문단·줄바꿈은 공백). 글자가 없으면 ''. */
export function drawingText(node: DrawingNode): string {
  if (node.kind === 'shape') {
    return (node.text?.paragraphs ?? [])
      .map((p) => p.runs.map((r) => (r.text === '\n' ? ' ' : r.text)).join(''))
      .join(' ')
      .trim()
  }
  if (node.kind === 'group') {
    return node.children.map((c) => drawingText(c.node)).filter(Boolean).join(' ')
  }
  return ''
}

const yieldToMain = () => new Promise<void>((resolve) => setTimeout(resolve, 0))

/**
 * 시트 하나를 검색한다.
 * @param sheet 숨김 행이 반영된(필터 적용된) 시트를 넘겨야 "가려진 행 제외"가 화면과 일치한다.
 */
export async function searchSheet(sheet: SheetModel, sheetIndex: number, opts: SearchOptions, ctl: SearchControl): Promise<SheetSearchResult> {
  const match = createMatcher(opts)
  const covered = nonMasterMergedCells(sheet)
  const hits: SearchHit[] = []
  let hiddenExcluded = 0
  let truncated = false
  let sinceYield = 0

  scan: for (let r = 1; r <= sheet.rows.length; r++) {
    const row = sheet.rows[r - 1]
    if (!row) continue
    const rowHidden = Boolean(sheet.hiddenRows[r - 1])

    for (let c = 1; c <= row.length; c++) {
      const cell = row[c - 1]
      if (!cell || covered.has(r * COL_KEY_STRIDE + c)) continue
      const found = matchCell(cell, match)
      if (!found) continue
      if (rowHidden || sheet.hiddenCols[c - 1]) {
        hiddenExcluded++
        continue
      }
      if (hits.length >= ctl.maxHits) {
        truncated = true
        break scan
      }
      hits.push({ kind: 'cell', sheetIndex, row: r, col: c, text: found.text, matchStart: found.start, matchLength: found.length })
    }

    sinceYield += row.length
    if (sinceYield >= YIELD_EVERY_CELLS) {
      sinceYield = 0
      await yieldToMain()
      if (ctl.cancelled()) return { hits, hiddenExcluded, truncated, cancelled: true }
    }
  }

  if (!truncated) {
    for (let i = 0; i < sheet.drawings.length; i++) {
      const item = sheet.drawings[i]
      const text = drawingText(item.node)
      const m = text ? match(text) : null
      if (!m) continue
      if (hits.length >= ctl.maxHits) {
        truncated = true
        break
      }
      const a = item.anchor
      const location = a.kind === 'absolute' ? '' : cellAddress(a.from.row + 1, a.from.col + 1)
      hits.push({ kind: 'shape', sheetIndex, drawingIndex: i, text, matchStart: m.start, matchLength: m.length, location })
    }
  }

  return { hits, hiddenExcluded, truncated, cancelled: ctl.cancelled() }
}

/** 결과 목록용 발췌: 일치 앞뒤 일부만 남기고 줄임표를 붙인다. 일치 위치를 모르면(-1) 앞부분만. */
export function makeSnippet(hit: Pick<HitBase, 'text' | 'matchStart' | 'matchLength'>): { before: string; match: string; after: string } {
  const flat = hit.text.replace(/\s+/g, ' ')
  if (hit.matchStart < 0) {
    return { before: '', match: '', after: flat.length > SNIPPET_CONTEXT * 2 ? `${flat.slice(0, SNIPPET_CONTEXT * 2)}…` : flat }
  }
  // 공백을 접으면 인덱스가 어긋날 수 있어, 접기 전 문자열에서 같은 위치를 쓰되 줄바꿈만 공백으로 바꾼다.
  const text = hit.text.replace(/\s/g, ' ')
  const start = hit.matchStart
  const end = start + hit.matchLength
  const from = Math.max(0, start - SNIPPET_CONTEXT)
  const to = Math.min(text.length, end + SNIPPET_CONTEXT)
  return {
    before: `${from > 0 ? '…' : ''}${text.slice(from, start)}`,
    match: text.slice(start, end),
    after: `${text.slice(end, to)}${to < text.length ? '…' : ''}`,
  }
}

/** 현재 시트의 결과를 그리드 강조용 자료구조로 바꾼다(셀 키 "행,열" 집합 + 도형 인덱스 집합). */
export function highlightsFor(hits: SearchHit[], sheetIndex: number): { cells: Set<string>; shapes: Set<number> } {
  const cells = new Set<string>()
  const shapes = new Set<number>()
  for (const hit of hits) {
    if (hit.sheetIndex !== sheetIndex) continue
    if (hit.kind === 'cell') cells.add(`${hit.row},${hit.col}`)
    else shapes.add(hit.drawingIndex)
  }
  return { cells, shapes }
}
