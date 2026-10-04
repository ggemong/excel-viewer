import { formatCellValue } from '../xlsx/formatValue'
import type { SheetModel } from '../xlsx/types'

export interface CellRange {
  r0: number
  c0: number
  r1: number
  c1: number
}

export function normalizeRange(a: { row: number; col: number }, b: { row: number; col: number }): CellRange {
  return {
    r0: Math.min(a.row, b.row),
    c0: Math.min(a.col, b.col),
    r1: Math.max(a.row, b.row),
    c1: Math.max(a.col, b.col),
  }
}

/**
 * @param skipRows 건너뛸 행(true인 인덱스 = r-1행). 필터로 걸러진 행을 복사에서 빼기 위한 것 —
 * Excel도 필터된 범위를 복사하면 보이는 행만 복사한다.
 */
function cellsInRange(sheet: SheetModel, range: CellRange, skipRows?: boolean[]): string[][] {
  const out: string[][] = []
  for (let r = range.r0; r <= range.r1; r++) {
    if (skipRows?.[r - 1]) continue
    const rowOut: string[] = []
    for (let c = range.c0; c <= range.c1; c++) {
      rowOut.push(formatCellValue(sheet.rows[r - 1]?.[c - 1]))
    }
    out.push(rowOut)
  }
  return out
}

/** Excel과 같은 방식: 탭·개행·큰따옴표가 든 셀은 큰따옴표로 감싸고 내부 큰따옴표는 두 번 씀. */
function tsvEscape(value: string): string {
  if (/[\t\n"]/.test(value)) {
    return `"${value.replace(/"/g, '""')}"`
  }
  return value
}

export function buildTsv(sheet: SheetModel, range: CellRange, skipRows?: boolean[]): string {
  return cellsInRange(sheet, range, skipRows)
    .map((row) => row.map(tsvEscape).join('\t'))
    .join('\n')
}

function htmlEscape(value: string): string {
  return value.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
}

export function buildHtmlTable(sheet: SheetModel, range: CellRange, skipRows?: boolean[]): string {
  const rows = cellsInRange(sheet, range, skipRows)
    .map((row) => `<tr>${row.map((v) => `<td>${htmlEscape(v)}</td>`).join('')}</tr>`)
    .join('')
  return `<table>${rows}</table>`
}
