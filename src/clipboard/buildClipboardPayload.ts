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

function cellsInRange(sheet: SheetModel, range: CellRange): string[][] {
  const out: string[][] = []
  for (let r = range.r0; r <= range.r1; r++) {
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

export function buildTsv(sheet: SheetModel, range: CellRange): string {
  return cellsInRange(sheet, range)
    .map((row) => row.map(tsvEscape).join('\t'))
    .join('\n')
}

function htmlEscape(value: string): string {
  return value.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
}

export function buildHtmlTable(sheet: SheetModel, range: CellRange): string {
  const rows = cellsInRange(sheet, range)
    .map((row) => `<tr>${row.map((v) => `<td>${htmlEscape(v)}</td>`).join('')}</tr>`)
    .join('')
  return `<table>${rows}</table>`
}

/**
 * 붙여넣기용 TSV 파서 — buildTsv가 만드는 형식(탭으로 열 구분, 필드가 탭·개행·
 * 큰따옴표를 담으면 전체를 큰따옴표로 감싸고 내부 큰따옴표는 두 번 씀)을 그대로
 * 되돌린다. 실제 Excel의 클립보드 TSV도 같은 규칙이라 Excel에서 복사해 붙여넣는
 * 것도 대부분 문제없이 읽힌다.
 *
 * 줄바꿈으로 먼저 나누고 그다음 각 줄을 따옴표 처리하면, 따옴표 안에 개행이
 * 담긴 필드(buildTsv가 그렇게 만든다)가 줄이 갈라지면서 깨진다 — 그래서
 * 텍스트 전체를 한 번에, 따옴표 상태를 유지하면서 훑어야 한다.
 */
export function parseTsv(text: string): string[][] {
  const normalized = text.replace(/\r\n/g, '\n').replace(/\r/g, '\n')
  const rows: string[][] = []
  let row: string[] = []
  let current = ''
  let inQuotes = false

  for (let i = 0; i < normalized.length; i++) {
    const ch = normalized[i]
    if (inQuotes) {
      if (ch === '"') {
        if (normalized[i + 1] === '"') {
          current += '"'
          i++
        } else {
          inQuotes = false
        }
      } else {
        current += ch
      }
    } else if (ch === '"' && current === '') {
      inQuotes = true
    } else if (ch === '\t') {
      row.push(current)
      current = ''
    } else if (ch === '\n') {
      row.push(current)
      current = ''
      rows.push(row)
      row = []
    } else {
      current += ch
    }
  }
  row.push(current)
  rows.push(row)

  // 맨 끝이 진짜 빈 줄(복사한 텍스트가 개행으로 끝나서 생긴 것)이면 버린다 —
  // 단, 유일한 셀이 빈 문자열인 한 칸짜리 붙여넣기까지 지워버리면 안 된다.
  const last = rows[rows.length - 1]
  if (rows.length > 1 && last.length === 1 && last[0] === '') rows.pop()

  return rows
}
