import { describe, expect, it } from 'vitest'
import { buildHtmlTable, buildTsv, normalizeRange } from './buildClipboardPayload'
import type { SheetModel } from '../xlsx/types'

function cell(value: string | number | null) {
  return { address: 'A1', value, formula: null, numFmt: null }
}

const sheet: SheetModel = {
  name: 'Sheet1',
  rowCount: 3,
  colCount: 2,
  merges: [],
  rows: [
    [cell('Name'), cell('Amount')],
    [cell('a "quote"\ttab'), cell(100)],
    [cell(null), cell('line1\nline2')],
  ],
}

describe('normalizeRange', () => {
  it('두 점 중 어느 쪽이 시작이든 좌상단-우하단으로 정규화한다', () => {
    expect(normalizeRange({ row: 3, col: 2 }, { row: 1, col: 1 })).toEqual({ r0: 1, c0: 1, r1: 3, c1: 2 })
  })
})

describe('buildTsv', () => {
  it('탭으로 열을, 개행으로 행을 구분한다', () => {
    const tsv = buildTsv(sheet, { r0: 1, c0: 1, r1: 1, c1: 2 })
    expect(tsv).toBe('Name\tAmount')
  })

  it('탭·큰따옴표가 든 값은 Excel 방식대로 따옴표로 감싸고 내부 따옴표를 두 번 쓴다', () => {
    const tsv = buildTsv(sheet, { r0: 2, c0: 1, r1: 2, c1: 1 })
    expect(tsv).toBe('"a ""quote""\ttab"')
  })

  it('빈 셀은 빈 문자열로 나온다', () => {
    const tsv = buildTsv(sheet, { r0: 3, c0: 1, r1: 3, c1: 1 })
    expect(tsv).toBe('')
  })
})

describe('buildHtmlTable', () => {
  it('<table><tr><td> 구조로 감싸고 HTML 특수문자를 이스케이프한다', () => {
    const html = buildHtmlTable(sheet, { r0: 1, c0: 1, r1: 2, c1: 2 })
    expect(html).toBe(
      '<table><tr><td>Name</td><td>Amount</td></tr><tr><td>a "quote"\ttab</td><td>100</td></tr></table>',
    )
  })
})
