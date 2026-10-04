import { describe, expect, it } from 'vitest'
import type { SheetModel } from '../xlsx/types'
import { computeSummary } from './summary'

function cell(value: string | number | null) {
  return { address: 'A1', value, formula: null, numFmt: null, style: null, hyperlink: null }
}

const sheet: SheetModel = {
  name: 'S',
  rowCount: 2,
  colCount: 2,
  merges: [],
  colWidths: [],
  rowHeights: [],
  hiddenCols: [],
  hiddenRows: [],
  frozen: null,
  filters: [],
  conditionalFormats: [],
  skippedConditionalFormats: {},
  drawings: [],
  skippedDrawings: {},
  rows: [
    [cell('Text'), cell(10)],
    [cell('More text'), cell(20)],
  ],
}

describe('computeSummary', () => {
  it('range 없이는 시트 전체를 본다', () => {
    const s = computeSummary(sheet)
    expect(s).toMatchObject({ sum: 30, avg: 15, max: 20, count: 2, selectedCount: null })
  })

  it('숫자가 없는 범위를 선택하면 count는 0이지만 selectedCount는 칸 수를 준다', () => {
    const s = computeSummary(sheet, { r0: 1, c0: 1, r1: 2, c1: 1 })
    expect(s).toMatchObject({ count: 0, selectedCount: 2 })
  })

  it('selectedCount는 선택 범위의 전체 칸 수(빈 칸 포함)다', () => {
    const s = computeSummary(sheet, { r0: 1, c0: 1, r1: 2, c1: 2 })
    expect(s.selectedCount).toBe(4)
  })
})

describe('computeSummary — 필터로 걸러진 행 제외', () => {
  it('skipRows인 행은 합계·개수·선택 칸 수에서 빠진다(Excel 상태 표시줄과 동일)', () => {
    const s = computeSummary(sheet, { r0: 1, c0: 1, r1: 2, c1: 2 }, [false, true])
    expect(s).toMatchObject({ sum: 10, count: 1, selectedCount: 2 })
  })
})
