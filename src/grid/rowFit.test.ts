import { describe, expect, it } from 'vitest'
import type { CellStyle } from '../xlsx/cellStyle'
import type { CellModel, SheetModel } from '../xlsx/types'
import { fitRowHeights } from './rowFit'
import type { MeasureWrappedHeights, WrapRequest } from './wrapMeasure'

const LINE = 20
const WRAP: CellStyle = { bg: null, color: null, bold: false, italic: false, border: null, align: { h: null, v: null, wrap: true } }

const cell = (value: string | number | null, style: CellStyle | null = WRAP): CellModel => ({ address: 'A1', value, formula: null, numFmt: null, style, hyperlink: null })

/**
 * 가짜 측정기: 글자 수 / (칸 폭 ÷ 10)을 올림해 줄 수로 보고, 줄 수 × LINE을 높이로 돌려준다.
 * 직접 줄바꿈(\n)도 줄로 센다. 한 줄 표본('가')은 LINE이 나온다.
 */
function fakeMeasure(calls: WrapRequest[][] = []): MeasureWrappedHeights {
  return (requests) => {
    calls.push(requests)
    return requests.map(({ text, cellWidth }) => {
      const perLine = Math.max(1, Math.floor(cellWidth / 10))
      const lines = text.split('\n').reduce((sum, part) => sum + Math.max(1, Math.ceil(part.length / perLine)), 0)
      return lines * LINE
    })
  }
}

function sheetOf(rows: (CellModel | undefined)[][], over: Partial<SheetModel> = {}): SheetModel {
  const colCount = Math.max(...rows.map((r) => r.length))
  return {
    name: 'S',
    rowCount: rows.length,
    colCount,
    rows,
    merges: [],
    colWidths: new Array(colCount).fill(100), // 한 줄 10글자
    rowHeights: new Array(rows.length).fill(22),
    autoHeightRows: new Array(rows.length).fill(true),
    hiddenCols: new Array(colCount).fill(false),
    hiddenRows: new Array(rows.length).fill(false),
    frozen: null,
    filters: [],
    conditionalFormats: [],
    skippedConditionalFormats: {},
    drawings: [],
    skippedDrawings: {},
    ...over,
  }
}

describe('fitRowHeights', () => {
  it('줄바꿈 글이 저장된 높이보다 길면 행을 키운다(줄 수 × 줄높이 + 여유)', () => {
    const sheet = sheetOf([[cell('가'.repeat(25))]]) // 3줄
    const [h] = fitRowHeights(sheet, fakeMeasure())
    expect(h).toBe(3 * LINE + 4)
  })

  it('한 줄에 들어가는 글은 행을 키우지 않는다 — 우리 줄 높이가 조금 커도 모든 행이 늘어나면 안 된다', () => {
    const sheet = sheetOf([[cell('짧은 글')]], { rowHeights: [15] })
    expect(fitRowHeights(sheet, fakeMeasure())).toEqual([15])
  })

  it('저장된 높이가 충분하면 줄이지 않는다', () => {
    const sheet = sheetOf([[cell('가'.repeat(15))]], { rowHeights: [200] })
    expect(fitRowHeights(sheet, fakeMeasure())).toEqual([200])
  })

  it('사용자가 높이를 직접 정한 행(autoHeightRows false)은 건드리지 않는다', () => {
    const sheet = sheetOf([[cell('가'.repeat(50))], [cell('가'.repeat(50))]], { autoHeightRows: [false, true] })
    const [first, second] = fitRowHeights(sheet, fakeMeasure())
    expect(first).toBe(22)
    expect(second).toBeGreaterThan(22)
  })

  it('병합된 칸은 맞추지 않는다(Excel도 병합 칸은 자동 높이에서 뺀다)', () => {
    const sheet = sheetOf([[cell('가'.repeat(50)), cell(null, null)]], { merges: ['A1:B1'] })
    expect(fitRowHeights(sheet, fakeMeasure())).toEqual([22])
  })

  it('줄바꿈 서식이 없는 칸, 숫자 칸, 숨긴 열/행은 무시한다', () => {
    const noWrap = cell('가'.repeat(50), { ...WRAP, align: { h: null, v: null, wrap: false } })
    const sheet = sheetOf([[noWrap, cell(12345), cell('가'.repeat(50))], [cell('가'.repeat(50))]], {
      hiddenCols: [false, false, true],
      hiddenRows: [false, true],
    })
    expect(fitRowHeights(sheet, fakeMeasure())).toEqual([22, 22])
  })

  it('한 행의 여러 칸 중 가장 긴 칸에 맞춘다', () => {
    const sheet = sheetOf([[cell('가'.repeat(15)), cell('가'.repeat(45)), cell('가'.repeat(5))]])
    expect(fitRowHeights(sheet, fakeMeasure())).toEqual([5 * LINE + 4])
  })

  it('직접 줄바꿈(\\n)도 줄로 센다', () => {
    const sheet = sheetOf([[cell('가\n나\n다')]])
    expect(fitRowHeights(sheet, fakeMeasure())).toEqual([3 * LINE + 4])
  })

  it('맞출 것이 없으면 같은 배열 참조를 돌려주고 측정도 하지 않는다', () => {
    const calls: WrapRequest[][] = []
    const none = sheetOf([[cell('가'.repeat(50))]], { autoHeightRows: [false] })
    expect(fitRowHeights(none, fakeMeasure(calls))).toBe(none.rowHeights)
    const noWrapCells = sheetOf([[cell('가'.repeat(50), null)]])
    expect(fitRowHeights(noWrapCells, fakeMeasure(calls))).toBe(noWrapCells.rowHeights)
    expect(calls).toHaveLength(0)
  })

  it('측정은 시트당 한 번만 부른다(레이아웃을 한 번만 돌리기 위해)', () => {
    const calls: WrapRequest[][] = []
    const sheet = sheetOf([[cell('가'.repeat(30)), cell('나'.repeat(30))], [cell('다'.repeat(30))]])
    fitRowHeights(sheet, fakeMeasure(calls))
    expect(calls).toHaveLength(1)
  })

  it('원본 rowHeights 배열을 바꾸지 않는다', () => {
    const sheet = sheetOf([[cell('가'.repeat(50))]])
    const before = [...sheet.rowHeights]
    fitRowHeights(sheet, fakeMeasure())
    expect(sheet.rowHeights).toEqual(before)
  })

  it('글자 크기가 다른 칸은 그 크기의 한 줄 높이로 줄 수를 센다', () => {
    // 큰 글자(2배)는 한 줄이 이미 2*LINE — 한 줄짜리라면 행을 키우지 않아야 한다.
    const calls: WrapRequest[][] = []
    const measure: MeasureWrappedHeights = (requests) => {
      calls.push(requests)
      return requests.map((r) => ((r.style?.fontScale ?? 1) * LINE))
    }
    const big: CellStyle = { ...WRAP, fontScale: 2 }
    const sheet = sheetOf([[cell('가', big)]], { rowHeights: [10] })
    expect(fitRowHeights(sheet, measure)).toEqual([10])
  })
})
