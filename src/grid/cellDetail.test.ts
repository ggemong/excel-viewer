import { describe, expect, it } from 'vitest'
import type { CellModel, SheetModel } from '../xlsx/types'
import { describeCell, hasDetail } from './cellDetail'

const cell = (value: string | number | null, over: Partial<CellModel> = {}): CellModel => ({
  address: 'A1',
  value,
  formula: null,
  numFmt: null,
  style: null,
  hyperlink: null,
  ...over,
})

function sheetOf(rows: (CellModel | undefined)[][], merges: string[] = []): SheetModel {
  return {
    name: 'S',
    rowCount: rows.length,
    colCount: 3,
    rows,
    merges,
    colWidths: [64, 64, 64],
    rowHeights: new Array(rows.length).fill(20),
    autoHeightRows: [],
    hiddenCols: [false, false, false],
    hiddenRows: new Array(rows.length).fill(false),
    frozen: null,
    filters: [],
    conditionalFormats: [],
    skippedConditionalFormats: {},
    drawings: [],
    skippedDrawings: {},
  }
}

describe('describeCell', () => {
  it('긴 글자 전체와 주소', () => {
    const long = '아주 긴 문장입니다. '.repeat(20)
    const d = describeCell(sheetOf([[cell('a'), cell(long)]]), 1, 2)
    expect(d.address).toBe('B1')
    expect(d.text).toBe(long)
  })

  it('보이는 글자와 원래 값이 다르면 원래 값도 알려준다', () => {
    const d = describeCell(sheetOf([[cell(1000, { numFmt: '#,##0"원"' })]]), 1, 1)
    expect(d).toMatchObject({ text: '1,000원', originalValue: '1000' })
  })

  it('같으면 원래 값은 null, 날짜 직렬번호는 보여주지 않는다', () => {
    expect(describeCell(sheetOf([[cell('가')]]), 1, 1).originalValue).toBeNull()
    expect(describeCell(sheetOf([[cell(46299, { numFmt: 'yyyy-mm-dd' })]]), 1, 1)).toMatchObject({ text: '2026-10-04', originalValue: null })
  })

  it('수식·링크·메모', () => {
    const d = describeCell(sheetOf([[cell(3, { formula: 'A1+B1', hyperlink: 'https://a.kr', note: '확인 필요' })]]), 1, 1)
    expect(d).toMatchObject({ formula: '=A1+B1', hyperlink: 'https://a.kr', note: '확인 필요' })
  })

  it('병합 범위 안쪽 칸을 골라도 마스터(왼쪽 위) 내용을 보여준다', () => {
    const s = sheetOf([[cell('제목'), cell('제목'), cell('제목')], [cell('x'), cell('x'), cell('x')]], ['A1:C2'])
    const d = describeCell(s, 2, 3)
    expect(d).toMatchObject({ address: 'A1', text: '제목' })
  })

  it('빈 칸은 내용 없음, 메모만 있는 칸은 내용 있음', () => {
    const empty = describeCell(sheetOf([[undefined]]), 1, 1)
    expect(empty).toMatchObject({ text: '', note: null })
    expect(hasDetail(empty)).toBe(false)
    expect(hasDetail(describeCell(sheetOf([[cell(null, { note: '메모만' })]]), 1, 1))).toBe(true)
  })
})
