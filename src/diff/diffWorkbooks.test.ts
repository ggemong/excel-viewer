import { describe, expect, it } from 'vitest'
import { diffWorkbooks } from './diffWorkbooks'
import type { WorkbookModel } from '../xlsx/types'

function cell(value: string | number | null) {
  return { address: 'A1', value, formula: null, numFmt: null, style: null, hyperlink: null }
}

function wb(sheets: WorkbookModel['sheets']): WorkbookModel {
  return { fileName: 'x.xlsx', sheets }
}

describe('diffWorkbooks', () => {
  it('값이 같으면 unchanged, 다르면 changed', () => {
    const base = wb([{ name: 'S', rowCount: 1, colCount: 2, merges: [], rows: [[cell('a'), cell(1)]] }])
    const compare = wb([{ name: 'S', rowCount: 1, colCount: 2, merges: [], rows: [[cell('a'), cell(2)]] }])

    const diff = diffWorkbooks(base, compare)
    const sheet = diff.sheets.get('S')!
    expect(sheet.cells[0][0].status).toBe('unchanged')
    expect(sheet.cells[0][1].status).toBe('changed')
    expect(sheet.cells[0][1]).toMatchObject({ oldValue: 2, newValue: 1 })
    expect(sheet.changedCount).toBe(1)
  })

  it('base에만 값이 있으면 added, compare에만 있으면 removed', () => {
    const base = wb([{ name: 'S', rowCount: 1, colCount: 2, merges: [], rows: [[cell('new'), undefined]] }])
    const compare = wb([{ name: 'S', rowCount: 1, colCount: 2, merges: [], rows: [[undefined, cell('gone')]] }])

    const diff = diffWorkbooks(base, compare)
    const sheet = diff.sheets.get('S')!
    expect(sheet.cells[0][0].status).toBe('added')
    expect(sheet.cells[0][1].status).toBe('removed')
    expect(sheet.addedCount).toBe(1)
    expect(sheet.removedCount).toBe(1)
  })

  it('양쪽 다 비어있으면 unchanged로 취급한다', () => {
    const base = wb([{ name: 'S', rowCount: 1, colCount: 1, merges: [], rows: [[undefined]] }])
    const compare = wb([{ name: 'S', rowCount: 1, colCount: 1, merges: [], rows: [[undefined]] }])

    const diff = diffWorkbooks(base, compare)
    expect(diff.sheets.get('S')!.cells[0][0].status).toBe('unchanged')
  })

  it('한쪽에만 있는 시트는 sheetsOnlyInBase/sheetsOnlyInCompare로 분리한다', () => {
    const base = wb([
      { name: 'A', rowCount: 1, colCount: 1, merges: [], rows: [[cell(1)]] },
      { name: 'B', rowCount: 1, colCount: 1, merges: [], rows: [[cell(1)]] },
    ])
    const compare = wb([
      { name: 'A', rowCount: 1, colCount: 1, merges: [], rows: [[cell(1)]] },
      { name: 'C', rowCount: 1, colCount: 1, merges: [], rows: [[cell(1)]] },
    ])

    const diff = diffWorkbooks(base, compare)
    expect(diff.sheetsOnlyInBase).toEqual(['B'])
    expect(diff.sheetsOnlyInCompare).toEqual(['C'])
    expect(diff.sheets.has('A')).toBe(true)
  })

  it('두 시트 중 더 큰 행/열 범위까지 비교한다', () => {
    const base = wb([{ name: 'S', rowCount: 1, colCount: 1, merges: [], rows: [[cell(1)]] }])
    const compare = wb([
      { name: 'S', rowCount: 2, colCount: 1, merges: [], rows: [[cell(1)], [cell('extra row')]] },
    ])

    const diff = diffWorkbooks(base, compare)
    const sheet = diff.sheets.get('S')!
    expect(sheet.rowCount).toBe(2)
    expect(sheet.cells[1][0].status).toBe('removed')
  })
})
