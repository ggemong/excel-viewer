import { describe, expect, it } from 'vitest'
import { diffWorkbooks } from './diffWorkbooks'
import type { SheetModel, WorkbookModel } from '../xlsx/types'

function cell(value: string | number | null) {
  return { address: 'A1', value, formula: null, numFmt: null, style: null, hyperlink: null }
}

/** 테스트마다 안 바뀌는 레이아웃 필드(colWidths 등)는 기본값으로 채운다. */
function sheet(partial: Pick<SheetModel, 'name' | 'rowCount' | 'colCount' | 'rows'>): SheetModel {
  return { merges: [], colWidths: [], rowHeights: [], hiddenCols: [], hiddenRows: [], frozen: null, drawings: [], skippedDrawings: {}, ...partial }
}

function wb(sheets: WorkbookModel['sheets']): WorkbookModel {
  return { fileName: 'x.xlsx', sheets, warnings: [] }
}

describe('diffWorkbooks', () => {
  it('값이 같으면 unchanged, 다르면 changed', () => {
    const base = wb([sheet({ name: 'S', rowCount: 1, colCount: 2, rows: [[cell('a'), cell(1)]] })])
    const compare = wb([sheet({ name: 'S', rowCount: 1, colCount: 2, rows: [[cell('a'), cell(2)]] })])

    const diff = diffWorkbooks(base, compare)
    const diffSheet = diff.sheets.get('S')!
    expect(diffSheet.cells[0][0].status).toBe('unchanged')
    expect(diffSheet.cells[0][1].status).toBe('changed')
    expect(diffSheet.cells[0][1]).toMatchObject({ oldValue: 2, newValue: 1 })
    expect(diffSheet.changedCount).toBe(1)
  })

  it('base에만 값이 있으면 added, compare에만 있으면 removed', () => {
    const base = wb([sheet({ name: 'S', rowCount: 1, colCount: 2, rows: [[cell('new'), undefined]] })])
    const compare = wb([sheet({ name: 'S', rowCount: 1, colCount: 2, rows: [[undefined, cell('gone')]] })])

    const diff = diffWorkbooks(base, compare)
    const diffSheet = diff.sheets.get('S')!
    expect(diffSheet.cells[0][0].status).toBe('added')
    expect(diffSheet.cells[0][1].status).toBe('removed')
    expect(diffSheet.addedCount).toBe(1)
    expect(diffSheet.removedCount).toBe(1)
  })

  it('양쪽 다 비어있으면 unchanged로 취급한다', () => {
    const base = wb([sheet({ name: 'S', rowCount: 1, colCount: 1, rows: [[undefined]] })])
    const compare = wb([sheet({ name: 'S', rowCount: 1, colCount: 1, rows: [[undefined]] })])

    const diff = diffWorkbooks(base, compare)
    expect(diff.sheets.get('S')!.cells[0][0].status).toBe('unchanged')
  })

  it('한쪽에만 있는 시트는 sheetsOnlyInBase/sheetsOnlyInCompare로 분리한다', () => {
    const base = wb([
      sheet({ name: 'A', rowCount: 1, colCount: 1, rows: [[cell(1)]] }),
      sheet({ name: 'B', rowCount: 1, colCount: 1, rows: [[cell(1)]] }),
    ])
    const compare = wb([
      sheet({ name: 'A', rowCount: 1, colCount: 1, rows: [[cell(1)]] }),
      sheet({ name: 'C', rowCount: 1, colCount: 1, rows: [[cell(1)]] }),
    ])

    const diff = diffWorkbooks(base, compare)
    expect(diff.sheetsOnlyInBase).toEqual(['B'])
    expect(diff.sheetsOnlyInCompare).toEqual(['C'])
    expect(diff.sheets.has('A')).toBe(true)
  })

  it('두 시트 중 더 큰 행/열 범위까지 비교한다', () => {
    const base = wb([sheet({ name: 'S', rowCount: 1, colCount: 1, rows: [[cell(1)]] })])
    const compare = wb([sheet({ name: 'S', rowCount: 2, colCount: 1, rows: [[cell(1)], [cell('extra row')]] })])

    const diff = diffWorkbooks(base, compare)
    const diffSheet = diff.sheets.get('S')!
    expect(diffSheet.rowCount).toBe(2)
    expect(diffSheet.cells[1][0].status).toBe('removed')
  })
})
