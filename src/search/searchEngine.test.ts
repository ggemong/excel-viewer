import { describe, expect, it } from 'vitest'
import type { DrawingItem, GroupNode, ShapeNode } from '../xlsx/drawingTypes'
import type { CellModel, SheetModel } from '../xlsx/types'
import { createMatcher, drawingText, highlightsFor, makeSnippet, MAX_HITS, searchSheet, type SearchOptions } from './searchEngine'

const cell = (value: string | number | boolean | null, numFmt: string | null = null): CellModel => ({
  address: 'A1',
  value,
  formula: null,
  numFmt,
  style: null,
  hyperlink: null,
})

function sheetOf(rows: (CellModel | undefined)[][], extra: Partial<SheetModel> = {}): SheetModel {
  const colCount = Math.max(0, ...rows.map((r) => r.length))
  return {
    name: 'S',
    rowCount: rows.length,
    colCount,
    rows,
    merges: [],
    colWidths: new Array(colCount).fill(64),
    rowHeights: new Array(rows.length).fill(20),
    hiddenCols: new Array(colCount).fill(false),
    hiddenRows: new Array(rows.length).fill(false),
    frozen: null,
    filters: [],
    drawings: [],
    skippedDrawings: {},
    ...extra,
  }
}

const opts = (query: string, extra: Partial<SearchOptions> = {}): SearchOptions => ({ query, matchCase: false, wholeCell: false, ...extra })
const ctl = { cancelled: () => false, maxHits: MAX_HITS }
const run = (sheet: SheetModel, o: SearchOptions, c = ctl) => searchSheet(sheet, 0, o, c)
const cells = (r: Awaited<ReturnType<typeof run>>) => r.hits.filter((h) => h.kind === 'cell').map((h) => (h.kind === 'cell' ? `${h.row},${h.col}` : ''))

describe('createMatcher', () => {
  it('기본은 대소문자 무시 부분 일치', () => {
    expect(createMatcher(opts('abc'))('xxABCxx')).toEqual({ start: 2, length: 3 })
    expect(createMatcher(opts('abc'))('xyz')).toBeNull()
  })
  it('대소문자 구분', () => {
    expect(createMatcher(opts('abc', { matchCase: true }))('ABC')).toBeNull()
    expect(createMatcher(opts('ABC', { matchCase: true }))('xABCx')).not.toBeNull()
  })
  it('셀 전체 일치는 부분 문자열로는 안 걸린다', () => {
    const m = createMatcher(opts('가', { wholeCell: true }))
    expect(m('가')).not.toBeNull()
    expect(m('가나')).toBeNull()
  })
})

describe('searchSheet — 셀', () => {
  it('행 우선 순서로 찾는다', async () => {
    const s = sheetOf([[cell('가'), cell('나가')], [cell('다'), cell('가다')]])
    expect(cells(await run(s, opts('가')))).toEqual(['1,1', '1,2', '2,2'])
  })

  it('보이는 글자로 찾는다: 1,000원', async () => {
    const s = sheetOf([[cell(1000, '#,##0"원"')]])
    const r = await run(s, opts('1,000원'))
    expect(cells(r)).toEqual(['1,1'])
    expect(r.hits[0]).toMatchObject({ text: '1,000원', matchStart: 0 })
  })

  it('원래 값으로도 찾는다: 1000 -> 1,000원 (보이는 글자엔 없는 부분 문자열)', async () => {
    const s = sheetOf([[cell(1000, '#,##0"원"')]])
    const r = await run(s, opts('1000'))
    expect(cells(r)).toEqual(['1,1'])
    expect(r.hits[0]).toMatchObject({ text: '1,000원', matchStart: -1 }) // 목록에는 보이는 글자를 보여주고 강조는 없음
  })

  it('날짜 셀의 원래 값(직렬번호)으로는 찾지 않는다', async () => {
    const s = sheetOf([[cell(46299, 'yyyy-mm-dd')]])
    expect(cells(await run(s, opts('46299')))).toEqual([])
    expect(cells(await run(s, opts('2026-10')))).toEqual(['1,1'])
  })

  it('불리언과 빈 셀', async () => {
    const s = sheetOf([[cell(true), undefined, cell(null)]])
    expect(cells(await run(s, opts('true')))).toEqual(['1,1'])
  })

  it('병합 셀은 왼쪽 위 칸만 센다(ExcelJS는 값을 범위 전체에 복제한다)', async () => {
    const s = sheetOf([[cell('제목'), cell('제목'), cell('제목')], [cell('x'), cell('x'), cell('제목')]], { merges: ['A1:C1'] })
    expect(cells(await run(s, opts('제목')))).toEqual(['1,1', '2,3'])
  })

  it('숨긴 행/열의 일치는 결과에서 빼고 개수만 센다', async () => {
    const s = sheetOf([[cell('a'), cell('a')], [cell('a'), cell('a')]], { hiddenRows: [false, true], hiddenCols: [false, true] })
    const r = await run(s, opts('a'))
    expect(cells(r)).toEqual(['1,1'])
    expect(r.hiddenExcluded).toBe(3)
  })

  it('결과 상한에서 멈추고 truncated로 알린다', async () => {
    const s = sheetOf([[cell('a'), cell('a'), cell('a'), cell('a')]])
    const r = await run(s, opts('a'), { ...ctl, maxHits: 2 })
    expect(r.hits).toHaveLength(2)
    expect(r.truncated).toBe(true)
  })

  it('취소를 확인하고 중간에 멈춘다', async () => {
    const big = Array.from({ length: 5000 }, () => Array.from({ length: 10 }, () => cell('a')))
    let calls = 0
    const r = await run(sheetOf(big), opts('a'), { cancelled: () => ++calls > 0, maxHits: MAX_HITS })
    expect(r.cancelled).toBe(true)
    expect(r.hits.length).toBeLessThan(50_000)
  })
})

describe('searchSheet — 도형 글자', () => {
  const shapeNode = (texts: string[][]): ShapeNode => ({
    kind: 'shape',
    preset: 'rect',
    adjust: {},
    fill: null,
    line: null,
    text: { paragraphs: texts.map((runs) => ({ align: 'left', runs: runs.map((t) => ({ text: t, sizePx: 14, bold: false, italic: false, color: null })) })), vAlign: 'top', wrap: true, insets: { l: 0, t: 0, r: 0, b: 0 } },
    rotation: 0,
    flipH: false,
    flipV: false,
  })
  const at = (col: number, row: number, node: DrawingItem['node']): DrawingItem => ({
    anchor: { kind: 'twoCell', from: { col, row, colOffset: 0, rowOffset: 0 }, to: { col: col + 2, row: row + 2, colOffset: 0, rowOffset: 0 } },
    node,
  })

  it('도형/텍스트상자 안 글자를 찾고 놓인 칸 주소를 알려준다', async () => {
    const s = sheetOf([[cell('없음')]], { drawings: [at(3, 4, shapeNode([['문구를 ', '수정해주세요']]))] })
    const r = await run(s, opts('수정'))
    expect(r.hits).toHaveLength(1)
    expect(r.hits[0]).toMatchObject({ kind: 'shape', drawingIndex: 0, location: 'D5', text: '문구를 수정해주세요' })
  })

  it('묶음 안 도형 글자는 하나로 합쳐 찾는다', () => {
    const rel = { x: 0, y: 0, w: 1, h: 1 }
    const group: GroupNode = { kind: 'group', rotation: 0, flipH: false, flipV: false, children: [{ rel, node: shapeNode([['가']]) }, { rel, node: shapeNode([['나']]) }] }
    expect(drawingText(group)).toBe('가 나')
  })

  it('글자 없는 도형/그림은 찾지 않는다', async () => {
    const empty = { ...shapeNode([]), text: null }
    const s = sheetOf([[cell('x')]], { drawings: [at(0, 0, empty)] })
    expect((await run(s, opts('x'))).hits.filter((h) => h.kind === 'shape')).toHaveLength(0)
  })
})

describe('makeSnippet / highlightsFor', () => {
  it('일치 앞뒤를 줄여서 보여준다', () => {
    const text = `${'가'.repeat(40)}찾는말${'나'.repeat(40)}`
    const s = makeSnippet({ text, matchStart: 40, matchLength: 3 })
    expect(s.match).toBe('찾는말')
    expect(s.before.startsWith('…')).toBe(true)
    expect(s.after.endsWith('…')).toBe(true)
  })
  it('일치 위치를 모르면(원래 값으로 일치) 앞부분만', () => {
    expect(makeSnippet({ text: '1,000원', matchStart: -1, matchLength: 0 })).toEqual({ before: '', match: '', after: '1,000원' })
  })
  it('현재 시트의 셀/도형만 강조 대상으로', () => {
    const hits = [
      { kind: 'cell', sheetIndex: 0, row: 2, col: 3, text: '', matchStart: 0, matchLength: 1 },
      { kind: 'cell', sheetIndex: 1, row: 9, col: 9, text: '', matchStart: 0, matchLength: 1 },
      { kind: 'shape', sheetIndex: 0, drawingIndex: 4, location: '', text: '', matchStart: 0, matchLength: 1 },
    ] as const
    const h = highlightsFor([...hits], 0)
    expect([...h.cells]).toEqual(['2,3'])
    expect([...h.shapes]).toEqual([4])
  })
})
