import { describe, expect, it } from 'vitest'
import type { SheetFilter } from '../xlsx/filter'
import type { CellModel, SheetModel } from '../xlsx/types'
import {
  applySelections,
  columnEntries,
  columnUniverse,
  countShownRows,
  fileFilteredOut,
  findFilterForColumn,
  initialSelections,
  setValues,
  toggleValue,
  type ColumnSelections,
} from './viewFilter'

const cell = (value: string | number | null, numFmt: string | null = null): CellModel => ({
  address: 'A1',
  value,
  formula: null,
  numFmt,
  style: null,
  hyperlink: null,
})

/** 1행 헤더(구분/금액) + 데이터 6행. A열 값: 가,나,가,나,가,(빈) / B열 숫자(서식 있음) */
function makeSheet(hiddenRows: boolean[] = new Array(7).fill(false)): SheetModel {
  const data: [string | null, number][] = [['가', 1000], ['나', 2000], ['가', 3000], ['나', 1000], ['가', 2000], [null, 1000]]
  return {
    name: 'S',
    rowCount: 7,
    colCount: 2,
    rows: [[cell('구분'), cell('금액')], ...data.map(([a, b]) => [cell(a), cell(b, '#,##0"원"')])],
    merges: [],
    colWidths: [80, 80],
    rowHeights: new Array(7).fill(20),
    hiddenCols: [false, false],
    hiddenRows,
    frozen: null,
    filters: [{ headerRow: 1, lastRow: 7, firstCol: 1, lastCol: 2, hiddenButtonCols: [], activeCols: [] }],
    conditionalFormats: [],
    skippedConditionalFormats: {},
    drawings: [],
    skippedDrawings: {},
  }
}
const filterOf = (s: SheetModel): SheetFilter => s.filters[0]
const sel = (entries: [number, string[]][]): ColumnSelections => new Map(entries.map(([c, v]) => [c, new Set(v)]))

describe('값 목록', () => {
  it('화면에 보이는 문자열 기준으로 묶고 개수를 센다(1,000원 처럼)', () => {
    const s = makeSheet()
    const entries = columnEntries(s, filterOf(s), new Map(), 2)
    expect(entries.map((e) => [e.value, e.count])).toEqual([['1,000원', 3], ['2,000원', 2], ['3,000원', 1]])
  })

  it('빈 값은 목록 맨 뒤', () => {
    const s = makeSheet()
    const entries = columnEntries(s, filterOf(s), new Map(), 1)
    expect(entries.map((e) => e.value)).toEqual(['가', '나', ''])
  })

  it('숫자가 든 문자열은 자연스러운 순서로 정렬한다(2 < 10)', () => {
    const s = makeSheet()
    s.rows = [[cell('h')], ...['10', '2', '1'].map((v) => [cell(v)])]
    s.rowCount = 4
    s.hiddenRows = [false, false, false, false]
    s.filters = [{ headerRow: 1, lastRow: 4, firstCol: 1, lastCol: 1, hiddenButtonCols: [], activeCols: [] }]
    expect(columnEntries(s, s.filters[0], new Map(), 1).map((e) => e.value)).toEqual(['1', '2', '10'])
  })

  it('다른 열 필터를 통과한 행의 값만 보인다(연쇄)', () => {
    const s = makeSheet()
    // A열을 '가'만 → B열 목록은 가의 행(1000,3000,2000)만
    const entries = columnEntries(s, filterOf(s), sel([[1, ['가']]]), 2)
    expect(entries.map((e) => [e.value, e.count])).toEqual([['1,000원', 1], ['2,000원', 1], ['3,000원', 1]])
  })

  it('자기 열의 선택은 목록을 줄이지 않고 체크 상태만 바꾼다', () => {
    const s = makeSheet()
    const entries = columnEntries(s, filterOf(s), sel([[1, ['가']]]), 1)
    expect(entries.map((e) => [e.value, e.included])).toEqual([['가', true], ['나', false], ['', false]])
  })
})

describe('행 숨김 계산', () => {
  it('선택한 값이 아닌 행만 숨기고, 헤더 행과 범위 밖은 건드리지 않는다', () => {
    const s = makeSheet()
    const view = applySelections(s, filterOf(s), sel([[1, ['가']]]))
    expect(view.hiddenRows).toEqual([false, false, true, false, true, false, true])
    expect(view.filteredOut).toEqual([false, false, true, false, true, false, true])
  })

  it('여러 열 필터는 AND', () => {
    const s = makeSheet()
    const view = applySelections(s, filterOf(s), sel([[1, ['가']], [2, ['1,000원']]]))
    expect(view.filteredOut.filter(Boolean)).toHaveLength(5) // 데이터 6행 중 1행만 통과
    expect(view.filteredOut[1]).toBe(false) // 2행(가,1000)
  })

  it('선택을 건드리면 파일에서 숨겨졌던 행도 다시 계산된다(모두 선택으로 되살릴 수 있다)', () => {
    const s = makeSheet([false, true, true, false, false, false, false])
    const view = applySelections(s, filterOf(s), new Map())
    expect(view.hiddenRows.every((h) => !h)).toBe(true)
  })

  it('필터 범위 밖 행의 파일 숨김 상태는 유지', () => {
    const s = makeSheet([false, false, false, false, false, false, true])
    s.rowCount = 7
    s.filters = [{ ...filterOf(s), lastRow: 5 }]
    const view = applySelections(s, s.filters[0], new Map())
    expect(view.hiddenRows[6]).toBe(true)
    expect(view.filteredOut[6]).toBe(false)
  })

  it('선택 전(파일 그대로)에는 범위 안의 숨김 행이 필터로 걸러진 행', () => {
    const s = makeSheet([false, false, true, false, false, false, false])
    expect(fileFilteredOut(s, filterOf(s))).toEqual([false, false, true, false, false, false, false])
  })
})

describe('파일에 저장된 필터 조건 복원', () => {
  it('조건이 걸려 있던 열의 선택을 지금 보이는 행의 값으로 만든다', () => {
    // 파일: A열 필터링 중, '나' 행(3,5행)이 숨겨져 있음
    const s = makeSheet([false, false, true, false, true, false, false])
    s.filters = [{ ...filterOf(s), activeCols: [1] }]
    const selections = initialSelections(s, s.filters[0])
    expect([...selections.get(1)!].sort()).toEqual(['', '가'])
    // 복원한 선택을 적용해도 같은 행이 숨겨진다(첫 클릭에 화면이 갑자기 바뀌지 않음)
    expect(applySelections(s, s.filters[0], selections).hiddenRows).toEqual([false, false, true, false, true, false, false])
  })
})

describe('선택 변경', () => {
  const s = makeSheet()
  const universe = columnUniverse(s, filterOf(s), 1)

  it('필터 없음 상태에서 하나를 끄면 나머지 전부가 선택된다', () => {
    expect([...toggleValue(universe, undefined, '나')!].sort()).toEqual(['', '가'])
  })
  it('전부 다시 켜지면 필터가 사라진다(null)', () => {
    const after = toggleValue(universe, undefined, '나')!
    expect(toggleValue(universe, after, '나')).toBeNull()
  })
  it('목록 전체 해제/선택', () => {
    expect([...setValues(universe, undefined, ['가', '나', ''], false)!]).toEqual([])
    expect(setValues(universe, new Set(), ['가', '나', ''], true)).toBeNull()
  })
  it('모두 해제하면 보이는 행이 없다', () => {
    const view = applySelections(s, filterOf(s), sel([[1, []]]))
    expect(view.filteredOut.filter(Boolean)).toHaveLength(6)
  })
})

describe('보조', () => {
  it('열이 속한 필터 찾기 / 표시 행 수', () => {
    const s = makeSheet()
    expect(findFilterForColumn(s, 2)).toBe(s.filters[0])
    expect(findFilterForColumn(s, 3)).toBeNull()
    const view = applySelections(s, filterOf(s), sel([[1, ['가']]]))
    expect(countShownRows(filterOf(s), view.filteredOut)).toEqual({ total: 6, shown: 3 })
  })
})
