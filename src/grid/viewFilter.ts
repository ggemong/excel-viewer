/**
 * 보기 전용 열 필터 — 헤더의 ▼를 눌러 값을 골라 행을 숨기는 기능의 순수 계산부.
 *
 * 파일은 절대 건드리지 않는다(이 앱은 읽기 전용, D-006). 필터는 화면에서만 행을 숨기는
 * "보기 상태"다. 선택한 값 집합만 상태로 들고, 어떤 행이 보이는지는 매번 여기서 계산한다.
 *
 * 모델: 열마다 "포함할 값(화면에 보이는 문자열)의 집합". 항목이 없는 열은 필터 없음(전부 포함).
 * 값을 셀의 원본이 아니라 **화면에 보이는 문자열**로 비교하는 건 Excel과 같다 — 사용자는 `1,000원`을
 * 보고 고르지 내부 숫자 1000을 고르지 않는다.
 *
 * 파일에 이미 저장된 필터 결과(조건에 안 맞아 hidden으로 저장된 행)와의 관계: 사용자가 필터를 처음
 * 건드리기 전(selections 없음)에는 파일의 숨김 상태를 그대로 보여주고, 건드리는 순간부터는
 * 필터 범위 안의 행 표시 여부를 전부 이 선택값으로 다시 계산한다. 그래야 "모두 선택"으로 파일에서
 * 걸러져 있던 행도 다시 볼 수 있다. 이때 파일에서 조건이 걸려 있던 열은 현재 보이는 행의 값으로
 * 선택을 복원한다(initialSelections) — 그렇지 않으면 처음 클릭하는 순간 화면이 갑자기 바뀐다.
 * 대가: 필터 범위 안에서 수동으로 숨긴 행도 이때 같이 풀린다(파일만으로는 둘을 구분할 수 없다).
 */
import type { SheetFilter } from '../xlsx/filter'
import { formatCellValue } from '../xlsx/formatValue'
import type { SheetModel } from '../xlsx/types'

export type ColumnSelections = Map<number, Set<string>>

export interface ValueEntry {
  /** 화면에 보이는 문자열. 빈 셀은 ''. */
  value: string
  /** 다른 열 필터를 통과한 행 중 이 값의 개수. */
  count: number
  included: boolean
}

export interface FilterView {
  hiddenRows: boolean[]
  /** 필터 때문에 숨겨진 행(필터 범위 안의 숨김 행). 복사·요약에서 제외하는 대상. */
  filteredOut: boolean[]
}

/** 열별 표시 문자열 캐시 — 체크박스를 누를 때마다 전체 셀을 다시 포맷하지 않기 위해서다. 키는 rows 배열(보기용 복사본과 공유). */
const displayCache = new WeakMap<SheetModel['rows'], Map<string, string[]>>()

const collator = new Intl.Collator('ko', { numeric: true })

export function findFilterForColumn(sheet: SheetModel, col: number): SheetFilter | null {
  return sheet.filters.find((f) => col >= f.firstCol && col <= f.lastCol) ?? null
}

/** 필터 범위의 데이터 행 수(헤더 제외). */
export function dataRowCount(filter: SheetFilter): number {
  return Math.max(0, filter.lastRow - filter.headerRow)
}

function columnDisplays(sheet: SheetModel, filter: SheetFilter, col: number): string[] {
  let perSheet = displayCache.get(sheet.rows)
  if (!perSheet) {
    perSheet = new Map()
    displayCache.set(sheet.rows, perSheet)
  }
  const key = `${col}:${filter.headerRow}:${filter.lastRow}`
  let cached = perSheet.get(key)
  if (!cached) {
    cached = []
    for (let r = filter.headerRow + 1; r <= filter.lastRow; r++) {
      cached.push(formatCellValue(sheet.rows[r - 1]?.[col - 1]))
    }
    perSheet.set(key, cached)
  }
  return cached
}

/** 각 데이터 행이 (exceptCol을 뺀) 모든 열 선택을 통과하는지. 인덱스 0 = headerRow+1행. */
function evaluateRows(sheet: SheetModel, filter: SheetFilter, selections: ColumnSelections, exceptCol?: number): boolean[] {
  const total = dataRowCount(filter)
  const pass = new Array<boolean>(total).fill(true)
  for (const [col, included] of selections) {
    if (col === exceptCol) continue
    const values = columnDisplays(sheet, filter, col)
    for (let i = 0; i < total; i++) {
      if (pass[i] && !included.has(values[i])) pass[i] = false
    }
  }
  return pass
}

/** 선택값을 적용했을 때의 행 표시 상태. */
export function applySelections(sheet: SheetModel, filter: SheetFilter, selections: ColumnSelections): FilterView {
  const pass = evaluateRows(sheet, filter, selections)
  const hiddenRows = sheet.hiddenRows.slice()
  const filteredOut = new Array<boolean>(sheet.rowCount).fill(false)
  pass.forEach((ok, i) => {
    const index = filter.headerRow + i // 0-based 행 인덱스 = (headerRow+1+i) - 1
    if (index >= hiddenRows.length) return
    hiddenRows[index] = !ok
    filteredOut[index] = !ok
  })
  return { hiddenRows, filteredOut }
}

/** 선택값이 없을 때(파일에 저장된 상태 그대로): 필터 범위 안의 숨김 행이 곧 "필터로 걸러진 행". */
export function fileFilteredOut(sheet: SheetModel, filter: SheetFilter | null): boolean[] {
  const out = new Array<boolean>(sheet.rowCount).fill(false)
  if (!filter) return out
  for (let r = filter.headerRow + 1; r <= Math.min(filter.lastRow, sheet.rowCount); r++) {
    out[r - 1] = Boolean(sheet.hiddenRows[r - 1])
  }
  return out
}

/** 파일에서 조건이 걸려 있던 열들의 선택을 "지금 보이는 행의 값"으로 복원한다. */
export function initialSelections(sheet: SheetModel, filter: SheetFilter): ColumnSelections {
  const selections: ColumnSelections = new Map()
  for (const col of filter.activeCols) {
    const values = columnDisplays(sheet, filter, col)
    const included = new Set<string>()
    values.forEach((v, i) => {
      if (!sheet.hiddenRows[filter.headerRow + i]) included.add(v)
    })
    selections.set(col, included)
  }
  return selections
}

/** 한 열의 모든 데이터 값(필터와 무관한 전체). "전부 선택됨" 판정의 기준이다. */
export function columnUniverse(sheet: SheetModel, filter: SheetFilter, col: number): Set<string> {
  return new Set(columnDisplays(sheet, filter, col))
}

/**
 * 체크박스 목록 — 이 열을 뺀 다른 열 필터를 통과한 행의 값만 보여준다(Excel과 같은 연쇄 동작).
 * 정렬은 숫자 인식 오름차순, 빈 값은 맨 뒤.
 */
export function columnEntries(sheet: SheetModel, filter: SheetFilter, selections: ColumnSelections, col: number): ValueEntry[] {
  const values = columnDisplays(sheet, filter, col)
  const pass = evaluateRows(sheet, filter, selections, col)
  const included = selections.get(col)
  const counts = new Map<string, number>()
  values.forEach((v, i) => {
    if (pass[i]) counts.set(v, (counts.get(v) ?? 0) + 1)
  })
  return [...counts.entries()]
    .map(([value, count]) => ({ value, count, included: included ? included.has(value) : true }))
    .sort((a, b) => (a.value === '' ? 1 : b.value === '' ? -1 : collator.compare(a.value, b.value)))
}

/** 선택이 전체 값을 다 포함하면 "필터 없음"(null)으로 돌려, 불필요한 필터 상태가 남지 않게 한다. */
function normalize(universe: Set<string>, included: Set<string>): Set<string> | null {
  for (const v of universe) if (!included.has(v)) return included
  return null
}

/** 값 하나를 켜고 끈 새 선택. 전부 포함이 되면 null(필터 해제). */
export function toggleValue(universe: Set<string>, current: Set<string> | undefined, value: string): Set<string> | null {
  const next = new Set(current ?? universe)
  if (next.has(value)) next.delete(value)
  else next.add(value)
  return normalize(universe, next)
}

/** 목록에 보이는 값들을 한꺼번에 켜거나 끈 새 선택("모두 선택"). */
export function setValues(universe: Set<string>, current: Set<string> | undefined, values: string[], include: boolean): Set<string> | null {
  const next = new Set(current ?? universe)
  for (const v of values) {
    if (include) next.add(v)
    else next.delete(v)
  }
  return normalize(universe, next)
}

export function countShownRows(filter: SheetFilter, filteredOut: boolean[]): { total: number; shown: number } {
  const total = dataRowCount(filter)
  let hidden = 0
  for (let r = filter.headerRow + 1; r <= filter.lastRow; r++) if (filteredOut[r - 1]) hidden++
  return { total, shown: total - hidden }
}
