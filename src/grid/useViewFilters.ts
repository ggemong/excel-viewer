import { useCallback, useMemo, useState } from 'react'
import type { SheetFilter } from '../xlsx/filter'
import type { SheetModel, WorkbookModel } from '../xlsx/types'
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
  type ValueEntry,
} from './viewFilter'

interface FilterState {
  /** 이 상태가 어느 파일(통합문서 객체)의 것인지 — 다른 파일을 열면 자동으로 무효가 된다. */
  owner: WorkbookModel | null
  /** 시트 이름 -> 필터 선택. 항목이 없으면 그 시트는 아직 건드리지 않은 것(파일 상태 그대로). */
  bySheet: Map<string, ColumnSelections>
}

export interface ViewFilters {
  /** 필터 숨김이 반영된 시트(Grid가 그린다). 필터가 없거나 안 건드렸으면 원본 그대로. 시트가 없으면 null. */
  viewSheet: SheetModel | null
  /** 필터로 걸러진 행 — 복사·요약에서 제외한다. */
  filteredOut: boolean[]
  /** 필터 버튼을 "활성"으로 강조할 열. */
  activeCols: Set<number>
  /** 사용자가 이 시트의 필터를 건드린 상태인가(파일 상태로 되돌리기 가능 여부). */
  touched: boolean
  stats: { total: number; shown: number } | null
  entriesFor: (col: number) => { entries: ValueEntry[]; filter: SheetFilter } | null
  toggle: (col: number, value: string) => void
  setListed: (col: number, values: string[], include: boolean) => void
  clearColumn: (col: number) => void
  clearAll: () => void
  resetToFile: () => void
}

/**
 * 시트별 보기 전용 열 필터 상태. 계산은 viewFilter.ts(순수 함수)가 하고 여기는 상태 보관과
 * React 연결만 한다. 필터 상태는 파일을 바꾸면(workbook 객체가 바뀌면) 버린다.
 */
export function useViewFilters(workbook: WorkbookModel | null, sheet: SheetModel | null): ViewFilters {
  const [state, setState] = useState<FilterState>({ owner: null, bySheet: new Map() })

  const filter = sheet?.filters[0] ?? null
  const sheetName = sheet?.name ?? ''
  const selections = state.owner === workbook ? state.bySheet.get(sheetName) : undefined

  const commit = useCallback(
    (next: ColumnSelections | undefined) => {
      setState((prev) => {
        const bySheet = prev.owner === workbook ? new Map(prev.bySheet) : new Map<string, ColumnSelections>()
        if (next) bySheet.set(sheetName, next)
        else bySheet.delete(sheetName)
        return { owner: workbook, bySheet }
      })
    },
    [workbook, sheetName],
  )

  /** 처음 건드릴 때는 파일에 저장된 조건을 선택값으로 복원한 뒤 수정한다. */
  const editable = useCallback((): ColumnSelections => {
    if (selections) return new Map(selections)
    return sheet && filter ? initialSelections(sheet, filter) : new Map()
  }, [selections, sheet, filter])

  const view = useMemo(() => {
    if (!sheet || !filter) return { viewSheet: sheet, filteredOut: new Array<boolean>(sheet?.rowCount ?? 0).fill(false) }
    if (!selections) return { viewSheet: sheet, filteredOut: fileFilteredOut(sheet, filter) }
    const applied = applySelections(sheet, filter, selections)
    return { viewSheet: { ...sheet, hiddenRows: applied.hiddenRows }, filteredOut: applied.filteredOut }
  }, [sheet, filter, selections])

  const activeCols = useMemo(() => new Set<number>(selections ? selections.keys() : (filter?.activeCols ?? [])), [selections, filter])

  const stats = useMemo(() => (filter ? countShownRows(filter, view.filteredOut) : null), [filter, view.filteredOut])

  const entriesFor = useCallback(
    (col: number) => {
      const colFilter = sheet ? findFilterForColumn(sheet, col) : null
      if (!sheet || !colFilter) return null
      // 목록의 체크 상태는 "건드린 뒤의 선택값" 기준 — 처음 열었을 때도 파일의 조건이 체크에 반영돼야 한다.
      const effective = selections ?? initialSelections(sheet, colFilter)
      return { entries: columnEntries(sheet, colFilter, effective, col), filter: colFilter }
    },
    [sheet, selections],
  )

  const update = useCallback(
    (col: number, compute: (universe: Set<string>, current: Set<string> | undefined) => Set<string> | null) => {
      if (!sheet || !filter) return
      const next = editable()
      const result = compute(columnUniverse(sheet, filter, col), next.get(col))
      if (result) next.set(col, result)
      else next.delete(col)
      commit(next)
    },
    [sheet, filter, editable, commit],
  )

  return {
    viewSheet: view.viewSheet,
    filteredOut: view.filteredOut,
    activeCols,
    touched: Boolean(selections),
    stats,
    entriesFor,
    toggle: (col, value) => update(col, (universe, current) => toggleValue(universe, current, value)),
    setListed: (col, values, include) => update(col, (universe, current) => setValues(universe, current, values, include)),
    clearColumn: (col) => update(col, () => null),
    clearAll: () => commit(new Map()),
    resetToFile: () => commit(undefined),
  }
}
