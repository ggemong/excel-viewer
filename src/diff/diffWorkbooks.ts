import type { CellModel, SheetModel, WorkbookModel } from '../xlsx/types'

export type CellDiffStatus = 'unchanged' | 'added' | 'removed' | 'changed'

export interface CellDiff {
  status: CellDiffStatus
  oldValue: CellModel['value'] | undefined
  newValue: CellModel['value'] | undefined
}

export interface SheetDiff {
  sheetName: string
  rowCount: number
  colCount: number
  /** cells[row][col], 0-based — base(현재 보고 있는 파일) 그리드와 같은 좌표계. */
  cells: CellDiff[][]
  changedCount: number
  addedCount: number
  removedCount: number
}

export interface WorkbookDiff {
  sheets: Map<string, SheetDiff>
  /** base에만 있는 시트 이름(비교 파일엔 없음). */
  sheetsOnlyInBase: string[]
  /** compare에만 있는 시트 이름(base엔 없어서 화면에 못 보여줌). */
  sheetsOnlyInCompare: string[]
}

function valuesEqual(a: CellModel['value'] | undefined, b: CellModel['value'] | undefined): boolean {
  return (a ?? null) === (b ?? null)
}

function diffSheet(base: SheetModel, compare: SheetModel): SheetDiff {
  const rowCount = Math.max(base.rowCount, compare.rowCount)
  const colCount = Math.max(base.colCount, compare.colCount)
  const cells: CellDiff[][] = []
  let changedCount = 0
  let addedCount = 0
  let removedCount = 0

  for (let r = 0; r < rowCount; r++) {
    const rowOut: CellDiff[] = []
    for (let c = 0; c < colCount; c++) {
      const newValue = base.rows[r]?.[c]?.value ?? undefined
      const oldValue = compare.rows[r]?.[c]?.value ?? undefined

      let status: CellDiffStatus = 'unchanged'
      if (oldValue === undefined && newValue !== undefined) {
        status = 'added'
        addedCount++
      } else if (oldValue !== undefined && newValue === undefined) {
        status = 'removed'
        removedCount++
      } else if (!valuesEqual(oldValue, newValue)) {
        status = 'changed'
        changedCount++
      }

      rowOut.push({ status, oldValue, newValue })
    }
    cells.push(rowOut)
  }

  return { sheetName: base.name, rowCount, colCount, cells, changedCount, addedCount, removedCount }
}

/**
 * base(현재 보고 있는 파일)를 기준으로 compare(비교 대상 파일)와 셀 단위로 비교한다.
 * 저장 경로(patch.ts)와는 완전히 분리된, 순수 비교 로직 — 읽기 경로가 만든
 * WorkbookModel만 본다. v1은 좌표(주소) 기준 비교라, 행이 통째로 삽입/삭제된
 * 경우 그 뒤 모든 셀이 "변경됨"으로 보일 수 있다(알려진 한계).
 */
export function diffWorkbooks(base: WorkbookModel, compare: WorkbookModel): WorkbookDiff {
  const compareByName = new Map(compare.sheets.map((s) => [s.name, s]))
  const baseNames = new Set(base.sheets.map((s) => s.name))

  const sheets = new Map<string, SheetDiff>()
  for (const sheet of base.sheets) {
    const compareSheet = compareByName.get(sheet.name)
    if (compareSheet) {
      sheets.set(sheet.name, diffSheet(sheet, compareSheet))
    }
  }

  const sheetsOnlyInBase = base.sheets.map((s) => s.name).filter((name) => !compareByName.has(name))
  const sheetsOnlyInCompare = compare.sheets.map((s) => s.name).filter((name) => !baseNames.has(name))

  return { sheets, sheetsOnlyInBase, sheetsOnlyInCompare }
}
