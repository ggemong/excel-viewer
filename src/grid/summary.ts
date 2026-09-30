import type { CellRange } from '../clipboard/buildClipboardPayload'
import type { SheetModel } from '../xlsx/types'

export interface RangeSummary {
  sum: number
  avg: number
  max: number
  /** 숫자 값이 든 셀 개수(합계/평균 계산 대상). */
  count: number
  /** 선택 범위 전체의 칸 수(빈 칸·문자 칸 포함) — range를 줬을 때만 값이 있다. */
  selectedCount: number | null
}

/**
 * 숫자 셀(수식 결과값 포함)에 대한 합계/평균/최고/개수. range를 주면 그 범위만,
 * 안 주면 시트 전체를 본다 — 그리드에서 블록을 선택하면 그 선택 범위로, 아니면
 * 시트 전체로 자동 요약바가 전환된다.
 */
export function computeSummary(sheet: SheetModel, range?: CellRange | null): RangeSummary {
  let sum = 0
  let max = -Infinity
  let count = 0

  const r0 = range?.r0 ?? 1
  const r1 = range?.r1 ?? sheet.rowCount
  const c0 = range?.c0 ?? 1
  const c1 = range?.c1 ?? sheet.colCount

  for (let r = r0; r <= r1; r++) {
    const row = sheet.rows[r - 1]
    if (!row) continue
    for (let c = c0; c <= c1; c++) {
      const value = row[c - 1]?.value
      if (typeof value === 'number') {
        sum += value
        if (value > max) max = value
        count++
      }
    }
  }

  return {
    sum,
    avg: count > 0 ? sum / count : 0,
    max: count > 0 ? max : 0,
    count,
    selectedCount: range ? (r1 - r0 + 1) * (c1 - c0 + 1) : null,
  }
}
