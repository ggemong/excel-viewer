import { useMemo } from 'react'
import type { SheetModel } from '../xlsx/types'

interface SummaryBarProps {
  sheet: SheetModel
}

function computeSummary(sheet: SheetModel) {
  let sum = 0
  let max = -Infinity
  let count = 0

  for (const row of sheet.rows) {
    for (const cell of row) {
      if (typeof cell?.value === 'number') {
        sum += cell.value
        if (cell.value > max) max = cell.value
        count++
      }
    }
  }

  return {
    sum,
    avg: count > 0 ? sum / count : 0,
    max: count > 0 ? max : 0,
    count,
  }
}

const numberFormat = new Intl.NumberFormat('ko-KR', { maximumFractionDigits: 2 })

export function SummaryBar({ sheet }: SummaryBarProps) {
  const summary = useMemo(() => computeSummary(sheet), [sheet])

  if (summary.count === 0) {
    return null
  }

  return (
    <div className="summary-bar">
      <span className="summary-label">자동 요약</span>
      <span className="summary-chip">합계 {numberFormat.format(summary.sum)}</span>
      <span className="summary-chip">평균 {numberFormat.format(summary.avg)}</span>
      <span className="summary-chip">최고 {numberFormat.format(summary.max)}</span>
      <span className="summary-chip">숫자 셀 {summary.count}개</span>
    </div>
  )
}
