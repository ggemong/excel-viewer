import { useMemo } from 'react'
import type { CellRange } from '../clipboard/buildClipboardPayload'
import { cellAddress } from '../xlsx/cellRef'
import type { SheetModel } from '../xlsx/types'
import { computeSummary } from './summary'

interface SummaryBarProps {
  sheet: SheetModel
  selection?: CellRange | null
  /** 필터로 걸러진 행 — 요약에서 제외한다. */
  filteredOut?: boolean[]
}

const numberFormat = new Intl.NumberFormat('ko-KR', { maximumFractionDigits: 2 })

/**
 * 셀(들)을 선택하면 — 한 칸이어도 — 그 선택 기준으로 보여준다(엑셀 상태
 * 표시줄과 같은 동작). 선택 안에 숫자가 없어도(문자만 선택) 막대 자체는
 * 유지하고 "선택된 셀 수"만 보여준다 — 선택했는데 막대가 사라지면 뭘
 * 선택했는지 감이 안 온다. 선택이 아예 없을 때만 시트 전체 기준으로
 * 돌아가고, 그때는 숫자 셀이 하나도 없으면 보여줄 게 없어 숨긴다.
 */
export function SummaryBar({ sheet, selection, filteredOut }: SummaryBarProps) {
  const summary = useMemo(() => computeSummary(sheet, selection, filteredOut), [sheet, selection, filteredOut])

  if (summary.count === 0 && !selection) {
    return null
  }

  const hasNumbers = summary.count > 0

  return (
    <div className="summary-bar">
      <span className="summary-label">
        {selection
          ? `${cellAddress(selection.r0, selection.c0)}:${cellAddress(selection.r1, selection.c1)} 요약`
          : '자동 요약'}
      </span>
      {hasNumbers ? (
        <>
          <span className="summary-chip">합계 {numberFormat.format(summary.sum)}</span>
          <span className="summary-chip">평균 {numberFormat.format(summary.avg)}</span>
          <span className="summary-chip">최고 {numberFormat.format(summary.max)}</span>
          <span className="summary-chip">숫자 셀 {summary.count}개</span>
        </>
      ) : (
        <span className="summary-chip">숫자 없음</span>
      )}
      {summary.selectedCount !== null && (
        <span className="summary-chip">선택 {summary.selectedCount}개</span>
      )}
    </div>
  )
}
