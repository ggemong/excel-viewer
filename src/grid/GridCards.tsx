import { useRef } from 'react'
import { useVirtualizer } from '@tanstack/react-virtual'
import { columnLetter } from '../xlsx/cellRef'
import { formatCellValue } from '../xlsx/formatValue'
import type { SheetModel } from '../xlsx/types'

interface GridCardsProps {
  sheet: SheetModel
}

/**
 * 모바일 카드뷰. 첫 번째 행을 헤더(필드 이름)로 취급하고, 그 아래 행들을
 * 각각 카드 하나로 렌더링한다 — 실사용 데이터 시트 대부분이 헤더 행을
 * 가진다는 가정을 v1 스코프로 채택했다 (구현 계획 참고).
 */
export function GridCards({ sheet }: GridCardsProps) {
  const scrollRef = useRef<HTMLDivElement>(null)
  const dataRowCount = Math.max(sheet.rowCount - 1, 0)

  const headerLabels = Array.from({ length: sheet.colCount }, (_, i) => {
    const label = formatCellValue(sheet.rows[0]?.[i])
    return label || columnLetter(i + 1)
  })

  const rowVirtualizer = useVirtualizer({
    count: dataRowCount,
    getScrollElement: () => scrollRef.current,
    estimateSize: () => 96,
    overscan: 8,
  })

  return (
    <div className="card-list" ref={scrollRef}>
      <div style={{ height: rowVirtualizer.getTotalSize(), position: 'relative' }}>
        {rowVirtualizer.getVirtualItems().map((virtualRow) => {
          const row = sheet.rows[virtualRow.index + 1]
          return (
            <div
              key={virtualRow.key}
              className="row-card"
              style={{
                position: 'absolute',
                top: 0,
                left: 0,
                right: 0,
                transform: `translateY(${virtualRow.start}px)`,
              }}
            >
              {headerLabels.map((label, colIndex) => {
                const cell = row?.[colIndex]
                if (!cell || cell.value === null) return null
                return (
                  <div className="row-card-field" key={colIndex}>
                    <span className="row-card-field-label">{label}</span>
                    <span className="row-card-field-value">{formatCellValue(cell)}</span>
                  </div>
                )
              })}
            </div>
          )
        })}
      </div>
    </div>
  )
}
