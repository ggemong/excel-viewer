import { useRef } from 'react'
import { useVirtualizer } from '@tanstack/react-virtual'
import { columnLetter } from '../xlsx/cellRef'
import { formatCellValue } from '../xlsx/formatValue'
import type { SheetModel } from '../xlsx/types'

const ROW_HEIGHT = 34
const HEADER_HEIGHT = 32
const ROW_NUM_WIDTH = 44
const COL_WIDTH = 120

interface GridProps {
  sheet: SheetModel
}

/**
 * 데스크톱 가상 스크롤 그리드. 행만 가상화한다(v1 스코프) — 대부분의 실사용
 * 파일에서 병목은 열 수가 아니라 행 수라서, 행 가상화만으로 대용량 파일의
 * 스크롤 성능 문제를 해결한다. 열이 극단적으로 많은 시트는 이후 확장 대상.
 */
export function Grid({ sheet }: GridProps) {
  const scrollRef = useRef<HTMLDivElement>(null)

  const rowVirtualizer = useVirtualizer({
    count: sheet.rowCount,
    getScrollElement: () => scrollRef.current,
    estimateSize: () => ROW_HEIGHT,
    overscan: 12,
  })

  const gridTemplateColumns = `${ROW_NUM_WIDTH}px repeat(${sheet.colCount}, ${COL_WIDTH}px)`

  return (
    <div className="grid-card">
      <div className="grid-scroll" ref={scrollRef}>
        <div style={{ display: 'grid', gridTemplateColumns, position: 'sticky', top: 0, zIndex: 2 }}>
          <div className="grid-cell grid-cell--colhead" style={{ height: HEADER_HEIGHT }} />
          {Array.from({ length: sheet.colCount }, (_, i) => (
            <div key={i} className="grid-cell grid-cell--colhead" style={{ height: HEADER_HEIGHT }}>
              {columnLetter(i + 1)}
            </div>
          ))}
        </div>

        <div style={{ height: rowVirtualizer.getTotalSize(), position: 'relative' }}>
          {rowVirtualizer.getVirtualItems().map((virtualRow) => {
            const row = sheet.rows[virtualRow.index]
            return (
              <div
                key={virtualRow.key}
                style={{
                  display: 'grid',
                  gridTemplateColumns,
                  position: 'absolute',
                  top: 0,
                  left: 0,
                  width: '100%',
                  transform: `translateY(${virtualRow.start}px)`,
                }}
              >
                <div className="grid-cell grid-cell--rownum">{virtualRow.index + 1}</div>
                {Array.from({ length: sheet.colCount }, (_, colIndex) => {
                  const cell = row?.[colIndex]
                  const isNumeric = typeof cell?.value === 'number'
                  return (
                    <div
                      key={colIndex}
                      className="grid-cell"
                      style={{ justifyContent: isNumeric ? 'flex-end' : 'flex-start' }}
                    >
                      {formatCellValue(cell)}
                    </div>
                  )
                })}
              </div>
            )
          })}
        </div>
      </div>
    </div>
  )
}
