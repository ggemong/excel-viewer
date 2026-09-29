import { useRef, useState } from 'react'
import { useVirtualizer } from '@tanstack/react-virtual'
import type { SheetDiff } from '../diff/diffWorkbooks'
import { cellAddress, columnLetter } from '../xlsx/cellRef'
import { formatCellValue } from '../xlsx/formatValue'
import type { CellModel, SheetModel } from '../xlsx/types'

interface GridCardsProps {
  sheet: SheetModel
  editMode: boolean
  onEditCell: (address: string, rawInput: string) => void
  diff?: SheetDiff | null
}

/**
 * 모바일 카드뷰. 첫 번째 행을 헤더(필드 이름)로 취급하고, 그 아래 행들을
 * 각각 카드 하나로 렌더링한다 — 실사용 데이터 시트 대부분이 헤더 행을
 * 가진다는 가정을 v1 스코프로 채택했다 (구현 계획 참고).
 */
export function GridCards({ sheet, editMode, onEditCell, diff }: GridCardsProps) {
  const scrollRef = useRef<HTMLDivElement>(null)
  const dataRowCount = Math.max(sheet.rowCount - 1, 0)
  const [editingAddress, setEditingAddress] = useState<string | null>(null)
  const [draft, setDraft] = useState('')

  const headerLabels = Array.from({ length: sheet.colCount }, (_, i) => {
    const label = formatCellValue(sheet.rows[0]?.[i])
    return label || columnLetter(i + 1)
  })

  const rowVirtualizer = useVirtualizer({
    count: dataRowCount,
    getScrollElement: () => scrollRef.current,
    estimateSize: () => (editMode ? 96 + headerLabels.length * 10 : 96),
    overscan: 8,
  })

  const startEdit = (address: string, cell: CellModel | undefined) => {
    if (!editMode || cell?.formula) return
    setEditingAddress(address)
    setDraft(cell && cell.value !== null ? String(cell.value) : '')
  }

  const commitEdit = () => {
    if (editingAddress) onEditCell(editingAddress, draft)
    setEditingAddress(null)
  }

  return (
    <div className="card-list" ref={scrollRef}>
      <div style={{ height: rowVirtualizer.getTotalSize(), position: 'relative' }}>
        {rowVirtualizer.getVirtualItems().map((virtualRow) => {
          const rowNum = virtualRow.index + 2 // 1-based, +1 for header row
          const row = sheet.rows[rowNum - 1]
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
                const cellDiff = diff?.cells[rowNum - 1]?.[colIndex]
                const hasDiff = cellDiff && cellDiff.status !== 'unchanged'
                if (!editMode && !hasDiff && (!cell || cell.value === null)) return null

                const address = cellAddress(rowNum, colIndex + 1)
                const isEditing = editingAddress === address

                if (isEditing) {
                  return (
                    <div className="row-card-field" key={colIndex}>
                      <span className="row-card-field-label">{label}</span>
                      <input
                        className="grid-cell--input"
                        style={{ width: '60%', textAlign: 'right' }}
                        autoFocus
                        value={draft}
                        onChange={(e) => setDraft(e.target.value)}
                        onBlur={commitEdit}
                        onKeyDown={(e) => {
                          if (e.key === 'Enter') commitEdit()
                          else if (e.key === 'Escape') setEditingAddress(null)
                        }}
                      />
                    </div>
                  )
                }

                return (
                  <div
                    className="row-card-field"
                    key={colIndex}
                    data-editable={editMode && !cell?.formula}
                    data-diff={hasDiff ? cellDiff.status : undefined}
                    onClick={() => startEdit(address, cell)}
                  >
                    <span className="row-card-field-label">{label}</span>
                    <span className="row-card-field-value">{formatCellValue(cell) || (editMode ? '—' : '')}</span>
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
