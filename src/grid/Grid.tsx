import { useRef, useState } from 'react'
import { useVirtualizer } from '@tanstack/react-virtual'
import { cellAddress, columnLetter } from '../xlsx/cellRef'
import { formatCellValue } from '../xlsx/formatValue'
import type { CellModel, SheetModel } from '../xlsx/types'

const ROW_HEIGHT = 34
const HEADER_HEIGHT = 32
const ROW_NUM_WIDTH = 44
const COL_WIDTH = 120

interface GridProps {
  sheet: SheetModel
  editMode: boolean
  onEditCell: (address: string, rawInput: string) => void
}

/**
 * 데스크톱 가상 스크롤 그리드. 행만 가상화한다(v1 스코프) — 대부분의 실사용
 * 파일에서 병목은 열 수가 아니라 행 수라서, 행 가상화만으로 대용량 파일의
 * 스크롤 성능 문제를 해결한다. 열이 극단적으로 많은 시트는 이후 확장 대상.
 */
export function Grid({ sheet, editMode, onEditCell }: GridProps) {
  const scrollRef = useRef<HTMLDivElement>(null)
  const [editingAddress, setEditingAddress] = useState<string | null>(null)
  const [draft, setDraft] = useState('')

  const rowVirtualizer = useVirtualizer({
    count: sheet.rowCount,
    getScrollElement: () => scrollRef.current,
    estimateSize: () => ROW_HEIGHT,
    overscan: 12,
  })

  const gridTemplateColumns = `${ROW_NUM_WIDTH}px repeat(${sheet.colCount}, ${COL_WIDTH}px)`

  const startEdit = (address: string, cell: CellModel | undefined) => {
    if (!editMode || cell?.formula) return
    setEditingAddress(address)
    setDraft(cell && cell.value !== null ? String(cell.value) : '')
  }

  const commitEdit = () => {
    if (editingAddress) onEditCell(editingAddress, draft)
    setEditingAddress(null)
  }

  const cancelEdit = () => setEditingAddress(null)

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
            const rowNum = virtualRow.index + 1
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
                <div className="grid-cell grid-cell--rownum">{rowNum}</div>
                {Array.from({ length: sheet.colCount }, (_, colIndex) => {
                  const cell = row?.[colIndex]
                  const address = cellAddress(rowNum, colIndex + 1)
                  const isNumeric = typeof cell?.value === 'number'
                  const isFormula = Boolean(cell?.formula)
                  const isEditing = editingAddress === address

                  if (isEditing) {
                    return (
                      <input
                        key={colIndex}
                        className="grid-cell grid-cell--input"
                        autoFocus
                        value={draft}
                        onChange={(e) => setDraft(e.target.value)}
                        onBlur={commitEdit}
                        onKeyDown={(e) => {
                          if (e.key === 'Enter') commitEdit()
                          else if (e.key === 'Escape') cancelEdit()
                        }}
                      />
                    )
                  }

                  return (
                    <div
                      key={colIndex}
                      className="grid-cell"
                      data-editable={editMode && !isFormula}
                      data-readonly={editMode && isFormula}
                      title={isFormula ? '수식 셀은 이 버전에서 수정할 수 없어요' : undefined}
                      style={{ justifyContent: isNumeric ? 'flex-end' : 'flex-start' }}
                      onClick={() => startEdit(address, cell)}
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
