import { useRef, useState } from 'react'
import { useVirtualizer } from '@tanstack/react-virtual'
import type { SheetDiff } from '../diff/diffWorkbooks'
import { isMergeMaster, useMergeLookup } from './useMergeLookup'
import { cellAddress, columnLetter } from '../xlsx/cellRef'
import { formatCellValue } from '../xlsx/formatValue'
import type { CellModel, SheetModel } from '../xlsx/types'

// 가상 스크롤용 카드 높이 추정치일 뿐 실제 렌더 높이와 정확히 일치할 필요는 없다
// (Grid.tsx의 ROW_HEIGHT와 달리 카드 높이는 CSS가 아니라 내용에 따라 자연스럽게
// 정해지므로, react-virtual이 실측 후 알아서 보정한다) — 그래도 이름 없는 숫자로
// 흩어놓지 않도록 상수로 뺀다.
const CARD_BASE_HEIGHT = 96
const CARD_FIELD_HEIGHT = 10

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
  const mergeLookup = useMergeLookup(sheet.merges)

  const headerLabels = Array.from({ length: sheet.colCount }, (_, i) => {
    const label = formatCellValue(sheet.rows[0]?.[i])
    return label || columnLetter(i + 1)
  })

  const rowVirtualizer = useVirtualizer({
    count: dataRowCount,
    getScrollElement: () => scrollRef.current,
    estimateSize: () => (editMode ? CARD_BASE_HEIGHT + headerLabels.length * CARD_FIELD_HEIGHT : CARD_BASE_HEIGHT),
    overscan: 8,
  })

  const startEdit = (address: string, cell: CellModel | undefined, merged: boolean) => {
    if (!editMode || cell?.formula || merged) return
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
                const col = colIndex + 1
                const merge = mergeLookup.get(`${rowNum},${col}`)
                // 병합 범위에서 마스터(왼쪽 위)가 아닌 칸은 카드에 아예 안 보여준다 —
                // ExcelJS가 마스터 값을 나머지 칸에도 그대로 돌려줘서, 안 걸러내면
                // 같은 값이 필드마다 중복으로 찍힌다.
                if (merge && !isMergeMaster(merge, rowNum, col)) return null

                const cell = row?.[colIndex]
                const cellDiff = diff?.cells[rowNum - 1]?.[colIndex]
                const hasDiff = cellDiff && cellDiff.status !== 'unchanged'
                if (!editMode && !hasDiff && (!cell || cell.value === null)) return null

                const address = cellAddress(rowNum, col)
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
                    data-editable={editMode && !cell?.formula && !merge}
                    data-diff={hasDiff ? cellDiff.status : undefined}
                    title={merge ? '병합된 셀은 이 버전에서 수정할 수 없어요' : undefined}
                    onClick={() => startEdit(address, cell, Boolean(merge))}
                  >
                    <span className="row-card-field-label">
                      {label}
                      {merge && merge.c1 > merge.c0 ? ` (${merge.c1 - merge.c0 + 1}칸 병합)` : ''}
                    </span>
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
