import { useRef, useState, type ReactNode } from 'react'
import { useVirtualizer } from '@tanstack/react-virtual'
import { normalizeRange, type CellRange } from '../clipboard/buildClipboardPayload'
import { useClipboardCopy } from '../clipboard/useClipboardCopy'
import type { SheetDiff } from '../diff/diffWorkbooks'
import { isMergeMaster, useMergeLookup } from './useMergeLookup'
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
  diff?: SheetDiff | null
}

function inRange(range: CellRange | null, row: number, col: number): boolean {
  return !!range && row >= range.r0 && row <= range.r1 && col >= range.c0 && col <= range.c1
}

/**
 * 데스크톱 가상 스크롤 그리드. 행만 가상화한다(v1 스코프) — 대부분의 실사용
 * 파일에서 병목은 열 수가 아니라 행 수라서, 행 가상화만으로 대용량 파일의
 * 스크롤 성능 문제를 해결한다. 열이 극단적으로 많은 시트는 이후 확장 대상.
 */
export function Grid({ sheet, editMode, onEditCell, diff }: GridProps) {
  const scrollRef = useRef<HTMLDivElement>(null)
  const [editingAddress, setEditingAddress] = useState<string | null>(null)
  const [draft, setDraft] = useState('')
  const [anchor, setAnchor] = useState<{ row: number; col: number } | null>(null)
  const [focus, setFocus] = useState<{ row: number; col: number } | null>(null)
  const { copyRange, copied } = useClipboardCopy()
  const mergeLookup = useMergeLookup(sheet.merges)

  const rowVirtualizer = useVirtualizer({
    count: sheet.rowCount,
    getScrollElement: () => scrollRef.current,
    estimateSize: () => ROW_HEIGHT,
    overscan: 12,
  })

  const gridTemplateColumns = `${ROW_NUM_WIDTH}px repeat(${sheet.colCount}, ${COL_WIDTH}px)`
  const selection = anchor && focus ? normalizeRange(anchor, focus) : null

  const selectCell = (row: number, col: number, extend: boolean) => {
    if (extend && anchor) {
      setFocus({ row, col })
    } else {
      setAnchor({ row, col })
      setFocus({ row, col })
    }
  }

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

  const handleCopy = () => {
    if (selection) void copyRange(sheet, selection)
  }

  return (
    <div className="grid-card">
      {diff && (diff.changedCount > 0 || diff.addedCount > 0 || diff.removedCount > 0) && (
        <div className="grid-selection-bar" style={{ left: 8, right: 'auto' }}>
          <span data-diff="changed">변경 {diff.changedCount}</span>
          <span data-diff="added">추가 {diff.addedCount}</span>
          <span data-diff="removed">삭제 {diff.removedCount}</span>
        </div>
      )}
      {selection && (
        <div className="grid-selection-bar">
          <span>
            {cellAddress(selection.r0, selection.c0)}:{cellAddress(selection.r1, selection.c1)} 선택됨
          </span>
          <button type="button" className="btn btn--ghost" onClick={handleCopy}>
            {copied ? '복사됨' : '복사'}
          </button>
        </div>
      )}
      <div
        className="grid-scroll"
        ref={scrollRef}
        tabIndex={0}
        onKeyDown={(e) => {
          if ((e.ctrlKey || e.metaKey) && e.key === 'c' && selection) {
            e.preventDefault()
            handleCopy()
          }
        }}
      >
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
                {(() => {
                  const cells: ReactNode[] = []
                  let colIndex = 0
                  while (colIndex < sheet.colCount) {
                    const col = colIndex + 1
                    const merge = mergeLookup.get(`${rowNum},${col}`)

                    // 병합 범위에 덮여있지만 이 칸이 마스터(왼쪽 위)가 아니면: 값을
                    // 중복으로 찍지 않고 빈 칸으로만 렌더링한다(세로 병합은 열
                    // 스팬만으로는 시각적으로 못 이어붙이므로 v1은 이 정도로 절충).
                    if (merge && !isMergeMaster(merge, rowNum, col)) {
                      cells.push(
                        <div
                          key={colIndex}
                          className="grid-cell"
                          data-merged="true"
                          title="병합된 셀이에요 — 왼쪽 위 셀에서 값을 확인/수정하세요"
                        />,
                      )
                      colIndex++
                      continue
                    }

                    const colSpan = merge ? merge.c1 - merge.c0 + 1 : 1
                    const cell = row?.[colIndex]
                    const address = cellAddress(rowNum, col)
                    const isNumeric = typeof cell?.value === 'number'
                    const isFormula = Boolean(cell?.formula)
                    const isEditing = editingAddress === address
                    const cellDiff = diff?.cells[rowNum - 1]?.[colIndex]
                    const diffTitle =
                      cellDiff?.status === 'changed'
                        ? `비교 파일 값: ${cellDiff.oldValue ?? '(없음)'}`
                        : cellDiff?.status === 'removed'
                          ? `비교 파일에만 있던 값: ${cellDiff.oldValue}`
                          : undefined

                    if (isEditing) {
                      cells.push(
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
                        />,
                      )
                    } else {
                      cells.push(
                        <div
                          key={colIndex}
                          className="grid-cell"
                          data-editable={editMode && !isFormula && !merge}
                          data-readonly={editMode && (isFormula || Boolean(merge))}
                          data-selected={inRange(selection, rowNum, col)}
                          data-diff={cellDiff && cellDiff.status !== 'unchanged' ? cellDiff.status : undefined}
                          title={
                            diffTitle ??
                            (isFormula
                              ? '수식 셀은 이 버전에서 수정할 수 없어요'
                              : merge
                                ? '병합된 셀은 이 버전에서 수정할 수 없어요'
                                : undefined)
                          }
                          style={{
                            justifyContent: isNumeric ? 'flex-end' : 'flex-start',
                            gridColumn: colSpan > 1 ? `span ${colSpan}` : undefined,
                          }}
                          onClick={(e) => {
                            selectCell(rowNum, col, e.shiftKey)
                            if (!e.shiftKey && !merge) startEdit(address, cell)
                          }}
                        >
                          {formatCellValue(cell)}
                        </div>,
                      )
                    }

                    colIndex += colSpan
                  }
                  return cells
                })()}
              </div>
            )
          })}
        </div>
      </div>
    </div>
  )
}
