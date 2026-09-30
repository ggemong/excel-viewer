import { useEffect, useMemo, useRef, useState, type ClipboardEvent as ReactClipboardEvent, type ReactNode } from 'react'
import { useVirtualizer } from '@tanstack/react-virtual'
import { normalizeRange, parseTsv, type CellRange } from '../clipboard/buildClipboardPayload'
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
  onSelectionChange?: (range: CellRange | null) => void
}

function inRange(range: CellRange | null, row: number, col: number): boolean {
  return !!range && row >= range.r0 && row <= range.r1 && col >= range.c0 && col <= range.c1
}

/**
 * 데스크톱 가상 스크롤 그리드. 행만 가상화한다(v1 스코프) — 대부분의 실사용
 * 파일에서 병목은 열 수가 아니라 행 수라서, 행 가상화만으로 대용량 파일의
 * 스크롤 성능 문제를 해결한다. 열이 극단적으로 많은 시트는 이후 확장 대상.
 */
export function Grid({ sheet, editMode, onEditCell, diff, onSelectionChange }: GridProps) {
  const scrollRef = useRef<HTMLDivElement>(null)
  const [editingAddress, setEditingAddress] = useState<string | null>(null)
  const [draft, setDraft] = useState('')
  const [anchor, setAnchor] = useState<{ row: number; col: number } | null>(null)
  const [focus, setFocus] = useState<{ row: number; col: number } | null>(null)
  const [dragging, setDragging] = useState(false)
  // anchor/focus를 state 클로저로만 읽으면, 같은 동기 구간 안에서 이벤트가
  // 연달아 여러 번 오는 경우(방향키를 빠르게 누르거나 프로그램이 연속
  // dispatch할 때) 리렌더 전이라 전부 같은 "이전" 값을 보고 서로를 덮어써
  // 이동이 씹힐 수 있다 — 그래서 항상 최신값인 ref를 같이 들고 다닌다.
  // dragging도 같은 이유로 게이트 체크는 ref로 하고, state는 mouseup
  // 이펙트를 구독/해제하는 용도로만 쓴다.
  const anchorRef = useRef<{ row: number; col: number } | null>(null)
  const focusRef = useRef<{ row: number; col: number } | null>(null)
  const draggingRef = useRef(false)
  const movedRef = useRef(false)
  const shiftRef = useRef(false)

  const setAnchorBoth = (pos: { row: number; col: number }) => {
    anchorRef.current = pos
    setAnchor(pos)
  }
  const setFocusBoth = (pos: { row: number; col: number }) => {
    focusRef.current = pos
    setFocus(pos)
  }
  const { copyRange, copied } = useClipboardCopy()
  const mergeLookup = useMergeLookup(sheet.merges)

  const rowVirtualizer = useVirtualizer({
    count: sheet.rowCount,
    getScrollElement: () => scrollRef.current,
    estimateSize: () => ROW_HEIGHT,
    overscan: 12,
  })

  const gridTemplateColumns = `${ROW_NUM_WIDTH}px repeat(${sheet.colCount}, ${COL_WIDTH}px)`
  const selection = useMemo(() => (anchor && focus ? normalizeRange(anchor, focus) : null), [anchor, focus])

  useEffect(() => {
    onSelectionChange?.(selection)
  }, [selection, onSelectionChange])

  const beginSelect = (row: number, col: number, shiftKey: boolean) => {
    movedRef.current = false
    shiftRef.current = shiftKey
    if (shiftKey && anchorRef.current) {
      setFocusBoth({ row, col })
    } else {
      setAnchorBoth({ row, col })
      setFocusBoth({ row, col })
    }
    draggingRef.current = true
    setDragging(true)
  }

  const extendSelect = (row: number, col: number) => {
    if (!draggingRef.current) return
    if (focusRef.current && focusRef.current.row === row && focusRef.current.col === col) return
    movedRef.current = true
    setFocusBoth({ row, col })
  }

  /** 행 번호/열 이름/모서리를 눌렀을 때: 드래그 판정 없이 그 행·열·전체를 바로 선택한다. */
  const selectWhole = (a: { row: number; col: number }, b: { row: number; col: number }) => {
    setAnchorBoth(a)
    setFocusBoth(b)
  }

  const startEdit = (address: string, cell: CellModel | undefined) => {
    if (!editMode || cell?.formula) return
    setEditingAddress(address)
    setDraft(cell && cell.value !== null ? String(cell.value) : '')
  }

  /** anchor가 가리키는 칸을 연다 — 수식/병합 비-마스터 칸이면 조용히 무시(startEdit이 수식은 이미 막음). */
  const openEditOnAnchor = () => {
    const pos = anchorRef.current
    if (!pos) return
    const merge = mergeLookup.get(`${pos.row},${pos.col}`)
    if (merge) return
    const cellData = sheet.rows[pos.row - 1]?.[pos.col - 1]
    startEdit(cellAddress(pos.row, pos.col), cellData)
  }

  // 마우스를 눌렀다 뗄 때까지 실제로 다른 칸으로 옮겨갔는지(드래그로 범위를
  // 그린 것)와 shift를 눌렀는지(범위 확장만 의도한 것)를 보고, 둘 다
  // 아니면(단순 클릭) 그 칸을 편집 모드로 연다 — mouseup 시점에 판단해야
  // 드래그 중간에 잘못 편집이 열리지 않는다.
  useEffect(() => {
    if (!dragging) return
    const onMouseUp = () => {
      draggingRef.current = false
      setDragging(false)
      if (!movedRef.current && !shiftRef.current && editMode) openEditOnAnchor()
    }
    window.addEventListener('mouseup', onMouseUp)
    return () => window.removeEventListener('mouseup', onMouseUp)
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [dragging])

  const commitEdit = () => {
    if (editingAddress) onEditCell(editingAddress, draft)
    setEditingAddress(null)
  }

  const cancelEdit = () => setEditingAddress(null)

  /**
   * 편집 중에 방향키를 누르면(실제 엑셀과 동일하게) 지금 입력을 커밋하고
   * 그 방향의 셀로 이동한다 — 입력창에 포커스가 있으면 그리드의 방향키
   * 이동 로직에 아예 안 닿기 때문에, 입력창 쪽에서 따로 처리해야 한다.
   */
  const commitAndMove = (dr: number, dc: number, shiftKey: boolean) => {
    if (editingAddress) onEditCell(editingAddress, draft)
    setEditingAddress(null)
    const base = focusRef.current ?? anchorRef.current
    if (!base) return
    const nextRow = Math.min(Math.max(base.row + dr, 1), Math.max(sheet.rowCount, 1))
    const nextCol = Math.min(Math.max(base.col + dc, 1), Math.max(sheet.colCount, 1))
    if (shiftKey && anchorRef.current) {
      setFocusBoth({ row: nextRow, col: nextCol })
    } else {
      setAnchorBoth({ row: nextRow, col: nextCol })
      setFocusBoth({ row: nextRow, col: nextCol })
    }
    rowVirtualizer.scrollToIndex(nextRow - 1, { align: 'auto' })
    scrollRef.current?.focus()
  }

  const handleCopy = () => {
    if (selection) void copyRange(sheet, selection)
  }

  const handleDeleteSelection = () => {
    if (!selection) return
    for (let r = selection.r0; r <= selection.r1; r++) {
      for (let c = selection.c0; c <= selection.c1; c++) {
        const cellData = sheet.rows[r - 1]?.[c - 1]
        if (!cellData || cellData.value === null) continue
        if (cellData.formula) continue
        const merge = mergeLookup.get(`${r},${c}`)
        if (merge && !isMergeMaster(merge, r, c)) continue
        onEditCell(cellAddress(r, c), '')
      }
    }
  }

  // 붙여넣기는 지금 있는 시트 범위 안의 칸만 대상으로 한다(원본 파일 범위를
  // 넘어서는 확장은 이번 스코프 밖) — 범위를 벗어나는 칸, 수식 셀, 병합
  // 비-마스터 칸은 조용히 건너뛴다.
  const handlePaste = (e: ReactClipboardEvent<HTMLDivElement>) => {
    if (!editMode || !anchor) return
    const text = e.clipboardData.getData('text/plain')
    if (!text) return
    e.preventDefault()

    const rows = parseTsv(text)
    for (let ri = 0; ri < rows.length; ri++) {
      const targetRow = anchor.row + ri
      if (targetRow > sheet.rowCount) break
      for (let ci = 0; ci < rows[ri].length; ci++) {
        const targetCol = anchor.col + ci
        if (targetCol > sheet.colCount) continue
        const merge = mergeLookup.get(`${targetRow},${targetCol}`)
        if (merge && !isMergeMaster(merge, targetRow, targetCol)) continue
        const existingCell = sheet.rows[targetRow - 1]?.[targetCol - 1]
        if (existingCell?.formula) continue
        onEditCell(cellAddress(targetRow, targetCol), rows[ri][ci])
      }
    }
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
          {editMode && (
            <button type="button" className="btn btn--ghost" onClick={handleDeleteSelection}>
              삭제
            </button>
          )}
        </div>
      )}
      <div
        className="grid-scroll"
        ref={scrollRef}
        tabIndex={0}
        onPaste={handlePaste}
        onKeyDown={(e) => {
          // 셀 편집 중(입력창에 포커스)일 땐 이 아래 단축키를 전부 무시한다 —
          // 안 그러면 입력창 안에서 텍스트 복사/방향키 이동하려던 게 그리드
          // 단축키로 가로채인다. Enter/Escape는 입력창 자체의 onKeyDown이 처리.
          if (editingAddress) return

          const arrowDelta: Record<string, [number, number]> = {
            ArrowUp: [-1, 0],
            ArrowDown: [1, 0],
            ArrowLeft: [0, -1],
            ArrowRight: [0, 1],
          }
          const delta = arrowDelta[e.key]

          if (delta) {
            e.preventDefault()
            const base = focusRef.current ?? anchorRef.current ?? { row: 1, col: 1 }
            const nextRow = Math.min(Math.max(base.row + delta[0], 1), Math.max(sheet.rowCount, 1))
            const nextCol = Math.min(Math.max(base.col + delta[1], 1), Math.max(sheet.colCount, 1))
            if (e.shiftKey && anchorRef.current) {
              setFocusBoth({ row: nextRow, col: nextCol })
            } else {
              setAnchorBoth({ row: nextRow, col: nextCol })
              setFocusBoth({ row: nextRow, col: nextCol })
            }
            // 가상 스크롤이라 화면 밖 행으로 이동하면 실제로 안 보인다 — 대상
            // 행이 보이게 스크롤한다(열은 가상화 안 해서 다 DOM에 있으므로 안 건드림).
            rowVirtualizer.scrollToIndex(nextRow - 1, { align: 'auto' })
          } else if ((e.key === 'Enter' || e.key === 'F2') && editMode && anchorRef.current) {
            e.preventDefault()
            openEditOnAnchor()
          } else if ((e.ctrlKey || e.metaKey) && e.key === 'c' && selection) {
            e.preventDefault()
            handleCopy()
          } else if ((e.key === 'Delete' || e.key === 'Backspace') && editMode && selection) {
            e.preventDefault()
            handleDeleteSelection()
          }
        }}
      >
        <div style={{ display: 'grid', gridTemplateColumns, position: 'sticky', top: 0, zIndex: 2 }}>
          <div
            className="grid-cell grid-cell--colhead grid-cell--corner"
            style={{ height: HEADER_HEIGHT }}
            title="전체 선택"
            onClick={() => selectWhole({ row: 1, col: 1 }, { row: sheet.rowCount, col: sheet.colCount })}
          />
          {Array.from({ length: sheet.colCount }, (_, i) => (
            <div
              key={i}
              className="grid-cell grid-cell--colhead"
              style={{ height: HEADER_HEIGHT }}
              title={`${columnLetter(i + 1)}열 전체 선택`}
              onClick={() => selectWhole({ row: 1, col: i + 1 }, { row: sheet.rowCount, col: i + 1 })}
            >
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
                <div
                  className="grid-cell grid-cell--rownum"
                  title={`${rowNum}행 전체 선택`}
                  onClick={() => selectWhole({ row: rowNum, col: 1 }, { row: rowNum, col: sheet.colCount })}
                >
                  {rowNum}
                </div>
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
                          data-selected={inRange(selection, rowNum, col)}
                          title="병합된 셀이에요 — 왼쪽 위 셀에서 값을 확인/수정하세요"
                          onMouseDown={(e) => beginSelect(rowNum, col, e.shiftKey)}
                          onMouseEnter={() => extendSelect(rowNum, col)}
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
                            if (e.key === 'Enter' || e.key === 'F2') {
                              commitEdit()
                              // 편집을 키보드로 끝냈을 땐 그리드로 포커스를 되돌려야
                              // 다음 방향키/Enter/F2/Ctrl+C가 계속 먹힌다 — blur로
                              // 끝난 경우(다른 칸 클릭 등)는 그 대상이 알아서 포커스를
                              // 받으므로 여기서 강제로 되돌리지 않는다.
                              scrollRef.current?.focus()
                            } else if (e.key === 'Escape') {
                              cancelEdit()
                              scrollRef.current?.focus()
                            } else if (e.key === 'ArrowUp') {
                              e.preventDefault()
                              commitAndMove(-1, 0, e.shiftKey)
                            } else if (e.key === 'ArrowDown') {
                              e.preventDefault()
                              commitAndMove(1, 0, e.shiftKey)
                            } else if (e.key === 'ArrowLeft') {
                              // 텍스트 커서가 맨 앞일 때만 셀 이동으로 취급한다 —
                              // 안 그러면 긴 값 수정 중에 커서를 옮길 수가 없다.
                              const el = e.currentTarget
                              if (el.selectionStart === 0 && el.selectionEnd === 0) {
                                e.preventDefault()
                                commitAndMove(0, -1, e.shiftKey)
                              }
                            } else if (e.key === 'ArrowRight') {
                              const el = e.currentTarget
                              if (el.selectionStart === el.value.length && el.selectionEnd === el.value.length) {
                                e.preventDefault()
                                commitAndMove(0, 1, e.shiftKey)
                              }
                            }
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
                          onMouseDown={(e) => beginSelect(rowNum, col, e.shiftKey)}
                          onMouseEnter={() => extendSelect(rowNum, col)}
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
