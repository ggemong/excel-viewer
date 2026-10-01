import { useEffect, useMemo, useRef, useState, type CSSProperties, type ReactNode } from 'react'
import { useVirtualizer } from '@tanstack/react-virtual'
import { normalizeRange, type CellRange } from '../clipboard/buildClipboardPayload'
import { useClipboardCopy } from '../clipboard/useClipboardCopy'
import type { SheetDiff } from '../diff/diffWorkbooks'
import { isMergeMaster, useMergeLookup } from './useMergeLookup'
import { cellAddress, columnLetter } from '../xlsx/cellRef'
import { cellStyleProps } from '../xlsx/cellStyle'
import { formatCellValue } from '../xlsx/formatValue'
import type { SheetModel } from '../xlsx/types'

/** 가상 스크롤의 행 높이 추정치 — components.css의 .grid-cell에 --grid-row-h로 주입해서 실제 렌더 높이와 묶는다(아래 return문). */
const ROW_HEIGHT = 34
const HEADER_HEIGHT = 32
const ROW_NUM_WIDTH = 44
const COL_WIDTH = 120

interface GridProps {
  sheet: SheetModel
  diff?: SheetDiff | null
  onSelectionChange?: (range: CellRange | null) => void
}

function inRange(range: CellRange | null, row: number, col: number): boolean {
  return !!range && row >= range.r0 && row <= range.r1 && col >= range.c0 && col <= range.c1
}

/**
 * 가상 스크롤 그리드 — 데스크톱/모바일 공용(터치는 네이티브 스크롤, 드래그 범위
 * 선택은 데스크톱 마우스 전용 — 터치 드래그는 브라우저가 스크롤 제스처로 먼저
 * 가로채서 못 만든다, 한 칸 탭 선택은 됨). 행만 가상화한다(v1 스코프) — 대부분의
 * 실사용 파일에서 병목은 열 수가 아니라 행 수라서, 행 가상화만으로 대용량 파일의
 * 스크롤 성능 문제를 해결한다. 열이 극단적으로 많은 시트는 이후 확장 대상.
 */
export function Grid({ sheet, diff, onSelectionChange }: GridProps) {
  const scrollRef = useRef<HTMLDivElement>(null)
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
    setFocusBoth({ row, col })
  }

  /** 행 번호/열 이름/모서리를 눌렀을 때: 드래그 판정 없이 그 행·열·전체를 바로 선택한다. */
  const selectWhole = (a: { row: number; col: number }, b: { row: number; col: number }) => {
    setAnchorBoth(a)
    setFocusBoth(b)
  }

  // 드래그가 끝났는지(마우스를 뗐는지)만 감지한다 — 끝나면 draggingRef를 꺼서
  // 그 이후의 mouseenter가 선택을 더 늘리지 않게 한다.
  useEffect(() => {
    if (!dragging) return
    const onMouseUp = () => {
      draggingRef.current = false
      setDragging(false)
    }
    window.addEventListener('mouseup', onMouseUp)
    return () => window.removeEventListener('mouseup', onMouseUp)
  }, [dragging])

  /**
   * focus(없으면 anchor, 둘 다 없으면 A1)에서 (dr, dc)만큼 이동한 칸으로 선택을
   * 옮긴다 — 시트 범위로 clamp하고, shift가 눌려있으면 focus만 넓혀 범위를
   * 늘리고 아니면 anchor/focus를 함께 옮긴다(새로 한 칸만 선택). 가상 스크롤이라
   * 화면 밖 행으로 이동하면 실제로 안 보이므로 대상 행이 보이게 스크롤한다(열은
   * 가상화 안 해서 다 DOM에 있으므로 안 건드림).
   */
  const moveFocus = (dr: number, dc: number, shiftKey: boolean) => {
    const base = focusRef.current ?? anchorRef.current ?? { row: 1, col: 1 }
    const nextRow = Math.min(Math.max(base.row + dr, 1), Math.max(sheet.rowCount, 1))
    const nextCol = Math.min(Math.max(base.col + dc, 1), Math.max(sheet.colCount, 1))
    if (shiftKey && anchorRef.current) {
      setFocusBoth({ row: nextRow, col: nextCol })
    } else {
      setAnchorBoth({ row: nextRow, col: nextCol })
      setFocusBoth({ row: nextRow, col: nextCol })
    }
    rowVirtualizer.scrollToIndex(nextRow - 1, { align: 'auto' })
  }

  const handleCopy = () => {
    if (selection) void copyRange(sheet, selection)
  }

  return (
    <div className="grid-card" style={{ '--grid-row-h': `${ROW_HEIGHT}px` } as CSSProperties}>
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
          const arrowDelta: Record<string, [number, number]> = {
            ArrowUp: [-1, 0],
            ArrowDown: [1, 0],
            ArrowLeft: [0, -1],
            ArrowRight: [0, 1],
          }
          const delta = arrowDelta[e.key]

          if (delta) {
            e.preventDefault()
            moveFocus(delta[0], delta[1], e.shiftKey)
          } else if ((e.ctrlKey || e.metaKey) && e.key === 'c' && selection) {
            e.preventDefault()
            handleCopy()
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
                          title="병합된 셀 — 값은 왼쪽 위 셀에 있어요"
                          onMouseDown={(e) => beginSelect(rowNum, col, e.shiftKey)}
                          onMouseEnter={() => extendSelect(rowNum, col)}
                        />,
                      )
                      colIndex++
                      continue
                    }

                    const colSpan = merge ? merge.c1 - merge.c0 + 1 : 1
                    const cell = row?.[colIndex]
                    const isNumeric = typeof cell?.value === 'number'
                    const cellDiff = diff?.cells[rowNum - 1]?.[colIndex]
                    const diffTitle =
                      cellDiff?.status === 'changed'
                        ? `비교 파일 값: ${cellDiff.oldValue ?? '(없음)'}`
                        : cellDiff?.status === 'removed'
                          ? `비교 파일에만 있던 값: ${cellDiff.oldValue}`
                          : undefined

                    cells.push(
                      <div
                        key={colIndex}
                        className="grid-cell"
                        data-selected={inRange(selection, rowNum, col)}
                        data-diff={cellDiff && cellDiff.status !== 'unchanged' ? cellDiff.status : undefined}
                        title={diffTitle}
                        style={{
                          justifyContent: isNumeric ? 'flex-end' : 'flex-start',
                          gridColumn: colSpan > 1 ? `span ${colSpan}` : undefined,
                          ...cellStyleProps(
                            cell?.style ?? null,
                            inRange(selection, rowNum, col) || Boolean(cellDiff && cellDiff.status !== 'unchanged'),
                          ),
                        }}
                        onMouseDown={(e) => beginSelect(rowNum, col, e.shiftKey)}
                        onMouseEnter={() => extendSelect(rowNum, col)}
                      >
                        {cell?.hyperlink ? (
                          <a href={cell.hyperlink} target="_blank" rel="noopener noreferrer" className="grid-cell-link">
                            {formatCellValue(cell)}
                          </a>
                        ) : (
                          formatCellValue(cell)
                        )}
                      </div>,
                    )

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
