import { useEffect, useMemo, useRef, useState, type CSSProperties, type ReactNode } from 'react'
import { useVirtualizer } from '@tanstack/react-virtual'
import { normalizeRange, type CellRange } from '../clipboard/buildClipboardPayload'
import { useClipboardCopy } from '../clipboard/useClipboardCopy'
import type { SheetDiff } from '../diff/diffWorkbooks'
import { createConditionalStyler } from './conditionalStyle'
import { DrawingLayer, type PlacedDrawing } from './DrawingLayer'
import { frozenColumnLayout } from './frozenColumns'
import { createSpillResolver } from './textSpill'
import { useFittedRowHeights } from './useFittedRowHeights'
import { useFontsVersion } from './useFontsVersion'
import { isMergeMaster, useMergeLookup, type MergeRange } from './useMergeLookup'
import { useBlobUrls } from './useBlobUrls'
import { cellHorizontalChrome } from './wrapMeasure'
import { cellAddress, columnLetter } from '../xlsx/cellRef'
import { applyConditionalStyle, cellStyleProps, type CellStyle } from '../xlsx/cellStyle'
import { axisOffset, buildAxis, collectPictureBlobs, placeItem } from '../xlsx/drawingLayout'
import { formatCellValue } from '../xlsx/formatValue'
import { textWidth } from '../xlsx/textMetrics'
import type { SheetModel } from '../xlsx/types'

/**
 * 파일에 행높이가 없을 때만 쓰는 비상 fallback(정상 동작이면 안 쓰임 — read.ts가
 * 항상 sheet.rowHeights를 rowCount만큼 채워서 준다). 열너비/행높이는 이제
 * sheet.colWidths/rowHeights(파일 값 기반, src/xlsx/columnWidth.ts)를 쓴다.
 */
const ROW_HEIGHT_FALLBACK = 34
const HEADER_HEIGHT = 32
const ROW_NUM_WIDTH = 44
/** 가로로 스크롤해도 왼쪽에 붙어 있는 칸(행 번호, 고정 열)이 스크롤되는 칸 위에 그려지도록 하는 쌓임 순서. */
const STICKY_COL_Z = 2
/**
 * 그림 배치용 축에서 데이터 범위 밖 칸의 크기. read.ts가 그림이 걸친 영역까지 시트 크기를 미리
 * 늘려 두므로(requiredExtent) 범위 밖 앵커는 생기지 않는다 — 혹시 생겨도 가장자리에 붙게 0으로 둔다.
 */
const OUT_OF_RANGE_CELL_SIZE = 0

interface GridProps {
  sheet: SheetModel
  diff?: SheetDiff | null
  onSelectionChange?: (range: CellRange | null) => void
  /** 필터 버튼을 강조할 열(조건이 걸린 열). 안 주면 파일에 저장된 조건 열을 따른다. */
  activeFilterCols?: Set<number>
  /** 헤더의 ▼ 버튼을 눌렀을 때. anchor는 화면 기준 버튼 위치(메뉴를 그 옆에 띄우는 용도). */
  onFilterButtonClick?: (col: number, anchor: DOMRect) => void
  /** 필터로 걸러진 행 — 복사에서 제외한다. */
  filteredOut?: boolean[]
  /** 검색 결과로 강조할 셀("행,열" 키)과 현재 위치. */
  searchCells?: Set<string>
  searchCurrentCell?: { row: number; col: number } | null
  /** 검색 결과로 강조할 도형(drawings 인덱스)과 현재 위치. */
  searchShapes?: Set<number>
  searchCurrentShape?: number | null
  /** 이 요청이 새로 오면(nonce가 바뀌면) 그 위치로 스크롤하고 선택한다. */
  focusRequest?: GridFocusRequest | null
}

/** 외부(검색)가 그리드에 "여기로 가 줘"라고 요청하는 모양. nonce가 바뀔 때마다 새 요청으로 본다. */
export type GridFocusRequest = { nonce: number } & ({ kind: 'cell'; row: number; col: number } | { kind: 'shape'; drawingIndex: number })

/** 검색 결과로 이동할 때 목표가 화면 가장자리에 붙지 않도록 남기는 여백(px). */
const SEARCH_SCROLL_MARGIN = 48

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
export function Grid({ sheet, diff, onSelectionChange, activeFilterCols, onFilterButtonClick, filteredOut, searchCells, searchCurrentCell, searchShapes, searchCurrentShape, focusRequest }: GridProps) {
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
  // 줄바꿈 글이 잘리는 행을 키운 높이(파일 값 + 화면 글꼴에서의 줄 수). 아래 모든 행 높이는 이 값을 쓴다.
  const fontsVersion = useFontsVersion()
  const rowHeights = useFittedRowHeights(sheet, fontsVersion)

  // 숨긴 열은 그리드 트랙 자체를 안 만든다(Excel이 숨긴 열을 아예 안 보여주는 것과
  // 동일) — 선택·복사는 여전히 숨긴 열을 포함한 논리적 범위(행/열 번호)로 동작하고,
  // 여기서 걸러내는 건 오직 "화면에 그릴 열의 목록"뿐이다.
  const visibleColNumbers = useMemo(
    () => Array.from({ length: sheet.colCount }, (_, i) => i + 1).filter((c) => !sheet.hiddenCols[c - 1]),
    [sheet.colCount, sheet.hiddenCols],
  )

  // 틀고정 행은 가상화 목록에서 빼서 열 헤더 바로 아래 별도의 고정 블록으로 그린다
  // (ySplit 아래부터만 가상 스크롤). 숨긴 행은 고정 블록/스크롤 블록 양쪽 모두에서
  // 제외한다.
  const frozenRowCount = sheet.frozen?.rows ?? 0
  // 열 고정: 앞쪽 N개 열(+ 항상 붙어 있는 행 번호 칸)이 가로 스크롤에서 왼쪽에 남는다. 열을 가상화하지 않으니 sticky left만 주면 된다.
  const frozenColCount = sheet.frozen?.cols ?? 0
  const frozenCols = useMemo(
    () => frozenColumnLayout(sheet.colWidths, visibleColNumbers, frozenColCount, ROW_NUM_WIDTH),
    [sheet.colWidths, visibleColNumbers, frozenColCount],
  )
  // sticky는 "포함하는 상자" 안에서만 붙어 있으므로, 행 상자의 폭이 화면 폭이 아니라 열 전체 폭이어야 끝까지 스크롤해도 유지된다.
  const totalWidth = useMemo(
    () => ROW_NUM_WIDTH + visibleColNumbers.reduce((sum, c) => sum + sheet.colWidths[c - 1], 0),
    [sheet.colWidths, visibleColNumbers],
  )
  /** 이 열이 고정이면 sticky 스타일. 고정 경계에 걸친 병합은 한 덩어리로 붙들 수 없어 고정하지 않는다(스크롤되어 나간다). */
  const stickyColStyle = (col: number, merge?: MergeRange | null): CSSProperties => {
    const left = frozenCols.left.get(col)
    if (left === undefined || (merge && merge.c1 > frozenColCount)) return {}
    return { position: 'sticky', left, zIndex: STICKY_COL_Z }
  }
  const cellClassName = (col: number, filterState: string | undefined) =>
    ['grid-cell', filterState ? 'grid-cell--filter' : '', col === frozenCols.edgeCol ? 'grid-cell--freeze-edge' : ''].filter(Boolean).join(' ')
  const frozenRowNumbers = useMemo(
    () =>
      Array.from({ length: frozenRowCount }, (_, i) => i + 1).filter((r) => !sheet.hiddenRows[r - 1]),
    [frozenRowCount, sheet.hiddenRows],
  )
  const scrollableRowNumbers = useMemo(
    () =>
      Array.from({ length: sheet.rowCount - frozenRowCount }, (_, i) => i + 1 + frozenRowCount).filter(
        (r) => !sheet.hiddenRows[r - 1],
      ),
    [sheet.rowCount, frozenRowCount, sheet.hiddenRows],
  )

  const rowVirtualizer = useVirtualizer({
    count: scrollableRowNumbers.length,
    getScrollElement: () => scrollRef.current,
    // 모든 행 높이를 파일 읽을 때 미리 다 알고 있어서(런타임 측정이 아님) 이
    // 추정치가 처음부터 정확하다 — 가변 가상화에서 흔한 "늦은 재측정으로 스크롤
    // 위치가 튀는" 문제가 없다.
    estimateSize: (index) => rowHeights[scrollableRowNumbers[index] - 1] ?? ROW_HEIGHT_FALLBACK,
    overscan: 12,
  })
  // 행 높이가 나중에 바뀌면(글꼴이 늦게 도착해 줄 수가 달라질 때) 가상 스크롤러가 옛 크기를 쥐고 있지 않게 비운다.
  useEffect(() => {
    rowVirtualizer.measure()
  }, [rowHeights, rowVirtualizer])

  // 방향키 이동 시 그 행이 스크롤 가상화 쪽 몇 번째 항목인지 바로 찾기 위한
  // 역방향 조회 — 매 키 입력마다 scrollableRowNumbers를 선형 탐색하지 않기 위해서다.
  const rowVirtualIndexByNumber = useMemo(() => {
    const map = new Map<number, number>()
    scrollableRowNumbers.forEach((r, i) => map.set(r, i))
    return map
  }, [scrollableRowNumbers])

  // 그림/도형: 시트 좌표(숨긴 열/행은 폭 0)로 배치하고, 틀고정 영역 안에 완전히 들어가는 것은
  // 고정 블록에, 나머지는 스크롤 본문에 올린다(틀고정 위에 그려진 그림이 같이 스크롤돼 어긋나지 않게).
  const colAxis = useMemo(() => buildAxis(sheet.colWidths, sheet.hiddenCols, OUT_OF_RANGE_CELL_SIZE), [sheet.colWidths, sheet.hiddenCols])
  const rowAxis = useMemo(() => buildAxis(rowHeights, sheet.hiddenRows, OUT_OF_RANGE_CELL_SIZE), [rowHeights, sheet.hiddenRows])
  const frozenHeight = axisOffset(rowAxis, frozenRowCount)
  const { frozenPlaced, scrollPlaced } = useMemo(() => {
    const placed: PlacedDrawing[] = sheet.drawings.map((item, index) => ({ item, index, box: placeItem(item, colAxis, rowAxis) }))
    const inFrozen = (p: PlacedDrawing) => frozenHeight > 0 && p.box.top + p.box.height <= frozenHeight + 1
    return { frozenPlaced: placed.filter(inFrozen), scrollPlaced: placed.filter((p) => !inFrozen(p)) }
  }, [sheet.drawings, colAxis, rowAxis, frozenHeight])
  // 자동 필터 버튼이 달릴 칸 -> 조건이 걸려 필터링 중인지. 버튼만 표시하는 용도(보기 전용)이고, 필터로
  // 숨겨진 행은 파일에 hidden으로 이미 저장돼 있어서 별도 계산 없이 숨김 행으로 처리된다.
  const filterButtons = useMemo(() => {
    const map = new Map<string, 'idle' | 'active'>()
    for (const f of sheet.filters) {
      for (let col = f.firstCol; col <= f.lastCol; col++) {
        if (f.hiddenButtonCols.includes(col)) continue
        const active = activeFilterCols ? activeFilterCols.has(col) : f.activeCols.includes(col)
        map.set(`${f.headerRow},${col}`, active ? 'active' : 'idle')
      }
    }
    return map
  }, [sheet.filters, activeFilterCols])
  // 한 줄 글이 칸보다 길고 옆 칸이 비어 있으면 그 위로 넘쳐 보이게 한다(규칙은 textSpill.ts). 글자 폭은 글꼴에 달려 있어서
  // 글꼴이 도착하면(fontsVersion) 다시 만든다.
  const spillAt = useMemo(
    () =>
      createSpillResolver({
        rows: sheet.rows,
        colWidths: sheet.colWidths,
        visibleCols: visibleColNumbers,
        merged: mergeLookup,
        // 고정 열은 스크롤에 따라 다른 칸 위에 얹히므로 글이 넘어가지도, 넘어오지도 않게 한다.
        isBlocked: (row, col) => col <= frozenColCount || filterButtons.has(`${row},${col}`),
        cellChrome: cellHorizontalChrome(),
        measure: textWidth,
      }),
    // oxlint-disable-next-line react-hooks/exhaustive-deps
    [sheet.rows, sheet.colWidths, visibleColNumbers, mergeLookup, filterButtons, frozenColCount, fontsVersion],
  )
  // 조건부서식은 칸 데이터와 규칙에만 의존한다(선택/스크롤/필터 숨김과 무관) — 그 둘이 같으면 평가 결과 캐시를 재사용한다.
  const conditionalStyler = useMemo(
    () => createConditionalStyler({ rows: sheet.rows, conditionalFormats: sheet.conditionalFormats }),
    [sheet.rows, sheet.conditionalFormats],
  )
  const styleAt = (row: number, col: number, base: CellStyle | null): CellStyle | null => {
    const cf = conditionalStyler?.(row, col)
    return cf ? applyConditionalStyle(base, cf) : base
  }

  const pictureBlobs = useMemo(() => collectPictureBlobs(sheet.drawings), [sheet.drawings])
  const urlFor = useBlobUrls(pictureBlobs)

  const gridTemplateColumns = `${ROW_NUM_WIDTH}px ${visibleColNumbers.map((c) => `${sheet.colWidths[c - 1]}px`).join(' ')}`
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
   * 늘리고 아니면 anchor/focus를 함께 옮긴다(새로 한 칸만 선택). 숨긴 행/열은
   * 화면에 없어 거기로 이동해도 아무것도 안 보이므로, Excel처럼 그 방향으로
   * 계속 건너뛰어 보이는 칸에 멈춘다. 가상 스크롤이라 화면 밖 행으로 이동하면
   * 실제로 안 보이므로 대상 행이 보이게 스크롤한다(틀고정 행은 항상 보이므로
   * 스크롤 대상이 아니다, 열은 가상화 안 해서 다 DOM에 있으므로 안 건드림).
   */
  const moveFocus = (dr: number, dc: number, shiftKey: boolean) => {
    const base = focusRef.current ?? anchorRef.current ?? { row: 1, col: 1 }
    let nextRow = Math.min(Math.max(base.row + dr, 1), Math.max(sheet.rowCount, 1))
    if (dr !== 0) {
      while (nextRow >= 1 && nextRow <= sheet.rowCount && sheet.hiddenRows[nextRow - 1]) {
        const stepped = nextRow + dr
        if (stepped < 1 || stepped > sheet.rowCount) break
        nextRow = stepped
      }
    }
    let nextCol = Math.min(Math.max(base.col + dc, 1), Math.max(sheet.colCount, 1))
    if (dc !== 0) {
      while (nextCol >= 1 && nextCol <= sheet.colCount && sheet.hiddenCols[nextCol - 1]) {
        const stepped = nextCol + dc
        if (stepped < 1 || stepped > sheet.colCount) break
        nextCol = stepped
      }
    }
    if (shiftKey && anchorRef.current) {
      setFocusBoth({ row: nextRow, col: nextCol })
    } else {
      setAnchorBoth({ row: nextRow, col: nextCol })
      setFocusBoth({ row: nextRow, col: nextCol })
    }
    const virtualIndex = rowVirtualIndexByNumber.get(nextRow)
    if (virtualIndex !== undefined) {
      rowVirtualizer.scrollToIndex(virtualIndex, { align: 'auto' })
    }
  }

  // 검색 결과로 이동: 셀이면 선택하고 가운데로 스크롤, 도형이면 그 위치로 스크롤. 키보드 포커스는
  // 검색 입력에 남겨 둔다(Enter로 계속 다음 결과로 넘어갈 수 있게).
  const focusNonce = focusRequest?.nonce
  useEffect(() => {
    if (!focusRequest) return
    const scroller = scrollRef.current
    if (!scroller) return

    if (focusRequest.kind === 'cell') {
      const { row, col } = focusRequest
      setAnchorBoth({ row, col })
      setFocusBoth({ row, col })
      const virtualIndex = rowVirtualIndexByNumber.get(row)
      if (virtualIndex !== undefined) rowVirtualizer.scrollToIndex(virtualIndex, { align: 'center' })
      // 열은 가상화하지 않으므로 가로 스크롤은 직접 계산한다(화면 밖일 때만 움직인다). 행 번호 칸과 고정 열은 늘 보이고
      // 그 밑으로 들어간 칸은 가려진 것이므로, 보이는 영역의 왼쪽 끝은 그 영역(frozenCols.width)의 오른쪽부터다.
      if (col > frozenColCount) {
        const left = ROW_NUM_WIDTH + axisOffset(colAxis, col - 1)
        const right = ROW_NUM_WIDTH + axisOffset(colAxis, col)
        if (left < scroller.scrollLeft + frozenCols.width || right > scroller.scrollLeft + scroller.clientWidth) {
          scroller.scrollTo({ left: Math.max(0, left - frozenCols.width - SEARCH_SCROLL_MARGIN) })
        }
      }
      return
    }

    const item = sheet.drawings[focusRequest.drawingIndex]
    if (!item) return
    const box = placeItem(item, colAxis, rowAxis)
    // 틀고정 영역 안의 도형은 항상 보이므로 세로로 움직이지 않는다.
    if (box.top + box.height > frozenHeight + 1) {
      scroller.scrollTo({ top: Math.max(0, box.top - frozenHeight - SEARCH_SCROLL_MARGIN) })
    }
    const left = ROW_NUM_WIDTH + box.left
    if (left < scroller.scrollLeft + frozenCols.width || left + box.width > scroller.scrollLeft + scroller.clientWidth) {
      scroller.scrollTo({ left: Math.max(0, left - frozenCols.width - SEARCH_SCROLL_MARGIN) })
    }
    // 의존성을 nonce 하나로 둔 이유: 같은 위치로 다시 이동(Enter 반복)하는 것도 새 요청이어야 하고,
    // 시트/축이 바뀐 것만으로는 이동하면 안 된다.
    // oxlint-disable-next-line react-hooks/exhaustive-deps
  }, [focusNonce])

  const searchShapeState = (index: number): 'match' | 'current' | undefined =>
    searchCurrentShape === index ? 'current' : searchShapes?.has(index) ? 'match' : undefined

  const handleCopy = () => {
    if (selection) void copyRange(sheet, selection, filteredOut)
  }

  /**
   * 한 행(row-number 칸 + 데이터 칸들)을 그린다 — 틀고정 블록과 가상 스크롤 블록이
   * 똑같은 셀/병합/선택 로직을 공유하도록 모아뒀다(두 블록이 각자 이 로직을 복제해
   * 들고 있으면 한쪽만 고치고 다른 쪽을 놓치기 쉽다). wrapperStyle만 호출부마다
   * 다르다(고정 블록은 평범한 블록 흐름, 스크롤 블록은 절대위치+translateY).
   */
  const renderRow = (rowNum: number, rowHeight: number, wrapperStyle: CSSProperties) => {
    const row = sheet.rows[rowNum - 1]
    return (
      <div
        key={rowNum}
        style={{
          display: 'grid',
          gridTemplateColumns,
          height: rowHeight,
          width: totalWidth,
          ...wrapperStyle,
        }}
      >
        <div
          className="grid-cell grid-cell--rownum"
          style={{ position: 'sticky', left: 0, zIndex: STICKY_COL_Z }}
          title={`${rowNum}행 전체 선택`}
          onClick={() => selectWhole({ row: rowNum, col: 1 }, { row: rowNum, col: sheet.colCount })}
        >
          {rowNum}
        </div>
        {(() => {
          const cells: ReactNode[] = []
          let vi = 0
          while (vi < visibleColNumbers.length) {
            const col = visibleColNumbers[vi]
            const merge = mergeLookup.get(`${rowNum},${col}`)

            // 병합 범위에 덮여있지만 이 칸이 마스터(왼쪽 위)가 아니면: 값을
            // 중복으로 찍지 않고 빈 칸으로만 렌더링한다. 각 행이 독립된 그리드로
            // 가상화되는 구조라 CSS의 grid-row: span으로 실제 행을 이어붙일 수는
            // 없지만(별개 그리드 컨테이너라 span이 작동 안 함), 마스터 셀의 실제
            // 서식은 가상화 여부와 무관하게 sheet.rows에 항상 남아있으므로 그대로
            // 가져와 배경/글자색을 이어붙이고, 마지막 행 전까지는 칸 사이 구분선
            // (border-bottom)을 지워 하나로 이어진 모양처럼 보이게 한다.
            if (merge && !isMergeMaster(merge, rowNum, col)) {
              const masterCell = sheet.rows[merge.r0 - 1]?.[merge.c0 - 1]
              const isLastMergedRow = rowNum === merge.r1
              cells.push(
                <div
                  key={col}
                  className={cellClassName(col, undefined)}
                  data-merged="true"
                  data-selected={inRange(selection, rowNum, col)}
                  title="병합된 셀 — 값은 왼쪽 위 셀에 있어요"
                  style={{
                    ...cellStyleProps(styleAt(merge.r0, merge.c0, masterCell?.style ?? null), inRange(selection, rowNum, col)),
                    ...(isLastMergedRow ? {} : { borderBottom: 'none' }),
                    ...stickyColStyle(col, merge),
                  }}
                  onMouseDown={(e) => beginSelect(rowNum, col, e.shiftKey)}
                  onMouseEnter={() => extendSelect(rowNum, col)}
                />,
              )
              vi++
              continue
            }

            // colSpan은 "보이는 열" 기준으로 센다 — 병합 범위 안에 숨긴 열이 섞여
            // 있으면 그만큼 적게 스팬해야 실제로 그려지는 그리드 트랙 수와 맞는다.
            let colSpan = 1
            if (merge) {
              colSpan = 0
              for (let k = vi; k < visibleColNumbers.length && visibleColNumbers[k] <= merge.c1; k++) colSpan++
            }

            const colIndex = col - 1
            const filterState = filterButtons.get(`${rowNum},${col}`)
            const searchState =
              searchCurrentCell && searchCurrentCell.row === rowNum && searchCurrentCell.col === col
                ? 'current'
                : searchCells?.has(`${rowNum},${col}`)
                  ? 'match'
                  : undefined
            const cell = row?.[colIndex]
            const spill = spillAt(rowNum, col)
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
                key={col}
                className={cellClassName(col, filterState)}
                data-selected={inRange(selection, rowNum, col)}
                data-search={searchState}
                data-diff={cellDiff && cellDiff.status !== 'unchanged' ? cellDiff.status : undefined}
                title={diffTitle ?? (filterState === 'active' ? '이 열에 필터 조건이 걸려 있어요' : undefined) ?? cell?.note}
                data-note={cell?.note ? 'true' : undefined}
                style={{
                  justifyContent: isNumeric ? 'flex-end' : 'flex-start',
                  gridColumn: colSpan > 1 ? `span ${colSpan}` : undefined,
                  ...cellStyleProps(
                    styleAt(rowNum, col, cell?.style ?? null),
                    inRange(selection, rowNum, col) || Boolean(cellDiff && cellDiff.status !== 'unchanged'),
                  ),
                  // 세로로 병합된 마스터 행이면(아래로 더 이어짐) 다음 칸과의 구분선을
                  // 지워 전체가 하나로 이어진 모양이 되게 한다(위 continuation 칸 처리와
                  // 짝을 이루는 로직).
                  ...(merge && merge.r1 > rowNum ? { borderBottom: 'none' } : {}),
                  // 글이 옆 칸 위로 넘칠 칸: 칸 밖도 보이게 하고, 뒤에 그려지는 옆 칸 배경 위에 올린다.
                  ...(spill ? { overflow: 'visible', zIndex: 1 } : {}),
                  ...stickyColStyle(col, merge),
                }}
                onMouseDown={(e) => beginSelect(rowNum, col, e.shiftKey)}
                onMouseEnter={() => extendSelect(rowNum, col)}
              >
                {cell?.hyperlink ? (
                  <a href={cell.hyperlink} target="_blank" rel="noopener noreferrer" className="grid-cell-link">
                    {formatCellValue(cell)}
                  </a>
                ) : spill ? (
                  <span className="grid-cell-spill" style={{ maxWidth: spill.maxWidth }}>
                    {formatCellValue(cell)}
                  </span>
                ) : (
                  formatCellValue(cell)
                )}
                {filterState && (
                  <button
                    type="button"
                    className="grid-filter-btn"
                    data-active={filterState === 'active'}
                    aria-label="필터"
                    aria-haspopup="dialog"
                    // 칸의 드래그 선택이 시작되지 않게 막고, 눌린 위치를 메뉴 앵커로 넘긴다.
                    onMouseDown={(e) => e.stopPropagation()}
                    onClick={(e) => {
                      e.stopPropagation()
                      onFilterButtonClick?.(col, e.currentTarget.getBoundingClientRect())
                    }}
                  />
                )}
              </div>,
            )

            vi += colSpan
          }
          return cells
        })()}
      </div>
    )
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
        <div style={{ display: 'grid', gridTemplateColumns, width: totalWidth, position: 'sticky', top: 0, zIndex: 2 }}>
          <div
            className="grid-cell grid-cell--colhead grid-cell--corner"
            style={{ height: HEADER_HEIGHT, position: 'sticky', left: 0, zIndex: STICKY_COL_Z }}
            title="전체 선택"
            onClick={() => selectWhole({ row: 1, col: 1 }, { row: sheet.rowCount, col: sheet.colCount })}
          />
          {visibleColNumbers.map((col) => (
            <div
              key={col}
              className={`${cellClassName(col, undefined)} grid-cell--colhead`}
              style={{ height: HEADER_HEIGHT, ...stickyColStyle(col) }}
              title={`${columnLetter(col)}열 전체 선택`}
              onClick={() => selectWhole({ row: 1, col }, { row: sheet.rowCount, col })}
            >
              {columnLetter(col)}
            </div>
          ))}
        </div>

        {frozenRowNumbers.length > 0 && (
          <div style={{ position: 'sticky', top: HEADER_HEIGHT, zIndex: 1 }}>
            {frozenRowNumbers.map((rowNum) => renderRow(rowNum, rowHeights[rowNum - 1] ?? ROW_HEIGHT_FALLBACK, {}))}
            <DrawingLayer placed={frozenPlaced} offsetX={ROW_NUM_WIDTH} offsetY={0} urlFor={urlFor} searchState={searchShapeState} />
          </div>
        )}

        {/* zIndex 0: 이 컨테이너가 자기만의 쌓임 맥락이 되어, 본문 그림이 위로 스크롤될 때 틀고정 블록(zIndex 1) 아래로 들어가게 한다. */}
        <div style={{ height: rowVirtualizer.getTotalSize(), position: 'relative', zIndex: 0 }}>
          {rowVirtualizer.getVirtualItems().map((virtualRow) => {
            const rowNum = scrollableRowNumbers[virtualRow.index]
            const rowHeight = rowHeights[rowNum - 1] ?? ROW_HEIGHT_FALLBACK
            return renderRow(rowNum, rowHeight, {
              position: 'absolute',
              top: 0,
              left: 0,
              transform: `translateY(${virtualRow.start}px)`,
            })
          })}
          <DrawingLayer placed={scrollPlaced} offsetX={ROW_NUM_WIDTH} offsetY={frozenHeight} urlFor={urlFor} searchState={searchShapeState} />
        </div>
      </div>
    </div>
  )
}
