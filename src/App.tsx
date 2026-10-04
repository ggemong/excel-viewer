import { useCallback, useEffect, useMemo, useRef, useState } from 'react'
import type { CellRange } from './clipboard/buildClipboardPayload'
import { DropZone } from './grid/DropZone'
import { CellDetailBar } from './grid/CellDetailBar'
import { describeCell } from './grid/cellDetail'
import { FilterMenu } from './grid/FilterMenu'
import { FilterStatusBar } from './grid/FilterStatusBar'
import { Grid, type GridFocusRequest } from './grid/Grid'
import { skippedDrawingNotice } from './grid/skippedDrawingNotice'
import { SummaryBar } from './grid/SummaryBar'
import { Toolbar } from './grid/Toolbar'
import { useViewFilters } from './grid/useViewFilters'
import { formatCellValue } from './xlsx/formatValue'
import { columnLetter } from './xlsx/cellRef'
import type { SheetModel } from './xlsx/types'
import { highlightsFor } from './search/searchEngine'
import { SearchBar } from './search/SearchBar'
import { SearchResults } from './search/SearchResults'
import { useSearch } from './search/useSearch'
import { useWorkbookController } from './state/useWorkbookController'

function App() {
  const {
    workbook,
    activeSheet,
    activeSheetIndex,
    setActiveSheetIndex,
    loading,
    error,
    openFile,
    closeFile,
    compareFileName,
    compareLoading,
    compareError,
    loadCompareFile,
    clearCompare,
    activeSheetDiff,
  } = useWorkbookController()
  const [selection, setSelection] = useState<CellRange | null>(null)
  // 셀 내용 줄의 펼침 상태 — 칸을 옮겨도 유지한다(여러 칸을 연달아 읽을 때 매번 펼치지 않게).
  const [detailExpanded, setDetailExpanded] = useState(false)
  const filters = useViewFilters(workbook, activeSheet)
  // 열린 필터 메뉴: 어느 시트의 어느 열이고 ▼ 버튼이 화면 어디에 있는지. 시트가 바뀌면 자연히 무효가 된다.
  const [filterMenu, setFilterMenu] = useState<{
    sheet: SheetModel
    col: number
    anchor: { left: number; top: number; bottom: number }
  } | null>(null)
  const closeFilterMenu = useCallback(() => setFilterMenu(null), [])
  const toggleFilterMenu = useCallback(
    (col: number, rect: DOMRect) => {
      if (!activeSheet) return
      setFilterMenu((open) =>
        open && open.sheet === activeSheet && open.col === col
          ? null
          : { sheet: activeSheet, col, anchor: { left: rect.left, top: rect.top, bottom: rect.bottom } },
      )
    },
    [activeSheet],
  )
  const menuData = filterMenu && filterMenu.sheet === activeSheet ? filters.entriesFor(filterMenu.col) : null

  // 시트 전환은 탭 클릭과 검색 결과 이동이 같은 경로를 쓴다(선택·열린 메뉴 초기화 포함).
  const activateSheet = useCallback(
    (index: number) => {
      setActiveSheetIndex(index)
      setSelection(null)
      setFilterMenu(null)
    },
    [setActiveSheetIndex],
  )

  const search = useSearch({ workbook, activeSheetIndex, viewSheetFor: filters.viewSheetFor, onActivateSheet: activateSheet })
  const searchInputRef = useRef<HTMLInputElement>(null)
  // 시트가 많을 때 검색 결과 이동 등으로 활성 탭이 바뀌면 탭 줄 안에서 그 탭이 보이게 스크롤한다.
  const tabsRef = useRef<HTMLDivElement>(null)
  useEffect(() => {
    tabsRef.current?.querySelector('[aria-current="true"]')?.scrollIntoView?.({ inline: 'nearest', block: 'nearest' })
  }, [activeSheetIndex])
  const { open: searchOpen, setOpen: setSearchOpen, go: searchGo } = search

  const openSearch = useCallback(() => {
    setSearchOpen(true)
    // 이미 열려 있으면 입력으로 포커스만 옮긴다. 새로 열릴 때는 아래 effect가 포커스한다.
    searchInputRef.current?.focus()
    searchInputRef.current?.select()
  }, [setSearchOpen])
  useEffect(() => {
    if (searchOpen) {
      searchInputRef.current?.focus()
      searchInputRef.current?.select()
    }
  }, [searchOpen])

  // Ctrl+F는 파일이 열려 있을 때만 가로챈다 — 브라우저 기본 찾기는 가상 스크롤이라 화면 밖 행을 못 찾는다.
  const filterMenuOpen = filterMenu !== null
  useEffect(() => {
    if (!workbook) return
    const onKeyDown = (e: KeyboardEvent) => {
      if ((e.ctrlKey || e.metaKey) && !e.altKey && e.key.toLowerCase() === 'f') {
        e.preventDefault()
        openSearch()
      } else if (searchOpen && e.key === 'F3') {
        e.preventDefault()
        searchGo(e.shiftKey ? -1 : 1)
      } else if (searchOpen && e.key === 'Escape' && !filterMenuOpen) {
        setSearchOpen(false)
      }
    }
    document.addEventListener('keydown', onKeyDown)
    return () => document.removeEventListener('keydown', onKeyDown)
  }, [workbook, searchOpen, searchGo, setSearchOpen, openSearch, filterMenuOpen])

  const cellDetail = useMemo(
    () => (activeSheet && selection ? describeCell(activeSheet, selection.r0, selection.c0) : null),
    [activeSheet, selection],
  )

  const searchHighlights = useMemo(() => (searchOpen ? highlightsFor(search.hits, activeSheetIndex) : null), [searchOpen, search.hits, activeSheetIndex])
  const currentHit = searchOpen ? search.currentHit : null
  const onActiveSheet = currentHit !== null && currentHit.sheetIndex === activeSheetIndex
  const focusRequest = useMemo<GridFocusRequest | null>(() => {
    if (!currentHit || currentHit.sheetIndex !== activeSheetIndex) return null
    return currentHit.kind === 'cell'
      ? { nonce: search.navNonce, kind: 'cell', row: currentHit.row, col: currentHit.col }
      : { nonce: search.navNonce, kind: 'shape', drawingIndex: currentHit.drawingIndex }
  }, [currentHit, activeSheetIndex, search.navNonce])

  return (
    <div className="app-shell">
      <Toolbar
        fileName={workbook?.fileName ?? null}
        onOpenAnother={closeFile}
        compareFileName={compareFileName}
        compareLoading={compareLoading}
        onPickCompareFile={loadCompareFile}
        onClearCompare={clearCompare}
        onOpenSearch={openSearch}
        searchOpen={searchOpen}
      />

      {!workbook || !activeSheet ? (
        <DropZone onFile={openFile} error={error} loading={loading} />
      ) : (
        <>
          {searchOpen && (
            <>
              <SearchBar search={search} inputRef={searchInputRef} onClose={() => setSearchOpen(false)} />
              {search.scope === 'all' && (
                <SearchResults hits={search.hits} sheetNames={workbook.sheets.map((s) => s.name)} currentIndex={search.currentIndex} onSelect={search.selectHit} />
              )}
            </>
          )}
          <SummaryBar sheet={activeSheet} selection={selection} filteredOut={filters.filteredOut} />
          {compareError && (
            <div className="summary-bar" style={{ background: 'var(--accent-soft)', color: 'var(--accent)' }}>
              {compareError}
            </div>
          )}
          {[...workbook.warnings, skippedDrawingNotice(activeSheet.skippedDrawings)].filter(Boolean).map((notice) => (
            <div key={notice} className="notice-bar" role="status">
              {notice}
            </div>
          ))}
          {filters.stats && filters.stats.shown < filters.stats.total && (
            <FilterStatusBar
              total={filters.stats.total}
              shown={filters.stats.shown}
              touched={filters.touched}
              onClearAll={filters.clearAll}
              onResetToFile={filters.resetToFile}
            />
          )}
          {cellDetail && <CellDetailBar detail={cellDetail} expanded={detailExpanded} onToggleExpanded={() => setDetailExpanded((v) => !v)} />}
          <Grid
            key={activeSheet.name}
            sheet={filters.viewSheet ?? activeSheet}
            diff={activeSheetDiff}
            onSelectionChange={setSelection}
            activeFilterCols={filters.activeCols}
            onFilterButtonClick={toggleFilterMenu}
            filteredOut={filters.filteredOut}
            searchCells={searchHighlights?.cells}
            searchCurrentCell={onActiveSheet && currentHit.kind === 'cell' ? { row: currentHit.row, col: currentHit.col } : null}
            searchShapes={searchHighlights?.shapes}
            searchCurrentShape={onActiveSheet && currentHit.kind === 'shape' ? currentHit.drawingIndex : null}
            focusRequest={focusRequest}
          />
          {filterMenu && menuData && (
            <FilterMenu
              title={formatCellValue(activeSheet.rows[menuData.filter.headerRow - 1]?.[filterMenu.col - 1]) || `${columnLetter(filterMenu.col)}열`}
              anchor={filterMenu.anchor}
              entries={menuData.entries}
              columnFiltered={filters.activeCols.has(filterMenu.col)}
              canResetToFile={filters.touched}
              onToggle={(value) => filters.toggle(filterMenu.col, value)}
              onSetListed={(values, include) => filters.setListed(filterMenu.col, values, include)}
              onClearColumn={() => filters.clearColumn(filterMenu.col)}
              onResetToFile={filters.resetToFile}
              onClose={closeFilterMenu}
            />
          )}
          {workbook.sheets.length > 1 && (
            <div className="sheet-tabs" ref={tabsRef}>
              {workbook.sheets.map((sheet, i) => (
                <button
                  key={sheet.name}
                  type="button"
                  className="sheet-tab"
                  aria-current={i === activeSheetIndex}
                  onClick={() => activateSheet(i)}
                >
                  {sheet.name}
                  {searchOpen && search.scope === 'all' && search.countsBySheet.has(i) && (
                    <span className="sheet-tab-badge">{search.countsBySheet.get(i)}</span>
                  )}
                </button>
              ))}
            </div>
          )}
        </>
      )}
    </div>
  )
}

export default App
