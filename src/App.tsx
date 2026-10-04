import { useCallback, useState } from 'react'
import type { CellRange } from './clipboard/buildClipboardPayload'
import { DropZone } from './grid/DropZone'
import { FilterMenu } from './grid/FilterMenu'
import { FilterStatusBar } from './grid/FilterStatusBar'
import { Grid } from './grid/Grid'
import { skippedDrawingNotice } from './grid/skippedDrawingNotice'
import { SummaryBar } from './grid/SummaryBar'
import { Toolbar } from './grid/Toolbar'
import { useViewFilters } from './grid/useViewFilters'
import { formatCellValue } from './xlsx/formatValue'
import { columnLetter } from './xlsx/cellRef'
import type { SheetModel } from './xlsx/types'
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

  return (
    <div className="app-shell">
      <Toolbar
        fileName={workbook?.fileName ?? null}
        onOpenAnother={closeFile}
        compareFileName={compareFileName}
        compareLoading={compareLoading}
        onPickCompareFile={loadCompareFile}
        onClearCompare={clearCompare}
      />

      {!workbook || !activeSheet ? (
        <DropZone onFile={openFile} error={error} loading={loading} />
      ) : (
        <>
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
          <Grid
            sheet={filters.viewSheet ?? activeSheet}
            diff={activeSheetDiff}
            onSelectionChange={setSelection}
            activeFilterCols={filters.activeCols}
            onFilterButtonClick={toggleFilterMenu}
            filteredOut={filters.filteredOut}
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
            <div className="sheet-tabs">
              {workbook.sheets.map((sheet, i) => (
                <button
                  key={sheet.name}
                  type="button"
                  className="sheet-tab"
                  aria-current={i === activeSheetIndex}
                  onClick={() => {
                    setActiveSheetIndex(i)
                    setSelection(null)
                    setFilterMenu(null)
                  }}
                >
                  {sheet.name}
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
