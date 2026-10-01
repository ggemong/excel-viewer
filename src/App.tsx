import { useState } from 'react'
import type { CellRange } from './clipboard/buildClipboardPayload'
import { DropZone } from './grid/DropZone'
import { Grid } from './grid/Grid'
import { SummaryBar } from './grid/SummaryBar'
import { Toolbar } from './grid/Toolbar'
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
          <SummaryBar sheet={activeSheet} selection={selection} />
          {compareError && (
            <div className="summary-bar" style={{ background: 'var(--accent-soft)', color: 'var(--accent)' }}>
              {compareError}
            </div>
          )}
          <Grid sheet={activeSheet} diff={activeSheetDiff} onSelectionChange={setSelection} />
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
