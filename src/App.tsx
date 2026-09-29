import { DropZone } from './grid/DropZone'
import { Grid } from './grid/Grid'
import { GridCards } from './grid/GridCards'
import { SummaryBar } from './grid/SummaryBar'
import { Toolbar } from './grid/Toolbar'
import { useResponsiveLayout } from './grid/useResponsiveLayout'
import { useWorkbookController } from './state/useWorkbookController'

function App() {
  const { workbook, activeSheet, activeSheetIndex, setActiveSheetIndex, loading, error, openFile, closeFile } =
    useWorkbookController()
  const layout = useResponsiveLayout()

  return (
    <div className="app-shell">
      <Toolbar fileName={workbook?.fileName ?? null} onOpenAnother={closeFile} />

      {!workbook || !activeSheet ? (
        <DropZone onFile={openFile} error={error} loading={loading} />
      ) : (
        <>
          <SummaryBar sheet={activeSheet} />
          {layout === 'mobile' ? <GridCards sheet={activeSheet} /> : <Grid sheet={activeSheet} />}
          {workbook.sheets.length > 1 && (
            <div className="sheet-tabs">
              {workbook.sheets.map((sheet, i) => (
                <button
                  key={sheet.name}
                  type="button"
                  className="sheet-tab"
                  aria-current={i === activeSheetIndex}
                  onClick={() => setActiveSheetIndex(i)}
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
