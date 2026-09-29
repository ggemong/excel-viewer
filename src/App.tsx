import { DropZone } from './grid/DropZone'
import { Grid } from './grid/Grid'
import { GridCards } from './grid/GridCards'
import { SummaryBar } from './grid/SummaryBar'
import { Toolbar } from './grid/Toolbar'
import { useResponsiveLayout } from './grid/useResponsiveLayout'
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
    editMode,
    setEditMode,
    editCell,
    isDirty,
    editedCount,
    saving,
    saveError,
    saveFile,
    compareFileName,
    compareLoading,
    compareError,
    loadCompareFile,
    clearCompare,
    activeSheetDiff,
  } = useWorkbookController()
  const layout = useResponsiveLayout()

  return (
    <div className="app-shell">
      <Toolbar
        fileName={workbook?.fileName ?? null}
        onOpenAnother={closeFile}
        editMode={editMode}
        onToggleEditMode={() => setEditMode((v) => !v)}
        isDirty={isDirty}
        editedCount={editedCount}
        saving={saving}
        onSave={saveFile}
        compareFileName={compareFileName}
        compareLoading={compareLoading}
        onPickCompareFile={loadCompareFile}
        onClearCompare={clearCompare}
      />

      {!workbook || !activeSheet ? (
        <DropZone onFile={openFile} error={error} loading={loading} />
      ) : (
        <>
          <SummaryBar sheet={activeSheet} />
          {(saveError || compareError) && (
            <div className="summary-bar" style={{ background: 'var(--accent-soft)', color: 'var(--accent)' }}>
              {saveError ?? compareError}
            </div>
          )}
          {layout === 'mobile' ? (
            <GridCards sheet={activeSheet} editMode={editMode} onEditCell={editCell} diff={activeSheetDiff} />
          ) : (
            <Grid sheet={activeSheet} editMode={editMode} onEditCell={editCell} diff={activeSheetDiff} />
          )}
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
