import { useEffect, useState } from 'react'
import type { CellRange } from './clipboard/buildClipboardPayload'
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
  const [selection, setSelection] = useState<CellRange | null>(null)

  // Ctrl/Cmd+S로 저장 — 그리드가 아니라 앱 전체에서 동작해야 하는 단축키라
  // (모바일 카드뷰일 때도 눌릴 수 있음) Grid 안이 아니라 여기서 처리한다.
  // 브라우저 기본 "페이지 저장" 동작은 항상 막는다.
  useEffect(() => {
    const onKeyDown = (e: KeyboardEvent) => {
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 's') {
        e.preventDefault()
        if (isDirty && !saving) void saveFile()
      }
    }
    window.addEventListener('keydown', onKeyDown)
    return () => window.removeEventListener('keydown', onKeyDown)
  }, [isDirty, saving, saveFile])

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
          <SummaryBar sheet={activeSheet} selection={layout === 'mobile' ? null : selection} />
          {(saveError || compareError) && (
            <div className="summary-bar" style={{ background: 'var(--accent-soft)', color: 'var(--accent)' }}>
              {saveError ?? compareError}
            </div>
          )}
          {layout === 'mobile' ? (
            <GridCards sheet={activeSheet} editMode={editMode} onEditCell={editCell} diff={activeSheetDiff} />
          ) : (
            <Grid
              sheet={activeSheet}
              editMode={editMode}
              onEditCell={editCell}
              diff={activeSheetDiff}
              onSelectionChange={setSelection}
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
