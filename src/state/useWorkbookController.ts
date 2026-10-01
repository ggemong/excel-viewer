import { useCallback, useMemo, useState } from 'react'
import { diffWorkbooks } from '../diff/diffWorkbooks'
import { cellAddress, parseCellAddress } from '../xlsx/cellRef'
import { isPlainNumberLiteral } from '../xlsx/cellValue'
import { findMergeAt, isMergeMaster, parseMergeRanges } from '../xlsx/mergeRange'
import type { CellEdit } from '../xlsx/patch'
import { readWorkbook } from '../xlsx/read'
import type { WorkbookModel } from '../xlsx/types'

/** 지원하는 파일 확장자 — 열기 검증(아래)과 파일 선택 UI(DropZone/Toolbar의 accept)가 같이 쓴다. */
export const SUPPORTED_EXTENSIONS = ['.xlsx', '.csv']

/** 입력값을 편집 로직이 쓸 리터럴로 정규화한다: 빈 문자열은 삭제(null), 숫자 패턴은 숫자로. */
function normalizeInput(raw: string): string | number | null {
  const trimmed = raw.trim()
  if (trimmed === '') return null
  if (isPlainNumberLiteral(trimmed)) return Number(trimmed)
  return trimmed
}

function updateSheetCell(
  workbook: WorkbookModel,
  sheetIndex: number,
  row: number,
  col: number,
  newValue: string | number | null,
): WorkbookModel {
  const sheets = workbook.sheets.slice()
  const sheet = sheets[sheetIndex]
  const rows = sheet.rows.slice()
  const rowArr = (rows[row - 1] ?? []).slice()
  const prev = rowArr[col - 1]

  rowArr[col - 1] = newValue === null ? undefined : { address: cellAddress(row, col), value: newValue, formula: null, numFmt: prev?.numFmt ?? null }

  rows[row - 1] = rowArr
  sheets[sheetIndex] = { ...sheet, rows }
  return { ...workbook, sheets }
}

function suggestedFileName(originalName: string): string {
  const dot = originalName.lastIndexOf('.')
  if (dot === -1) return `${originalName}-edited`
  return `${originalName.slice(0, dot)}-edited${originalName.slice(dot)}`
}

function downloadBlob(blob: Blob, fileName: string) {
  const url = URL.createObjectURL(blob)
  const a = document.createElement('a')
  a.href = url
  a.download = fileName
  a.click()
  URL.revokeObjectURL(url)
}

export function useWorkbookController() {
  const [file, setFile] = useState<File | null>(null)
  const [workbook, setWorkbook] = useState<WorkbookModel | null>(null)
  const [activeSheetIndex, setActiveSheetIndex] = useState(0)
  const [loading, setLoading] = useState(false)
  const [error, setError] = useState<string | null>(null)
  const [editMode, setEditMode] = useState(false)
  const [edits, setEdits] = useState<Map<string, CellEdit>>(new Map())
  const [saving, setSaving] = useState(false)
  const [saveError, setSaveError] = useState<string | null>(null)
  const [compareWorkbook, setCompareWorkbook] = useState<WorkbookModel | null>(null)
  const [compareFileName, setCompareFileName] = useState<string | null>(null)
  const [compareLoading, setCompareLoading] = useState(false)
  const [compareError, setCompareError] = useState<string | null>(null)

  /** 파일을 새로 열거나 닫을 때 공통으로 되돌려야 하는 상태 — 이전 파일의 편집/비교 상태가 새 파일로 새어 들어가지 않게 한다. */
  const resetEditAndCompareState = () => {
    setActiveSheetIndex(0)
    setEdits(new Map())
    setEditMode(false)
    setSaveError(null)
    setCompareWorkbook(null)
    setCompareFileName(null)
    setCompareError(null)
  }

  const openFile = useCallback(async (nextFile: File) => {
    const lower = nextFile.name.toLowerCase()
    if (!SUPPORTED_EXTENSIONS.some((ext) => lower.endsWith(ext))) {
      setError('.xlsx 또는 .csv 파일만 열 수 있어요.')
      return
    }

    setLoading(true)
    setError(null)
    try {
      const model = await readWorkbook(nextFile)
      setWorkbook(model)
      setFile(nextFile)
      resetEditAndCompareState()
    } catch {
      setError('파일을 읽지 못했어요. 손상되었거나 지원하지 않는 형식일 수 있어요.')
    } finally {
      setLoading(false)
    }
  }, [])

  const closeFile = useCallback(() => {
    setWorkbook(null)
    setFile(null)
    setError(null)
    resetEditAndCompareState()
  }, [])

  const loadCompareFile = useCallback(async (compareCandidate: File) => {
    setCompareLoading(true)
    setCompareError(null)
    try {
      const model = await readWorkbook(compareCandidate)
      setCompareWorkbook(model)
      setCompareFileName(compareCandidate.name)
    } catch {
      setCompareError('비교할 파일을 읽지 못했어요.')
    } finally {
      setCompareLoading(false)
    }
  }, [])

  const clearCompare = useCallback(() => {
    setCompareWorkbook(null)
    setCompareFileName(null)
    setCompareError(null)
  }, [])

  const activeSheet = workbook?.sheets[activeSheetIndex] ?? null

  const editCell = useCallback(
    (address: string, rawInput: string) => {
      if (!workbook || !activeSheet) return
      const { row, col } = parseCellAddress(address)
      const existing = activeSheet.rows[row - 1]?.[col - 1]
      if (existing?.formula) {
        setSaveError(`${address} 셀은 수식이 있어서 수정할 수 없어요.`)
        return
      }
      const merge = findMergeAt(parseMergeRanges(activeSheet.merges), row, col)
      if (merge && !isMergeMaster(merge, row, col)) {
        setSaveError(`${address} 셀은 병합된 셀이에요 — 왼쪽 위 기준 셀에서 수정하세요.`)
        return
      }

      const newValue = normalizeInput(rawInput)

      setWorkbook((wb) => (wb ? updateSheetCell(wb, activeSheetIndex, row, col, newValue) : wb))
      setEdits((prev) => {
        const next = new Map(prev)
        next.set(`${activeSheet.name}::${address}`, { sheetName: activeSheet.name, address, newValue })
        return next
      })
    },
    [workbook, activeSheet, activeSheetIndex],
  )

  const saveFile = useCallback(async () => {
    if (!file || edits.size === 0) return
    setSaving(true)
    setSaveError(null)
    try {
      const { patchWorkbook } = await import('../xlsx/patch')
      const blob = await patchWorkbook(file, Array.from(edits.values()))
      downloadBlob(blob, suggestedFileName(file.name))
    } catch (e) {
      setSaveError(e instanceof Error ? e.message : '저장하지 못했어요.')
    } finally {
      setSaving(false)
    }
  }, [file, edits])

  const isDirty = useMemo(() => edits.size > 0, [edits])

  const diff = useMemo(
    () => (workbook && compareWorkbook ? diffWorkbooks(workbook, compareWorkbook) : null),
    [workbook, compareWorkbook],
  )
  const activeSheetDiff = activeSheet ? (diff?.sheets.get(activeSheet.name) ?? null) : null

  return {
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
    editedCount: edits.size,
    saving,
    saveError,
    saveFile,
    compareFileName,
    compareLoading,
    compareError,
    loadCompareFile,
    clearCompare,
    diff,
    activeSheetDiff,
  }
}
