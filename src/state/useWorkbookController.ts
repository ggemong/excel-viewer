import { useCallback, useMemo, useState } from 'react'
import { diffWorkbooks } from '../diff/diffWorkbooks'
import { readWorkbook } from '../xlsx/read'
import type { WorkbookModel } from '../xlsx/types'

/** 지원하는 파일 확장자 — 열기 검증(아래)과 파일 선택 UI(DropZone/Toolbar의 accept)가 같이 쓴다. */
export const SUPPORTED_EXTENSIONS = ['.xlsx', '.csv']

export function useWorkbookController() {
  const [workbook, setWorkbook] = useState<WorkbookModel | null>(null)
  const [activeSheetIndex, setActiveSheetIndex] = useState(0)
  const [loading, setLoading] = useState(false)
  const [error, setError] = useState<string | null>(null)
  const [compareWorkbook, setCompareWorkbook] = useState<WorkbookModel | null>(null)
  const [compareFileName, setCompareFileName] = useState<string | null>(null)
  const [compareLoading, setCompareLoading] = useState(false)
  const [compareError, setCompareError] = useState<string | null>(null)

  /** 새 파일을 열 때 이전 파일의 비교 상태가 새어 들어가지 않게 되돌린다. */
  const resetCompareState = () => {
    setActiveSheetIndex(0)
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
      resetCompareState()
    } catch {
      setError('파일을 읽지 못했어요. 손상되었거나 지원하지 않는 형식일 수 있어요.')
    } finally {
      setLoading(false)
    }
  }, [])

  const closeFile = useCallback(() => {
    setWorkbook(null)
    setError(null)
    resetCompareState()
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
    compareFileName,
    compareLoading,
    compareError,
    loadCompareFile,
    clearCompare,
    activeSheetDiff,
  }
}
