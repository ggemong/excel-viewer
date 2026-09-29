import { useCallback, useState } from 'react'
import { readWorkbook } from '../xlsx/read'
import type { WorkbookModel } from '../xlsx/types'

const SUPPORTED_EXTENSIONS = ['.xlsx', '.csv']

export function useWorkbookController() {
  const [workbook, setWorkbook] = useState<WorkbookModel | null>(null)
  const [activeSheetIndex, setActiveSheetIndex] = useState(0)
  const [loading, setLoading] = useState(false)
  const [error, setError] = useState<string | null>(null)

  const openFile = useCallback(async (file: File) => {
    const lower = file.name.toLowerCase()
    if (!SUPPORTED_EXTENSIONS.some((ext) => lower.endsWith(ext))) {
      setError('.xlsx 또는 .csv 파일만 열 수 있어요.')
      return
    }

    setLoading(true)
    setError(null)
    try {
      const model = await readWorkbook(file)
      setWorkbook(model)
      setActiveSheetIndex(0)
    } catch {
      setError('파일을 읽지 못했어요. 손상되었거나 지원하지 않는 형식일 수 있어요.')
    } finally {
      setLoading(false)
    }
  }, [])

  const closeFile = useCallback(() => {
    setWorkbook(null)
    setActiveSheetIndex(0)
    setError(null)
  }, [])

  return {
    workbook,
    activeSheet: workbook?.sheets[activeSheetIndex] ?? null,
    activeSheetIndex,
    setActiveSheetIndex,
    loading,
    error,
    openFile,
    closeFile,
  }
}
