import { useCallback, useState } from 'react'
import { buildHtmlTable, buildTsv, type CellRange } from './buildClipboardPayload'
import type { SheetModel } from '../xlsx/types'

/**
 * text/plain(TSV) + text/html(<table>)를 동시에 클립보드에 올린다 — Excel/Sheets는
 * html을 읽어서 셀 경계를 정확히 복원하고, 그 외(Slack, 메모장 등)는 TSV로 받는다.
 * 클립보드 API는 사용자 제스처(클릭 등) 핸들러 안에서 동기적으로 불러야 한다.
 */
/** "복사됨" 표시를 띄워두는 시간 — 너무 짧으면 못 보고, 너무 길면 다음 복사와 헷갈린다. */
const COPIED_FLASH_MS = 1500

export function useClipboardCopy() {
  const [copied, setCopied] = useState(false)

  const copyRange = useCallback(async (sheet: SheetModel, range: CellRange) => {
    const tsv = buildTsv(sheet, range)
    const html = buildHtmlTable(sheet, range)

    try {
      if (typeof ClipboardItem !== 'undefined') {
        await navigator.clipboard.write([
          new ClipboardItem({
            'text/plain': new Blob([tsv], { type: 'text/plain' }),
            'text/html': new Blob([html], { type: 'text/html' }),
          }),
        ])
      } else {
        await navigator.clipboard.writeText(tsv)
      }
      setCopied(true)
      setTimeout(() => setCopied(false), COPIED_FLASH_MS)
      return true
    } catch {
      return false
    }
  }, [])

  return { copyRange, copied }
}
