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

  const flashCopied = useCallback(() => {
    setCopied(true)
    setTimeout(() => setCopied(false), COPIED_FLASH_MS)
  }, [])

  /** 서식 없는 글자 그대로 복사(셀 상세 줄용). 셀 범위 복사(copyRange)와 달리 탭/따옴표 이스케이프를 하지 않는다. */
  const copyText = useCallback(
    async (text: string) => {
      try {
        await navigator.clipboard.writeText(text)
        flashCopied()
        return true
      } catch {
        return false
      }
    },
    [flashCopied],
  )

  const copyRange = useCallback(async (sheet: SheetModel, range: CellRange, skipRows?: boolean[]) => {
    const tsv = buildTsv(sheet, range, skipRows)
    const html = buildHtmlTable(sheet, range, skipRows)

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
      flashCopied()
      return true
    } catch {
      return false
    }
  }, [flashCopied])

  return { copyRange, copyText, copied }
}
