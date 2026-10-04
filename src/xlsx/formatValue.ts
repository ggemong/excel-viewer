/**
 * 셀 하나를 화면에 보여줄 문자열로 만든다. 실제 서식 해석은 numberFormat.ts가 하고,
 * 여기는 값 종류(불리언/숫자/문자)에 따라 어느 경로로 보낼지만 정한다.
 */

import { formatNumber, formatText } from './numberFormat'
import type { CellModel } from './types'

export function formatCellValue(cell: CellModel | undefined): string {
  if (!cell || cell.value === null || cell.value === undefined) return ''

  const { value, numFmt } = cell

  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE'
  if (typeof value === 'number') return formatNumber(value, numFmt)
  return formatText(value, numFmt)
}
