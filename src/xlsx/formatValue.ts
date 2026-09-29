/**
 * numFmt 코드의 자주 쓰이는 부분집합만 해석해서 표시용 문자열을 만든다.
 * 조건부 색상 섹션([Red] 등)이나 특이 커스텀 포맷은 다루지 않는다 — 이 프로젝트는
 * 엑셀을 픽셀 단위로 복제하는 게 목적이 아니라 "읽기 좋게 보여주는" 게 목적이기
 * 때문에 흔한 케이스(천단위 구분, 소수점, 퍼센트, 통화, 기본 날짜)만 다룬다.
 */

import type { CellModel } from './types'

export function formatCellValue(cell: CellModel | undefined): string {
  if (!cell || cell.value === null || cell.value === undefined) return ''

  const { value, numFmt } = cell

  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE'

  if (typeof value === 'number') {
    return formatNumber(value, numFmt)
  }

  if (typeof value === 'string') {
    // ISO 문자열로 정규화된 날짜값
    if (numFmt && isDateFormat(numFmt) && !Number.isNaN(Date.parse(value))) {
      return formatDate(new Date(value), numFmt)
    }
    return value
  }

  return String(value)
}

function isDateFormat(numFmt: string): boolean {
  return /[ymdh]/i.test(numFmt.replace(/"[^"]*"/g, ''))
}

function formatNumber(value: number, numFmt: string | null): string {
  if (!numFmt || numFmt === 'General') {
    return String(value)
  }

  if (numFmt.includes('%')) {
    const decimals = countDecimals(numFmt)
    return `${(value * 100).toFixed(decimals)}%`
  }

  const currencyMatch = numFmt.match(/^[$₩¥€]|"[^"]*"/)
  const prefix = currencyMatch ? currencyMatch[0].replace(/"/g, '') : ''
  const decimals = countDecimals(numFmt)
  const useThousands = numFmt.includes('#,##0') || numFmt.includes('#,#')

  const formatted = useThousands
    ? value.toLocaleString('en-US', { minimumFractionDigits: decimals, maximumFractionDigits: decimals })
    : value.toFixed(decimals)

  return prefix ? `${prefix}${formatted}` : formatted
}

function countDecimals(numFmt: string): number {
  const match = numFmt.match(/0\.(0+)/)
  return match ? match[1].length : 0
}

function formatDate(date: Date, numFmt: string): string {
  const yyyy = date.getFullYear()
  const mm = String(date.getMonth() + 1).padStart(2, '0')
  const dd = String(date.getDate()).padStart(2, '0')
  const hasTime = /h/i.test(numFmt)

  if (!hasTime) {
    if (numFmt.includes('yyyy')) return `${yyyy}-${mm}-${dd}`
    return `${mm}/${dd}/${String(yyyy).slice(2)}`
  }

  const hh = String(date.getHours()).padStart(2, '0')
  const min = String(date.getMinutes()).padStart(2, '0')
  return `${yyyy}-${mm}-${dd} ${hh}:${min}`
}
