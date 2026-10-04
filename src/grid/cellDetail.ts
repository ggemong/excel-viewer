/**
 * 선택한 칸의 "전체 내용"을 계산하는 순수 함수. 칸이 좁아 글자가 잘려도(실제 파일에서 글자 있는
 * 칸의 32%가 잘려 있었다) 여기서 만든 내용으로 전부 읽을 수 있다 — Excel의 수식 입력줄과 같은 역할.
 *
 * 병합 범위의 아무 칸을 선택해도 값이 있는 왼쪽 위 칸(마스터) 내용을 보여준다. 화면은 마스터만 그리고
 * 나머지 칸은 선택이 안 되지만, 범위 선택의 왼쪽 위가 병합 안쪽일 수 있기 때문이다.
 */
import { cellAddress } from '../xlsx/cellRef'
import { formatCellValue } from '../xlsx/formatValue'
import { findMergeAt, parseMergeRanges } from '../xlsx/mergeRange'
import { isDateTimeFormat } from '../xlsx/numberFormat'
import type { SheetModel } from '../xlsx/types'

export interface CellDetail {
  /** 내용을 가진 칸의 주소(병합이면 마스터 주소). */
  address: string
  /** 화면에 보이는 글자(서식 적용 후) 전체. */
  text: string
  /** 보이는 글자와 다른 원래 값(예: 보이는 `1,000원` ↔ 원래 `1000`). 같거나 의미 없으면 null. */
  originalValue: string | null
  /** 수식 텍스트(`=A1+B1`). 없거나 알 수 없으면 null(공유 수식의 따라가는 칸은 원본이 안 준다). */
  formula: string | null
  hyperlink: string | null
  note: string | null
}

export function describeCell(sheet: SheetModel, row: number, col: number): CellDetail {
  const merge = findMergeAt(parseMergeRanges(sheet.merges), row, col)
  const r = merge ? merge.r0 : row
  const c = merge ? merge.c0 : col
  const cell = sheet.rows[r - 1]?.[c - 1]
  const address = cellAddress(r, c)
  if (!cell) return { address, text: '', originalValue: null, formula: null, hyperlink: null, note: null }

  const text = formatCellValue(cell)
  // 날짜 셀의 원래 값은 직렬번호(46299)라 사용자에게 의미가 없어 보여주지 않는다.
  const raw = cell.value === null || isDateTimeFormat(cell.numFmt) ? null : String(cell.value)
  return {
    address,
    text,
    originalValue: raw !== null && raw !== text ? raw : null,
    formula: cell.formula ? `=${cell.formula}` : null,
    hyperlink: cell.hyperlink,
    note: cell.note ?? null,
  }
}

/** 보여줄 게 하나라도 있는가 — 빈 칸이면 상세 줄에 "빈 셀"만 보여준다. */
export function hasDetail(detail: CellDetail): boolean {
  return detail.text !== '' || detail.hyperlink !== null || detail.note !== null || detail.formula !== null
}
