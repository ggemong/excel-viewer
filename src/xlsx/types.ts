/**
 * 화면 렌더링에만 쓰이는 워크북 모델. ExcelJS로 읽은 결과를 이 형태로 뽑아낸다.
 * 이 프로젝트는 읽기 전용 뷰어라(D-006) 이 모델을 다시 파일로 쓰는 경로는 없다.
 */
import type { CellStyle } from './cellStyle'

export type { CellStyle } from './cellStyle'

export interface CellModel {
  address: string // e.g. "B4"
  value: string | number | boolean | null
  formula: string | null // '=' 없는 순수 수식 텍스트, 없으면 null
  numFmt: string | null
  /** 배경색·글자색·굵게/기울임(표시 전용) — src/xlsx/cellStyle.ts 참고. CSV는 항상 null. */
  style: CellStyle | null
  /**
   * http/https/mailto 스킴의 하이퍼링크 URL(그 외 스킴·내부(같은 통합문서) 링크는
   * null — src/xlsx/read.ts의 cellToModel 참고). CSV는 항상 null.
   */
  hyperlink: string | null
}

export interface SheetModel {
  name: string
  rowCount: number
  colCount: number
  /** rows[rowIndex][colIndex], 0-based. 빈 셀은 undefined. */
  rows: (CellModel | undefined)[][]
  merges: string[] // "B2:C3" 형태
  /** px, 0-based(colWidths[0] = A열). CSV는 전부 기본값. */
  colWidths: number[]
  /** px, 0-based(rowHeights[0] = 1행). CSV는 전부 기본값. */
  rowHeights: number[]
  /** true인 인덱스는 숨긴 열/행 — Grid.tsx가 건너뛴다(Phase D). */
  hiddenCols: boolean[]
  hiddenRows: boolean[]
  /** 틀고정된 행 수(화면 최상단에 항상 보이는 행). 없으면 null. 열고정은 범위 밖(D-006 계획 참고). */
  frozen: { rows: number } | null
}

export interface WorkbookModel {
  fileName: string
  sheets: SheetModel[]
}
