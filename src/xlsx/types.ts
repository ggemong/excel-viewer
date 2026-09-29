/**
 * 화면 렌더링에만 쓰이는 워크북 모델. ExcelJS로 읽은 결과를 이 형태로 뽑아내며,
 * 이 모델은 절대 파일로 다시 쓰지 않는다 — 저장 경로(src/xlsx/patch.ts, 아직 없음)는
 * 원본 파일의 raw XML을 직접 patch하는 별도 경로를 쓴다.
 */

export interface CellModel {
  address: string // e.g. "B4"
  value: string | number | boolean | null
  formula: string | null // '=' 없는 순수 수식 텍스트, 없으면 null
  numFmt: string | null
}

export interface SheetModel {
  name: string
  rowCount: number
  colCount: number
  /** rows[rowIndex][colIndex], 0-based. 빈 셀은 undefined. */
  rows: (CellModel | undefined)[][]
  merges: string[] // "B2:C3" 형태
}

export interface WorkbookModel {
  fileName: string
  sheets: SheetModel[]
}
