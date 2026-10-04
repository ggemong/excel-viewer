/**
 * 줄바꿈 글이 저장된 행높이보다 길면 행을 키워서 다 보이게 한다.
 *
 * 왜 필요한가: 파일의 행높이는 파일을 만든 쪽의 글꼴(맑은 고딕 등)로 맞춰 둔 값인데, 이 앱은 다른 글꼴로
 * 그려서 같은 글이라도 줄 수가 달라진다. 가만히 두면 줄바꿈 칸의 아랫부분이 잘려 읽을 수 없다.
 *
 * 지키는 규칙(Excel의 "자동 높이"와 같은 규칙):
 *  - 사용자가 높이를 직접 정한 행(customHeight, SheetModel.autoHeightRows가 false)은 건드리지 않는다.
 *    제목 배너나 얇은 구분 행처럼 의도한 높이를 키우면 모양이 깨진다.
 *  - 병합된 칸은 맞추지 않는다. Excel도 병합 칸은 자동 높이 계산에서 뺀다.
 *  - 행을 줄이지는 않는다. 저장된 높이가 충분하면 그대로 둔다.
 *  - 한 줄짜리 글은 행을 키우지 않는다. 우리 줄 높이가 Excel보다 조금 커서 모든 행이 1~2px씩 늘어나는 걸 막는다.
 */
import { formatCellValue } from '../xlsx/formatValue'
import { buildMergeLookup } from '../xlsx/mergeRange'
import type { CellStyle } from '../xlsx/cellStyle'
import type { SheetModel } from '../xlsx/types'
import type { MeasureWrappedHeights, WrapRequest } from './wrapMeasure'

/** 글 높이 위/아래에 두는 여유(px) — 아랫 테두리(1px)와 글이 칸 가장자리에 닿지 않을 정도의 숨 쉴 틈. */
const ROW_FIT_PAD_PX = 4
/** "한 줄의 높이"를 재는 표본 글자. 어떤 폭에서도 한 줄이다. */
const LINE_SAMPLE_TEXT = '가'
const LINE_SAMPLE_CELL_WIDTH = 1000

interface Candidate {
  row: number
  request: WrapRequest
  /** 이 칸의 "한 줄 높이"를 담은 표본 요청의 번호. */
  sampleKey: string
}

/** 한 줄 높이는 글자 크기에만 달려 있다 — 굵기·기울임은 줄 높이를 바꾸지 않는다. */
function sampleKeyOf(style: CellStyle | null): string {
  return String(style?.fontScale ?? 1)
}

function sampleStyleOf(style: CellStyle | null): CellStyle {
  return {
    bg: null,
    color: null,
    bold: false,
    italic: false,
    border: null,
    align: { h: null, v: null, wrap: true },
    ...(style?.fontScale ? { fontScale: style.fontScale } : {}),
  }
}

/**
 * @param sheet 줄바꿈 칸과 행높이를 가진 시트.
 * @param measure 줄바꿈된 글의 높이를 재는 함수(실제로는 wrapMeasure.ts의 DOM 측정, 테스트에서는 가짜).
 * @returns 행마다 화면에서 쓸 높이(px). 키울 행이 없으면 `sheet.rowHeights` 자체를 돌려준다(같은 참조 — 호출부가
 * 바뀐 게 없음을 참조 비교로 알 수 있다).
 */
export function fitRowHeights(sheet: SheetModel, measure: MeasureWrappedHeights): number[] {
  const { rowHeights, autoHeightRows, hiddenRows, hiddenCols, colWidths } = sheet
  if (!autoHeightRows.some(Boolean)) return rowHeights

  const merged = buildMergeLookup(sheet.merges)
  const candidates: Candidate[] = []
  const rowLimit = Math.min(sheet.rows.length, rowHeights.length)
  for (let r = 1; r <= rowLimit; r++) {
    if (!autoHeightRows[r - 1] || hiddenRows[r - 1]) continue
    const rowCells = sheet.rows[r - 1]
    for (let c = 1; c <= rowCells.length; c++) {
      const cell = rowCells[c - 1]
      if (!cell?.style?.align?.wrap || typeof cell.value !== 'string' || hiddenCols[c - 1] || merged.has(`${r},${c}`)) continue
      const text = formatCellValue(cell)
      const cellWidth = colWidths[c - 1]
      if (text === '' || !cellWidth) continue
      candidates.push({ row: r, request: { text, style: cell.style, cellWidth }, sampleKey: sampleKeyOf(cell.style) })
    }
  }
  if (candidates.length === 0) return rowHeights

  // 한 번의 측정 호출로 끝낸다: 앞쪽은 글자 크기별 "한 줄 높이" 표본, 뒤쪽은 실제 칸들.
  const samples = new Map<string, WrapRequest>()
  for (const { request, sampleKey } of candidates) {
    if (!samples.has(sampleKey)) {
      samples.set(sampleKey, { text: LINE_SAMPLE_TEXT, style: sampleStyleOf(request.style), cellWidth: LINE_SAMPLE_CELL_WIDTH })
    }
  }
  const sampleKeys = [...samples.keys()]
  const heights = measure([...samples.values(), ...candidates.map((c) => c.request)])
  const lineHeightOf = new Map(sampleKeys.map((key, i) => [key, heights[i]]))

  const fitted = rowHeights.slice()
  candidates.forEach((candidate, i) => {
    const textHeight = heights[sampleKeys.length + i]
    const lineHeight = lineHeightOf.get(candidate.sampleKey)
    if (!lineHeight || lineHeight <= 0) return
    if (Math.round(textHeight / lineHeight) <= 1) return
    const needed = Math.ceil(textHeight + ROW_FIT_PAD_PX)
    if (needed > fitted[candidate.row - 1]) fitted[candidate.row - 1] = needed
  })
  return fitted
}
