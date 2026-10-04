import { columnLetter } from './cellRef'
import { extractCellStyle } from './cellStyle'
import { interpretCellValue, interpretNote } from './cellValue'
import { excelColumnWidthToPx, excelPointsToPx, measureDefaultFontWidth } from './columnWidth'
import { readDefaultFontPt } from './defaultFont'
import { isLiteralHyperlinkFormula, linkFromFormula, sanitizeHyperlink } from './hyperlink'
import { isLinkFormula } from '../formula/linkTarget'
import { readWorkbookDrawings, type SheetDrawings } from './drawing'
import { extractConditionalFormats, type RawConditionalFormatting } from './conditionalFormat'
import { buildAxis, requiredExtent } from './drawingLayout'
import { autoFilterRangeText, buildSheetFilter, readAutoFilterDetails, type AutoFilterDetails } from './filter'
import { readCustomHeightRows } from './rowHeights'
import { parseTheme, type ThemeColors } from './themeColor'
import type { CellModel, SheetModel, WorkbookModel } from './types'
import { listSheetParts, readThemeXml } from './workbookParts'
import { ZipArchive } from './zipReader'

// 파일이 기본 열너비/행높이(sheetFormatPr)조차 안 적었을 때만 쓰는 Excel 통상 기본값.
// 실제 파일은 보통 적어 두고(예: 맑은 고딕 기반 파일은 행높이 16.5pt, 열너비 9), 그 값을
// 쓰지 않고 15pt/8.43으로 고정하면 명시 안 된 모든 행/열이 어긋난다.
const DEFAULT_COLUMN_CHAR_WIDTH = 8.43
const DEFAULT_ROW_HEIGHT_PT = 15

/**
 * ExcelJS가 읽을 때 건너뛸 XML 노드. 데이터 유효성 검사(드롭다운 목록 등)는 읽기 전용 뷰어가 쓰지 않는데, ExcelJS는
 * 그 적용 범위를 칸 하나하나로 펼쳐 저장한다. 열 전체(`L4:L1048576`)에 걸린 규칙이 있으면 시트 하나가 100만 칸이 된다 —
 * 실제 업무 파일(칸 4천 개, 시트 59개)에서 열기가 6.2초·메모리 450MB·화면 멈춤 3.7초였고 이것을 건너뛰면 2.6초·130MB·0.05초가
 * 됐다(운영 빌드에서 실측). 이 값을 늘릴 때는 "우리가 쓰지 않는 노드인가"를 먼저 확인한다.
 */
export const EXCELJS_IGNORED_NODES = ['dataValidations']

/**
 * 파일을 읽어서 화면 표시용 모델을 만든다. ExcelJS는 파일을 열 때만 동적으로
 * 불러온다 — 초기 로딩 번들을 가볍게 유지하기 위해서다 (구현 계획 참고).
 */
export async function readWorkbook(file: File): Promise<WorkbookModel> {
  const lower = file.name.toLowerCase()
  if (lower.endsWith('.csv')) {
    return readCsv(file)
  }
  return readXlsx(file)
}

async function readXlsx(file: File): Promise<WorkbookModel> {
  const ExcelJS = await import('exceljs')
  const buffer = await file.arrayBuffer()
  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.load(buffer, { ignoreNodes: EXCELJS_IGNORED_NODES })

  const mdw = measureDefaultFontWidth()

  // 그림/도형은 ExcelJS가 일부만 읽고 나머지는 버려서 원본 XML을 직접 읽는다(src/xlsx/drawing.ts).
  // 이쪽이 실패해도 셀 데이터는 이미 읽혔으므로 파일 열기 자체는 막지 않고 경고로만 알린다.
  const warnings: string[] = []
  const describe = (err: unknown) => (err instanceof Error ? err.message : String(err))
  let drawingsBySheet = new Map<string, SheetDrawings>()
  // 자동 필터의 "조건"(어느 열이 필터링 중인지)도 ExcelJS가 버리는 정보라 같은 ZIP에서 읽는다.
  // 필터 범위 자체는 ExcelJS가 주므로, 필터가 있는 시트만 원본 XML을 연다(큰 시트 이중 해제 방지).
  const filterDetailsBySheet = new Map<string, AutoFilterDetails | null>()
  // 시트 이름 -> 시트 XML 경로(필터 조건, 행높이 정보를 원본에서 읽을 때 쓴다).
  let sheetPaths = new Map<string, string>()
  // 통합문서 기본 글꼴 크기(pt). 못 읽으면 null이고, 그러면 칸별 글자 크기를 조정하지 않는다(src/xlsx/defaultFont.ts).
  let defaultFontPt: number | null = null
  // 아래 원본 XML 읽기들은 서로 독립이다 — 하나가 실패해도 나머지와 셀 데이터는 그대로 보여준다.
  let zip: ZipArchive | null = null
  try {
    zip = ZipArchive.open(buffer)
  } catch (err) {
    warnings.push(`파일 구조를 읽지 못해 그림/도형·필터 조건·글자 크기·테마 색을 반영하지 못했어요: ${describe(err)}`)
  }
  // 테마는 워크북 하나당 한 번만 파싱한다(셀마다 다시 파싱하지 않음). 테마 색을 모르면 테마 색을 쓰는 칸의 색이 빠진다.
  let theme: ThemeColors | null = null
  if (zip) {
    try {
      const themeXml = await readThemeXml(zip)
      theme = themeXml ? parseTheme(themeXml) : null
    } catch (err) {
      warnings.push(`테마 색을 읽지 못해 일부 색이 다르게 보일 수 있어요: ${describe(err)}`)
    }
    try {
      defaultFontPt = await readDefaultFontPt(zip)
    } catch (err) {
      warnings.push(`기본 글꼴 크기를 읽지 못해 글자 크기를 반영하지 못했어요: ${describe(err)}`)
    }
    try {
      const result = await readWorkbookDrawings(zip, theme)
      drawingsBySheet = result.bySheet
      warnings.push(...result.warnings)
    } catch (err) {
      warnings.push(`그림/도형을 읽지 못했어요: ${describe(err)}`)
    }
    try {
      sheetPaths = new Map((await listSheetParts(zip)).map((part) => [part.name, part.path]))
      for (const ws of workbook.worksheets) {
        const path = sheetPaths.get(ws.name)
        if (!ws.autoFilter || ws.state !== 'visible' || !path) continue
        try {
          filterDetailsBySheet.set(ws.name, await readAutoFilterDetails(zip, path))
        } catch (err) {
          warnings.push(`"${ws.name}" 시트의 필터 조건을 읽지 못했어요: ${describe(err)}`)
        }
      }
    } catch (err) {
      warnings.push(`시트 목록을 읽지 못해 필터 조건·행 높이 정보를 반영하지 못했어요: ${describe(err)}`)
    }
  }

  const sheets: SheetModel[] = []
  for (const worksheet of workbook.worksheets) {
    // 숨김/완전숨김 시트는 Excel도 탭에 안 보여준다(Excel의 "숨기기" 의도 존중).
    if (worksheet.state !== 'visible') continue

    const defaultRowPt = worksheet.properties?.defaultRowHeight ?? DEFAULT_ROW_HEIGHT_PT
    const defaultColChars = worksheet.properties?.defaultColWidth ?? DEFAULT_COLUMN_CHAR_WIDTH
    const defaultRowPx = excelPointsToPx(defaultRowPt)
    const defaultColPx = excelColumnWidthToPx(defaultColChars, mdw)

    const rowCount = worksheet.rowCount
    const colCount = worksheet.columnCount
    const rows: (CellModel | undefined)[][] = []
    const rowHeights: number[] = []
    const hiddenRows: boolean[] = []
    let hasWrappedText = false

    for (let r = 1; r <= rowCount; r++) {
      const row = worksheet.getRow(r)
      rowHeights.push(row.height === undefined ? defaultRowPx : excelPointsToPx(row.height))
      hiddenRows.push(Boolean(row.hidden))

      const rowCells: (CellModel | undefined)[] = []
      for (let c = 1; c <= colCount; c++) {
        const cell = row.getCell(c)
        // 값이 없어도 메모가 있는 칸이 오면 남긴다(메모 표시와 상세 보기에 필요). 다만 ExcelJS는 값 없는 칸의
        // 메모를 쓰기/읽기 왕복에서 보존하지 않아서 이 경로를 시험으로는 확인하지 못했다.
        const note = interpretNote(cell.note)
        if ((cell.type === undefined || cell.value === null || cell.value === undefined) && !note) {
          rowCells.push(undefined)
          continue
        }
        const model = cellToModel(cell, theme, note, defaultFontPt)
        if (model.style?.align?.wrap && typeof model.value === 'string') hasWrappedText = true
        rowCells.push(model)
      }
      rows.push(rowCells)
    }

    const colWidths: number[] = []
    const hiddenCols: boolean[] = []
    for (let c = 1; c <= colCount; c++) {
      const col = worksheet.getColumn(c)
      colWidths.push(col.width === undefined ? defaultColPx : excelColumnWidthToPx(col.width, mdw))
      hiddenCols.push(Boolean(col.hidden))
    }

    // 그림/도형이 데이터 영역 밖까지 걸쳐 있으면(셀은 0행인데 스크린샷만 있는 시트 등) 그리드가
    // 그 영역을 빈 칸으로 이어서 그려야 하므로 시트 크기를 늘리고, 모자란 칸은 기본 크기로 채운다.
    const drawing = drawingsBySheet.get(worksheet.name)
    const extent = drawing
      ? requiredExtent(drawing.items, buildAxis(colWidths, hiddenCols, defaultColPx), buildAxis(rowHeights, hiddenRows, defaultRowPx))
      : { cols: 0, rows: 0 }
    while (colWidths.length < extent.cols) {
      colWidths.push(defaultColPx)
      hiddenCols.push(false)
    }
    while (rowHeights.length < extent.rows) {
      rowHeights.push(defaultRowPx)
      hiddenRows.push(false)
    }

    // 줄바꿈 글이 있는 시트만 "어느 행이 직접 정한 높이인가"를 원본에서 읽는다(큰 시트 XML을 불필요하게 또 읽지 않는다).
    const autoHeightRows: boolean[] = new Array(rowHeights.length).fill(false)
    const sheetPath = sheetPaths.get(worksheet.name)
    if (hasWrappedText && zip && sheetPath) {
      try {
        const customRows = await readCustomHeightRows(zip, sheetPath)
        for (let r = 1; r <= rowCount; r++) autoHeightRows[r - 1] = !customRows.has(r)
      } catch (err) {
        warnings.push(`"${worksheet.name}" 시트의 행 높이 정보를 읽지 못해 줄바꿈 글의 행 높이를 자동으로 맞추지 못했어요: ${describe(err)}`)
      }
    }

    // ExcelJS의 views 타입 선언(Array<Partial<WorksheetView>>)은 state로 좁혀도
    // ySplit이 안 보인다(frozen 전용 필드인데 Partial이 판별 유니온 좁히기를 못
    // 살림) — 필요한 필드만 최소 타입으로 캐스팅한다.
    const frozenView = worksheet.views?.find((v) => v.state === 'frozen') as { xSplit?: number; ySplit?: number } | undefined
    const frozenRows = frozenView?.ySplit ?? 0
    const frozenCols = frozenView?.xSplit ?? 0
    const frozen = frozenRows > 0 || frozenCols > 0 ? { rows: frozenRows, cols: frozenCols } : null

    const merges = mergeRangesOf(worksheet)

    const conditional = extractConditionalFormats(
      // ExcelJS의 타입 선언에는 이 속성이 빠져 있다(런타임에는 있다).
      (worksheet as unknown as { conditionalFormattings: RawConditionalFormatting[] }).conditionalFormattings,
      theme,
    )

    const filterRange = autoFilterRangeText(worksheet.autoFilter)
    const sheetFilter = filterRange ? buildSheetFilter(filterRange, filterDetailsBySheet.get(worksheet.name) ?? null) : null

    sheets.push({
      name: worksheet.name,
      rowCount: Math.max(rowCount, extent.rows),
      colCount: Math.max(colCount, extent.cols),
      rows,
      merges,
      colWidths,
      rowHeights,
      autoHeightRows,
      hiddenCols,
      hiddenRows,
      frozen,
      filters: sheetFilter ? [sheetFilter] : [],
      conditionalFormats: conditional.formats,
      skippedConditionalFormats: conditional.skipped,
      drawings: drawing?.items ?? [],
      skippedDrawings: drawing?.skipped ?? {},
    })
  }

  return { fileName: file.name, sheets, warnings }
}

/**
 * 시트의 병합 범위("B2:C3" 목록). ExcelJS의 `worksheet.model.merges`와 같은 값이지만, 그 getter는 시트의 모든 칸을 모델로
 * 새로 만들어서 시트 59개에 0.64초가 들었다(운영 빌드 실측, 실제 칸 순회는 30ms). ExcelJS가 내부에 쥔 병합 목록(`_merges`)을
 * 바로 읽는다. ExcelJS를 올려서 이 내부 구조가 바뀌면 조용히 병합을 놓치지 않고, 느려도 정확한 model로 되돌아간다.
 */
export function mergeRangesOf(worksheet: import('exceljs').Worksheet): string[] {
  const internal = (worksheet as unknown as { _merges?: Record<string, { range?: unknown }> })._merges
  if (internal && typeof internal === 'object') {
    const ranges = Object.values(internal).map((merge) => merge.range)
    if (ranges.every((range): range is string => typeof range === 'string')) return ranges
  }
  return worksheet.model.merges ?? []
}

/**
 * 칸의 링크. 파일이 칸에 직접 건 외부 링크가 우선이고, 없으면 `=HYPERLINK(...)` 수식에서 읽는다(외부 주소 또는 같은 통합문서 안 이동).
 * 외부 주소는 허용된 스킴만 남긴다.
 */
function linkOf(hyperlink: string | null, formula: string | null, value: CellModel['value']): Pick<CellModel, 'hyperlink' | 'internalLink' | 'computedLink'> {
  const direct = hyperlink ? sanitizeHyperlink(hyperlink) : null
  if (direct) return { hyperlink: direct }
  const fromFormula = formula ? linkFromFormula(formula) : null
  if (fromFormula?.kind === 'internal') return { hyperlink: null, internalLink: fromFormula.link }
  if (fromFormula?.kind === 'external') return { hyperlink: fromFormula.url }
  // 대상이 글자 그대로가 아니라 계산식인 HYPERLINK: 링크로 보이고, 눌렀을 때 계산한다. 리터럴인데 안전하지 않은 주소면 링크로 만들지 않는다.
  // 저장된 결과가 비어 있으면 눌러도 보이는 것이 없다 — 대상 계산이 오류라서 IFERROR가 빈 글자를 낸 칸이므로 링크로 만들지 않는다.
  const shown = value !== null && value !== ''
  if (formula && shown && isLinkFormula(formula) && !isLiteralHyperlinkFormula(formula)) return { hyperlink: null, computedLink: true }
  return { hyperlink: null }
}

function cellToModel(cell: import('exceljs').Cell, theme: ThemeColors | null, note: string | undefined, defaultFontPt: number | null): CellModel {
  const { value, formula, hyperlink } = interpretCellValue(cell.value)

  return {
    address: cell.address,
    value,
    formula,
    numFmt: cell.numFmt ?? null,
    style: extractCellStyle(cell, theme, defaultFontPt),
    ...linkOf(hyperlink, formula, value),
    ...(note ? { note } : {}),
  }
}

async function readCsv(file: File): Promise<WorkbookModel> {
  const text = await file.text()
  const delimiter = sniffDelimiter(text)
  const lines = text.split(/\r\n|\n|\r/).filter((line, i, arr) => !(i === arr.length - 1 && line === ''))

  const rows: (CellModel | undefined)[][] = lines.map((line, rowIndex) =>
    splitCsvLine(line, delimiter).map((value, colIndex) => ({
      address: `${columnLetter(colIndex + 1)}${rowIndex + 1}`,
      value: value === '' ? null : value,
      formula: null,
      numFmt: null,
      style: null,
      hyperlink: null,
    })),
  )

  const rowCount = rows.length
  const colCount = rows.reduce((max, row) => Math.max(max, row.length), 0)
  // CSV는 열너비/행높이/숨김/틀고정 개념 자체가 없다 — 전부 Excel 기본값 하나로 균일하게 채운다.
  const defaultColWidth = excelColumnWidthToPx(DEFAULT_COLUMN_CHAR_WIDTH, measureDefaultFontWidth())
  const defaultRowHeight = excelPointsToPx(DEFAULT_ROW_HEIGHT_PT)

  return {
    fileName: file.name,
    sheets: [
      {
        name: 'Sheet1',
        rowCount,
        colCount,
        rows,
        merges: [],
        colWidths: Array(colCount).fill(defaultColWidth),
        rowHeights: Array(rowCount).fill(defaultRowHeight),
        autoHeightRows: Array(rowCount).fill(false),
        hiddenCols: Array(colCount).fill(false),
        hiddenRows: Array(rowCount).fill(false),
        frozen: null,
        filters: [],
        conditionalFormats: [],
        skippedConditionalFormats: {},
        drawings: [],
        skippedDrawings: {},
      },
    ],
    warnings: [],
  }
}

function sniffDelimiter(text: string): string {
  const firstLine = text.slice(0, text.indexOf('\n') === -1 ? text.length : text.indexOf('\n'))
  const comma = (firstLine.match(/,/g) ?? []).length
  const semicolon = (firstLine.match(/;/g) ?? []).length
  return semicolon > comma ? ';' : ','
}

function splitCsvLine(line: string, delimiter: string): string[] {
  const cells: string[] = []
  let current = ''
  let inQuotes = false
  for (let i = 0; i < line.length; i++) {
    const ch = line[i]
    if (inQuotes) {
      if (ch === '"') {
        if (line[i + 1] === '"') {
          current += '"'
          i++
        } else {
          inQuotes = false
        }
      } else {
        current += ch
      }
    } else if (ch === '"') {
      inQuotes = true
    } else if (ch === delimiter) {
      cells.push(current)
      current = ''
    } else {
      current += ch
    }
  }
  cells.push(current)
  return cells
}
