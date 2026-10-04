import { columnLetter } from './cellRef'
import { extractCellStyle } from './cellStyle'
import { interpretCellValue, interpretNote } from './cellValue'
import { excelColumnWidthToPx, excelPointsToPx, measureDefaultFontWidth } from './columnWidth'
import { readWorkbookDrawings, type SheetDrawings } from './drawing'
import { extractConditionalFormats, type RawConditionalFormatting } from './conditionalFormat'
import { buildAxis, requiredExtent } from './drawingLayout'
import { autoFilterRangeText, buildSheetFilter, readAutoFilterDetails, type AutoFilterDetails } from './filter'
import { parseTheme, type ThemeColors } from './themeColor'
import type { CellModel, SheetModel, WorkbookModel } from './types'
import { listSheetParts } from './workbookParts'
import { ZipArchive } from './zipReader'

// 파일이 기본 열너비/행높이(sheetFormatPr)조차 안 적었을 때만 쓰는 Excel 통상 기본값.
// 실제 파일은 보통 적어 두고(예: 맑은 고딕 기반 파일은 행높이 16.5pt, 열너비 9), 그 값을
// 쓰지 않고 15pt/8.43으로 고정하면 명시 안 된 모든 행/열이 어긋난다.
const DEFAULT_COLUMN_CHAR_WIDTH = 8.43
const DEFAULT_ROW_HEIGHT_PT = 15

/** 사용자가 올린 임의 파일을 그대로 열어주는 뷰어라, 하이퍼링크는 이 스킴만 신뢰한다 —
 * javascript:/file: 같은 스킴으로 된 악성 링크가 그대로 클릭 가능한 <a>가 되는 걸 막는다. */
const ALLOWED_HYPERLINK_SCHEMES = ['http:', 'https:', 'mailto:']

function sanitizeHyperlink(url: string): string | null {
  try {
    const scheme = new URL(url).protocol
    return ALLOWED_HYPERLINK_SCHEMES.includes(scheme) ? url : null
  } catch {
    return null
  }
}

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
  await workbook.xlsx.load(buffer)

  // 테마는 워크북 하나당 한 번만 파싱한다(셀마다 다시 파싱하지 않음) — ExcelJS는
  // 공식 지원은 안 하지만 원본 XML 문자열은 model.themes.theme1에 그대로 들어있다.
  // (ExcelJS의 타입 선언은 themes를 string[]로 잘못 적어뒀다 — 실제로는 테마 이름을
  // 키로 쓰는 객체라 unknown을 거쳐 캐스팅한다.)
  const themeXml = (workbook.model as unknown as { themes?: Record<string, string> }).themes?.theme1
  const theme = themeXml ? parseTheme(themeXml) : null

  const mdw = measureDefaultFontWidth()

  // 그림/도형은 ExcelJS가 일부만 읽고 나머지는 버려서 원본 XML을 직접 읽는다(src/xlsx/drawing.ts).
  // 이쪽이 실패해도 셀 데이터는 이미 읽혔으므로 파일 열기 자체는 막지 않고 경고로만 알린다.
  const warnings: string[] = []
  let drawingsBySheet = new Map<string, SheetDrawings>()
  // 자동 필터의 "조건"(어느 열이 필터링 중인지)도 ExcelJS가 버리는 정보라 같은 ZIP에서 읽는다.
  // 필터 범위 자체는 ExcelJS가 주므로, 필터가 있는 시트만 원본 XML을 연다(큰 시트 이중 해제 방지).
  const filterDetailsBySheet = new Map<string, AutoFilterDetails | null>()
  try {
    const zip = ZipArchive.open(buffer)
    const result = await readWorkbookDrawings(zip, theme)
    drawingsBySheet = result.bySheet
    warnings.push(...result.warnings)

    const sheetPaths = new Map((await listSheetParts(zip)).map((part) => [part.name, part.path]))
    for (const ws of workbook.worksheets) {
      const path = sheetPaths.get(ws.name)
      if (!ws.autoFilter || ws.state !== 'visible' || !path) continue
      try {
        filterDetailsBySheet.set(ws.name, await readAutoFilterDetails(zip, path))
      } catch (err) {
        warnings.push(`"${ws.name}" 시트의 필터 조건을 읽지 못했어요: ${err instanceof Error ? err.message : String(err)}`)
      }
    }
  } catch (err) {
    warnings.push(`그림/도형을 읽지 못했어요: ${err instanceof Error ? err.message : String(err)}`)
  }

  const sheets: SheetModel[] = []
  workbook.eachSheet((worksheet) => {
    // 숨김/완전숨김 시트는 Excel도 탭에 안 보여준다(Excel의 "숨기기" 의도 존중).
    if (worksheet.state !== 'visible') return

    const defaultRowPt = worksheet.properties?.defaultRowHeight ?? DEFAULT_ROW_HEIGHT_PT
    const defaultColChars = worksheet.properties?.defaultColWidth ?? DEFAULT_COLUMN_CHAR_WIDTH
    const defaultRowPx = excelPointsToPx(defaultRowPt)
    const defaultColPx = excelColumnWidthToPx(defaultColChars, mdw)

    const rowCount = worksheet.rowCount
    const colCount = worksheet.columnCount
    const rows: (CellModel | undefined)[][] = []
    const rowHeights: number[] = []
    const hiddenRows: boolean[] = []

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
        rowCells.push(cellToModel(cell, theme, note))
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

    // ExcelJS의 views 타입 선언(Array<Partial<WorksheetView>>)은 state로 좁혀도
    // ySplit이 안 보인다(frozen 전용 필드인데 Partial이 판별 유니온 좁히기를 못
    // 살림) — 필요한 필드만 최소 타입으로 캐스팅한다.
    const frozenView = worksheet.views?.find((v) => v.state === 'frozen') as { ySplit?: number } | undefined
    const frozen = frozenView?.ySplit ? { rows: frozenView.ySplit } : null

    const merges = worksheet.model.merges ?? []

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
      hiddenCols,
      hiddenRows,
      frozen,
      filters: sheetFilter ? [sheetFilter] : [],
      conditionalFormats: conditional.formats,
      skippedConditionalFormats: conditional.skipped,
      drawings: drawing?.items ?? [],
      skippedDrawings: drawing?.skipped ?? {},
    })
  })

  return { fileName: file.name, sheets, warnings }
}

function cellToModel(cell: import('exceljs').Cell, theme: ThemeColors | null, note: string | undefined): CellModel {
  const { value, formula, hyperlink } = interpretCellValue(cell.value)

  return {
    address: cell.address,
    value,
    formula,
    numFmt: cell.numFmt ?? null,
    style: extractCellStyle(cell, theme),
    hyperlink: hyperlink ? sanitizeHyperlink(hyperlink) : null,
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
