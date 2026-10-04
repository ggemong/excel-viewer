import { columnLetter } from './cellRef'
import { extractCellStyle } from './cellStyle'
import { interpretCellValue } from './cellValue'
import { excelColumnWidthToPx, excelPointsToPx, measureDefaultFontWidth } from './columnWidth'
import { parseTheme, type ThemeColors } from './themeColor'
import type { CellModel, SheetModel, WorkbookModel } from './types'

// 파일에 열너비/행높이가 명시 안 된(사용자가 한 번도 손 안 댄) 열/행에 쓰는 Excel
// 통상 기본값 — 실제로 <col>/<row> 항목 자체가 없는 경우가 흔해서(기본값인 열/행은
// Excel이 아예 안 적음) ExcelJS가 undefined를 돌려줄 때를 대비해 필요하다.
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

  const sheets: SheetModel[] = []
  workbook.eachSheet((worksheet) => {
    // 숨김/완전숨김 시트는 Excel도 탭에 안 보여준다(Excel의 "숨기기" 의도 존중).
    if (worksheet.state !== 'visible') return

    const rowCount = worksheet.rowCount
    const colCount = worksheet.columnCount
    const rows: (CellModel | undefined)[][] = []
    const rowHeights: number[] = []
    const hiddenRows: boolean[] = []

    for (let r = 1; r <= rowCount; r++) {
      const row = worksheet.getRow(r)
      rowHeights.push(excelPointsToPx(row.height ?? DEFAULT_ROW_HEIGHT_PT))
      hiddenRows.push(Boolean(row.hidden))

      const rowCells: (CellModel | undefined)[] = []
      for (let c = 1; c <= colCount; c++) {
        const cell = row.getCell(c)
        if (cell.type === undefined || cell.value === null || cell.value === undefined) {
          rowCells.push(undefined)
          continue
        }
        rowCells.push(cellToModel(cell, theme))
      }
      rows.push(rowCells)
    }

    const colWidths: number[] = []
    const hiddenCols: boolean[] = []
    for (let c = 1; c <= colCount; c++) {
      const col = worksheet.getColumn(c)
      colWidths.push(excelColumnWidthToPx(col.width ?? DEFAULT_COLUMN_CHAR_WIDTH, mdw))
      hiddenCols.push(Boolean(col.hidden))
    }

    // ExcelJS의 views 타입 선언(Array<Partial<WorksheetView>>)은 state로 좁혀도
    // ySplit이 안 보인다(frozen 전용 필드인데 Partial이 판별 유니온 좁히기를 못
    // 살림) — 필요한 필드만 최소 타입으로 캐스팅한다.
    const frozenView = worksheet.views?.find((v) => v.state === 'frozen') as { ySplit?: number } | undefined
    const frozen = frozenView?.ySplit ? { rows: frozenView.ySplit } : null

    const merges = worksheet.model.merges ?? []

    sheets.push({
      name: worksheet.name,
      rowCount,
      colCount,
      rows,
      merges,
      colWidths,
      rowHeights,
      hiddenCols,
      hiddenRows,
      frozen,
    })
  })

  return { fileName: file.name, sheets }
}

function cellToModel(cell: import('exceljs').Cell, theme: ThemeColors | null): CellModel {
  const { value, formula, hyperlink } = interpretCellValue(cell.value)

  return {
    address: cell.address,
    value,
    formula,
    numFmt: cell.numFmt ?? null,
    style: extractCellStyle(cell, theme),
    hyperlink: hyperlink ? sanitizeHyperlink(hyperlink) : null,
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
      },
    ],
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
