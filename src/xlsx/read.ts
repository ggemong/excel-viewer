import { columnLetter } from './cellRef'
import { extractCellStyle } from './cellStyle'
import { parseTheme, type ThemeColors } from './themeColor'
import type { CellModel, SheetModel, WorkbookModel } from './types'

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

  const sheets: SheetModel[] = []
  workbook.eachSheet((worksheet) => {
    const rowCount = worksheet.rowCount
    const colCount = worksheet.columnCount
    const rows: (CellModel | undefined)[][] = []

    for (let r = 1; r <= rowCount; r++) {
      const row = worksheet.getRow(r)
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

    const merges = worksheet.model.merges ?? []

    sheets.push({
      name: worksheet.name,
      rowCount,
      colCount,
      rows,
      merges,
    })
  })

  return { fileName: file.name, sheets }
}

function cellToModel(cell: import('exceljs').Cell, theme: ThemeColors | null): CellModel {
  let formula: string | null = null
  let value: CellModel['value'] = null
  let hyperlink: string | null = null

  const raw = cell.value as unknown
  if (raw && typeof raw === 'object' && 'formula' in (raw as Record<string, unknown>)) {
    const f = raw as { formula?: string; result?: unknown }
    formula = f.formula ?? null
    value = normalizeValue(f.result)
  } else if (raw && typeof raw === 'object' && 'richText' in (raw as Record<string, unknown>)) {
    const rt = raw as { richText: { text: string }[] }
    value = rt.richText.map((t) => t.text).join('')
  } else if (raw && typeof raw === 'object' && 'hyperlink' in (raw as Record<string, unknown>)) {
    // 외부(관계 ID 기반) 링크만 이 모양으로 온다 — 같은 통합문서 안 다른 시트/셀로
    // 가는 내부 링크는 ExcelJS가 애초에 이 모양으로 안 올려준다(알려진 한계).
    const h = raw as { text?: unknown; hyperlink?: string }
    value = normalizeValue(h.text)
    hyperlink = h.hyperlink ? sanitizeHyperlink(h.hyperlink) : null
  } else {
    value = normalizeValue(raw)
  }

  return {
    address: cell.address,
    value,
    formula,
    numFmt: cell.numFmt ?? null,
    style: extractCellStyle(cell, theme),
    hyperlink,
  }
}

function normalizeValue(raw: unknown): CellModel['value'] {
  if (raw === null || raw === undefined) return null
  if (raw instanceof Date) return raw.toISOString()
  if (typeof raw === 'string' || typeof raw === 'number' || typeof raw === 'boolean') return raw
  return String(raw)
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

  return {
    fileName: file.name,
    sheets: [
      {
        name: 'Sheet1',
        rowCount: rows.length,
        colCount: rows.reduce((max, row) => Math.max(max, row.length), 0),
        rows,
        merges: [],
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
