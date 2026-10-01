import { columnLetter } from './cellRef'
import { extractCellStyle } from './cellStyle'
import type { CellModel, SheetModel, WorkbookModel } from './types'

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
        rowCells.push(cellToModel(cell))
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

function cellToModel(cell: import('exceljs').Cell): CellModel {
  let formula: string | null = null
  let value: CellModel['value'] = null

  const raw = cell.value as unknown
  if (raw && typeof raw === 'object' && 'formula' in (raw as Record<string, unknown>)) {
    const f = raw as { formula?: string; result?: unknown }
    formula = f.formula ?? null
    value = normalizeValue(f.result)
  } else if (raw && typeof raw === 'object' && 'richText' in (raw as Record<string, unknown>)) {
    const rt = raw as { richText: { text: string }[] }
    value = rt.richText.map((t) => t.text).join('')
  } else {
    value = normalizeValue(raw)
  }

  return {
    address: cell.address,
    value,
    formula,
    numFmt: cell.numFmt ?? null,
    style: extractCellStyle(cell),
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
