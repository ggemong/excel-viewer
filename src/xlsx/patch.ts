import { parseCellAddress } from './cellRef'
import { compressEntry, decompressEntry } from './deflate'
import { findMergeAt, isMergeMaster, type MergeRange } from './mergeRange'
import { buildZip, parseZip, type ZipEntry } from './zip'

const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
const textDecoder = new TextDecoder()
const textEncoder = new TextEncoder()

export interface CellEdit {
  sheetName: string
  address: string // "B4"
  /** null = 셀을 비운다(삭제). 그 외에는 텍스트/숫자 리터럴로 저장 — 절대 수식으로 해석하지 않는다. */
  newValue: string | number | null
}

/**
 * 서식·차트·다른 셀은 전혀 건드리지 않고, 지정한 셀들의 값만 원본 zip에 직접
 * patch해서 새 xlsx Blob을 만든다. ExcelJS의 저장 기능은 쓰지 않는다 — 구현
 * 계획의 "읽기/저장 경로 분리" 참고.
 */
export async function patchWorkbook(file: File, edits: CellEdit[]): Promise<Blob> {
  if (edits.length === 0) {
    throw new Error('저장할 변경사항이 없어요.')
  }

  const buffer = await file.arrayBuffer()
  const entries = parseZip(buffer)
  const entryByName = new Map(entries.map((e) => [e.name, e]))

  const sheetPaths = resolveSheetPaths(entryByName)

  const editsBySheetPath = new Map<string, CellEdit[]>()
  for (const edit of edits) {
    const path = sheetPaths.get(edit.sheetName)
    if (!path) {
      throw new Error(`"${edit.sheetName}" 시트를 파일 안에서 못 찾았어요.`)
    }
    const list = editsBySheetPath.get(path) ?? []
    list.push(edit)
    editsBySheetPath.set(path, list)
  }

  const patchedEntries = entries.map((entry) => {
    const sheetEdits = editsBySheetPath.get(entry.name)
    if (!sheetEdits) return entry
    return patchSheetEntry(entry, sheetEdits)
  })

  const zipBytes = buildZip(patchedEntries)
  return new Blob([zipBytes as BlobPart], { type: XLSX_MIME })
}

function resolveSheetPaths(entryByName: Map<string, ZipEntry>): Map<string, string> {
  const workbookEntry = entryByName.get('xl/workbook.xml')
  const relsEntry = entryByName.get('xl/_rels/workbook.xml.rels')
  if (!workbookEntry || !relsEntry) {
    throw new Error('xlsx 구조가 예상과 달라요(workbook.xml 또는 rels를 못 찾음).')
  }

  const workbookXml = parseXml(decodeEntry(workbookEntry))
  const relsXml = parseXml(decodeEntry(relsEntry))

  const idToTarget = new Map<string, string>()
  for (const rel of Array.from(relsXml.documentElement.children)) {
    const id = rel.getAttribute('Id')
    const target = rel.getAttribute('Target')
    if (id && target) idToTarget.set(id, target)
  }

  const nameToPath = new Map<string, string>()
  const sheetsEl = workbookXml.getElementsByTagName('sheets')[0]
  if (!sheetsEl) throw new Error('workbook.xml에 <sheets>가 없어요.')

  for (const sheetEl of Array.from(sheetsEl.children)) {
    const name = sheetEl.getAttribute('name')
    const rId = sheetEl.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id') || sheetEl.getAttribute('r:id')
    if (!name || !rId) continue
    const target = idToTarget.get(rId)
    if (!target) continue
    const path = target.startsWith('/') ? target.slice(1) : `xl/${target}`
    nameToPath.set(name, path)
  }

  return nameToPath
}

function patchSheetEntry(entry: ZipEntry, edits: CellEdit[]): ZipEntry {
  const originalText = decodeEntry(entry)
  const declMatch = originalText.match(/^<\?xml[^?]*\?>\s*/)
  const decl = declMatch ? declMatch[0] : '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'

  const doc = parseXml(originalText)
  const sheetData = doc.getElementsByTagName('sheetData')[0]
  if (!sheetData) {
    throw new Error('시트 XML에 <sheetData>가 없어요.')
  }
  const mergeRanges = parseMergeRangesFromDoc(doc)

  for (const edit of edits) {
    applyEdit(doc, sheetData, edit, mergeRanges)
  }

  const serialized = decl + new XMLSerializer().serializeToString(doc.documentElement)
  const newRaw = textEncoder.encode(serialized)
  const { data, crc32, compressedSize, uncompressedSize } = compressEntry(newRaw, entry.compressionMethod)

  return { ...entry, data, crc32, compressedSize, uncompressedSize }
}

function parseMergeRangesFromDoc(doc: Document): MergeRange[] {
  const ranges: MergeRange[] = []
  for (const mergeCell of Array.from(doc.getElementsByTagName('mergeCell'))) {
    const ref = mergeCell.getAttribute('ref')
    if (!ref) continue
    const [a, b] = ref.split(':')
    const pa = parseCellAddress(a)
    const pb = parseCellAddress(b ?? a)
    ranges.push({
      r0: Math.min(pa.row, pb.row),
      c0: Math.min(pa.col, pb.col),
      r1: Math.max(pa.row, pb.row),
      c1: Math.max(pa.col, pb.col),
    })
  }
  return ranges
}

function applyEdit(doc: Document, sheetData: Element, edit: CellEdit, mergeRanges: MergeRange[]) {
  const { row: rowNum, col: colNum } = parseCellAddress(edit.address)

  const merge = findMergeAt(mergeRanges, rowNum, colNum)
  if (merge && !isMergeMaster(merge, rowNum, colNum)) {
    throw new Error(`${edit.address} 셀은 병합된 셀이에요 — 왼쪽 위 기준 셀에서 수정하세요.`)
  }

  const row = findOrInsertRow(doc, sheetData, rowNum)

  const existingCell = findCell(row, edit.address)

  if (edit.newValue === null) {
    if (existingCell) row.removeChild(existingCell)
    return
  }

  const cell = existingCell ?? insertCell(doc, row, edit.address, colNum)

  if (cell.getElementsByTagName('f').length > 0) {
    throw new Error(`${edit.address} 셀은 수식이 있어서 수정할 수 없어요.`)
  }

  while (cell.firstChild) cell.removeChild(cell.firstChild)

  if (isPlainNumber(edit.newValue)) {
    cell.removeAttribute('t')
    const v = doc.createElement('v')
    v.textContent = String(edit.newValue)
    cell.appendChild(v)
  } else {
    cell.setAttribute('t', 'inlineStr')
    const is = doc.createElement('is')
    const t = doc.createElement('t')
    t.textContent = String(edit.newValue)
    is.appendChild(t)
    cell.appendChild(is)
  }
}

function isPlainNumber(value: string | number): boolean {
  if (typeof value === 'number') return Number.isFinite(value)
  return /^-?\d+(\.\d+)?$/.test(value.trim())
}

function findOrInsertRow(doc: Document, sheetData: Element, rowNum: number): Element {
  const rows = Array.from(sheetData.children)
  for (const row of rows) {
    const r = Number(row.getAttribute('r'))
    if (r === rowNum) return row
    if (r > rowNum) {
      const newRow = doc.createElement('row')
      newRow.setAttribute('r', String(rowNum))
      sheetData.insertBefore(newRow, row)
      return newRow
    }
  }
  const newRow = doc.createElement('row')
  newRow.setAttribute('r', String(rowNum))
  sheetData.appendChild(newRow)
  return newRow
}

function findCell(row: Element, address: string): Element | null {
  for (const cell of Array.from(row.children)) {
    if (cell.getAttribute('r') === address) return cell
  }
  return null
}

function insertCell(doc: Document, row: Element, address: string, colNum: number): Element {
  const newCell = doc.createElement('c')
  newCell.setAttribute('r', address)

  for (const cell of Array.from(row.children)) {
    const existingCol = parseCellAddress(cell.getAttribute('r') ?? 'A1').col
    if (existingCol > colNum) {
      row.insertBefore(newCell, cell)
      return newCell
    }
  }
  row.appendChild(newCell)
  return newCell
}

function decodeEntry(entry: ZipEntry): string {
  return textDecoder.decode(decompressEntry(entry.data, entry.compressionMethod))
}

function parseXml(text: string): Document {
  const doc = new DOMParser().parseFromString(text, 'application/xml')
  const parseError = doc.getElementsByTagName('parsererror')[0]
  if (parseError) {
    throw new Error(`xlsx 내부 XML을 파싱하지 못했어요: ${parseError.textContent}`)
  }
  return doc
}
