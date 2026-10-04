import ExcelJS from 'exceljs'
import { describe, expect, it } from 'vitest'
import { ZipArchive, ZipUnsupportedError } from './zipReader'

async function makeXlsxBuffer(): Promise<ArrayBuffer> {
  const wb = new ExcelJS.Workbook()
  const ws = wb.addWorksheet('시트1')
  ws.getCell('A1').value = '안녕'
  const out = await wb.xlsx.writeBuffer()
  const u8 = new Uint8Array(out as ArrayBuffer)
  return u8.buffer.slice(u8.byteOffset, u8.byteOffset + u8.byteLength) as ArrayBuffer
}

describe('ZipArchive', () => {
  it('실제 xlsx에서 압축된 항목을 풀어 읽는다', async () => {
    const zip = ZipArchive.open(await makeXlsxBuffer())
    expect(zip.has('xl/workbook.xml')).toBe(true)
    const xml = await zip.readText('xl/workbook.xml')
    expect(xml).toContain('시트1')
  })

  it('없는 항목은 null', async () => {
    const zip = ZipArchive.open(await makeXlsxBuffer())
    expect(await zip.read('nope.xml')).toBeNull()
    expect(zip.has('nope.xml')).toBe(false)
  })

  it('ZIP이 아닌 데이터는 명시적 오류', () => {
    expect(() => ZipArchive.open(new ArrayBuffer(100))).toThrow(ZipUnsupportedError)
  })
})
