import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import ExcelJS from 'exceljs'
import { describe, expect, it } from 'vitest'
import { patchWorkbook } from './patch'
import { parseZip } from './zip'

const fixturePath = resolve(process.cwd(), 'src/xlsx/__fixtures__/sample.xlsx')

function loadFixtureFile(): File {
  const buffer = readFileSync(fixturePath)
  return new File([buffer], 'sample.xlsx', {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  })
}

async function readBack(blob: Blob) {
  const buffer = await blob.arrayBuffer()
  const wb = new ExcelJS.Workbook()
  await wb.xlsx.load(buffer)
  return wb
}

describe('patchWorkbook', () => {
  it('공유 문자열 셀 값을 수정한다', async () => {
    const file = loadFixtureFile()
    const blob = await patchWorkbook(file, [{ sheetName: 'Data', address: 'A2', newValue: 'Alpha!!' }])
    const wb = await readBack(blob)
    expect(wb.getWorksheet('Data')!.getCell('A2').value).toBe('Alpha!!')
  })

  it('숫자 셀 값을 수정한다', async () => {
    const file = loadFixtureFile()
    const blob = await patchWorkbook(file, [{ sheetName: 'Data', address: 'B2', newValue: 250 }])
    const wb = await readBack(blob)
    expect(wb.getWorksheet('Data')!.getCell('B2').value).toBe(250)
  })

  it('빈 셀에 값을 추가한다 (기존 행 안의 새 셀)', async () => {
    const file = loadFixtureFile()
    const blob = await patchWorkbook(file, [{ sheetName: 'Data', address: 'C1', newValue: 'New' }])
    const wb = await readBack(blob)
    expect(wb.getWorksheet('Data')!.getCell('C1').value).toBe('New')
  })

  it('빈 셀에 값을 추가한다 (새 행 자체가 필요한 경우)', async () => {
    const file = loadFixtureFile()
    const blob = await patchWorkbook(file, [{ sheetName: 'Data', address: 'A10', newValue: 'Far row' }])
    const wb = await readBack(blob)
    expect(wb.getWorksheet('Data')!.getCell('A10').value).toBe('Far row')
  })

  it('값이 있는 셀을 삭제하면 빈 셀이 된다', async () => {
    const file = loadFixtureFile()
    const blob = await patchWorkbook(file, [{ sheetName: 'Data', address: 'A4', newValue: null }])
    const wb = await readBack(blob)
    const cell = wb.getWorksheet('Data')!.getCell('A4')
    expect(cell.value == null).toBe(true)
  })

  it('수식이 있는 셀은 수정을 거부한다', async () => {
    const file = loadFixtureFile()
    await expect(patchWorkbook(file, [{ sheetName: 'Data', address: 'B3', newValue: 999 }])).rejects.toThrow()
  })

  it('수정한 셀이 든 시트 XML 말고는 모든 zip 항목이 바이트 단위로 원본과 동일하다', async () => {
    const file = loadFixtureFile()
    const originalBuffer = await file.arrayBuffer()
    const originalEntries = parseZip(originalBuffer)

    const blob = await patchWorkbook(file, [{ sheetName: 'Data', address: 'A2', newValue: 'Alpha!!' }])
    const patchedBuffer = await blob.arrayBuffer()
    const patchedEntries = parseZip(patchedBuffer)

    expect(patchedEntries.length).toBe(originalEntries.length)

    const originalByName = new Map(originalEntries.map((e) => [e.name, e]))
    let changedCount = 0

    for (const patched of patchedEntries) {
      const original = originalByName.get(patched.name)
      expect(original).toBeDefined()
      const same =
        original!.data.length === patched.data.length &&
        original!.data.every((byte, i) => byte === patched.data[i])
      if (!same) changedCount++
    }

    // 수정 대상 시트 XML 하나만 바뀌어야 한다.
    expect(changedCount).toBe(1)
  })

  it('무관한 서식(굵게+배경색, 병합, 열 너비)은 그대로 보존된다', async () => {
    const file = loadFixtureFile()
    const blob = await patchWorkbook(file, [{ sheetName: 'Data', address: 'A2', newValue: 'Alpha!!' }])
    const wb = await readBack(blob)
    const ws = wb.getWorksheet('Data')!

    expect(ws.getCell('A1').font?.bold).toBe(true)
    expect(ws.getCell('A1').fill).toMatchObject({ fgColor: { argb: 'FFEFEFEF' } })
    expect(ws.getColumn(1).width).toBe(18)
    expect(ws.getColumn(2).width).toBe(12)

    expect(ws.model.merges).toContain('A6:B6')

    // 수식 캐시값도 안 건드렸는지 확인
    const b3 = ws.getCell('B3')
    expect(b3.formula).toBe('B2*2')
  })
})
