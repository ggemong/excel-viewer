import { describe, expect, it } from 'vitest'
import type { CellModel, SheetModel, WorkbookModel } from '../xlsx/types'
import { isLinkFormula, resolveLinkFormula, workbookEvalContext } from './linkTarget'

const cell = (value: string | number | null, formula: string | null = null): CellModel => ({ address: 'A1', value, formula, numFmt: null, style: null, hyperlink: null })

/** 테스트에 필요한 필드(이름, 칸, 크기)만 가진 시트. 나머지 필드는 이 계산이 읽지 않는다. */
function sheetOf(name: string, rows: (CellModel | undefined)[][]): SheetModel {
  return { name, rows, rowCount: rows.length, colCount: Math.max(0, ...rows.map((r) => r.length)) } as unknown as SheetModel
}
const workbookOf = (fileName: string, sheets: SheetModel[]): WorkbookModel => ({ fileName, sheets, warnings: [] })

/** 실제 업무 파일의 "돌아가기" 링크와 같은 모양: 지금 시트 이름으로 요청 시트의 F열에서 행을 찾아 그 A열로 이동한다. */
const BACK_LINK =
  `IFERROR(HYPERLINK("#'"&SUBSTITUTE(C1,"'","''")&"'!"&"A"&MATCH(MID(CELL("filename",$A$1),FIND("]",CELL("filename",$A$1))+1,255),INDIRECT("'"& $C$1 & "'!f:f"),0),"요청사항으로 이동"),"")`

function backLinkWorkbook() {
  const request = sheetOf('요청 목록', [
    [cell('순번'), undefined, undefined, undefined, undefined, cell('참조시트')],
    [cell(1), undefined, undefined, undefined, undefined, cell('참조 A')],
    [cell(2), undefined, undefined, undefined, undefined, cell('참조 B')],
    [cell(3), undefined, undefined, undefined, undefined, cell('참조 C')],
  ])
  const refB = sheetOf('참조 B', [[cell('요청사항으로 이동', BACK_LINK), undefined, cell('요청 목록')]])
  const refZ = sheetOf('참조 Z', [[cell('요청사항으로 이동', BACK_LINK), undefined, cell('요청 목록')]])
  return workbookOf('통합문서.xlsx', [request, refB, refZ])
}

describe('resolveLinkFormula: 계산으로 만들어지는 대상', () => {
  it('실제 파일의 돌아가기 링크: 지금 시트 이름을 요청 시트 F열에서 찾아 그 행의 A열로', () => {
    const wb = backLinkWorkbook()
    const sheet = wb.sheets[1]
    const result = resolveLinkFormula(BACK_LINK, workbookEvalContext(wb, sheet, 1, 1))
    expect(result).toEqual({ ok: true, target: { kind: 'internal', link: { sheet: '요청 목록', row: 3, col: 1 } } })
  })

  it('목록에 없는 시트에서는 틀린 곳으로 보내지 않고 이유를 돌려준다(#N/A)', () => {
    const wb = backLinkWorkbook()
    const result = resolveLinkFormula(BACK_LINK, workbookEvalContext(wb, wb.sheets[2], 1, 1))
    expect(result.ok).toBe(false)
    if (!result.ok) expect(result.reason).toContain('#N/A')
  })

  it('INDIRECT가 가리키는 시트가 없으면(#REF!) 실패', () => {
    const wb = backLinkWorkbook()
    wb.sheets[1].rows[0][2] = cell('없는 시트')
    const result = resolveLinkFormula(BACK_LINK, workbookEvalContext(wb, wb.sheets[1], 1, 1))
    expect(result.ok).toBe(false)
    if (!result.ok) expect(result.reason).toContain('#REF!')
  })

  it('시트 이름은 대소문자를 구분하지 않고 찾는다', () => {
    const wb = backLinkWorkbook()
    wb.sheets[1].rows[0][2] = cell('요청 목록'.toUpperCase())
    expect(resolveLinkFormula(BACK_LINK, workbookEvalContext(wb, wb.sheets[1], 1, 1)).ok).toBe(true)
  })

  it('단순한 글자 대상과 외부 주소, 허용되지 않은 주소', () => {
    const wb = workbookOf('a.xlsx', [sheetOf('S', [[cell(1)]])])
    const ctx = workbookEvalContext(wb, wb.sheets[0], 1, 1)
    expect(resolveLinkFormula(`HYPERLINK("#S!B2","x")`, ctx)).toEqual({ ok: true, target: { kind: 'internal', link: { sheet: 'S', row: 2, col: 2 } } })
    expect(resolveLinkFormula(`HYPERLINK("https://a.example","x")`, ctx)).toEqual({ ok: true, target: { kind: 'external', url: 'https://a.example' } })
    expect(resolveLinkFormula(`HYPERLINK("javascript:alert(1)","x")`, ctx).ok).toBe(false)
  })

  it('이동할 곳이 비었거나 칸 주소가 아니면 실패', () => {
    const wb = workbookOf('a.xlsx', [sheetOf('S', [[cell(null)]])])
    const ctx = workbookEvalContext(wb, wb.sheets[0], 1, 1)
    expect(resolveLinkFormula('HYPERLINK(A1,"x")', ctx).ok).toBe(false)
    expect(resolveLinkFormula(`HYPERLINK("#이름정의","x")`, ctx).ok).toBe(false)
  })

  it('HYPERLINK가 아닌 수식이나 해석할 수 없는 수식은 실패', () => {
    const wb = workbookOf('a.xlsx', [sheetOf('S', [[cell(1)]])])
    const ctx = workbookEvalContext(wb, wb.sheets[0], 1, 1)
    expect(resolveLinkFormula('SUM(1,2)', ctx).ok).toBe(false)
    expect(resolveLinkFormula('HYPERLINK(', ctx).ok).toBe(false)
  })
})

describe('isLinkFormula', () => {
  it('HYPERLINK로 시작하거나 IFERROR로 감싼 것만', () => {
    expect(isLinkFormula('HYPERLINK("#A1","x")')).toBe(true)
    expect(isLinkFormula(' iferror( hyperlink("#A1"),"")')).toBe(true)
    expect(isLinkFormula('IF(A1,HYPERLINK("#A1"),"")')).toBe(false)
    expect(isLinkFormula('SUM(A1)')).toBe(false)
  })
})
