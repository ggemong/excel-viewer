import { describe, expect, it } from 'vitest'
import { extractConditionalFormats, formulaSupport, parseSqref, type RawConditionalFormatting } from './conditionalFormat'
import type { ThemeColors } from './themeColor'

const theme: ThemeColors = { slots: ['#FFFFFF', '#000000', '#EEECE1', '#1F497D', '#4F81BD', '#C0504D', '#9BBB59', '#8064A2', '', '', '', ''] }

describe('parseSqref', () => {
  it('공백으로 구분된 여러 범위와 단일 칸, 엑셀 최대 행까지의 범위', () => {
    expect(parseSqref('A1:N3 H4:N5 F4 A7:N1048576')).toEqual([
      { r0: 1, c0: 1, r1: 3, c1: 14 },
      { r0: 4, c0: 8, r1: 5, c1: 14 },
      { r0: 4, c0: 6, r1: 4, c1: 6 },
      { r0: 7, c0: 1, r1: 1048576, c1: 14 },
    ])
  })
  it('해석 못 하는 조각은 건너뛰고 나머지는 살린다', () => {
    expect(parseSqref('A1:B2 ???')).toEqual([{ r0: 1, c0: 1, r1: 2, c1: 2 }])
  })
})

describe('formulaSupport', () => {
  it('지원 함수만 쓰면 ok, 모르는 함수/다른 시트 참조/문법 오류는 사유와 함께 거절', () => {
    expect(formulaSupport('AND($L1="",$J1<TODAY())').ok).toBe(true)
    expect(formulaSupport('VLOOKUP(A1,B:C,2,0)>1')).toMatchObject({ ok: false })
    expect(formulaSupport('Sheet2!A1=1')).toMatchObject({ ok: false, reason: '다른 시트 참조' })
    expect(formulaSupport('1+')).toMatchObject({ ok: false })
  })
})

describe('extractConditionalFormats — 실제 파일에서 ExcelJS가 주는 모양', () => {
  const raw: RawConditionalFormatting[] = [
    {
      ref: 'A1:N3 H4:N5',
      rules: [
        { type: 'expression', priority: 25, formulae: ['AND($L1="",$J1<TODAY(),$J1<>"")'], style: { fill: null, font: { color: { argb: 'FFFF0000' } } } },
        { type: 'expression', priority: 26, formulae: ['OR($L1="불가",$L1="보류")'], style: { fill: { bgColor: { theme: 7, tint: 0.8 } }, font: { strike: true } } },
      ],
    },
  ]

  it('채우기(bgColor의 테마색+tint)·글자색·취소선을 CSS 색/불리언으로 바꾸고 우선순위를 보존', () => {
    const { formats, skipped } = extractConditionalFormats(raw, theme)
    expect(skipped).toEqual({})
    expect(formats).toHaveLength(1)
    expect(formats[0].ranges).toHaveLength(2)
    const [overdue, hold] = formats[0].rules
    expect(overdue).toMatchObject({ priority: 25, kind: 'expression', style: { color: '#FF0000' } })
    expect(hold.priority).toBe(26)
    expect(hold.style.strike).toBe(true)
    expect(hold.style.bg).toMatch(/^#[0-9A-F]{6}$/)
    expect(hold.style.bg).not.toBe('#8064A2')
  })

  it('지원하지 않는 규칙 종류는 조용히 버리지 않고 종류별로 센다', () => {
    const { formats, skipped } = extractConditionalFormats(
      [
        { ref: 'A1:A9', rules: [{ type: 'colorScale', priority: 1 }, { type: 'dataBar', priority: 2 }, { type: 'iconSet', priority: 3 }, { type: 'top10', priority: 4 }] },
        { ref: 'B1:B9', rules: [{ type: 'duplicateValues', priority: 5, style: { fill: { bgColor: { argb: 'FFFFC7CE' } } } }] },
      ],
      theme,
    )
    expect(formats).toEqual([])
    expect(skipped).toEqual({ colorScale: 1, dataBar: 1, iconSet: 1, top10: 1, duplicateValues: 1 })
  })

  it('지원하지 않는 수식(모르는 함수)을 쓰는 규칙은 건너뛰고 unsupportedFormula로 센다', () => {
    const { formats, skipped } = extractConditionalFormats(
      [{ ref: 'A1', rules: [{ type: 'expression', priority: 1, formulae: ['VLOOKUP(A1,B:C,2,0)>1'], style: { fill: { bgColor: { argb: 'FFFF0000' } } } }] }],
      theme,
    )
    expect(formats).toEqual([])
    expect(skipped).toEqual({ unsupportedFormula: 1 })
  })

  it('수식을 같이 주는 containsText 계열은 expression처럼, cellIs는 연산자와 함께 읽는다', () => {
    const { formats } = extractConditionalFormats(
      [
        {
          ref: 'A1:A5',
          rules: [
            { type: 'containsText', priority: 1, operator: 'containsText', formulae: ['NOT(ISERROR(SEARCH("수정",A1)))'], style: { font: { bold: true } } },
            { type: 'cellIs', priority: 2, operator: 'greaterThan', formulae: ['100'], style: { fill: { bgColor: { argb: 'FFFFFF00' } } } },
          ],
        },
      ],
      theme,
    )
    expect(formats[0].rules.map((r) => [r.kind, r.operator])).toEqual([
      ['expression', 'containsText'],
      ['cellIs', 'greaterThan'],
    ])
  })

  it('테두리·표시 형식만 바꾸는 규칙은 "못 보여준 것"으로 알린다', () => {
    const { formats, skipped } = extractConditionalFormats(
      [{ ref: 'A1', rules: [{ type: 'expression', priority: 1, formulae: ['A1>1'], style: { border: { top: { style: 'thin' } } } }] }],
      theme,
    )
    expect(formats).toEqual([])
    expect(skipped).toEqual({ borderOrNumberFormat: 1 })
  })
})
