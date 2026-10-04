import { describe, expect, it } from 'vitest'
import { linkFromFormula, parseInternalTarget, sanitizeHyperlink } from './hyperlink'

describe('sanitizeHyperlink', () => {
  it('http/https/mailto만 허용한다', () => {
    expect(sanitizeHyperlink('https://a.example/x')).toBe('https://a.example/x')
    expect(sanitizeHyperlink('mailto:a@b.example')).toBe('mailto:a@b.example')
    expect(sanitizeHyperlink('javascript:alert(1)')).toBeNull()
    expect(sanitizeHyperlink('file:///c:/x')).toBeNull()
    expect(sanitizeHyperlink('not a url')).toBeNull()
  })
})

describe('parseInternalTarget', () => {
  it('따옴표 있는 시트 이름(공백·한글 포함)과 칸', () => {
    expect(parseInternalTarget("#'참조 2607-1'!A1")).toEqual({ sheet: '참조 2607-1', row: 1, col: 1 })
  })

  it('따옴표 없는 시트 이름, 절대 참조($), 소문자', () => {
    expect(parseInternalTarget('#Sheet2!$C$5')).toEqual({ sheet: 'Sheet2', row: 5, col: 3 })
    expect(parseInternalTarget('#sheet2!c5')).toEqual({ sheet: 'sheet2', row: 5, col: 3 })
  })

  it('시트 이름 안의 작은따옴표는 두 개로 쓴다', () => {
    expect(parseInternalTarget("#'Bob''s Sheet'!B2")).toEqual({ sheet: "Bob's Sheet", row: 2, col: 2 })
  })

  it('시트 없이 칸만 있으면 지금 시트(sheet null)', () => {
    expect(parseInternalTarget('#B7')).toEqual({ sheet: null, row: 7, col: 2 })
  })

  it('범위면 왼쪽 위 칸', () => {
    expect(parseInternalTarget("#'S'!B2:D9")).toEqual({ sheet: 'S', row: 2, col: 2 })
  })

  it('이름 정의나 이상한 대상은 null — 어디로 갈지 모른다', () => {
    expect(parseInternalTarget('#MyNamedRange')).toBeNull()
    expect(parseInternalTarget("#'S'!NamedThing")).toBeNull()
    expect(parseInternalTarget('#!A1')).toBeNull()
    expect(parseInternalTarget('https://a.example')).toBeNull()
    expect(parseInternalTarget('#A0')).toBeNull()
  })
})

describe('linkFromFormula', () => {
  it('실제 파일의 시트 이동 링크 수식', () => {
    expect(linkFromFormula(`HYPERLINK("#'참조 2607-1'!A1","참조 2607-1")`)).toEqual({ kind: 'internal', link: { sheet: '참조 2607-1', row: 1, col: 1 } })
  })

  it('친근한 이름 인자가 없어도, 소문자·공백이 있어도 읽는다', () => {
    expect(linkFromFormula(`hyperlink( "#Sheet2!A1" )`)).toEqual({ kind: 'internal', link: { sheet: 'Sheet2', row: 1, col: 1 } })
  })

  it('외부 주소는 허용된 스킴만', () => {
    expect(linkFromFormula(`HYPERLINK("https://a.example/p?q=1","열기")`)).toEqual({ kind: 'external', url: 'https://a.example/p?q=1' })
    expect(linkFromFormula(`HYPERLINK("javascript:alert(1)","x")`)).toBeNull()
  })

  it('따옴표 두 개는 따옴표 하나로 푼다', () => {
    expect(linkFromFormula('HYPERLINK("https://a.example/?q=""x""","t")')).toEqual({ kind: 'external', url: 'https://a.example/?q="x"' })
  })

  it('대상이 셀 값으로 만들어지는 수식이나 다른 수식은 링크로 만들지 않는다', () => {
    expect(linkFromFormula(`HYPERLINK("#'"&A1&"'!A1","x")`)).toBeNull()
    expect(linkFromFormula(`HYPERLINK(B2,"x")`)).toBeNull()
    expect(linkFromFormula('SUM(A1:A3)')).toBeNull()
    expect(linkFromFormula(`IFERROR(HYPERLINK("#S!A1","x"),"")`)).toBeNull()
  })
})
