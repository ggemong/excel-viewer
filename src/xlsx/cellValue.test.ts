import { describe, expect, it } from 'vitest'
import { dateToSerial, interpretCellValue, UNSUPPORTED_VALUE } from './cellValue'

describe('interpretCellValue', () => {
  it('숫자/문자/불리언/빈값은 그대로', () => {
    expect(interpretCellValue(5).value).toBe(5)
    expect(interpretCellValue('가').value).toBe('가')
    expect(interpretCellValue(false).value).toBe(false)
    expect(interpretCellValue(null).value).toBeNull()
  })

  it('일반 수식은 formula와 캐시된 result를 같이 준다', () => {
    expect(interpretCellValue({ formula: 'A1+1', result: 3 })).toEqual({ value: 3, formula: 'A1+1', hyperlink: null })
  })

  it('공유 수식의 따라가는 칸(sharedFormula)도 result를 값으로 쓴다 — "[object Object]"가 나오던 버그', () => {
    const r = interpretCellValue({ sharedFormula: 'A4', result: 2 })
    expect(r.value).toBe(2)
    expect(r.formula).toBeNull()
  })

  it('계산값이 안 저장된 수식은 빈 값', () => {
    expect(interpretCellValue({ formula: 'A1+B1' }).value).toBeNull()
  })

  it('오류값은 리터럴이든 수식 결과든 오류 문자열 그대로', () => {
    expect(interpretCellValue({ error: '#N/A' }).value).toBe('#N/A')
    expect(interpretCellValue({ formula: '1/0', result: { error: '#DIV/0!' } }).value).toBe('#DIV/0!')
  })

  it('richText는 글자만 이어 붙인다', () => {
    expect(interpretCellValue({ richText: [{ text: '빨강' }, { text: ' 보통' }] }).value).toBe('빨강 보통')
  })

  it('하이퍼링크는 text와 URL을 분리하고, text가 richText여도 푼다', () => {
    expect(interpretCellValue({ text: '링크', hyperlink: 'https://a.kr' })).toEqual({
      value: '링크',
      formula: null,
      hyperlink: 'https://a.kr',
    })
    expect(interpretCellValue({ text: { richText: [{ text: 'a' }, { text: 'b' }] }, hyperlink: 'https://a.kr' }).value).toBe('ab')
  })

  it('Date는 시간대와 무관한 Excel 직렬번호로 바꾼다', () => {
    expect(interpretCellValue(new Date(Date.UTC(1900, 0, 1))).value).toBe(2) // 1900-01-01 = 직렬 2(1900 윤년 버그 기준 EPOCH=1899-12-30)
    expect(dateToSerial(new Date(Date.UTC(2026, 9, 4, 12)))).toBe(dateToSerial(new Date(Date.UTC(2026, 9, 4))) + 0.5)
  })

  it('알 수 없는 모양은 [object Object] 대신 눈에 띄는 표식으로', () => {
    expect(interpretCellValue({ something: 1 }).value).toBe(UNSUPPORTED_VALUE)
    expect(interpretCellValue({ formula: 'X', result: { weird: true } }).value).toBe(UNSUPPORTED_VALUE)
  })
})
