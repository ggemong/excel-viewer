import { describe, expect, it } from 'vitest'
import { formatNumber, formatText, generalNumber, isDateTimeFormat } from './numberFormat'

/** 시간대와 무관하게 Excel 직렬번호를 만든다(기준일 1899-12-30, cellValue.ts의 dateToSerial과 동일 기준). */
function serial(y: number, mo: number, d: number, h = 0, mi = 0, s = 0): number {
  return (Date.UTC(y, mo - 1, d, h, mi, s) - Date.UTC(1899, 11, 30)) / 86_400_000
}

describe('General', () => {
  it('부동소수 잡음을 걷는다', () => {
    expect(generalNumber(0.1 + 0.2)).toBe('0.3')
    expect(generalNumber(1 / 3)).toBe('0.3333333333')
  })
  it('큰 수는 Excel처럼 지수 표기, 11자리까지는 그대로', () => {
    expect(generalNumber(12345678901)).toBe('12345678901')
    expect(generalNumber(123456789012)).toBe('1.23457E+11')
  })
  it('아주 작은 수도 지수 표기 경계 안에서는 소수로', () => {
    expect(generalNumber(0.00000015)).toBe('0.00000015')
    expect(generalNumber(0)).toBe('0')
    expect(generalNumber(-2.5)).toBe('-2.5')
  })
  it('코드가 없거나 General이면 General로', () => {
    expect(formatNumber(0.1 + 0.2, null)).toBe('0.3')
    expect(formatNumber(5, 'General')).toBe('5')
  })
})

describe('숫자 서식', () => {
  it('접미사 리터럴은 숫자 뒤에 붙는다 — 한글 파일에서 흔한 단위 표기', () => {
    expect(formatNumber(1234567, '#,##0"원"')).toBe('1,234,567원')
    expect(formatNumber(5, '0"개"')).toBe('5개')
    expect(formatNumber(3, '0" 명"')).toBe('3 명')
  })
  it('접두 통화 기호', () => {
    expect(formatNumber(1500000, '₩#,##0')).toBe('₩1,500,000')
    expect(formatNumber(1500000, '[$₩-412]#,##0')).toBe('₩1,500,000')
    expect(formatNumber(1234.5, '$#,##0.00')).toBe('$1,234.50')
  })
  it('0 자리표시는 앞자리를 0으로 채운다', () => {
    expect(formatNumber(7, '0000')).toBe('0007')
    expect(formatNumber(12, '00000')).toBe('00012')
  })
  it('리터럴이 섞인 자리표시(전화번호 등)는 오른쪽부터 채운다', () => {
    expect(formatNumber(1012345678, '000-0000-0000')).toBe('010-1234-5678')
    expect(formatNumber(123456, '###-###')).toBe('123-456')
  })
  it('자릿수보다 큰 수는 왼쪽으로 넘쳐서 전부 보인다', () => {
    expect(formatNumber(123456, '00')).toBe('123456')
  })
  it('소수 자리: 0은 항상, #은 필요할 때만', () => {
    expect(formatNumber(1234567.891, '#,##0.00')).toBe('1,234,567.89')
    expect(formatNumber(2.5, '0.##')).toBe('2.5')
    expect(formatNumber(2, '0.0#')).toBe('2.0')
  })
  it('반올림은 보이는 십진수 기준(2.675 -> 2.68)', () => {
    expect(formatNumber(2.675, '0.00')).toBe('2.68')
    expect(formatNumber(0.5, '0')).toBe('1')
  })
  it('퍼센트는 100을 곱한다', () => {
    expect(formatNumber(0.256, '0.0%')).toBe('25.6%')
    expect(formatNumber(1.5, '0%')).toBe('150%')
  })
  it('지수 서식', () => {
    expect(formatNumber(12345.678, '0.00E+00')).toBe('1.23E+04')
    expect(formatNumber(0.000123, '0.00E+00')).toBe('1.23E-04')
  })
  it('마지막 자리표시 뒤 쉼표는 천 단위로 나눈다', () => {
    expect(formatNumber(1234567, '#,##0,')).toBe('1,235')
  })
  it('_x는 공백, 따옴표 없는 \\x는 리터럴', () => {
    expect(formatNumber(3.14159, '0.00_)')).toBe('3.14 ')
    expect(formatNumber(5, '0\\m')).toBe('5m')
  })
  it('#만 있는 형식에서 0은 비어 있다', () => {
    expect(formatNumber(0, '#')).toBe('')
    expect(formatNumber(0.5, '#.00')).toBe('.50')
  })
})

describe('구간(양수;음수;0;텍스트)', () => {
  it('음수 구간이 있으면 부호 없이 그 구간 모양대로 — 괄호 음수', () => {
    expect(formatNumber(-1234, '#,##0;[Red](#,##0)')).toBe('(1,234)')
    expect(formatNumber(1234, '#,##0;[Red](#,##0)')).toBe('1,234')
  })
  it('구간이 하나뿐이면 음수는 앞에 -', () => {
    expect(formatNumber(-1.5, '0.00')).toBe('-1.50')
    expect(formatNumber(-1500, '₩#,##0')).toBe('-₩1,500')
  })
  it('세 번째 구간은 0일 때, 비어 있으면 0을 숨긴다', () => {
    expect(formatNumber(0, '0;-0;;@')).toBe('')
    expect(formatNumber(0, '0;-0;"-"')).toBe('-')
  })
  it('색상/조건 대괄호는 무시한다', () => {
    expect(formatNumber(5, '[Blue]0')).toBe('5')
  })
})

describe('날짜/시간 서식', () => {
  const oct4 = serial(2026, 10, 4) // 일요일

  it('기본 날짜', () => {
    expect(formatNumber(oct4, 'yyyy-mm-dd')).toBe('2026-10-04')
    expect(formatNumber(oct4, 'yy/m/d')).toBe('26/10/4')
  })
  it('내장 로케일 의존 코드(mm-dd-yy)는 한국식 날짜로', () => {
    expect(formatNumber(oct4, 'mm-dd-yy')).toBe('2026-10-04')
  })
  it('한글 리터럴이 섞인 날짜', () => {
    expect(formatNumber(oct4, 'm"월" d"일"')).toBe('10월 4일')
    expect(formatNumber(oct4, 'yyyy"년" m"월" d"일"')).toBe('2026년 10월 4일')
  })
  it('요일(한글 aaa/aaaa, 기본 로케일 ddd/dddd)', () => {
    expect(formatNumber(oct4, 'yyyy-mm-dd (aaa)')).toBe('2026-10-04 (일)')
    expect(formatNumber(oct4, 'aaaa')).toBe('일요일')
    expect(formatNumber(oct4, 'ddd')).toBe('일')
  })
  it('영문 로케일 표기가 있으면 영어 이름', () => {
    expect(formatNumber(oct4, '[$-409]dddd')).toBe('Sunday')
    expect(formatNumber(oct4, '[$-409]mmm d')).toBe('Oct 4')
    expect(formatNumber(oct4, '[$-409]mmmm')).toBe('October')
  })
  it('시간대와 무관하게 입력한 시각 그대로 — 12:00이 21:00으로 밀리던 버그', () => {
    expect(formatNumber(serial(2026, 10, 4, 12), 'yyyy-mm-dd hh:mm')).toBe('2026-10-04 12:00')
    expect(formatNumber(0.5, 'h:mm')).toBe('12:00')
  })
  it('m은 시 뒤/초 앞이면 분, 아니면 월', () => {
    expect(formatNumber(serial(2026, 3, 5, 9, 7, 3), 'h:mm:ss')).toBe('9:07:03')
    expect(formatNumber(serial(2026, 3, 5, 9, 7, 3), 'mm/dd')).toBe('03/05')
    expect(formatNumber(serial(2026, 3, 5, 9, 7, 3), 'm:ss')).toBe('7:03')
  })
  it('오전/오후 표기', () => {
    expect(formatNumber(0.5, 'h:mm AM/PM')).toBe('12:00 오후')
    expect(formatNumber(serial(2026, 10, 4, 0, 5), 'h:mm AM/PM')).toBe('12:05 오전')
    expect(formatNumber(serial(2026, 10, 4, 15, 30), '[$-409]h:mm AM/PM')).toBe('3:30 PM')
  })
  it('경과 시간 [h]:mm', () => {
    expect(formatNumber(1.5, '[h]:mm')).toBe('36:00')
    expect(formatNumber(2 / 24 + 5 / 1440, '[mm]')).toBe('125')
  })
  it('초 반올림 잡음 제거(11:59:59.9999 -> 12:00:00)', () => {
    expect(formatNumber(0.49999999, 'hh:mm:ss')).toBe('12:00:00')
  })
  it('날짜 서식에서 음수 직렬번호는 ########', () => {
    expect(formatNumber(-1, 'yyyy-mm-dd')).toBe('########')
  })
})

describe('텍스트 구간', () => {
  it('@ 구간이 있으면 문자열에 적용', () => {
    expect(formatText('abc', '@')).toBe('abc')
    expect(formatText('abc', '0;-0;;"["@"]"')).toBe('[abc]')
  })
  it('@가 없는 서식이면 원문 그대로', () => {
    expect(formatText('abc', '#,##0')).toBe('abc')
    expect(formatText('abc', null)).toBe('abc')
  })
})

describe('isDateTimeFormat', () => {
  it('날짜/시간 서식만 true', () => {
    expect(isDateTimeFormat('yyyy-mm-dd')).toBe(true)
    expect(isDateTimeFormat('h:mm AM/PM')).toBe(true)
    expect(isDateTimeFormat('[h]:mm')).toBe(true)
    expect(isDateTimeFormat('mm-dd-yy')).toBe(true)
    expect(isDateTimeFormat('#,##0"원"')).toBe(false)
    expect(isDateTimeFormat('0.0%')).toBe(false)
    expect(isDateTimeFormat('General')).toBe(false)
    expect(isDateTimeFormat(null)).toBe(false)
  })
})
