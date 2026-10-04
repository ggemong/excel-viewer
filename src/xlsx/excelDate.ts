/**
 * Excel 날짜 직렬번호 <-> JS Date 변환의 단일 출처. 셀 값 해석(cellValue.ts), 표시 형식(numberFormat.ts),
 * 수식 평가(src/formula)가 모두 같은 기준을 써야 해서 상수를 한 곳에 둔다.
 *
 * 기준일은 1899-12-30(ExcelJS와 같은 기준). 시간대 시차를 피하려고 모든 계산은 UTC 기준의 "달력 날짜"로만
 * 다룬다 — 직렬번호는 시간대와 무관한 순수 숫자이기 때문이다.
 */

export const EXCEL_EPOCH_UTC_MS = Date.UTC(1899, 11, 30)
export const MS_PER_DAY = 86_400_000

export function dateToSerial(date: Date): number {
  return (date.getTime() - EXCEL_EPOCH_UTC_MS) / MS_PER_DAY
}

/** 직렬번호의 날짜 부분을 UTC Date(그날 00:00)로. 시간(소수) 부분은 버린다. */
export function serialToUtcDate(serial: number): Date {
  return new Date(EXCEL_EPOCH_UTC_MS + Math.floor(serial) * MS_PER_DAY)
}

/** 지금 이 순간(사용자의 로컬 시계 기준)을 직렬번호로: 정수부는 오늘 날짜, 소수부는 현재 시각. */
export function nowSerial(now: Date = new Date()): number {
  const localDate = Date.UTC(now.getFullYear(), now.getMonth(), now.getDate())
  const msIntoDay = now.getHours() * 3_600_000 + now.getMinutes() * 60_000 + now.getSeconds() * 1000 + now.getMilliseconds()
  return (localDate - EXCEL_EPOCH_UTC_MS) / MS_PER_DAY + msIntoDay / MS_PER_DAY
}

/** TODAY(): 사용자 로컬 기준 오늘의 직렬번호(정수). */
export function todaySerial(now: Date = new Date()): number {
  return Math.floor(nowSerial(now))
}
