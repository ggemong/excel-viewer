/**
 * A1 표기 <-> 0-based {row, col} 변환. 저장 경로(patch.ts)의 공유 수식
 * 마스터 판별에도 이 모듈을 그대로 재사용한다.
 */

export function columnLetter(oneBasedCol: number): string {
  let s = ''
  let num = oneBasedCol
  while (num > 0) {
    const rem = (num - 1) % 26
    s = String.fromCharCode(65 + rem) + s
    num = Math.floor((num - 1) / 26)
  }
  return s
}

export function cellAddress(oneBasedRow: number, oneBasedCol: number): string {
  return `${columnLetter(oneBasedCol)}${oneBasedRow}`
}
