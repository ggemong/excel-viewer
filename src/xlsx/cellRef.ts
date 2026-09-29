/** A1 표기 <-> 1-based {row, col} 변환. 읽기 경로(read.ts/grid)와 저장 경로(patch.ts)가 같이 쓴다. */

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

export function parseCellAddress(address: string): { row: number; col: number } {
  const match = address.match(/^([A-Z]+)(\d+)$/)
  if (!match) throw new Error(`셀 주소가 이상해요: ${address}`)
  const [, colLetters, rowDigits] = match
  let col = 0
  for (const ch of colLetters) {
    col = col * 26 + (ch.charCodeAt(0) - 64)
  }
  return { row: Number(rowDigits), col }
}
