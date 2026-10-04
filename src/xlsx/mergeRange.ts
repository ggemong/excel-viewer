import { parseCellAddress } from './cellRef'

export interface MergeRange {
  r0: number
  c0: number
  r1: number
  c1: number
}

export function parseMergeRanges(merges: string[]): MergeRange[] {
  return merges.filter(Boolean).map((range) => {
    const [a, b] = range.split(':')
    const pa = parseCellAddress(a)
    const pb = parseCellAddress(b ?? a)
    return {
      r0: Math.min(pa.row, pb.row),
      c0: Math.min(pa.col, pb.col),
      r1: Math.max(pa.row, pb.row),
      c1: Math.max(pa.col, pb.col),
    }
  })
}

export function isMergeMaster(range: MergeRange, row: number, col: number): boolean {
  return range.r0 === row && range.c0 === col
}

export function findMergeAt(ranges: MergeRange[], row: number, col: number): MergeRange | null {
  for (const range of ranges) {
    if (row >= range.r0 && row <= range.r1 && col >= range.c0 && col <= range.c1) return range
  }
  return null
}

/**
 * (row, col) -> 그 칸이 속한 병합 범위("행,열" 키). 병합 범위 안의 모든 칸(마스터 포함)이 들어 있다.
 * 화면(useMergeLookup)과 행 높이 맞춤(src/grid/rowFit.ts)이 같은 규칙을 쓰도록 한 곳에 둔다.
 */
export function buildMergeLookup(merges: string[]): Map<string, MergeRange> {
  const lookup = new Map<string, MergeRange>()
  for (const range of parseMergeRanges(merges)) {
    for (let r = range.r0; r <= range.r1; r++) {
      for (let c = range.c0; c <= range.c1; c++) {
        lookup.set(`${r},${c}`, range)
      }
    }
  }
  return lookup
}
