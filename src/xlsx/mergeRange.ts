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
