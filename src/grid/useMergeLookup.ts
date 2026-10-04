import { useMemo } from 'react'
import { buildMergeLookup, type MergeRange } from '../xlsx/mergeRange'

export type { MergeRange } from '../xlsx/mergeRange'
export { isMergeMaster } from '../xlsx/mergeRange'

/**
 * (row, col) -> 그 셀이 속한 병합 범위. ExcelJS는 병합된 셀을 읽으면 마스터(왼쪽
 * 위) 값을 나머지 칸에도 그대로 돌려주기 때문에, 그리드가 값을 중복으로 찍지
 * 않으려면 어떤 칸이 마스터고 어떤 칸이 "덮인" 칸인지 직접 판별해야 한다.
 */
export function useMergeLookup(merges: string[]): Map<string, MergeRange> {
  return useMemo(() => buildMergeLookup(merges), [merges])
}
