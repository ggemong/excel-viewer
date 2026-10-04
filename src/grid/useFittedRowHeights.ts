import { useMemo } from 'react'
import type { SheetModel } from '../xlsx/types'
import { fitRowHeights } from './rowFit'
import { measureWrappedHeights } from './wrapMeasure'

/**
 * 시트의 행높이에서 줄바꿈 글이 잘리는 행을 키운 결과(규칙은 rowFit.ts). 시트가 같고 글꼴 로딩 상태가 같으면
 * 다시 계산하지 않는다 — 시트가 바뀔 때와 글꼴이 도착할 때만 한 번씩 잰다.
 *
 * @param fontsVersion useFontsVersion()의 값. 글꼴이 도착하면 줄 수가 달라지므로 바뀔 때 다시 잰다.
 */
export function useFittedRowHeights(sheet: SheetModel, fontsVersion: number): number[] {
  // fontsVersion은 값으로 쓰지 않지만 "글꼴이 바뀌었으니 다시 재라"는 신호라 의존성에 둔다.
  // oxlint-disable-next-line react-hooks/exhaustive-deps
  return useMemo(() => fitRowHeights(sheet, measureWrappedHeights), [sheet, fontsVersion])
}
