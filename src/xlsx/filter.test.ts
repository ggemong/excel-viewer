import { describe, expect, it } from 'vitest'
import { autoFilterRangeText, buildSheetFilter, parseAutoFilterDetails } from './filter'

describe('parseAutoFilterDetails', () => {
  it('조건이 없는 자동 필터(버튼만)는 활성 열 없음 — 실제 파일에 가장 흔한 모양', () => {
    expect(parseAutoFilterDetails('<autoFilter ref="A3:N18" xr:uid="{00000000}"/>')).toEqual({ activeColIds: [], hiddenButtonColIds: [] })
  })

  it('조건(filters/customFilters/top10 등)이 있는 열만 활성으로 센다', () => {
    const xml =
      '<autoFilter ref="A1:D9"><filterColumn colId="1"><filters><filter val="A"/></filters></filterColumn>' +
      '<filterColumn colId="3"><customFilters><customFilter operator="greaterThan" val="5"/></customFilters></filterColumn></autoFilter>'
    expect(parseAutoFilterDetails(xml).activeColIds).toEqual([1, 3])
  })

  it('버튼을 숨긴 열(hiddenButton / showButton=0)', () => {
    const xml = '<autoFilter ref="A1:C9"><filterColumn colId="0" hiddenButton="1"/><filterColumn colId="2" showButton="0"/></autoFilter>'
    const d = parseAutoFilterDetails(xml)
    expect(d.hiddenButtonColIds).toEqual([0, 2])
    expect(d.activeColIds).toEqual([])
  })

  it('접두사가 붙은 요소도 읽는다', () => {
    const xml = '<x:autoFilter ref="A1:B2"><x:filterColumn colId="0"><x:filters><x:filter val="1"/></x:filters></x:filterColumn></x:autoFilter>'
    expect(parseAutoFilterDetails(xml).activeColIds).toEqual([0])
  })
})

describe('autoFilterRangeText / buildSheetFilter', () => {
  it('ExcelJS의 문자열/{from,to} 두 표기를 모두 범위 문자열로', () => {
    expect(autoFilterRangeText('A3:N18')).toBe('A3:N18')
    expect(autoFilterRangeText({ from: { row: 3, column: 1 }, to: { row: 18, column: 14 } })).toBe('A3:N18')
    expect(autoFilterRangeText({ from: 'B2', to: 'D9' })).toBe('B2:D9')
    expect(autoFilterRangeText(undefined)).toBeNull()
  })

  it('colId는 범위 첫 열 기준 상대값이라 절대 열 번호로 바꾼다', () => {
    const f = buildSheetFilter('C3:F20', { activeColIds: [1], hiddenButtonColIds: [3] })!
    expect(f).toEqual({ headerRow: 3, lastRow: 20, firstCol: 3, lastCol: 6, activeCols: [4], hiddenButtonCols: [6] })
  })

  it('해석 못 하는 범위는 null', () => {
    expect(buildSheetFilter('???', null)).toBeNull()
  })
})
