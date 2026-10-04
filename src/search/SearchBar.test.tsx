import { act, createRef } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { SearchBar } from './SearchBar'
import type { SearchController } from './useSearch'

;(globalThis as unknown as { IS_REACT_ACT_ENVIRONMENT: boolean }).IS_REACT_ACT_ENVIRONMENT = true

let root: Root
let host: HTMLDivElement

function makeSearch(over: Partial<SearchController> = {}): SearchController {
  return {
    open: true, setOpen: vi.fn(), query: '가', setQuery: vi.fn(), scope: 'sheet', setScope: vi.fn(),
    matchCase: false, setMatchCase: vi.fn(), wholeCell: false, setWholeCell: vi.fn(),
    hits: [], hiddenExcluded: 0, truncated: false, searching: false, currentIndex: -1, currentHit: null,
    navNonce: 0, countsBySheet: new Map(), selectHit: vi.fn(), go: vi.fn(),
    ...over,
  } as SearchController
}

function mount(search: SearchController, onClose = vi.fn()) {
  act(() => root.render(<SearchBar search={search} inputRef={createRef<HTMLInputElement>()} onClose={onClose} />))
  return onClose
}
const q = (sel: string) => host.querySelector(sel) as HTMLElement
const buttonByText = (text: string) => Array.from(host.querySelectorAll('button')).find((b) => b.textContent === text) as HTMLButtonElement

beforeEach(() => {
  host = document.createElement('div')
  document.body.appendChild(host)
  root = createRoot(host)
})
afterEach(() => {
  act(() => root.unmount())
  host.remove()
})

describe('SearchBar', () => {
  it('상태 문구: 현재 위치/전체, 검색 중, 결과 없음', () => {
    const hit = { kind: 'cell', sheetIndex: 0, row: 1, col: 1, text: '가', matchStart: 0, matchLength: 1 } as const
    mount(makeSearch({ hits: [hit, hit, hit], currentIndex: 1 }))
    expect(q('.search-bar__count').textContent).toBe('2 / 3')
    mount(makeSearch({ searching: true }))
    expect(q('.search-bar__count').textContent).toBe('검색 중…')
    mount(makeSearch())
    expect(q('.search-bar__count').textContent).toBe('결과 없음')
    mount(makeSearch({ query: '' }))
    expect(q('.search-bar__count').textContent).toBe('')
  })

  it('Enter는 다음, Shift+Enter는 이전', () => {
    const s = makeSearch()
    mount(s)
    const input = q('.search-bar__input') as HTMLInputElement
    act(() => input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true })))
    act(() => input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', shiftKey: true, bubbles: true })))
    expect(s.go).toHaveBeenNthCalledWith(1, 1)
    expect(s.go).toHaveBeenNthCalledWith(2, -1)
  })

  it('범위 토글과 옵션 토글, 닫기', () => {
    const s = makeSearch()
    const onClose = mount(s)
    act(() => buttonByText('전체 시트').click())
    expect(s.setScope).toHaveBeenCalledWith('all')
    act(() => buttonByText('Aa').click())
    expect(s.setMatchCase).toHaveBeenCalledWith(true)
    act(() => buttonByText('전체 일치').click())
    expect(s.setWholeCell).toHaveBeenCalledWith(true)
    act(() => (q('.search-bar__close') as HTMLButtonElement).click())
    expect(onClose).toHaveBeenCalled()
    expect(buttonByText('현재 시트').getAttribute('aria-pressed')).toBe('true')
  })

  it('현재 시트에 결과가 없으면 전체 시트로 넓히는 링크를 보여준다', () => {
    const s = makeSearch()
    mount(s)
    act(() => buttonByText('이 시트에는 없어요 — 전체 시트에서 찾기').click())
    expect(s.setScope).toHaveBeenCalledWith('all')
  })

  it('숨겨서 제외한 건수와 상한 도달을 알려준다', () => {
    mount(makeSearch({ hiddenExcluded: 3, truncated: true, hits: [{ kind: 'cell', sheetIndex: 0, row: 1, col: 1, text: 'a', matchStart: 0, matchLength: 1 }] }))
    const note = q('.search-bar__note').textContent ?? ''
    expect(note).toContain('숨겨진 행·열의 일치 3건은 제외했어요')
    expect(note).toContain('처음 1건만 보여요')
  })
})
