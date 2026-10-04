import { useCallback, useEffect, useMemo, useRef, useState } from 'react'
import type { SheetModel, WorkbookModel } from '../xlsx/types'
import { MAX_HITS, searchSheet, type SearchHit } from './searchEngine'

export type SearchScope = 'sheet' | 'all'

/** 입력이 멈추고 이만큼 지나야 검색을 시작한다 — 글자마다 큰 시트를 훑지 않게. */
const DEBOUNCE_MS = 200

interface SearchState {
  hits: SearchHit[]
  hiddenExcluded: number
  truncated: boolean
  searching: boolean
}

const EMPTY: SearchState = { hits: [], hiddenExcluded: 0, truncated: false, searching: false }

interface UseSearchArgs {
  workbook: WorkbookModel | null
  activeSheetIndex: number
  /** 시트에 보기 필터(가려진 행)를 적용한 결과 — 검색이 화면과 같은 기준으로 "가려진 행 제외"를 하게 한다. */
  viewSheetFor: (sheet: SheetModel) => SheetModel
  /** 다른 시트의 결과로 이동할 때 그 시트를 활성화한다(선택·메뉴 초기화는 호출부 몫). */
  onActivateSheet: (index: number) => void
}

/**
 * 시트 검색의 상태와 동작. 실제 일치 판단은 searchEngine.ts, 여기는 "언제 검색을 돌리고, 결과 중
 * 어디에 있는지(현재 위치)"만 관리한다.
 *
 * 설계 메모:
 *  - 결과는 워크북 순서(시트 → 행 → 열)로 쌓는다. 전체 시트 검색은 시트를 순서대로 처리하며 끝난
 *    시트부터 결과를 보여준다(큰 파일에서 기다리게 하지 않음).
 *  - 입력/옵션이 바뀌면 이전 검색을 취소하고 새로 시작한다.
 *  - 새 검색의 첫 현재 위치는 "지금 보고 있는 시트의 첫 결과"다. 다른 시트에만 있으면 자동으로
 *    시트를 넘기지 않는다(타이핑 중 화면이 갑자기 바뀌는 걸 막기 위해) — Enter/클릭할 때 이동한다.
 */
export function useSearch({ workbook, activeSheetIndex, viewSheetFor, onActivateSheet }: UseSearchArgs) {
  const [open, setOpen] = useState(false)
  const [query, setQuery] = useState('')
  const [scope, setScope] = useState<SearchScope>('sheet')
  const [matchCase, setMatchCase] = useState(false)
  const [wholeCell, setWholeCell] = useState(false)
  const [state, setState] = useState<SearchState>(EMPTY)
  const [current, setCurrent] = useState(-1)
  /** 이동할 때마다 올라가는 번호 — 같은 결과로 다시 이동(Enter 반복)해도 그리드가 다시 스크롤하게 한다. */
  const [navNonce, setNavNonce] = useState(0)

  const activeSheetRef = useRef(activeSheetIndex)
  useEffect(() => {
    activeSheetRef.current = activeSheetIndex
  }, [activeSheetIndex])

  const trimmed = query.trim()
  const active = open && workbook !== null && trimmed !== ''
  // 전체 시트 검색은 보고 있는 시트가 바뀌어도 결과가 같으므로 다시 돌리지 않는다.
  const sheetKey = scope === 'sheet' ? activeSheetIndex : -1

  useEffect(() => {
    if (!active || !workbook) return
    let cancelled = false
    const timer = setTimeout(async () => {
      setState({ ...EMPTY, searching: true })
      setCurrent(-1)

      const order = scope === 'sheet' ? [activeSheetRef.current] : workbook.sheets.map((_, i) => i)
      const all: SearchHit[] = []
      let hiddenExcluded = 0
      let truncated = false
      let needAutoSelect = true

      for (const index of order) {
        const sheet = workbook.sheets[index]
        if (!sheet) continue
        const result = await searchSheet(viewSheetFor(sheet), index, { query, matchCase, wholeCell }, { cancelled: () => cancelled, maxHits: MAX_HITS - all.length })
        if (cancelled || result.cancelled) return

        const firstOfSheet = all.length
        all.push(...result.hits)
        hiddenExcluded += result.hiddenExcluded
        truncated ||= result.truncated

        if (needAutoSelect && index === activeSheetRef.current && result.hits.length > 0) {
          needAutoSelect = false
          setCurrent(firstOfSheet)
          setNavNonce((n) => n + 1)
        }
        setState({ hits: [...all], hiddenExcluded, truncated, searching: true })
        if (truncated) break
      }
      setState({ hits: all, hiddenExcluded, truncated, searching: false })
    }, DEBOUNCE_MS)

    return () => {
      cancelled = true
      clearTimeout(timer)
    }
    // query는 trimmed가 아니라 원문을 쓴다(공백이 든 검색어 "홍 길동" 보존). active는 trimmed에 의존한다.
  }, [active, workbook, query, scope, sheetKey, matchCase, wholeCell, viewSheetFor])

  const view = active ? state : EMPTY
  const currentIndex = active && current >= 0 && current < view.hits.length ? current : -1
  const currentHit = currentIndex >= 0 ? view.hits[currentIndex] : null

  const countsBySheet = useMemo(() => {
    const counts = new Map<number, number>()
    for (const hit of view.hits) counts.set(hit.sheetIndex, (counts.get(hit.sheetIndex) ?? 0) + 1)
    return counts
  }, [view.hits])

  const selectHit = useCallback(
    (index: number) => {
      const hit = view.hits[index]
      if (!hit) return
      setCurrent(index)
      setNavNonce((n) => n + 1)
      if (hit.sheetIndex !== activeSheetIndex) onActivateSheet(hit.sheetIndex)
    },
    [view.hits, activeSheetIndex, onActivateSheet],
  )

  const go = useCallback(
    (delta: 1 | -1) => {
      const n = view.hits.length
      if (n === 0) return
      if (currentIndex < 0) {
        // 아직 현재 위치가 없으면: 다음은 보고 있는 시트의 첫 결과(없으면 맨 처음), 이전은 맨 끝.
        const inActive = view.hits.findIndex((h) => h.sheetIndex === activeSheetIndex)
        selectHit(delta > 0 ? Math.max(inActive, 0) : n - 1)
        return
      }
      selectHit((currentIndex + delta + n) % n)
    },
    [view.hits, currentIndex, activeSheetIndex, selectHit],
  )

  return {
    open,
    setOpen,
    query,
    setQuery,
    scope,
    setScope,
    matchCase,
    setMatchCase,
    wholeCell,
    setWholeCell,
    hits: view.hits,
    hiddenExcluded: view.hiddenExcluded,
    truncated: view.truncated,
    searching: active && (state.searching || view === EMPTY),
    currentIndex,
    currentHit,
    navNonce,
    countsBySheet,
    selectHit,
    go,
  }
}

export type SearchController = ReturnType<typeof useSearch>
