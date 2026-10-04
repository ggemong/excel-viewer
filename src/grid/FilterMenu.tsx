import { useEffect, useLayoutEffect, useMemo, useRef, useState } from 'react'
import { createPortal } from 'react-dom'
import type { ValueEntry } from './viewFilter'

/** 메뉴 너비(px). 위치 계산(화면 밖으로 안 나가게)과 실제 너비가 어긋나지 않도록 한 곳에서 정한다. */
const MENU_WIDTH_PX = 264
const VIEWPORT_MARGIN_PX = 8
const ANCHOR_GAP_PX = 4
/** 목록이 최소 몇 줄은 보이도록 메뉴가 가져야 할 높이. 아래 공간이 이보다 작으면 앵커 위로 띄운다. */
const MIN_MENU_HEIGHT_PX = 300
/** 이 폭 이하에서는 앵커 옆 팝오버 대신 화면 아래에서 올라오는 시트로 보여준다(터치로 누르기 쉽게). */
const NARROW_VIEWPORT_QUERY = '(max-width: 640px)'
/** 값이 수만 개여도 화면이 멈추지 않게 한 번에 그리는 개수를 제한한다 — 검색으로 좁히게 안내. */
const MAX_RENDERED_ENTRIES = 500

export const BLANK_LABEL = '(빈 셀)'

interface FilterMenuProps {
  /** 헤더 칸의 글자 — 어느 열의 필터인지 알려주는 제목. */
  title: string
  anchor: { left: number; top: number; bottom: number }
  entries: ValueEntry[]
  /** 이 열에 선택(필터)이 걸려 있는가. */
  columnFiltered: boolean
  /** 사용자가 필터를 건드린 뒤인가 — 파일에 저장돼 있던 상태로 되돌리는 버튼 노출 조건. */
  canResetToFile: boolean
  onToggle: (value: string) => void
  onSetListed: (values: string[], include: boolean) => void
  onClearColumn: () => void
  onResetToFile: () => void
  onClose: () => void
}

function labelOf(value: string): string {
  return value === '' ? BLANK_LABEL : value
}

/**
 * 헤더 ▼를 눌렀을 때 뜨는 값 선택 메뉴. 체크할 때마다 바로 적용된다(확인 버튼 없음) —
 * 보기 전용이라 되돌리기 부담이 없고, 결과를 바로 보면서 고르는 편이 빠르다.
 * 문서 body로 포털을 띄운다: 그리드는 overflow가 잘리고 스크롤되는 컨테이너라
 * 그 안에 두면 메뉴가 잘리거나 같이 스크롤된다.
 */
export function FilterMenu({ title, anchor, entries, columnFiltered, canResetToFile, onToggle, onSetListed, onClearColumn, onResetToFile, onClose }: FilterMenuProps) {
  const menuRef = useRef<HTMLDivElement>(null)
  const selectAllRef = useRef<HTMLInputElement>(null)
  const [query, setQuery] = useState('')
  const narrow = typeof window !== 'undefined' && window.matchMedia?.(NARROW_VIEWPORT_QUERY).matches

  const listed = useMemo(() => {
    const q = query.trim().toLowerCase()
    return q ? entries.filter((e) => labelOf(e.value).toLowerCase().includes(q)) : entries
  }, [entries, query])
  const rendered = listed.slice(0, MAX_RENDERED_ENTRIES)
  const allChecked = listed.length > 0 && listed.every((e) => e.included)
  const someChecked = listed.some((e) => e.included)

  useLayoutEffect(() => {
    if (selectAllRef.current) selectAllRef.current.indeterminate = someChecked && !allChecked
  }, [someChecked, allChecked])

  // 바깥을 누르거나 스크롤/크기 변경/Esc면 닫는다. ▼ 버튼 자체를 누르는 건 버튼이 열기/닫기를 토글하게 둔다.
  useEffect(() => {
    const onPointerDown = (e: PointerEvent) => {
      const target = e.target as Element | null
      if (menuRef.current?.contains(target) || target?.closest?.('.grid-filter-btn')) return
      onClose()
    }
    // 그리드가 스크롤되면 앵커(▼)가 움직여 메뉴 위치가 어긋나므로 닫는다. 다른 스크롤(모바일 가상
    // 키보드가 열릴 때의 페이지 스크롤, 메뉴 자체의 목록 스크롤)에는 닫지 않는다.
    const onScroll = (e: Event) => {
      if ((e.target as Element | null)?.classList?.contains('grid-scroll')) onClose()
    }
    // 가상 키보드가 열려도 resize가 오지만 너비는 그대로다 — 너비가 바뀔 때(회전 등)만 닫는다.
    const initialWidth = window.innerWidth
    const onResize = () => {
      if (window.innerWidth !== initialWidth) onClose()
    }
    const onKeyDown = (e: KeyboardEvent) => {
      if (e.key === 'Escape') onClose()
    }
    document.addEventListener('pointerdown', onPointerDown)
    document.addEventListener('keydown', onKeyDown)
    window.addEventListener('scroll', onScroll, true)
    window.addEventListener('resize', onResize)
    return () => {
      document.removeEventListener('pointerdown', onPointerDown)
      document.removeEventListener('keydown', onKeyDown)
      window.removeEventListener('scroll', onScroll, true)
      window.removeEventListener('resize', onResize)
    }
  }, [onClose])

  // 아래 공간이 모자라고 위가 더 넓으면 ▼ 위로 띄운다(화면 아래로 잘려 목록이 안 보이는 것 방지).
  const spaceBelow = typeof window === 'undefined' ? 0 : window.innerHeight - anchor.bottom - ANCHOR_GAP_PX - VIEWPORT_MARGIN_PX
  const spaceAbove = anchor.top - ANCHOR_GAP_PX - VIEWPORT_MARGIN_PX
  const placeAbove = spaceBelow < MIN_MENU_HEIGHT_PX && spaceAbove > spaceBelow
  const style = narrow
    ? undefined
    : {
        width: MENU_WIDTH_PX,
        left: Math.max(VIEWPORT_MARGIN_PX, Math.min(anchor.left, window.innerWidth - MENU_WIDTH_PX - VIEWPORT_MARGIN_PX)),
        ...(placeAbove
          ? { bottom: window.innerHeight - anchor.top + ANCHOR_GAP_PX, maxHeight: spaceAbove }
          : { top: anchor.bottom + ANCHOR_GAP_PX, maxHeight: Math.max(spaceBelow, Math.min(MIN_MENU_HEIGHT_PX, window.innerHeight - 2 * VIEWPORT_MARGIN_PX)) }),
      }

  return createPortal(
    <>
      {narrow && <div className="filter-menu-backdrop" aria-hidden="true" />}
      <div ref={menuRef} className={narrow ? 'filter-menu filter-menu--sheet' : 'filter-menu'} style={style} role="dialog" aria-label={`${title} 필터`}>
        <div className="filter-menu__title">{title || '필터'}</div>
        <input
          type="search"
          className="filter-menu__search"
          placeholder="값 검색"
          value={query}
          onChange={(e) => setQuery(e.target.value)}
          autoFocus={!narrow}
          aria-label="값 검색"
        />
        <label className="filter-menu__row filter-menu__row--all">
          <input ref={selectAllRef} type="checkbox" checked={allChecked} disabled={listed.length === 0} onChange={() => onSetListed(listed.map((e) => e.value), !allChecked)} />
          <span className="filter-menu__value">(모두 선택)</span>
          <span className="filter-menu__count">{listed.length}</span>
        </label>
        <div className="filter-menu__list">
          {rendered.map((e) => (
            <label key={e.value} className="filter-menu__row">
              <input type="checkbox" checked={e.included} onChange={() => onToggle(e.value)} />
              <span className="filter-menu__value" title={labelOf(e.value)}>
                {labelOf(e.value)}
              </span>
              <span className="filter-menu__count">{e.count}</span>
            </label>
          ))}
          {listed.length === 0 && <div className="filter-menu__empty">일치하는 값이 없어요</div>}
          {listed.length > rendered.length && (
            <div className="filter-menu__empty">{listed.length - rendered.length}개 더 있어요 — 검색으로 좁혀 보세요</div>
          )}
        </div>
        <div className="filter-menu__footer">
          <button type="button" className="btn btn--ghost" disabled={!columnFiltered} onClick={onClearColumn}>
            이 열 필터 해제
          </button>
          {canResetToFile && (
            <button type="button" className="btn btn--ghost" onClick={onResetToFile}>
              파일 원본 상태로
            </button>
          )}
          <button type="button" className="btn btn--ghost" onClick={onClose}>
            닫기
          </button>
        </div>
      </div>
    </>,
    document.body,
  )
}
