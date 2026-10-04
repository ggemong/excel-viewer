import { useClipboardCopy } from '../clipboard/useClipboardCopy'
import type { CellDetail } from './cellDetail'

/** 이보다 긴 글자(또는 줄바꿈이 있는 글자)는 한 줄에 다 안 들어간다고 보고 "펼치기"를 보여준다. */
const EXPAND_THRESHOLD_CHARS = 60

interface CellDetailBarProps {
  detail: CellDetail
  expanded: boolean
  onToggleExpanded: () => void
}

function isLong(text: string | null): boolean {
  return text !== null && (text.length > EXPAND_THRESHOLD_CHARS || text.includes('\n'))
}

/**
 * 선택한 칸의 전체 내용을 보여주는 줄(Excel 수식 입력줄에 해당). 칸이 좁아 글자가 잘려도 여기서
 * 전부 읽고 복사할 수 있다. 접힌 상태는 한 줄(높이 고정 — 방향키로 칸을 옮길 때 화면이 출렁이지
 * 않게)이고, 펼치면 줄바꿈을 살려 전부 보여준다. 펼침 상태는 칸을 옮겨도 유지된다(여러 칸을 연달아
 * 읽는 사용을 위해) — 상태는 호출부(App)가 들고 있다.
 */
export function CellDetailBar({ detail, expanded, onToggleExpanded }: CellDetailBarProps) {
  const { copyText, copied } = useClipboardCopy()
  const empty = detail.text === '' && !detail.note && !detail.hyperlink && !detail.formula
  const canExpand = isLong(detail.text) || isLong(detail.note)

  return (
    <div className="cell-detail" data-expanded={expanded} role="region" aria-label="선택한 칸의 내용">
      <div className="cell-detail__main">
        <span className="cell-detail__addr">{detail.address}</span>
        <div className="cell-detail__text" data-empty={detail.text === ''}>
          {detail.text !== '' ? detail.text : empty ? '(빈 셀)' : ''}
        </div>
        <div className="cell-detail__actions">
          {canExpand && (
            <button type="button" className="btn btn--ghost" onClick={onToggleExpanded} aria-expanded={expanded}>
              {expanded ? '접기' : '펼치기'}
            </button>
          )}
          <button type="button" className="btn btn--ghost" onClick={() => void copyText(detail.text)} disabled={detail.text === ''}>
            {copied ? '복사됨' : '복사'}
          </button>
        </div>
      </div>

      {(detail.originalValue || detail.formula || detail.hyperlink || detail.note) && (
        <div className="cell-detail__extras">
          {detail.originalValue && (
            <div>
              <span className="cell-detail__label">원래 값</span> <code>{detail.originalValue}</code>
            </div>
          )}
          {detail.formula && (
            <div>
              <span className="cell-detail__label">수식</span> <code>{detail.formula}</code>
            </div>
          )}
          {detail.hyperlink && (
            <div>
              <span className="cell-detail__label">링크</span>{' '}
              <a href={detail.hyperlink} target="_blank" rel="noopener noreferrer">
                {detail.hyperlink}
              </a>
            </div>
          )}
          {detail.note && (
            <div className="cell-detail__note">
              <span className="cell-detail__label">메모</span> {detail.note}
            </div>
          )}
        </div>
      )}
    </div>
  )
}
