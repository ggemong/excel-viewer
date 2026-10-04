interface FilterStatusBarProps {
  total: number
  shown: number
  /** 사용자가 필터를 건드린 뒤인가 — 아니면 파일에 저장돼 있던 필터 결과를 보고 있는 것. */
  touched: boolean
  onClearAll: () => void
  onResetToFile: () => void
}

/**
 * 필터로 일부 행이 가려져 있다는 걸 항상 알려주는 줄. 이게 없으면 "데이터가 원래 이것뿐인가?"와
 * "필터로 가려졌나?"를 구분할 수 없다(Excel도 상태 표시줄에 "N개 레코드 중 M개 찾음"을 띄운다).
 */
export function FilterStatusBar({ total, shown, touched, onClearAll, onResetToFile }: FilterStatusBarProps) {
  return (
    <div className="notice-bar notice-bar--action" role="status">
      <span>
        {touched ? '필터 적용 중' : '파일에 저장된 필터가 적용돼 있어요'} · 전체 {total}행 중 {shown}행 표시
      </span>
      <button type="button" className="btn btn--ghost" onClick={onClearAll}>
        필터 모두 해제
      </button>
      {touched && (
        <button type="button" className="btn btn--ghost" onClick={onResetToFile}>
          파일 원본 상태로
        </button>
      )}
    </div>
  )
}
