import { useRef } from 'react'
import tigerFace from '../assets/tiger/tiger-face.png'

interface ToolbarProps {
  fileName: string | null
  onOpenAnother: () => void
  editMode: boolean
  onToggleEditMode: () => void
  isDirty: boolean
  editedCount: number
  saving: boolean
  onSave: () => void
  compareFileName: string | null
  compareLoading: boolean
  onPickCompareFile: (file: File) => void
  onClearCompare: () => void
}

export function Toolbar({
  fileName,
  onOpenAnother,
  editMode,
  onToggleEditMode,
  isDirty,
  editedCount,
  saving,
  onSave,
  compareFileName,
  compareLoading,
  onPickCompareFile,
  onClearCompare,
}: ToolbarProps) {
  const compareInputRef = useRef<HTMLInputElement>(null)

  return (
    <div className="toolbar">
      <div className="toolbar-group toolbar-brand">
        <span className="tiger-disk tiger-disk--40">
          <img src={tigerFace} alt="" />
        </span>
        <span className="wm">
          w<span className="wm-i">i</span>thv<span className="wm-i">i</span>be
        </span>
        {fileName && (
          <>
            <span className="toolbar-brand-sub" aria-hidden="true">
              ·
            </span>
            <span className="toolbar-brand-sub">{fileName}</span>
          </>
        )}
      </div>
      <div className="toolbar-group">
        {fileName && (
          <>
            <button
              type="button"
              className="pill pill--outline"
              aria-pressed={editMode}
              onClick={onToggleEditMode}
              style={editMode ? { color: 'var(--accent)', borderColor: 'var(--accent)' } : undefined}
            >
              {editMode ? '편집 중' : '편집 모드'}
            </button>
            {isDirty && (
              <button type="button" className="btn btn--primary" onClick={onSave} disabled={saving}>
                {saving ? '저장 중…' : `변경 ${editedCount}개 저장`}
              </button>
            )}

            {compareFileName ? (
              <span className="pill pill--outline" style={{ color: 'var(--accent)', borderColor: 'var(--accent)' }}>
                {compareFileName}와 비교 중
                <button
                  type="button"
                  onClick={onClearCompare}
                  aria-label="비교 종료"
                  style={{ border: 'none', background: 'none', color: 'inherit', cursor: 'pointer', padding: 0, marginLeft: 4 }}
                >
                  ×
                </button>
              </span>
            ) : (
              <button
                type="button"
                className="pill pill--outline"
                onClick={() => compareInputRef.current?.click()}
                disabled={compareLoading}
              >
                {compareLoading ? '불러오는 중…' : '버전 비교'}
              </button>
            )}
            <input
              ref={compareInputRef}
              type="file"
              accept=".xlsx,.csv"
              style={{ display: 'none' }}
              onChange={(e) => {
                const f = e.target.files?.[0]
                if (f) onPickCompareFile(f)
                e.target.value = ''
              }}
            />
          </>
        )}
        <span className="pill pill--live">브라우저에서만 열림</span>
        {fileName && (
          <button className="btn btn--ghost" onClick={onOpenAnother} type="button">
            다른 파일 열기
          </button>
        )}
      </div>
    </div>
  )
}
