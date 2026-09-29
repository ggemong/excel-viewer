import tigerFace from '../assets/tiger/tiger-face.png'

interface ToolbarProps {
  fileName: string | null
  onOpenAnother: () => void
}

export function Toolbar({ fileName, onOpenAnother }: ToolbarProps) {
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
